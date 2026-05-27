/**
 * 「生成手填本」隱藏 manifest 的共用邏輯（reader / writer 共用，避免反向依賴）。
 *
 * manifest 設計：
 *   - writer 寫出 .xlsx 時加一個 veryHidden 分頁（HANDFILL_MANIFEST_SHEET），
 *     A1 存 JSON：{ version, book, layoutHash }
 *   - layoutHash 是「上」分頁版面文字的雜湊，於寫檔時透過 round-trip
 *     （writeBuffer → load → 抽 matrix）計算，使寫/讀路徑對稱，合併格與數字格式
 *     皆由 ExcelJS 反序列化統一處理，避免 in-memory 與 loaded 不一致造成偽陽性。
 *   - reader 讀檔時重算「上」分頁 hash 與 manifest.layoutHash 比對：
 *       相符 → 使用者沒在 Excel 改過版面 → 直接用 JSON 還原（快、準）
 *       不符 → 使用者改過版面 → 改走版面分析，並從 JSON 補回版面表達不了的內部欄位
 *
 * 注意：hash 只涵蓋「上」分頁，與 reader 版面分析範圍一致（reader 僅讀「上」）。
 * 使用者若只改「下」分頁，reader 本來就不採納，hash 也不偵測，行為一致。
 */

import ExcelJS from 'exceljs';

import {HANDFILL_MANIFEST_SHEET, type HandfillBook} from '@/domain/models/handfill-book';

export const HANDFILL_MANIFEST_VERSION = 1;

/** A1 JSON 的新格式。舊格式（af24e5d）A1 直接是 HandfillBook，無此包裝。 */
export interface HandfillManifest {
    version: number;
    book: HandfillBook;
    layoutHash: string;
}

export type Cell = string;
export type Matrix = Cell[][];   // matrix[rowIdx][colIdx]，0-indexed

/* ====================== ExcelJS cell → 字串 ======================= */

export function textOfExcelJs(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    if (value instanceof Date) return value.toISOString();
    const obj = value as { richText?: Array<{ text: string }>; text?: string; result?: unknown };
    if (obj.richText) return obj.richText.map((t) => t.text).join('');
    if (obj.text) return obj.text;
    if (obj.result !== undefined) return textOfExcelJs(obj.result as ExcelJS.CellValue);
    return '';
}

/* ====================== 抽「上」分頁 matrix ======================= */

/**
 * 取「上」分頁（沒有則取第一個非 manifest 分頁）讀為 2D 字串矩陣。
 * reader 版面分析與 writer round-trip hash 共用此函式，確保兩邊抽法一致。
 */
export function buildMatrixFromWorkbook(wb: ExcelJS.Workbook): Matrix {
    const sheet = wb.getWorksheet('上')
        ?? wb.worksheets.find((s) => s.name !== HANDFILL_MANIFEST_SHEET);
    if (!sheet) throw new Error('找不到資料分頁');

    const matrix: Matrix = [];
    const totalRows = sheet.actualRowCount || sheet.rowCount;
    const totalCols = sheet.actualColumnCount || sheet.columnCount;

    for (let r = 1; r <= totalRows; r++) {
        const row: Cell[] = [];
        for (let c = 1; c <= totalCols; c++) {
            row.push(textOfExcelJs(sheet.getRow(r).getCell(c).value));
        }
        matrix.push(row);
    }
    return matrix;
}

/* ====================== 版面 hash ======================= */

/**
 * 正規化 matrix：每格 trim，並剔除尾端整列／整欄空白，
 * 讓「無意義的尾端空白差異」不影響 hash（中間任何格的改動仍會偵測得到）。
 */
function normalizeMatrix(matrix: Matrix): Matrix {
    const trimmed = matrix.map((row) => row.map((c) => (c ?? '').trim()));

    let lastRow = trimmed.length;
    while (lastRow > 0 && trimmed[lastRow - 1].every((c) => c === '')) lastRow--;
    const rows = trimmed.slice(0, lastRow);

    let lastCol = 0;
    for (const row of rows) {
        for (let c = row.length - 1; c >= 0; c--) {
            if (row[c] !== '') {
                if (c + 1 > lastCol) lastCol = c + 1;
                break;
            }
        }
    }
    return rows.map((row) => row.slice(0, lastCol));
}

/** djb2 非加密雜湊（僅用於偵測版面差異，非安全用途）。 */
function djb2(str: string): string {
    let h = 5381;
    for (let i = 0; i < str.length; i++) {
        h = ((h << 5) + h + str.charCodeAt(i)) >>> 0;
    }
    return h.toString(36);
}

export function hashMatrix(matrix: Matrix): string {
    return djb2(JSON.stringify(normalizeMatrix(matrix)));
}

/** 直接對 workbook 的「上」分頁算版面 hash。 */
export function layoutHashOfWorkbook(wb: ExcelJS.Workbook): string {
    return hashMatrix(buildMatrixFromWorkbook(wb));
}

/* ====================== manifest 讀取（相容新舊格式） ======================= */

/**
 * 從 workbook 讀出 manifest 原始資料，相容兩種格式：
 *   - 新格式：{ version, book, layoutHash }
 *   - 舊格式（af24e5d）：A1 直接是 HandfillBook（頂層有 customers 陣列），無 layoutHash
 * 回傳統一形狀 { book, layoutHash? }；無 manifest 或無法解析則回 null。
 */
export function readManifestRaw(wb: ExcelJS.Workbook): { book: HandfillBook; layoutHash?: string } | null {
    const sheet = wb.getWorksheet(HANDFILL_MANIFEST_SHEET);
    if (!sheet) return null;
    const raw = textOfExcelJs(sheet.getCell(1, 1).value);
    if (!raw) return null;
    try {
        const parsed = JSON.parse(raw) as unknown;
        if (!parsed || typeof parsed !== 'object') return null;

        // 新格式：有 book 包裝
        const wrapped = parsed as Partial<HandfillManifest>;
        if (wrapped.book && Array.isArray(wrapped.book.customers)) {
            return {
                book: wrapped.book,
                layoutHash: typeof wrapped.layoutHash === 'string' ? wrapped.layoutHash : undefined,
            };
        }

        // 舊格式：頂層即 book
        const legacy = parsed as Partial<HandfillBook>;
        if (Array.isArray(legacy.customers)) {
            return {book: legacy as HandfillBook, layoutHash: undefined};
        }

        return null;
    } catch {
        return null;
    }
}
