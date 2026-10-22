/**
 * 收支表（例：115-收支表.xlsx）月工作表解析器。純函式，不依賴 ExcelJS，方便測試。
 *
 * 月工作表結構（列號不固定，一律以內容辨識）：
 * - A1：「114/1月」→ 民國年 / 月份
 * - 標題列：同一列出現 ≥ 3 個大類名稱（原物料、營運、車輛、人事、其他）。
 *   每個大類名稱所在欄 = 區塊起點；區塊延伸到下一個大類起點前一欄。
 *   區塊內：第 1 欄 = 品項、第 2 欄 = 金額、其餘欄 = 廠商／月份／人名等備註。
 * - 區塊內往下掃：
 *   - 金額欄是 SUM(...) 公式 → 目前小節的小計（Excel 年度總表就是引用這格）
 *   - 品項欄是大類名稱且金額空白 → 開啟新小節（例：「其他」下方的「銀行」、
 *     1 月「車輛」下方的「人事」紅包區塊）
 *   - 金額是數字 → 一筆明細
 * - 標題列左側（A/B 欄）為收入，直到「合計」列。
 */

import {
    EXPENSE_CATEGORIES,
    type ExpenseCategory,
    type ExpenseItem,
    type IncomeItem,
    isExpenseCategory,
    monthSortKey,
    type MonthExpense,
    emptyTotals,
} from './expense-model';

export interface GridCell {
    /** 顯示值（公式取快取結果） */
    value: string | number | null;
    /** 公式文字（不含 =），沒有公式時為 undefined */
    formula?: string;
}

export interface SheetGrid {
    name: string;
    rowCount: number;
    colCount: number;

    /** 1-based */
    get(row: number, col: number): GridCell | null;
}

export class ExpenseParseError extends Error {}

const HEADER_SCAN_ROWS = 10;
const MIN_HEADER_CATEGORIES = 3;

export function colLetter(col: number): string {
    let s = '';
    let n = col;
    while (n > 0) {
        const m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

function colIndex(letters: string): number {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
}

function text(cell: GridCell | null): string {
    if (!cell || cell.value === null) return '';
    return String(cell.value).trim();
}

function num(cell: GridCell | null): number | null {
    if (!cell || cell.value === null || cell.value === '') return null;
    if (typeof cell.value === 'number') return Number.isFinite(cell.value) ? cell.value : null;
    const n = Number(String(cell.value).replace(/,/g, '').trim());
    return Number.isFinite(n) && String(cell.value).trim() !== '' ? n : null;
}

function isSumFormula(cell: GridCell | null): boolean {
    return !!cell?.formula && /^\s*SUM\s*\(/i.test(cell.formula);
}

/** 自行計算 SUM(A1:B2, C3) 的值（快取值缺漏時的備援，只支援同工作表範圍）。 */
function evalSum(grid: SheetGrid, formula: string): number | null {
    const m = /^\s*SUM\s*\((.*)\)\s*$/i.exec(formula);
    if (!m) return null;
    let total = 0;
    for (const part of m[1].split(',')) {
        const ref = /^\s*\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?\s*$/i.exec(part);
        if (!ref) return null;
        const c1 = colIndex(ref[1]);
        const r1 = Number(ref[2]);
        const c2 = ref[3] ? colIndex(ref[3]) : c1;
        const r2 = ref[4] ? Number(ref[4]) : r1;
        for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
            for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
                total += num(grid.get(r, c)) ?? 0;
            }
        }
    }
    return total;
}

function subtotalValue(grid: SheetGrid, cell: GridCell): number {
    const cached = num(cell);
    if (cached !== null) return cached;
    return evalSum(grid, cell.formula ?? '') ?? 0;
}

/** 「114/1月」「114年1月」→ {year, month}；只有月份（工作表名「1月」）時 year 為 null。 */
export function parseMonthTitle(title: string): { year: number | null; month: number } | null {
    const full = /(\d{2,4})\s*[/年.\-]\s*(\d{1,2})\s*月?/.exec(title);
    if (full) {
        let year = Number(full[1]);
        if (year > 1911) year -= 1911;
        const month = Number(full[2]);
        if (month >= 1 && month <= 12) return {year, month};
    }
    const only = /^\s*(\d{1,2})\s*月\s*$/.exec(title);
    if (only) {
        const month = Number(only[1]);
        if (month >= 1 && month <= 12) return {year: null, month};
    }
    return null;
}

function findHeaderRow(grid: SheetGrid): { row: number; starts: { col: number; category: ExpenseCategory }[] } | null {
    for (let r = 1; r <= Math.min(HEADER_SCAN_ROWS, grid.rowCount); r++) {
        const starts: { col: number; category: ExpenseCategory }[] = [];
        for (let c = 1; c <= grid.colCount; c++) {
            const t = text(grid.get(r, c));
            if (isExpenseCategory(t)) starts.push({col: c, category: t});
        }
        if (starts.length >= MIN_HEADER_CATEGORIES) return {row: r, starts};
    }
    return null;
}

/** 判斷這張工作表是不是收支表的月工作表（有月份標題 + 大類標題列）。 */
export function looksLikeMonthSheet(grid: SheetGrid): boolean {
    const title = parseMonthTitle(text(grid.get(1, 1))) ?? parseMonthTitle(grid.name);
    return !!title && findHeaderRow(grid) !== null;
}

interface Section {
    category: ExpenseCategory;
    items: ExpenseItem[];
    subtotal: number | null;
    subtotalCell: string;
}

export function parseMonthSheet(
    grid: SheetGrid,
    opts: { fallbackYear: number; sourceFile: string },
): MonthExpense {
    const title = parseMonthTitle(text(grid.get(1, 1))) ?? parseMonthTitle(grid.name);
    if (!title) throw new ExpenseParseError(`「${grid.name}」找不到月份標題（A1 應為如「114/1月」）`);
    const header = findHeaderRow(grid);
    if (!header) throw new ExpenseParseError(`「${grid.name}」找不到支出大類標題列（原物料／營運／車輛／人事／其他）`);

    const year = title.year ?? opts.fallbackYear;
    const month = title.month;
    const warnings: string[] = [];
    const sections: Section[] = [];

    header.starts.forEach((start, idx) => {
        const endCol = idx + 1 < header.starts.length ? header.starts[idx + 1].col - 1 : grid.colCount;
        const labelCol = start.col;
        const amountCol = start.col + 1;
        let current: Section = {category: start.category, items: [], subtotal: null, subtotalCell: ''};
        sections.push(current);

        for (let r = header.row + 1; r <= grid.rowCount; r++) {
            const label = text(grid.get(r, labelCol));
            const amtCell = grid.get(r, amountCol);

            if (isSumFormula(amtCell)) {
                const value = subtotalValue(grid, amtCell!);
                if (current.subtotal !== null) {
                    warnings.push(`${grid.name}「${current.category}」有多個小計格（${current.subtotalCell}、${colLetter(amountCol)}${r}），以後者為準`);
                }
                current.subtotal = value;
                current.subtotalCell = `${colLetter(amountCol)}${r}`;
                continue;
            }

            const amount = num(amtCell);
            if (isExpenseCategory(label) && amount === null) {
                current = {category: label, items: [], subtotal: null, subtotalCell: ''};
                sections.push(current);
                continue;
            }
            if (amount === null) continue;

            const extras: string[] = [];
            for (let c = amountCol + 1; c <= endCol; c++) {
                const t = text(grid.get(r, c));
                if (t) extras.push(t);
            }
            const name = label || extras.shift() || '（未命名）';
            current.items.push({
                category: current.category,
                name,
                amount,
                note: extras.join('・'),
                cell: `${colLetter(amountCol)}${r}`,
            });
        }
    });

    const totals = emptyTotals();
    const items: ExpenseItem[] = [];
    for (const s of sections) {
        const itemSum = s.items.reduce((acc, it) => acc + it.amount, 0);
        if (s.subtotal === null) {
            if (s.items.length > 0) {
                warnings.push(`${grid.name}「${s.category}」找不到 SUM 小計格，改用明細加總 ${itemSum.toLocaleString('zh-TW')}`);
            }
            totals[s.category] += itemSum;
        } else {
            if (Math.abs(s.subtotal - itemSum) > 0.5) {
                warnings.push(
                    `${grid.name}「${s.category}」小計 ${s.subtotalCell}=${s.subtotal.toLocaleString('zh-TW')}，` +
                    `但明細加總為 ${itemSum.toLocaleString('zh-TW')}（可能有明細沒被 SUM 範圍涵蓋）`,
                );
            }
            totals[s.category] += s.subtotal;
        }
        items.push(...s.items);
    }

    const {income, incomeTotal} = parseIncome(grid, header.row, header.starts[0].col);

    return {
        year,
        month,
        sortKey: monthSortKey(year, month),
        sheetName: grid.name,
        totals,
        items,
        income,
        incomeTotal,
        warnings,
        sourceFile: opts.sourceFile,
        uploadedAt: Date.now(),
    };
}

/** 收入：標題列起、支出區塊左側的「名稱 | 金額」兩欄，到「合計」為止。 */
function parseIncome(grid: SheetGrid, headerRow: number, firstExpenseCol: number): { income: IncomeItem[]; incomeTotal: number } {
    const income: IncomeItem[] = [];
    let total: number | null = null;
    if (firstExpenseCol < 3) return {income, incomeTotal: 0};
    for (let r = headerRow; r <= grid.rowCount; r++) {
        const label = text(grid.get(r, 1));
        const cell = grid.get(r, 2);
        if (/^(合計|總收入|小計)$/.test(label)) {
            if (cell) total = cell.formula && num(cell) === null ? evalSum(grid, cell.formula) : num(cell);
            break;
        }
        const amount = num(cell);
        if (label && amount !== null && !cell?.formula) income.push({name: label, amount});
    }
    return {income, incomeTotal: total ?? income.reduce((s, it) => s + it.amount, 0)};
}

export interface SummaryMismatch {
    month: number;
    category: ExpenseCategory;
    summary: number;
    parsed: number;
}

/**
 * 讀「年度總表」的支出總表（B 欄＝大類、表頭列＝1月~12月），回傳 month → 各類金額。
 * 只用來核對解析結果，不作為資料來源。
 */
export function parseAnnualSummary(grid: SheetGrid): Map<number, Partial<Record<ExpenseCategory, number>>> {
    const result = new Map<number, Partial<Record<ExpenseCategory, number>>>();
    let headerRow = 0;
    const monthCols = new Map<number, number>();
    for (let r = 1; r <= Math.min(HEADER_SCAN_ROWS, grid.rowCount) && !headerRow; r++) {
        for (let c = 1; c <= grid.colCount; c++) {
            const m = /^(\d{1,2})月$/.exec(text(grid.get(r, c)));
            if (m) monthCols.set(c, Number(m[1]));
        }
        if (monthCols.size >= 6) headerRow = r;
        else monthCols.clear();
    }
    if (!headerRow) return result;
    for (let r = headerRow + 1; r <= grid.rowCount; r++) {
        for (let c = 1; c <= 3; c++) {
            const label = text(grid.get(r, c));
            if (!isExpenseCategory(label)) continue;
            for (const [col, month] of monthCols) {
                const v = num(grid.get(r, col));
                if (v === null) continue;
                const entry = result.get(month) ?? {};
                entry[label] = v;
                result.set(month, entry);
            }
            break;
        }
    }
    return result;
}

export function compareWithSummary(
    months: ReadonlyArray<MonthExpense>,
    summary: Map<number, Partial<Record<ExpenseCategory, number>>>,
): SummaryMismatch[] {
    const out: SummaryMismatch[] = [];
    for (const m of months) {
        const row = summary.get(m.month);
        if (!row) continue;
        for (const cat of EXPENSE_CATEGORIES) {
            const s = row[cat];
            if (s === undefined) continue;
            if (Math.abs(s - m.totals[cat]) > 0.5) out.push({month: m.month, category: cat, summary: s, parsed: m.totals[cat]});
        }
    }
    return out;
}

export interface WorkbookParseResult {
    months: MonthExpense[];
    skippedSheets: string[];
    errors: string[];
    summaryMismatches: SummaryMismatch[];
    hasSummary: boolean;
}

export function parseExpenseWorkbook(grids: ReadonlyArray<SheetGrid>, sourceFile: string, defaultYear: number): WorkbookParseResult {
    const months: MonthExpense[] = [];
    const skippedSheets: string[] = [];
    const errors: string[] = [];
    let summaryGrid: SheetGrid | null = null;

    // 有些月工作表 A1 可能沒寫年份，先從其他月份的標題推出年份
    let fallbackYear = defaultYear;
    for (const g of grids) {
        const t = parseMonthTitle(text(g.get(1, 1)));
        if (t?.year) {
            fallbackYear = t.year;
            break;
        }
    }

    for (const g of grids) {
        if (/年度|總表/.test(g.name)) {
            summaryGrid = g;
            continue;
        }
        if (!looksLikeMonthSheet(g)) {
            skippedSheets.push(g.name);
            continue;
        }
        try {
            months.push(parseMonthSheet(g, {fallbackYear, sourceFile}));
        } catch (err) {
            errors.push(err instanceof Error ? err.message : String(err));
        }
    }

    // 同一份檔案內同月份出現兩次：保留後面的，並提醒
    const byKey = new Map<string, MonthExpense>();
    for (const m of months) {
        if (byKey.has(m.sortKey)) errors.push(`「${m.sheetName}」與「${byKey.get(m.sortKey)!.sheetName}」是同一個月份，以「${m.sheetName}」為準`);
        byKey.set(m.sortKey, m);
    }
    const unique = [...byKey.values()].sort((a, b) => a.sortKey.localeCompare(b.sortKey));

    const summaryMismatches = summaryGrid ? compareWithSummary(unique, parseAnnualSummary(summaryGrid)) : [];
    return {months: unique, skippedSheets, errors, summaryMismatches, hasSummary: !!summaryGrid};
}
