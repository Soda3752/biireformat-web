import type ExcelJS from 'exceljs';
import type {BankInfo} from '@/domain/models/bank-info';
import {collectAccounts, findMatchesFromCandidates} from '@/domain/bank-match-service';
import {buildStyle, createWorkbook, workbookToBlob,} from '@/infra/excel-service';

/**
 * 對帳結果輸出
 *
 * 精簡欄位設計（一列一筆入帳）：
 *   狀態 / 客戶 / 線別 / 日期 / 摘要 / 存入 / 配對來源
 *
 * - 多客戶配對：客戶與線別欄各自以「、」串接，仍維持一列一筆
 * - 未配對列整列染淺黃底色 + 狀態顯示 ✗，方便人工追查
 * - 凍結表頭、欄寬依內容自動
 */

const HEADER = ['狀態', '客戶', '線別', '日期', '摘要', '存入', '配對來源'] as const;

const ARGB_HEADER_BG = 'FF3E5F8A';
const ARGB_HEADER_FG = 'FFFFFFFF';
const ARGB_UNMATCHED_BG = 'FFFFF4D6';
const ARGB_BORDER = 'FFD0D7DE';

const HEADER_STYLE = buildStyle({
    font: {bold: true, color: ARGB_HEADER_FG, size: 12},
    fill: ARGB_HEADER_BG,
    align: 'center',
});

const BASE_CELL_STYLE = buildStyle({
    font: {size: 11},
    align: {vertical: 'middle', wrapText: true} as Partial<ExcelJS.Alignment>,
});

const UNMATCHED_FILL = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: {argb: ARGB_UNMATCHED_BG},
} as ExcelJS.FillPattern;

const STANDARD_HEADER_FIRST = '日期';

export const writeBankNameMerged = async (
  bankInfos: ReadonlyArray<BankInfo>,
  originData: ReadonlyArray<ReadonlyArray<string>>
): Promise<Blob> => {
  const workbook = createWorkbook();
  const sheet = workbook.addWorksheet('對帳結果');

    const headerRow = sheet.addRow([...HEADER]);
    headerRow.eachCell((cell) => {
        cell.style = HEADER_STYLE;
        cell.border = thinBorderAll();
    });
    headerRow.height = 22;

  for (const rowInfo of originData) {
    const trimmed = dropTrailingBlanks(rowInfo);
      if (trimmed.length === 0) continue;
      if (isHeaderRow(trimmed)) continue;

      const accounts = collectAccounts(trimmed);
      const matches = findMatchesFromCandidates(accounts, bankInfos);
      const isMatched = matches.length > 0;

      const customers = isMatched
          ? matches.map((m) => m.customerName).join('、')
          : '(未配對)';
      const lines = isMatched
          ? matches.map((m) => m.customerLine).filter((s) => s.length > 0).join('、')
          : '';

      const date = trimmed[0] ?? '';
      const summary = trimmed[1] ?? '';
      const deposit = parseAmount(trimmed[4] ?? '');
      const source = accounts.join('、');

      const row = sheet.addRow([
          isMatched ? '✓' : '✗',
          customers,
          lines,
          date,
          summary,
          deposit,
          source,
      ]);

      row.eachCell((cell, colNumber) => {
          cell.style = {...BASE_CELL_STYLE};
          cell.border = thinBorderAll();
          if (colNumber === 1) {
              cell.alignment = {...cell.alignment, horizontal: 'center'};
              cell.font = {...cell.font, bold: true, color: {argb: isMatched ? 'FF1F8A4C' : 'FFC23E3E'}};
          } else if (colNumber === 6) {
              cell.alignment = {...cell.alignment, horizontal: 'right'};
              cell.numFmt = '#,##0';
          }
          if (!isMatched) cell.fill = UNMATCHED_FILL;
      });
  }

    sheet.views = [{state: 'frozen', ySplit: 1}];
    autoFitColumns(sheet);

  return workbookToBlob(workbook);
};

const dropTrailingBlanks = (row: ReadonlyArray<string>): string[] => {
  const out = [...row];
  while (out.length > 0 && (out[out.length - 1] ?? '').trim() === '') {
    out.pop();
  }
  return out;
};

const isHeaderRow = (row: ReadonlyArray<string>): boolean =>
    (row[0] ?? '').trim() === STANDARD_HEADER_FIRST && (row[1] ?? '').trim() === '摘要';

const parseAmount = (raw: string): number | string => {
    const s = raw.replace(/,/g, '').trim();
    if (s === '') return '';
    const n = Number(s);
    return Number.isFinite(n) ? n : raw;
};

const thinBorderAll = (): Partial<ExcelJS.Borders> => ({
    top: {style: 'thin', color: {argb: ARGB_BORDER}},
    left: {style: 'thin', color: {argb: ARGB_BORDER}},
    bottom: {style: 'thin', color: {argb: ARGB_BORDER}},
    right: {style: 'thin', color: {argb: ARGB_BORDER}},
});

/**
 * 依每欄實際內容寬度自動調整 width。
 * CJK 字寬以 2 計，半形以 1 計，再加 2 字 padding；夾在 [min, max] 範圍。
 */
const autoFitColumns = (
    sheet: ExcelJS.Worksheet,
    minWidth = 6,
    maxWidth = 40
): void => {
    const widths: number[] = [];
    sheet.eachRow({includeEmpty: false}, (row) => {
        row.eachCell({includeEmpty: false}, (cell, colNumber) => {
            const text = cellValueText(cell.value);
            const w = displayWidth(text) + 2;
            const idx = colNumber - 1;
            if (widths[idx] === undefined || w > widths[idx]) widths[idx] = w;
        });
    });
    widths.forEach((w, i) => {
        if (w === undefined) return;
        sheet.getColumn(i + 1).width = Math.min(Math.max(w, minWidth), maxWidth);
    });
};

const cellValueText = (v: ExcelJS.CellValue): string => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return v.toLocaleString();
    if (typeof v === 'boolean') return String(v);
    if (typeof v === 'object') {
        const obj = v as unknown as Record<string, unknown>;
        if ('text' in obj) return String(obj.text ?? '');
        if (Array.isArray(obj.richText)) {
            return (obj.richText as Array<{ text?: string }>).map((rt) => rt.text ?? '').join('');
        }
    }
    return String(v);
};

const displayWidth = (s: string): number => {
    let w = 0;
    for (const ch of s) {
        const code = ch.codePointAt(0)!;
        if (
            (code >= 0x4e00 && code <= 0x9fff) ||
            (code >= 0x3000 && code <= 0x303f) ||
            (code >= 0x3040 && code <= 0x30ff) ||
            (code >= 0xac00 && code <= 0xd7af) ||
            (code >= 0xff00 && code <= 0xffef)
        ) {
            w += 2;
        } else {
            w += 1;
        }
    }
    return w;
};

/** 對帳結果_${YYYYMMDD}.xlsx */
export const buildBankResultFilename = (date = new Date()): string => {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `對帳結果_${yyyy}${mm}${dd}.xlsx`;
};
