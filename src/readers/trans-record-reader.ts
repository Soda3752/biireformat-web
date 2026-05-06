import ExcelJS from 'exceljs';
import Papa from 'papaparse';

/**
 * 對應桌面版 TransRecordReader.kt
 * 讀取銀行對帳單，支援兩種格式：
 *   - .csv（桌面版預設 Big5 編碼，亦相容 UTF-8 / UTF-8 BOM）
 *   - .xlsx（網路銀行常見匯出格式，內含 formula cell）
 *
 * 兩種格式都輸出統一的 string[][]（與桌面版 List<List<String>> 等價），
 * 第一列為標題（日期, 摘要, 幣別, 支出金額, 存入金額, 餘額, 備註, 轉出入帳號, 註記）。
 * 後續 bank-match-service / bank-name-writer 對 CSV 與 XLSX 來源一視同仁。
 */
export const parseTransRecord = async (file: File): Promise<string[][]> => {
  const lower = file.name.toLowerCase();
  if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
    return parseTransRecordXlsx(file);
  }
  return parseTransRecordCsv(file);
};

const parseTransRecordCsv = async (file: File): Promise<string[][]> => {
  const buffer = await file.arrayBuffer();
  const text = decodeCsv(buffer);

  const result = Papa.parse<string[]>(text, {
    header: false,
    skipEmptyLines: 'greedy',
    transform: (value) => (typeof value === 'string' ? value : String(value ?? '')),
  });

  if (result.errors.length > 0) {
    const firstFatal = result.errors.find((e) => e.type !== 'FieldMismatch');
    if (firstFatal) {
      throw new Error(`CSV 解析失敗：${firstFatal.message}`);
    }
  }

  return (result.data as string[][]).filter((row) => Array.isArray(row));
};

const parseTransRecordXlsx = async (file: File): Promise<string[][]> => {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('xlsx 未含任何工作表');
  }

  const colCount = Math.max(sheet.columnCount, sheet.actualColumnCount ?? 0);
  if (colCount === 0) return [];

  const out: string[][] = [];
  sheet.eachRow({includeEmpty: false}, (row) => {
    const cells: string[] = [];
    for (let c = 1; c <= colCount; c++) {
      cells.push(cellValueToString(row.getCell(c).value));
    }
    if (cells.every((s) => s.trim() === '')) return;
    out.push(cells);
  });
  return out;
};

/**
 * 偵測編碼：UTF-8 BOM 直接走 UTF-8；
 * 否則先嘗試 Big5，若 replacement char 比例過高就退回 UTF-8。
 */
const decodeCsv = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer);
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }

  try {
    const big5 = new TextDecoder('big5', { fatal: false }).decode(bytes);
    if (replacementRatio(big5) < 0.02) return big5;
  } catch {
    // ignore，走 utf-8
  }

  return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
};

const replacementRatio = (s: string): number => {
  if (s.length === 0) return 0;
  let count = 0;
  for (const ch of s) if (ch === '�') count++;
  return count / s.length;
};

/**
 * 將 ExcelJS cell value 轉字串。
 * - null / undefined → ''
 * - Date → yyyy/M/d（與 CSV `2026/4/20` 寫法對齊，不補零）
 * - Formula cell → 取 result 後遞迴
 * - RichText → 串接 text
 */
const cellValueToString = (value: ExcelJS.CellValue): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'boolean') return String(value);
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = value.getMonth() + 1;
    const d = value.getDate();
    return `${y}/${m}/${d}`;
  }
  if (typeof value === 'object') {
    const obj = value as unknown as Record<string, unknown>;
    if (Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>).map((rt) => rt.text ?? '').join('');
    }
    if ('result' in obj) {
      return cellValueToString(obj.result as ExcelJS.CellValue);
    }
    if ('text' in obj) {
      return String(obj.text ?? '');
    }
    if ('error' in obj) {
      return '';
    }
  }
  return String(value);
};
