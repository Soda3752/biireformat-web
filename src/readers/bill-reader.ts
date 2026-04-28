/**
 * 對應桌面版 `billReformat/core/Reader.kt`。
 * 用 ExcelJS 讀第一張 sheet，逐列把 cell 串成 string[]，
 * 再交給 ExcelRowType 判別「列類型」並觸發 callback。
 */

import ExcelJS from 'exceljs';

import { getRowType, type ExcelRowType } from '@/domain/excel-row-data';

export interface ParsedRow {
  type: ExcelRowType;
  values: string[];
}

export type RowHandler = (row: ParsedRow) => void;

/**
 * 解析帳單 .xlsx，僅針對符合 ExcelRowType 的列觸發 handler。
 */
export async function parseBillFile(
  source: File | ArrayBuffer | Blob,
  handler: RowHandler
): Promise<void> {
  const buf = await toArrayBuffer(source);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  const ws = wb.worksheets[0];
  if (!ws) throw new Error('找不到任何 worksheet');

  ws.eachRow({ includeEmpty: false }, (row) => {
    const values: string[] = [];
    row.eachCell({ includeEmpty: false }, (cell) => {
      const s = cellValueToString(cell.value);
      // 對齊桌面版：null cell 略過，不保留空字串占位
      if (s !== null) values.push(s);
    });
    if (values.length === 0) return;
    const type = getRowType(values);
    if (type) handler({ type, values });
  });
}

async function toArrayBuffer(source: File | ArrayBuffer | Blob): Promise<ArrayBuffer> {
  if (source instanceof ArrayBuffer) return source;
  if (typeof Blob !== 'undefined' && source instanceof Blob) return source.arrayBuffer();
  // File 是 Blob 的子類，理論上前一條已涵蓋；保底再呼叫一次
  return (source as File).arrayBuffer();
}

/**
 * 將 ExcelJS cell value 對齊桌面版 POI 的轉字串行為：
 * - null/undefined → null（呼叫端會略過）
 * - string → 原樣
 * - number / boolean → toString
 * - Date → 格式化為 yyyy/MM/dd（與帳單預期格式一致）
 * - RichText / Hyperlink / Formula → 取顯示文字
 */
function cellValueToString(value: ExcelJS.CellValue): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'boolean') return String(value);
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, '0');
    const d = String(value.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  }
  // 物件型 cell value
  if (typeof value === 'object') {
    const obj = value as unknown as Record<string, unknown>;
    // RichText
    if (Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>).map((rt) => rt.text ?? '').join('');
    }
    // Formula：result 可能還是 string / number / Date
    if ('result' in obj) {
      return cellValueToString(obj.result as ExcelJS.CellValue);
    }
    // Hyperlink
    if ('text' in obj) {
      return String(obj.text ?? '');
    }
    // Error cell
    if ('error' in obj) {
      return '';
    }
  }
  return String(value);
}
