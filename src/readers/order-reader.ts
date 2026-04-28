/**
 * 對應桌面版 `billReformat/core/OrderReader.kt`。
 * 排序檔每列第一欄為一個客戶代碼，輸出客戶代碼字串陣列。
 * 桌面版讀的是 `cell.rawValue`（原始字串、無格式化），
 * 網頁版這邊把 number 也轉成不含結尾 .0 的字串以對齊。
 */

import ExcelJS from 'exceljs';

export async function parseOrderList(source: File | ArrayBuffer | Blob): Promise<string[]> {
  const buf = source instanceof ArrayBuffer ? source : await source.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);

  const ws = wb.worksheets[0];
  if (!ws) throw new Error('找不到任何 worksheet');

  const result: string[] = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const v = row.getCell(1).value;
    const s = rawStringOf(v);
    if (s && s.trim().length > 0) result.push(s);
  });
  return result;
}

function rawStringOf(value: ExcelJS.CellValue): string | null {
  if (value === null || value === undefined) return null;
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'boolean') return String(value);
  if (typeof value === 'object') {
    const obj = value as unknown as Record<string, unknown>;
    if (Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>).map((rt) => rt.text ?? '').join('');
    }
    if ('result' in obj) return rawStringOf(obj.result as ExcelJS.CellValue);
    if ('text' in obj) return String(obj.text ?? '');
  }
  return String(value);
}
