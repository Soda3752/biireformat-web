/**
 * Excel 讀取共用 helper：給設定頁的 cargo / daily / lastFive 三個 pane 重用，
 * 以一致的方式把第一張 worksheet 轉成「字串二維陣列」。
 *
 * - 完全空白列會被略過，但列內中間的空白儲存格會以空字串保留位置（重要：
 *   單日數量的「分類」欄常為空，必須能維持「分類|編號|名稱」的三欄定位）。
 * - 數值會去掉尾巴 `.0`，與桌面版 BankInfoReader.kt 對齊。
 */

import ExcelJS from 'exceljs';

export async function readXlsxAsRows(file: File): Promise<string[][]> {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const sheet = wb.worksheets[0];
    if (!sheet) throw new Error('找不到工作表');

    const rows: string[][] = [];
    sheet.eachRow({includeEmpty: false}, (row) => {
        const values = row.values as Array<ExcelJS.CellValue | undefined>;
        // values[0] 是 ExcelJS 的 1-based 佔位 undefined，從 index 1 起為實際儲存格
        const cells: string[] = [];
        for (let i = 1; i < values.length; i++) {
            cells.push(readCellAsText(values[i] as ExcelJS.CellValue));
        }
        if (cells.some((c) => c.length > 0)) rows.push(cells);
    });
    return rows;
}

export function readCellAsText(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return formatNumber(value);
    if (typeof value === 'boolean') return value.toString();
    if (value instanceof Date) return value.toISOString();
    if (typeof value === 'object') {
        if ('text' in value && typeof value.text === 'string') return value.text;
        if ('richText' in value && Array.isArray(value.richText)) {
            return value.richText.map((rt) => rt.text).join('');
        }
        if ('result' in value) return readCellAsText(value.result as ExcelJS.CellValue);
        if ('formula' in value) return '';
    }
    return String(value);
}

function formatNumber(n: number): string {
    if (Number.isInteger(n)) return n.toString();
    return n.toString().replace(/\.0$/, '');
}
