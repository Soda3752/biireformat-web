import ExcelJS from 'exceljs';

import type { BankInfo } from '@/domain/models/bank-info';
import { equalsBankInfo } from '@/domain/models/bank-info';

/**
 * 對應桌面版 BankInfoReader.kt
 * 讀取末五碼對照 .xlsx：第一欄=客戶名、第二欄=線別、第三欄=末五碼。
 * 第三欄空白的列會被略過；重複者不重複加入。
 */
export const parseBankInfo = async (file: File): Promise<BankInfo[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('找不到工作表');

  const result: BankInfo[] = [];

  sheet.eachRow({ includeEmpty: false }, (row) => {
    const customerName = readCell(row.getCell(1).value);
    const customerLine = readCell(row.getCell(2).value);
    const lastFiveDigit = readCell(row.getCell(3).value);

    if (lastFiveDigit.trim().length === 0) return;

    const info: BankInfo = { customerName, customerLine, lastFiveDigit };
    if (!result.some((existing) => equalsBankInfo(existing, info))) {
      result.push(info);
    }
  });

  return result;
};

const readCell = (value: ExcelJS.CellValue): string => {
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
    if ('result' in value) return readCell(value.result as ExcelJS.CellValue);
    if ('formula' in value) return '';
  }
  return String(value);
};

/** 對應 cell.numericCellValue.toString().replace(".0", "") */
const formatNumber = (n: number): string => {
  if (Number.isInteger(n)) return n.toString();
  return n.toString().replace(/\.0$/, '');
};
