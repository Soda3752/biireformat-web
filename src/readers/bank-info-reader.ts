import ExcelJS from 'exceljs';

import type {BankInfo} from '@/domain/models/bank-info';
import {equalsBankInfo} from '@/domain/models/bank-info';

/**
 * 對應桌面版 BankInfoReader.kt
 * 讀取末五碼對照 .xlsx。支援兩種 layout：
 *   - 新版（4 欄）：客戶名稱 / 店家編號 / 線別 / 末五碼
 *   - 舊版（3 欄）：客戶名稱 / 線別 / 末五碼（storeCode 補空字串）
 * 透過第一列 header 偵測；header 一律跳過。末五碼空白者略過；重複者去重。
 */
export const parseBankInfo = async (file: File): Promise<BankInfo[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('找不到工作表');

  const result: BankInfo[] = [];
    const headerRow = sheet.getRow(1);
    const isNewLayout = readCell(headerRow.getCell(2).value).trim() === '店家編號';

    sheet.eachRow({includeEmpty: false}, (row, rowNumber) => {
        if (rowNumber === 1) return; // 跳過 header

        let info: BankInfo;
        if (isNewLayout) {
            info = {
                customerName: readCell(row.getCell(1).value),
                storeCode: readCell(row.getCell(2).value),
                customerLine: readCell(row.getCell(3).value),
                lastFiveDigit: readCell(row.getCell(4).value),
            };
        } else {
            info = {
                customerName: readCell(row.getCell(1).value),
                storeCode: '',
                customerLine: readCell(row.getCell(2).value),
                lastFiveDigit: readCell(row.getCell(3).value),
            };
        }

        if (info.lastFiveDigit.trim().length === 0) return;
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
