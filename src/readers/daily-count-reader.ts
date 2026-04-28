import ExcelJS from 'exceljs';

import { loadDailyReportTemplate } from '@/domain/daily-report-loader';
import {
  cloneProductMap,
  type DailyProduct,
  type DailyProductMap,
} from '@/domain/models/daily-product';

export interface DailyCountResult {
  readonly map: DailyProductMap;
  readonly matched: number;
  readonly otherCount: number;
}

/**
 * 對應桌面版 DailyCountViewModel.setDataByInputFile()
 * 讀取單日銷售 .xlsx：
 * - 第 0 列為 header 略過
 * - 每列：code(欄1)、name(欄2)、count(欄3)
 * - 若 code 在模板中存在 → 更新 count
 * - 否則丟入「其他」分組（建立新 Product）
 */
export const parseDailyCount = async (file: File): Promise<DailyCountResult> => {
  const template = await loadDailyReportTemplate();
  const productMap = cloneProductMap(template);

  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('找不到工作表');

  // 建立 code → product 索引以加速查找
  const codeIndex = new Map<string, DailyProduct>();
  for (const products of productMap.values()) {
    for (const p of products) codeIndex.set(p.code, p);
  }

  let matched = 0;
  let otherCount = 0;

  // POI lastRowNum 從 0 起，ExcelJS 從 1 起；rowIndex=0 (header) 略過 → 從第 2 列開始
  const lastRow = sheet.actualRowCount > 0 ? sheet.rowCount : 0;
  for (let rowIndex = 2; rowIndex <= lastRow; rowIndex++) {
    const row = sheet.getRow(rowIndex);
    if (!row) continue;

    const code = readCell(row.getCell(1).value);
    const name = readCell(row.getCell(2).value);
    const countStr = readCell(row.getCell(3).value);

    if (code.length === 0 || name.length === 0) continue;

    const count = parseCount(countStr);
    const existing = codeIndex.get(code);
    if (existing) {
      existing.count = count;
      matched++;
    } else {
      const newProduct: DailyProduct = { code, name, groupName: '其他', count };
      let otherGroup = productMap.get('其他');
      if (!otherGroup) {
        otherGroup = [];
        productMap.set('其他', otherGroup);
      }
      otherGroup.push(newProduct);
      codeIndex.set(code, newProduct);
      otherCount++;
    }
  }

  return { map: productMap, matched, otherCount };
};

const readCell = (value: ExcelJS.CellValue): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return value.toString();
  if (typeof value === 'boolean') return value.toString();
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'object') {
    if ('text' in value && typeof value.text === 'string') return value.text;
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map((rt) => rt.text).join('');
    }
    if ('result' in value) return readCell(value.result as ExcelJS.CellValue);
  }
  return String(value);
};

/** 對應 countStr.toDoubleOrNull()?.toInt() ?: 0 */
const parseCount = (s: string): number => {
  const n = Number.parseFloat(s);
  if (!Number.isFinite(n)) return 0;
  return Math.trunc(n);
};
