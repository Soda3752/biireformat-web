/**
 * 對應桌面版 `models/BillDateInfo.kt`。
 * 從帳單第一列「客戶區間：xxx 至：xxx 日期區間：yyyy/MM/dd 至：yyyy/MM/dd」解析出年/月/日期範圍。
 */

import { parseDateOfTaiwanMinguoYear, parseDayOfMonth, parseMonth } from '../date-utility';

const PARAM_START_DATE = '日期區間';
const PARAM_END_DATE = '至';

export interface BillDateInfo {
  /** 民國年（字串格式以維持桌面版輸出一致） */
  year: string;
  /** 月份（字串格式以維持桌面版輸出一致） */
  month: string;
  /** 日期範圍，含起訖兩端（如 [16,17,18,19,20]） */
  dateRange: number[];
}

export function parseBillDateInfo(rowData: ReadonlyArray<string>): BillDateInfo {
  const startIdx = rowData.findIndex((v) => v.includes(PARAM_START_DATE));
  if (startIdx < 0) throw new Error('找不到「日期區間」欄位');

  // 對應 indexOfLast { it.contains("至") }
  let endIdx = -1;
  for (let i = rowData.length - 1; i >= 0; i--) {
    if (rowData[i].includes(PARAM_END_DATE)) {
      endIdx = i;
      break;
    }
  }
  if (endIdx < 0) throw new Error('找不到「至」欄位');

  const startDay = rowData[startIdx + 1];
  const endDay = rowData[endIdx + 1];

  const year = String(parseDateOfTaiwanMinguoYear(startDay));
  const month = String(parseMonth(startDay));
  const start = parseDayOfMonth(startDay);
  const end = parseDayOfMonth(endDay);

  const dateRange: number[] = [];
  for (let i = start; i <= end; i++) dateRange.push(i);

  return { year, month, dateRange };
}
