/**
 * 對應桌面版 `models/BillDateInfo.kt`。
 * 從帳單第一列「客戶區間：xxx 至：xxx 日期區間：yyyy/MM/dd 至：yyyy/MM/dd」解析出年/月/日期範圍。
 *
 * 與桌面版差異：支援跨月日期區間（例如 3/31~4/14 → dateRange=[31,1,2..14], dates=[3/31,4/1..4/14]）。
 */

import {parseDate} from '../date-utility';

const PARAM_START_DATE = '日期區間';
const PARAM_END_DATE = '至';
const ONE_DAY_MS = 24 * 60 * 60 * 1000;

export interface BillDateInfo {
  /** 民國年（字串格式以維持桌面版輸出一致），來自起始日 */
  year: string;
  /** 月份（字串格式以維持桌面版輸出一致），來自起始日 */
  month: string;
  /** 顯示用 day 序列，跨月時會出現「31, 1, 2..」這種非單調序列 */
  dateRange: number[];
  /** 真實日期序列（length 與 dateRange 相同），用於支援跨月與日期校正 */
  dates: Date[];
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

  const startDateStr = rowData[startIdx + 1];
  const endDateStr = rowData[endIdx + 1];

  const startDate = parseDate(startDateStr);
  const endDate = parseDate(endDateStr);
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error(`日期區間結束日早於起始日: ${startDateStr} ~ ${endDateStr}`);
  }

  const dates: Date[] = [];
  for (let t = startDate.getTime(); t <= endDate.getTime(); t += ONE_DAY_MS) {
    dates.push(new Date(t));
  }

  const year = String(startDate.getFullYear() - 1911);
  const month = String(startDate.getMonth() + 1);
  const dateRange = dates.map((d) => d.getDate());

  return {year, month, dateRange, dates};
}
