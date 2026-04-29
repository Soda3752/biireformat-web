/**
 * 日期工具，對應桌面版 `billReformat/core/DateUtility.kt`。
 * 桌面版輸入字串格式為 `yyyy/MM/dd`，民國年 = 西元年 - 1911。
 */

const ISO_LIKE = /^(\d{4})\/(\d{1,2})\/(\d{1,2})/;

export function parseDate(dateString: string): Date {
  const m = ISO_LIKE.exec(dateString.trim());
  if (!m) {
    throw new Error(`日期格式錯誤: ${dateString}`);
  }
  const [, y, mo, d] = m;
  return new Date(Number(y), Number(mo) - 1, Number(d));
}

export function parseDateOfTaiwanMinguoYear(dateString: string): number {
  return parseDate(dateString).getFullYear() - 1911;
}

export function parseDayOfMonth(dateString: string): number {
  return parseDate(dateString).getDate();
}

export function parseMonth(dateString: string): number {
  return parseDate(dateString).getMonth() + 1;
}

/**
 * 從民國年 + 月 + 日推算星期（0=週日, 1=週一, ..., 6=週六）。
 * 民國年 + 1911 = 西元年。
 */
export function getWeekday(minguoYear: number | string, month: number, day: number): number {
  const y = Number(minguoYear) + 1911;
  return new Date(y, month - 1, day).getDay();
}

/** 0~6 對應的中文短名 */
export const WEEKDAY_NAMES = ['日', '一', '二', '三', '四', '五', '六'] as const;
