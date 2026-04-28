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
