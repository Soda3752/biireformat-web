// 移植自 invoice-helper/src/dates.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 民國年日期工具
export type Ymd = { y: number; m: number; d: number }; // y 為西元年

export function rocString(ymd: Ymd): string {
  const roc = ymd.y - 1911;
  return `${roc}/${String(ymd.m).padStart(2, "0")}/${String(ymd.d).padStart(2, "0")}`;
}

export function lastDayOfMonth(y: number, m: number): number {
  return new Date(Date.UTC(y, m, 0)).getUTCDate();
}

/** 上期 → 15 日；下期 → 月底 */
export function invoiceDate(y: number, m: number, half: "上期" | "下期"): Ymd {
  return { y, m, d: half === "上期" ? 15 : lastDayOfMonth(y, m) };
}

/**
 * 解析三聯清單裡的日期：Excel 日期、"115/09/15"、"2026-09-15"、"115.9.15"、"9/15"、"9月15日"。
 * 沒有年份時用 defaultYear（西元）。
 */
export function parseDateLoose(v: { date: Date | null; text: string; num: number | null }, defaultYear: number): Ymd | null {
  if (v.date) return { y: v.date.getUTCFullYear(), m: v.date.getUTCMonth() + 1, d: v.date.getUTCDate() };
  if (v.num !== null && !v.text) {
    // Excel 序號日期
    if (v.num > 20000 && v.num < 80000) {
      const dt = new Date(Date.UTC(1899, 11, 30) + v.num * 86400000);
      return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
    }
    return null;
  }
  const t = v.text.replace(/\s/g, "");
  let m = /^(\d{2,4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/.exec(t);
  if (m) {
    let y = Number(m[1]);
    if (y < 1911) y += 1911;
    return { y, m: Number(m[2]), d: Number(m[3]) };
  }
  m = /^(\d{1,2})[\/\-.月](\d{1,2})日?$/.exec(t);
  if (m) return { y: defaultYear, m: Number(m[1]), d: Number(m[2]) };
  return null;
}

/** 從檔名抓月份，例如「10月_已自動對帳_最新.xlsx」→ 10 */
export function monthFromFilename(file: string): number | null {
  const base = file.split(/[\\/]/).pop() ?? file;
  const m = /(\d{1,2})\s*月/.exec(base);
  if (!m) return null;
  const n = Number(m[1]);
  return n >= 1 && n <= 12 ? n : null;
}
