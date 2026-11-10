/**
 * 「支出」分頁下方「分析」區塊的計算（純函式，方便測試）。
 * 所有分析只用「有網站營收」的月份；呼叫端傳入全部月份也沒關係，這裡會先過濾。
 */

import type {MonthSummary} from './expense-summary';

/** 支出比「該類其他月份平均」高出這個比例以上就標紅（0.2 = 20%）。 */
export const ANOMALY_THRESHOLD = 0.2;
/** 食品業一般參考：原物料＋人事占營收在這個比例以內算健康（0.65 = 65%）。 */
export const HEALTHY_MATERIAL_LABOR_RATIO = 0.65;
/** 滾動平均毛利率的月數（含當月）。 */
export const ROLLING_MONTHS = 3;

export function revenueRows(rows: ReadonlyArray<MonthSummary>): MonthSummary[] {
    return rows.filter((r) => r.revenue !== null);
}

/* ---------- 1. 各項占營收 % ---------- */

export const SHARE_KEYS = ['原物料', '人事', '營運', '車輛', '其他店務', '還款', '家用', '營業結餘'] as const;
export type ShareKey = (typeof SHARE_KEYS)[number];

export interface ShareRow {
    /** 年度列為 null */
    sortKey: string | null;
    month: number | null;
    revenue: number;
    amounts: Record<ShareKey, number>;
    /** 各項 ÷ 營收（營收為 0 時為 null） */
    pct: Record<ShareKey, number | null>;
    /** （原物料＋人事）÷ 營收 */
    materialLabor: number | null;
}

type ShareSource = Pick<MonthSummary, 'totals' | 'otherShop' | 'repayment' | 'household'> & {revenue: number; operatingBalance: number};

function shareRow(sortKey: string | null, month: number | null, s: ShareSource): ShareRow {
    const amounts: Record<ShareKey, number> = {
        原物料: s.totals['原物料'],
        人事: s.totals['人事'],
        營運: s.totals['營運'],
        車輛: s.totals['車輛'],
        其他店務: s.otherShop,
        還款: s.repayment,
        家用: s.household,
        營業結餘: s.operatingBalance,
    };
    const ratio = (v: number) => (s.revenue ? v / s.revenue : null);
    const pct = Object.fromEntries(SHARE_KEYS.map((k) => [k, ratio(amounts[k])])) as Record<ShareKey, number | null>;
    return {sortKey, month, revenue: s.revenue, amounts, pct, materialLabor: ratio(amounts['原物料'] + amounts['人事'])};
}

export function revenueShares(rows: ReadonlyArray<MonthSummary>): {months: ShareRow[]; year: ShareRow | null} {
    const list = revenueRows(rows);
    const months = list.map((r) =>
        shareRow(r.sortKey, r.month, {...r, revenue: r.revenue!, operatingBalance: r.operatingBalance!}),
    );
    if (list.length === 0) return {months, year: null};
    const totals = {...list[0].totals};
    for (const k of Object.keys(totals) as Array<keyof typeof totals>) totals[k] = list.reduce((s, r) => s + r.totals[k], 0);
    const sum = (f: (r: MonthSummary) => number) => list.reduce((s, r) => s + f(r), 0);
    const year = shareRow(null, null, {
        totals,
        otherShop: sum((r) => r.otherShop),
        repayment: sum((r) => r.repayment),
        household: sum((r) => r.household),
        revenue: sum((r) => r.revenue!),
        operatingBalance: sum((r) => r.operatingBalance!),
    });
    return {months, year};
}

/* ---------- 2. 近 3 個月平均毛利率 ---------- */

export interface RollingMargin {
    sortKey: string;
    month: number;
    /** 單月毛利率 */
    margin: number | null;
    /** 近 N 個月（含當月）合計毛利 ÷ 合計營收 */
    average: number | null;
    /** 實際用到幾個有營收的月份（年初或缺帳單時可能少於 N） */
    monthsUsed: number;
}

const monthIndex = (r: MonthSummary) => r.year * 12 + r.month;

export function rollingMargins(rows: ReadonlyArray<MonthSummary>, window = ROLLING_MONTHS): RollingMargin[] {
    const list = revenueRows(rows);
    return list.map((r) => {
        const idx = monthIndex(r);
        const inWindow = list.filter((x) => monthIndex(x) <= idx && monthIndex(x) > idx - window);
        const rev = inWindow.reduce((s, x) => s + x.revenue!, 0);
        const gross = inWindow.reduce((s, x) => s + x.grossProfit!, 0);
        return {
            sortKey: r.sortKey,
            month: r.month,
            margin: r.revenue ? r.grossProfit! / r.revenue : null,
            average: rev ? gross / rev : null,
            monthsUsed: inWindow.length,
        };
    });
}

/* ---------- 3. 支出異常 ---------- */

export const ANOMALY_KEYS = ['營運', '車輛', '人事', '其他店務', '家用'] as const;
export type AnomalyKey = (typeof ANOMALY_KEYS)[number];

export function anomalyAmount(r: MonthSummary, key: AnomalyKey): number {
    if (key === '其他店務') return r.otherShop;
    if (key === '家用') return r.household;
    return r.totals[key];
}

export interface Anomaly {
    sortKey: string;
    month: number;
    key: AnomalyKey;
    amount: number;
    /** 該類「其他月份」的平均 */
    average: number;
    /** 高出平均的比例（0.25 = 高 25%） */
    over: number;
}

export function detectAnomalies(rows: ReadonlyArray<MonthSummary>, threshold = ANOMALY_THRESHOLD): Anomaly[] {
    const list = revenueRows(rows);
    if (list.length < 2) return [];
    const out: Anomaly[] = [];
    for (const r of list) {
        for (const key of ANOMALY_KEYS) {
            const others = list.filter((x) => x !== r);
            const average = others.reduce((s, x) => s + anomalyAmount(x, key), 0) / others.length;
            const amount = anomalyAmount(r, key);
            if (average <= 0) continue;
            const over = amount / average - 1;
            if (over >= threshold - 1e-9) out.push({sortKey: r.sortKey, month: r.month, key, amount, average, over});
        }
    }
    return out;
}

/* ---------- 4. 每月至少要賣多少才不虧 ---------- */

export interface Breakeven {
    months: number;
    /** 年度累計毛利率（扣原物料） */
    grossMargin: number | null;
    avgRevenue: number;
    /** 每月平均固定支出：營運＋車輛＋人事＋其他店務 */
    avgFixed: number;
    avgRepayment: number;
    avgHousehold: number;
    /** 營業不虧所需月營收；毛利率 ≤ 0 時為 null（怎麼賣都不夠） */
    operating: number | null;
    /** 要還款也不虧 */
    withRepayment: number | null;
    /** 連家用也打平 */
    withHousehold: number | null;
}

export function breakeven(rows: ReadonlyArray<MonthSummary>): Breakeven | null {
    const list = revenueRows(rows);
    if (list.length === 0) return null;
    const n = list.length;
    const sum = (f: (r: MonthSummary) => number) => list.reduce((s, r) => s + f(r), 0);
    const revenue = sum((r) => r.revenue!);
    const grossMargin = revenue ? sum((r) => r.grossProfit!) / revenue : null;
    const avgFixed = sum((r) => r.totals['營運'] + r.totals['車輛'] + r.totals['人事'] + r.otherShop) / n;
    const avgRepayment = sum((r) => r.repayment) / n;
    const avgHousehold = sum((r) => r.household) / n;
    const need = (cost: number) => (grossMargin !== null && grossMargin > 0 ? cost / grossMargin : null);
    return {
        months: n,
        grossMargin,
        avgRevenue: revenue / n,
        avgFixed,
        avgRepayment,
        avgHousehold,
        operating: need(avgFixed),
        withRepayment: need(avgFixed + avgRepayment),
        withHousehold: need(avgFixed + avgRepayment + avgHousehold),
    };
}

/* ---------- 5. 家用占還款後結餘 % ---------- */

export interface HouseholdShare {
    sortKey: string | null;
    month: number | null;
    household: number;
    /** 分母：還款後結餘 */
    afterRepayBalance: number;
    /** 家用 ÷ 還款後結餘；還款後結餘 ≤ 0 時為 null（家用全靠老本／借錢） */
    ratio: number | null;
}

export function householdShares(rows: ReadonlyArray<MonthSummary>): {months: HouseholdShare[]; year: HouseholdShare | null} {
    const list = revenueRows(rows);
    const make = (sortKey: string | null, month: number | null, household: number, base: number): HouseholdShare => ({
        sortKey,
        month,
        household,
        afterRepayBalance: base,
        ratio: base > 0 ? household / base : null,
    });
    const months = list.map((r) => make(r.sortKey, r.month, r.household, r.afterRepayBalance!));
    if (list.length === 0) return {months, year: null};
    return {
        months,
        year: make(
            null,
            null,
            list.reduce((s, r) => s + r.household, 0),
            list.reduce((s, r) => s + r.afterRepayBalance!, 0),
        ),
    };
}
