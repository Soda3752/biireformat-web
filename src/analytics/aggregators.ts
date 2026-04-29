/**
 * 共用聚合工具：所有 chart adapter 從這裡拿聚合結果，
 * 避免每個 adapter 自己寫一遍 group-by。
 */

import type {AnalyticsRow} from './dataset-builder';

export interface GroupSum {
    key: string;
    amount: number;
    count: number;
    costAmount: number;
    profit: number;
    /** 該分組是否所有 row 皆未填成本（首列 isCostUnset 為 true 後，遇到任一已填即翻回 false）。 */
    allCostUnset: boolean;
}

/** 依任意欄位分組加總 amount + count + costAmount + profit，回傳排序好的陣列。 */
export function groupBy<K extends keyof AnalyticsRow>(
    rows: ReadonlyArray<AnalyticsRow>,
    key: K,
    sortBy: 'amount' | 'count' | 'profit' | 'key' = 'amount',
    desc = true
): GroupSum[] {
    const map = new Map<string, GroupSum>();
    for (const r of rows) {
        const k = String(r[key]);
        let g = map.get(k);
        if (!g) {
            g = {key: k, amount: 0, count: 0, costAmount: 0, profit: 0, allCostUnset: true};
            map.set(k, g);
        }
        g.amount += r.amount;
        g.count += r.count;
        g.costAmount += r.costAmount;
        g.profit += r.profit;
        if (!r.isCostUnset) g.allCostUnset = false;
    }
    const arr = [...map.values()];
    arr.sort((a, b) => {
        if (sortBy === 'key') return desc ? b.key.localeCompare(a.key) : a.key.localeCompare(b.key);
        const av = a[sortBy];
        const bv = b[sortBy];
        return desc ? bv - av : av - bv;
    });
    return arr;
}

/** Top N */
export function topN(groups: ReadonlyArray<GroupSum>, n: number): GroupSum[] {
    return groups.slice(0, n);
}

/** 客戶名稱+編號顯示用 */
export function customerLabel(rows: ReadonlyArray<AnalyticsRow>): Map<string, string> {
    const m = new Map<string, string>();
    for (const r of rows) {
        if (!m.has(r.customerCode)) m.set(r.customerCode, `${r.customerName}(${r.customerCode})`);
    }
    return m;
}

/** 依「客戶」分組（用 customerCode 為 key，但顯示用 name+code） */
export function groupByCustomer(
    rows: ReadonlyArray<AnalyticsRow>,
    sortBy: 'amount' | 'count' | 'profit' = 'amount',
    desc = true
): GroupSum[] {
    const labelMap = customerLabel(rows);
    const groups = groupBy(rows, 'customerCode', sortBy, desc);
    return groups.map((g) => ({...g, key: labelMap.get(g.key) ?? g.key}));
}

/** 每日金額/數量序列；x 為 day 1~31 字串 */
export interface DailyPoint {
    day: number;
    amount: number;
    count: number;
    costAmount: number;
    profit: number;
    allCostUnset: boolean;
}

function newDailyPoint(day: number): DailyPoint {
    return {day, amount: 0, count: 0, costAmount: 0, profit: 0, allCostUnset: true};
}

export function dailySeries(rows: ReadonlyArray<AnalyticsRow>): DailyPoint[] {
    const map = new Map<number, DailyPoint>();
    for (const r of rows) {
        let p = map.get(r.day);
        if (!p) {
            p = newDailyPoint(r.day);
            map.set(r.day, p);
        }
        p.amount += r.amount;
        p.count += r.count;
        p.costAmount += r.costAmount;
        p.profit += r.profit;
        if (!r.isCostUnset) p.allCostUnset = false;
    }
    return [...map.values()].sort((a, b) => a.day - b.day);
}

/** 星期金額/數量序列（0=日 ~ 6=六） */
export function weekdaySeries(rows: ReadonlyArray<AnalyticsRow>): DailyPoint[] {
    const map = new Map<number, DailyPoint>();
    for (let i = 0; i < 7; i++) map.set(i, newDailyPoint(i));
    for (const r of rows) {
        const p = map.get(r.weekday)!;
        p.amount += r.amount;
        p.count += r.count;
        p.costAmount += r.costAmount;
        p.profit += r.profit;
        if (!r.isCostUnset) p.allCostUnset = false;
    }
    return [...map.values()];
}

/** distinct count */
export function distinctCount<K extends keyof AnalyticsRow>(
    rows: ReadonlyArray<AnalyticsRow>,
    key: K
): number {
    const s = new Set<string>();
    for (const r of rows) s.add(String(r[key]));
    return s.size;
}

export function sumAmount(rows: ReadonlyArray<AnalyticsRow>): number {
    let s = 0;
    for (const r of rows) s += r.amount;
    return s;
}

export function sumCount(rows: ReadonlyArray<AnalyticsRow>): number {
    let s = 0;
    for (const r of rows) s += r.count;
    return s;
}

export function sumCostAmount(rows: ReadonlyArray<AnalyticsRow>): number {
    let s = 0;
    for (const r of rows) s += r.costAmount;
    return s;
}

export function sumProfit(rows: ReadonlyArray<AnalyticsRow>): number {
    let s = 0;
    for (const r of rows) s += r.profit;
    return s;
}

/** 毛利率（%）。amount=0 回傳 null（不適用）。 */
export function marginPct(profit: number, amount: number): number | null {
    if (amount <= 0) return null;
    return (profit / amount) * 100;
}

