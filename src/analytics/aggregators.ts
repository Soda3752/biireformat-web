/**
 * 共用聚合工具：所有 chart adapter 從這裡拿聚合結果，
 * 避免每個 adapter 自己寫一遍 group-by。
 */

import type {AnalyticsRow} from './dataset-builder';

export interface GroupSum {
    key: string;
    amount: number;
    count: number;
}

/** 依任意欄位分組加總 amount + count，回傳排序好的陣列。 */
export function groupBy<K extends keyof AnalyticsRow>(
    rows: ReadonlyArray<AnalyticsRow>,
    key: K,
    sortBy: 'amount' | 'count' | 'key' = 'amount',
    desc = true
): GroupSum[] {
    const map = new Map<string, GroupSum>();
    for (const r of rows) {
        const k = String(r[key]);
        let g = map.get(k);
        if (!g) {
            g = {key: k, amount: 0, count: 0};
            map.set(k, g);
        }
        g.amount += r.amount;
        g.count += r.count;
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
    sortBy: 'amount' | 'count' = 'amount',
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
}

export function dailySeries(rows: ReadonlyArray<AnalyticsRow>): DailyPoint[] {
    const map = new Map<number, DailyPoint>();
    for (const r of rows) {
        let p = map.get(r.day);
        if (!p) {
            p = {day: r.day, amount: 0, count: 0};
            map.set(r.day, p);
        }
        p.amount += r.amount;
        p.count += r.count;
    }
    return [...map.values()].sort((a, b) => a.day - b.day);
}

/** 星期金額/數量序列（0=日 ~ 6=六） */
export function weekdaySeries(rows: ReadonlyArray<AnalyticsRow>): DailyPoint[] {
    const map = new Map<number, DailyPoint>();
    for (let i = 0; i < 7; i++) map.set(i, {day: i, amount: 0, count: 0});
    for (const r of rows) {
        const p = map.get(r.weekday)!;
        p.amount += r.amount;
        p.count += r.count;
    }
    return [...map.values()];
}

/** 結帳模式重疊計數：回傳每個模式的客戶數（distinct customerCode） */
export function paymentModeCounts(
    rows: ReadonlyArray<AnalyticsRow>
): { monthly: number; needTex: number; cash: number; total: number } {
    const monthly = new Set<string>();
    const needTex = new Set<string>();
    const cash = new Set<string>();
    const all = new Set<string>();
    for (const r of rows) {
        all.add(r.customerCode);
        if (r.isMonthly) monthly.add(r.customerCode);
        if (r.isNeedTex) needTex.add(r.customerCode);
        if (r.isCashUser) cash.add(r.customerCode);
    }
    return {monthly: monthly.size, needTex: needTex.size, cash: cash.size, total: all.size};
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

