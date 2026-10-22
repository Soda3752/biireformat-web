/**
 * 「支出」分頁 localStorage 儲存層。
 *
 * Schema：
 *  - `bii.expense.months.v1`：Record<sortKey, MonthExpense>，同月份重新上傳即覆蓋
 *  - `bii.expense.household.v1`：string[]，被勾為「家用」的「其他」項目名稱（同名項目跨月自動套用）
 *  - `bii.expense.revenue.v1`：Record<fileId, RevenueFile>，帳單檔解析後的每月商品銷售快照。
 *    只存數量與金額，成本在顯示時才依當下設定的商品成本計算，改了成本不必重傳帳單。
 */

import type {MonthExpense} from './expense-model';

const MONTHS_KEY = 'bii.expense.months.v1';
const HOUSEHOLD_KEY = 'bii.expense.household.v1';
const REVENUE_KEY = 'bii.expense.revenue.v1';

export interface ProductSale {
    name: string;
    count: number;
    amount: number;
}

export interface RevenueFile {
    /** `${fileName}__${size}` */
    id: string;
    fileName: string;
    /** sortKey → 該月各商品銷售彙總 */
    months: Record<string, ProductSale[]>;
    savedAt: number;
}

function readJson<T>(key: string, fallback: T): T {
    try {
        const raw = localStorage.getItem(key);
        return raw ? (JSON.parse(raw) as T) : fallback;
    } catch {
        return fallback;
    }
}

function writeJson(key: string, value: unknown): void {
    localStorage.setItem(key, JSON.stringify(value));
}

/* ---------- 月支出 ---------- */

export function loadExpenseMonths(): Record<string, MonthExpense> {
    const data = readJson<Record<string, MonthExpense>>(MONTHS_KEY, {});
    return data && typeof data === 'object' ? data : {};
}

/** 寫入（覆蓋）多個月份 */
export function saveExpenseMonths(months: ReadonlyArray<MonthExpense>): void {
    const all = loadExpenseMonths();
    for (const m of months) all[m.sortKey] = m;
    writeJson(MONTHS_KEY, all);
}

export function removeExpenseMonth(sortKey: string): void {
    const all = loadExpenseMonths();
    delete all[sortKey];
    writeJson(MONTHS_KEY, all);
}

/* ---------- 家用標記 ---------- */

export function householdKey(name: string): string {
    return name.trim();
}

export function loadHouseholdNames(): Set<string> {
    const arr = readJson<unknown>(HOUSEHOLD_KEY, []);
    return new Set(Array.isArray(arr) ? arr.filter((x): x is string => typeof x === 'string') : []);
}

export function setHousehold(name: string, household: boolean): void {
    const set = loadHouseholdNames();
    const key = householdKey(name);
    if (household) set.add(key);
    else set.delete(key);
    try {
        writeJson(HOUSEHOLD_KEY, [...set]);
    } catch (err) {
        console.warn('[expense-store] persist household failed', err);
    }
}

/* ---------- 帳單營收快照 ---------- */

export function loadRevenueFiles(): RevenueFile[] {
    const data = readJson<Record<string, RevenueFile>>(REVENUE_KEY, {});
    return data && typeof data === 'object' ? Object.values(data) : [];
}

export function saveRevenueFiles(files: ReadonlyArray<RevenueFile>): void {
    const all = readJson<Record<string, RevenueFile>>(REVENUE_KEY, {});
    for (const f of files) all[f.id] = f;
    writeJson(REVENUE_KEY, all);
}

export function removeRevenueFile(id: string): void {
    const all = readJson<Record<string, RevenueFile>>(REVENUE_KEY, {});
    delete all[id];
    writeJson(REVENUE_KEY, all);
}
