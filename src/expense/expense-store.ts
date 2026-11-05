/**
 * 「支出」分頁 localStorage 儲存層。
 *
 * Schema：
 *  - `bii.expense.months.v1`：Record<sortKey, MonthExpense>，同月份重新上傳即覆蓋
 *  - `bii.expense.shop.v1`：string[]，被取消「家用」勾選（=店務）的「其他」項目名稱（同名項目跨月自動套用）。
 *    「其他」項目預設都算家用；不在此清單的一律勾選為家用，新月份新項目也一樣。
 *    （舊版 `bii.expense.household.v1` 記的是「勾為家用」的名稱，平時不再讀取，只供下方一次性轉換使用，保留不刪。）
 *  - `bii.expense.shop.migrated.v1`：一次性轉換完成旗標（值為 '1'）。
 *  - `bii.expense.revenue.v1`：Record<fileId, RevenueFile>，帳單檔解析後的每月商品銷售快照。
 *    只存數量與金額，成本在顯示時才依當下設定的商品成本計算，改了成本不必重傳帳單。
 */

import type {MonthExpense} from './expense-model';

const MONTHS_KEY = 'bii.expense.months.v1';
const SHOP_KEY = 'bii.expense.shop.v1';
const LEGACY_HOUSEHOLD_KEY = 'bii.expense.household.v1';
const SHOP_MIGRATED_KEY = 'bii.expense.shop.migrated.v1';
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

/* ---------- 家用標記（預設家用，記住被取消勾選＝店務的名稱） ---------- */

export function householdKey(name: string): string {
    return name.trim();
}

function readStringSet(key: string): Set<string> {
    const arr = readJson<unknown>(key, []);
    return new Set(Array.isArray(arr) ? arr.filter((x): x is string => typeof x === 'string') : []);
}

/**
 * 一次性轉換：舊版（v2.12.38 以前）預設不勾、只記「勾為家用」的名稱；
 * v2.12.45 起改成預設全勾、只記「取消勾選＝店務」的名稱，導致舊版沒勾的項目在新版全變成勾選。
 *
 * 規則：
 *  - 瀏覽器有舊 key `bii.expense.household.v1`，且還沒寫過旗標 `bii.expense.shop.migrated.v1` 才執行。
 *  - 取出目前已存在瀏覽器裡所有月份「其他」類的項目名稱（householdKey 正規化），
 *    不在舊家用清單裡的 → 加進店務清單；與 v2.12.45 後使用者已取消勾選的店務清單取聯集，不覆蓋。
 *  - 做完寫旗標，之後不再執行；舊 key 保留不刪。
 *  - 沒有舊 key 的新使用者不做任何事（維持預設全勾）。
 * 由 loadShopNames() 自動觸發，所以畫面計算店務／家用之前一定已轉換完成。
 */
export function migrateLegacyHousehold(): void {
    try {
        if (localStorage.getItem(SHOP_MIGRATED_KEY)) return;
        if (localStorage.getItem(LEGACY_HOUSEHOLD_KEY) === null) return;
    } catch {
        return;
    }
    const legacy = new Set([...readStringSet(LEGACY_HOUSEHOLD_KEY)].map(householdKey));
    const shop = readStringSet(SHOP_KEY);
    for (const m of Object.values(loadExpenseMonths())) {
        for (const it of m?.items ?? []) {
            if (it.category !== '其他') continue;
            const key = householdKey(it.name);
            if (key && !legacy.has(key)) shop.add(key);
        }
    }
    try {
        writeJson(SHOP_KEY, [...shop]);
        localStorage.setItem(SHOP_MIGRATED_KEY, '1');
    } catch (err) {
        console.warn('[expense-store] migrate legacy household names failed', err);
    }
}

/** 被取消家用勾選（=店務）的「其他」項目名稱（第一次讀取時會先做舊版資料轉換） */
export function loadShopNames(): Set<string> {
    migrateLegacyHousehold();
    return readStringSet(SHOP_KEY);
}

/** 「其他」項目是否算家用：不在店務清單裡就是家用 */
export function isHouseholdName(name: string, shopNames: ReadonlySet<string>): boolean {
    return !shopNames.has(householdKey(name));
}

export function setHousehold(name: string, household: boolean): void {
    const set = loadShopNames();
    const key = householdKey(name);
    if (household) set.delete(key);
    else set.add(key);
    try {
        writeJson(SHOP_KEY, [...set]);
    } catch (err) {
        console.warn('[expense-store] persist shop names failed', err);
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
