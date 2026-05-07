/**
 * 跨月分析聚合層：把 AnalyticsRow 依「民國年+月」拆組，
 * 提供月度總計、MoM 比較、客戶分群、商品漲價偵測等資料模型。
 *
 * 月份 key 規約：使用 row.year + row.month（不依賴檔案邊界），
 * 排序鍵 sortKey = `${year}-${MM}`（補零，字串比較 = 時間順序）。
 */

import type {AnalyticsRow} from '@/analytics/dataset-builder';

export interface MonthKey {
    /** 民國年 */
    year: string;
    /** 1~12 */
    month: number;
    /** 民國 + 西元 對照標籤，如 "114年3月"（西元 2025） */
    label: string;
    /** 排序鍵，如 "114-03" */
    sortKey: string;
}

export function makeMonthKey(year: string, month: number): MonthKey {
    const mm = String(month).padStart(2, '0');
    return {
        year,
        month,
        label: `${year}年${month}月`,
        sortKey: `${year}-${mm}`,
    };
}

export function rowMonthSortKey(row: AnalyticsRow): string {
    return `${row.year}-${String(row.month).padStart(2, '0')}`;
}

export interface MonthlyTotals {
    key: MonthKey;
    amount: number;
    count: number;
    costAmount: number;
    profit: number;
    customerCount: number;
    productCount: number;
    rowCount: number;
    /** 該月所有列皆未填成本 */
    allCostUnset: boolean;
}

/** 把 rows 依「年-月」分組，回傳排序好的月度總計（時間遞增）。 */
export function monthlyTotals(rows: ReadonlyArray<AnalyticsRow>): MonthlyTotals[] {
    interface Acc {
        key: MonthKey;
        amount: number;
        count: number;
        costAmount: number;
        profit: number;
        customers: Set<string>;
        products: Set<string>;
        rowCount: number;
        anyCostSet: boolean;
    }

    const map = new Map<string, Acc>();
    for (const r of rows) {
        const key = makeMonthKey(r.year, r.month);
        let acc = map.get(key.sortKey);
        if (!acc) {
            acc = {
                key,
                amount: 0,
                count: 0,
                costAmount: 0,
                profit: 0,
                customers: new Set(),
                products: new Set(),
                rowCount: 0,
                anyCostSet: false,
            };
            map.set(key.sortKey, acc);
        }
        acc.amount += r.amount;
        acc.count += r.count;
        acc.costAmount += r.costAmount;
        acc.profit += r.profit;
        acc.customers.add(r.customerCode);
        acc.products.add(r.productName);
        acc.rowCount += 1;
        if (!r.isCostUnset) acc.anyCostSet = true;
    }

    const arr: MonthlyTotals[] = [...map.values()].map((a) => ({
        key: a.key,
        amount: a.amount,
        count: a.count,
        costAmount: a.costAmount,
        profit: a.profit,
        customerCount: a.customers.size,
        productCount: a.products.size,
        rowCount: a.rowCount,
        allCostUnset: !a.anyCostSet,
    }));

    arr.sort((a, b) => a.key.sortKey.localeCompare(b.key.sortKey));
    return arr;
}

/** 把 rows 依月份切片，給定 sortKey 取出該月的列。 */
export function rowsOfMonth(
    rows: ReadonlyArray<AnalyticsRow>,
    sortKey: string
): AnalyticsRow[] {
    return rows.filter((r) => rowMonthSortKey(r) === sortKey);
}

/* ============== MoM 環比 ============== */

export interface DeltaValue {
    /** 絕對差值 */
    delta: number;
    /** 百分比變化（基準=上月）；上月為 0 時回 null */
    pct: number | null;
}

export interface MoMComparison {
    current: MonthlyTotals;
    previous: MonthlyTotals | null;
    diffs: {
        amount: DeltaValue;
        count: DeltaValue;
        profit: DeltaValue;
        customerCount: DeltaValue;
        productCount: DeltaValue;
        marginPct: DeltaValue;
    };
}

function deltaOf(curr: number, prev: number | null): DeltaValue {
    if (prev === null) return {delta: curr, pct: null};
    const delta = curr - prev;
    const pct = prev !== 0 ? (delta / Math.abs(prev)) * 100 : null;
    return {delta, pct};
}

function marginOf(t: MonthlyTotals): number {
    return t.amount > 0 ? (t.profit / t.amount) * 100 : 0;
}

/** 給定本月與上月（任一可為 null 代表不存在），回傳 MoM 結構；本月必須存在。 */
export function computeMoM(
    current: MonthlyTotals,
    previous: MonthlyTotals | null
): MoMComparison {
    const prev = previous;
    return {
        current,
        previous: prev,
        diffs: {
            amount: deltaOf(current.amount, prev?.amount ?? null),
            count: deltaOf(current.count, prev?.count ?? null),
            profit: deltaOf(current.profit, prev?.profit ?? null),
            customerCount: deltaOf(current.customerCount, prev?.customerCount ?? null),
            productCount: deltaOf(current.productCount, prev?.productCount ?? null),
            marginPct: prev ? deltaOf(marginOf(current), marginOf(prev)) : {
                delta: marginOf(current),
                pct: null,
            },
        },
    };
}

/* ============== B1 客戶分群（new / churn / retain） ============== */

export interface CustomerMonthStat {
    customerCode: string;
    customerName: string;
    line: string;
    amount: number;
    count: number;
    profit: number;
}

function aggregateCustomersOfMonth(rows: ReadonlyArray<AnalyticsRow>): Map<string, CustomerMonthStat> {
    const map = new Map<string, CustomerMonthStat>();
    for (const r of rows) {
        let s = map.get(r.customerCode);
        if (!s) {
            s = {
                customerCode: r.customerCode,
                customerName: r.customerName,
                line: r.line,
                amount: 0,
                count: 0,
                profit: 0,
            };
            map.set(r.customerCode, s);
        }
        s.amount += r.amount;
        s.count += r.count;
        s.profit += r.profit;
    }
    return map;
}

export interface RetainedCustomer extends CustomerMonthStat {
    prevAmount: number;
    prevCount: number;
    prevProfit: number;
    amountDelta: number;
    amountDeltaPct: number | null;
    countDelta: number;
    countDeltaPct: number | null;
}

export interface CustomerSegmentation {
    /** 本月新出現（上月沒下單） */
    newCustomers: CustomerMonthStat[];
    /** 上月有但本月沒下單 */
    churnedCustomers: CustomerMonthStat[];
    /** 兩月都有，含 MoM 變化 */
    retainedCustomers: RetainedCustomer[];
}

/**
 * @param currentRows  本月完整 rows
 * @param previousRows 上月完整 rows
 */
export function customerSegmentation(
    currentRows: ReadonlyArray<AnalyticsRow>,
    previousRows: ReadonlyArray<AnalyticsRow>
): CustomerSegmentation {
    const cur = aggregateCustomersOfMonth(currentRows);
    const prev = aggregateCustomersOfMonth(previousRows);

    const newCustomers: CustomerMonthStat[] = [];
    const retainedCustomers: RetainedCustomer[] = [];
    for (const [code, s] of cur) {
        const p = prev.get(code);
        if (!p) {
            newCustomers.push(s);
        } else {
            const amountDelta = s.amount - p.amount;
            const countDelta = s.count - p.count;
            retainedCustomers.push({
                ...s,
                prevAmount: p.amount,
                prevCount: p.count,
                prevProfit: p.profit,
                amountDelta,
                amountDeltaPct: p.amount !== 0 ? (amountDelta / Math.abs(p.amount)) * 100 : null,
                countDelta,
                countDeltaPct: p.count !== 0 ? (countDelta / Math.abs(p.count)) * 100 : null,
            });
        }
    }

    const churnedCustomers: CustomerMonthStat[] = [];
    for (const [code, p] of prev) {
        if (!cur.has(code)) churnedCustomers.push(p);
    }

    newCustomers.sort((a, b) => b.amount - a.amount);
    churnedCustomers.sort((a, b) => b.amount - a.amount);
    retainedCustomers.sort((a, b) => b.amountDelta - a.amountDelta);

    return {newCustomers, churnedCustomers, retainedCustomers};
}

/* ============== C1 商品漲價影響 ============== */

interface ProductMonthAgg {
    productName: string;
    count: number;
    amount: number;
    costAmount: number;
    profit: number;
    customers: Set<string>;
    /** 該月份內出現過的 distinct 單價（一個商品在不同客戶可能不同價）。 */
    distinctPrices: Set<number>;
}

function aggregateProductsOfMonth(rows: ReadonlyArray<AnalyticsRow>): Map<string, ProductMonthAgg> {
    const map = new Map<string, ProductMonthAgg>();
    for (const r of rows) {
        let s = map.get(r.productName);
        if (!s) {
            s = {
                productName: r.productName,
                count: 0,
                amount: 0,
                costAmount: 0,
                profit: 0,
                customers: new Set(),
                distinctPrices: new Set(),
            };
            map.set(r.productName, s);
        }
        s.count += r.count;
        s.amount += r.amount;
        s.costAmount += r.costAmount;
        s.profit += r.profit;
        s.customers.add(r.customerCode);
        s.distinctPrices.add(r.price);
    }
    return map;
}

export interface ProductPriceChange {
    productName: string;
    /** 加權平均單價 = amount / count */
    prevAvgPrice: number;
    currentAvgPrice: number;
    priceChange: number;
    priceChangePct: number | null;
    /** 各客戶使用的單價數（>1 表示同月份內就有混價） */
    prevDistinctPrices: number;
    currentDistinctPrices: number;

    prevCount: number;
    currentCount: number;
    countChange: number;
    countChangePct: number | null;

    prevAmount: number;
    currentAmount: number;
    amountChange: number;
    amountChangePct: number | null;

    prevProfit: number;
    currentProfit: number;
    profitChange: number;
    profitChangePct: number | null;

    prevCustomerCount: number;
    currentCustomerCount: number;
}

function avgPrice(amount: number, count: number): number {
    return count > 0 ? amount / count : 0;
}

function pctOf(curr: number, prev: number): number | null {
    return prev !== 0 ? ((curr - prev) / Math.abs(prev)) * 100 : null;
}

/**
 * 找出兩月都出現、且加權均價變動超過 minPriceChangePct 的商品。
 * 若 minPriceChangePct = 0，則回傳所有兩月都有的商品。
 */
export function detectProductPriceChanges(
    currentRows: ReadonlyArray<AnalyticsRow>,
    previousRows: ReadonlyArray<AnalyticsRow>,
    minPriceChangePct = 0.01
): ProductPriceChange[] {
    const cur = aggregateProductsOfMonth(currentRows);
    const prev = aggregateProductsOfMonth(previousRows);

    const out: ProductPriceChange[] = [];
    for (const [name, c] of cur) {
        const p = prev.get(name);
        if (!p) continue;

        const prevAvg = avgPrice(p.amount, p.count);
        const currAvg = avgPrice(c.amount, c.count);
        const priceChange = currAvg - prevAvg;
        const priceChangePct = pctOf(currAvg, prevAvg);

        if (priceChangePct === null) {
            // 上月均價為 0，跳過
            continue;
        }
        if (Math.abs(priceChangePct) < minPriceChangePct) continue;

        out.push({
            productName: name,
            prevAvgPrice: prevAvg,
            currentAvgPrice: currAvg,
            priceChange,
            priceChangePct,
            prevDistinctPrices: p.distinctPrices.size,
            currentDistinctPrices: c.distinctPrices.size,

            prevCount: p.count,
            currentCount: c.count,
            countChange: c.count - p.count,
            countChangePct: pctOf(c.count, p.count),

            prevAmount: p.amount,
            currentAmount: c.amount,
            amountChange: c.amount - p.amount,
            amountChangePct: pctOf(c.amount, p.amount),

            prevProfit: p.profit,
            currentProfit: c.profit,
            profitChange: c.profit - p.profit,
            profitChangePct: pctOf(c.profit, p.profit),

            prevCustomerCount: p.customers.size,
            currentCustomerCount: c.customers.size,
        });
    }

    // 預設：依價格變動絕對值由大到小（最劇烈的在前）
    out.sort((a, b) => Math.abs(b.priceChangePct ?? 0) - Math.abs(a.priceChangePct ?? 0));
    return out;
}
