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

/** 三大金額型指標的數值組（amount / count / profit）。供平均、中位數等表示一致。 */
export interface MetricTriple {
    /** 營收 */
    amount: number;
    /** 銷售數量 */
    count: number;
    /** 毛利 */
    profit: number;
}

/** 沿用既有命名：每日平均沿用同一介面 */
export type DailyMetrics = MetricTriple;

/** 單一客戶在某月的彙總（用於跨客戶中位數計算） */
interface CustomerAgg {
    amount: number;
    count: number;
    profit: number;
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
    /** 該月份的日曆天數（依民國年 + 月份計算，2 月含閏年判斷） */
    daysInMonth: number;
    /**
     * 該月「跨客戶中位數」：以每位客戶的當月彙總為樣本，
     * 對 amount / count / profit 各取中位數。代表「典型客戶的月表現」。
     * （注：本月有交易的客戶才納入樣本，本月零交易的客戶不會列入分母）
     */
    customerMedian: MetricTriple;
    /** 該月所有列皆未填成本 */
    allCostUnset: boolean;
}

/** 民國年 + 月份 → 該月日曆天數（ROC year + 1911 = 西元年）。 */
function daysInRocMonth(rocYear: string, month: number): number {
    const adYear = Number(rocYear) + 1911;
    if (!Number.isFinite(adYear) || month < 1 || month > 12) return 30;
    return new Date(adYear, month, 0).getDate();
}

/** 中位數（空陣列回 0）。偶數筆取中間兩值平均。 */
function median(values: ReadonlyArray<number>): number {
    if (values.length === 0) return 0;
    const sorted = [...values].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 === 0
        ? (sorted[mid - 1] + sorted[mid]) / 2
        : sorted[mid];
}

/** 對該月所有客戶的當月彙總取中位數（跨客戶代表典型客戶月表現） */
function customerMedianOf(
    customerStats: Map<string, CustomerAgg>
): MetricTriple {
    const stats = [...customerStats.values()];
    return {
        amount: median(stats.map((s) => s.amount)),
        count: median(stats.map((s) => s.count)),
        profit: median(stats.map((s) => s.profit)),
    };
}

/** 把 rows 依「年-月」分組，回傳排序好的月度總計（時間遞增）。 */
export function monthlyTotals(rows: ReadonlyArray<AnalyticsRow>): MonthlyTotals[] {
    interface Acc {
        key: MonthKey;
        amount: number;
        count: number;
        costAmount: number;
        profit: number;
        /** 該月內 customerCode → 該客戶當月彙總，用於跨客戶中位數計算 */
        customerStats: Map<string, CustomerAgg>;
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
                customerStats: new Map(),
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
        acc.products.add(r.productName);

        let cust = acc.customerStats.get(r.customerCode);
        if (!cust) {
            cust = {amount: 0, count: 0, profit: 0};
            acc.customerStats.set(r.customerCode, cust);
        }
        cust.amount += r.amount;
        cust.count += r.count;
        cust.profit += r.profit;

        acc.rowCount += 1;
        if (!r.isCostUnset) acc.anyCostSet = true;
    }

    const arr: MonthlyTotals[] = [...map.values()].map((a) => ({
        key: a.key,
        amount: a.amount,
        count: a.count,
        costAmount: a.costAmount,
        profit: a.profit,
        customerCount: a.customerStats.size,
        productCount: a.products.size,
        rowCount: a.rowCount,
        daysInMonth: daysInRocMonth(a.key.year, a.key.month),
        customerMedian: customerMedianOf(a.customerStats),
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

/* ============== 期段（每 N 日為一段）細粒度聚合，用於趨勢圖 ============== */

/** 趨勢圖期段天數預設值（每段 N 日；可由 UI 動態調整為 1~31）。 */
export const DEFAULT_TREND_DAYS_PER_PERIOD = 5;
export const MIN_TREND_DAYS_PER_PERIOD = 1;
export const MAX_TREND_DAYS_PER_PERIOD = 31;

/** 把使用者輸入夾擠到合法範圍並轉為整數；NaN/越界皆退回預設值。 */
export function clampTrendDaysPerPeriod(value: number): number {
    if (!Number.isFinite(value)) return DEFAULT_TREND_DAYS_PER_PERIOD;
    const v = Math.floor(value);
    if (v < MIN_TREND_DAYS_PER_PERIOD) return MIN_TREND_DAYS_PER_PERIOD;
    if (v > MAX_TREND_DAYS_PER_PERIOD) return MAX_TREND_DAYS_PER_PERIOD;
    return v;
}

/** 民國年（字串）→ 西曆年（number）。 */
function rocToAdYear(rocYear: string): number {
    return Number(rocYear) + 1911;
}

/** 把 (民國年, 月, 日) 轉為自 Unix epoch 起算的「絕對天數」，用於跨月連續切段。 */
function rocDateToAbsoluteDay(rocYear: string, month: number, day: number): number {
    return Math.floor(Date.UTC(rocToAdYear(rocYear), month - 1, day) / 86400000);
}

/** 絕對天數 → (民國年字串, 月, 日)。 */
function absoluteDayToRocCalendar(absDay: number): { year: string; month: number; day: number } {
    const d = new Date(absDay * 86400000);
    return {
        year: String(d.getUTCFullYear() - 1911),
        month: d.getUTCMonth() + 1,
        day: d.getUTCDate(),
    };
}

/** 段標籤（x 軸用，緊湊）。同月：「3月1-5日」；跨月：「3/31-4/4」；跨年：「114/12/30-115/1/3」。 */
function buildPeriodLabel(
    start: { year: string; month: number; day: number },
    end: { year: string; month: number; day: number }
): string {
    const sameYear = start.year === end.year;
    const sameMonth = sameYear && start.month === end.month;
    if (sameMonth && start.day === end.day) return `${start.month}月${start.day}日`;
    if (sameMonth) return `${start.month}月${start.day}-${end.day}日`;
    if (sameYear) return `${start.month}/${start.day}-${end.month}/${end.day}`;
    return `${start.year}/${start.month}/${start.day}-${end.year}/${end.month}/${end.day}`;
}

/** 段完整標籤（tooltip 用）。 */
function buildPeriodFullLabel(
    start: { year: string; month: number; day: number },
    end: { year: string; month: number; day: number }
): string {
    const sameYear = start.year === end.year;
    const sameMonth = sameYear && start.month === end.month;
    if (sameMonth && start.day === end.day) return `${start.year}年${start.month}月${start.day}日`;
    if (sameMonth) return `${start.year}年${start.month}月${start.day}-${end.day}日`;
    if (sameYear) return `${start.year}年${start.month}月${start.day}日 ~ ${end.month}月${end.day}日`;
    return `${start.year}年${start.month}月${start.day}日 ~ ${end.year}年${end.month}月${end.day}日`;
}

export interface PeriodKey {
    /** 段索引（0 起算，相對於資料的最早月份 1 號） */
    periodIndex: number;
    /** 段起始日（民國年） */
    startYear: string;
    startMonth: number;
    startDay: number;
    /** 段結束日（民國年）。最後一段保持 daysPerPeriod 寬度，可能落在資料末日之後（僅標籤用）。 */
    endYear: string;
    endMonth: number;
    endDay: number;
    /** x 軸用緊湊標籤，例：「3月1-5日」、「3/31-4/4」 */
    label: string;
    /** 含年的完整標籤，例：「114年3月1-5日」、「114年3月31日 ~ 4月4日」 */
    fullLabel: string;
    /** 排序鍵，使用起始絕對天數補零至 8 位 */
    sortKey: string;
}

export interface PeriodTotals {
    key: PeriodKey;
    amount: number;
    count: number;
    costAmount: number;
    profit: number;
    customerCount: number;
    productCount: number;
    rowCount: number;
    /** 該段所有列皆未填成本 */
    allCostUnset: boolean;
}

function makePeriodKey(
    baseAbsDay: number,
    periodIndex: number,
    daysPerPeriod: number
): PeriodKey {
    const startAbsDay = baseAbsDay + periodIndex * daysPerPeriod;
    const endAbsDay = startAbsDay + daysPerPeriod - 1;
    const start = absoluteDayToRocCalendar(startAbsDay);
    const end = absoluteDayToRocCalendar(endAbsDay);
    return {
        periodIndex,
        startYear: start.year,
        startMonth: start.month,
        startDay: start.day,
        endYear: end.year,
        endMonth: end.month,
        endDay: end.day,
        label: buildPeriodLabel(start, end),
        fullLabel: buildPeriodFullLabel(start, end),
        sortKey: String(startAbsDay).padStart(8, '0'),
    };
}

/**
 * 把 rows 依連續 N 日段分組（不在月底切割：如 daysPerPeriod=5 時 3/31~4/4 為同一段）。
 * 段錨點為「資料中最早一筆的所屬月 1 號」，由此向後每 N 日為一段。
 * @param daysPerPeriod 每段天數（1~31，會自動 clamp）。
 */
export function periodTotals(
    rows: ReadonlyArray<AnalyticsRow>,
    daysPerPeriod: number = DEFAULT_TREND_DAYS_PER_PERIOD
): PeriodTotals[] {
    const dpp = clampTrendDaysPerPeriod(daysPerPeriod);
    if (rows.length === 0) return [];

    // 找出最早資料月份，以該月 1 號為段錨點
    let earliest = rows[0];
    let earliestAbsDay = rocDateToAbsoluteDay(earliest.year, earliest.month, earliest.day);
    for (const r of rows) {
        const ad = rocDateToAbsoluteDay(r.year, r.month, r.day);
        if (ad < earliestAbsDay) {
            earliest = r;
            earliestAbsDay = ad;
        }
    }
    const baseAbsDay = rocDateToAbsoluteDay(earliest.year, earliest.month, 1);

    interface Acc {
        key: PeriodKey;
        amount: number;
        count: number;
        costAmount: number;
        profit: number;
        customers: Set<string>;
        products: Set<string>;
        rowCount: number;
        anyCostSet: boolean;
    }

    const map = new Map<number, Acc>();
    for (const r of rows) {
        const ad = rocDateToAbsoluteDay(r.year, r.month, r.day);
        const idx = Math.floor((ad - baseAbsDay) / dpp);
        let acc = map.get(idx);
        if (!acc) {
            acc = {
                key: makePeriodKey(baseAbsDay, idx, dpp),
                amount: 0,
                count: 0,
                costAmount: 0,
                profit: 0,
                customers: new Set(),
                products: new Set(),
                rowCount: 0,
                anyCostSet: false,
            };
            map.set(idx, acc);
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

    const arr: PeriodTotals[] = [...map.values()].map((a) => ({
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

/* ============== MoM 環比 ============== */

export interface DeltaValue {
    /** 絕對差值 */
    delta: number;
    /** 百分比變化（基準=上月）；上月為 0 時回 null */
    pct: number | null;
}

/** 每日平均（amount / count / profit ÷ daysInMonth） */
export type DailyAverages = DailyMetrics;

export interface MoMComparison {
    current: MonthlyTotals;
    previous: MonthlyTotals | null;
    /** 本月每日平均（依 daysInMonth） */
    currentDaily: DailyAverages;
    /** 上月每日平均（無上月時為 null） */
    previousDaily: DailyAverages | null;
    /** 本月跨客戶中位數（每位客戶月彙總取中位） */
    currentMedian: MetricTriple;
    /** 上月跨客戶中位數（無上月時為 null） */
    previousMedian: MetricTriple | null;
    diffs: {
        amount: DeltaValue;
        count: DeltaValue;
        profit: DeltaValue;
        customerCount: DeltaValue;
        productCount: DeltaValue;
        marginPct: DeltaValue;
        dailyAmount: DeltaValue;
        dailyCount: DeltaValue;
        dailyProfit: DeltaValue;
        medianAmount: DeltaValue;
        medianCount: DeltaValue;
        medianProfit: DeltaValue;
    };
}

function dailyOf(t: MonthlyTotals): DailyAverages {
    const d = t.daysInMonth > 0 ? t.daysInMonth : 1;
    return {
        amount: t.amount / d,
        count: t.count / d,
        profit: t.profit / d,
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
    const curDaily = dailyOf(current);
    const prevDaily = prev ? dailyOf(prev) : null;
    const curMedian = current.customerMedian;
    const prevMedian = prev?.customerMedian ?? null;
    return {
        current,
        previous: prev,
        currentDaily: curDaily,
        previousDaily: prevDaily,
        currentMedian: curMedian,
        previousMedian: prevMedian,
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
            dailyAmount: deltaOf(curDaily.amount, prevDaily?.amount ?? null),
            dailyCount: deltaOf(curDaily.count, prevDaily?.count ?? null),
            dailyProfit: deltaOf(curDaily.profit, prevDaily?.profit ?? null),
            medianAmount: deltaOf(curMedian.amount, prevMedian?.amount ?? null),
            medianCount: deltaOf(curMedian.count, prevMedian?.count ?? null),
            medianProfit: deltaOf(curMedian.profit, prevMedian?.profit ?? null),
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
 *
 * 「活躍」定義：該月有訂單且營收 > 0；amount ≤ 0（無訂單或退貨抵銷後為零/負）視為當月不活躍。
 *   - 上月活躍、本月不活躍 → 流失
 *   - 上月不活躍、本月活躍 → 新增
 *   - 兩月皆活躍 → 留存
 *   - 兩月皆不活躍 → 不納入任何分群
 */
export function customerSegmentation(
    currentRows: ReadonlyArray<AnalyticsRow>,
    previousRows: ReadonlyArray<AnalyticsRow>
): CustomerSegmentation {
    const cur = aggregateCustomersOfMonth(currentRows);
    const prev = aggregateCustomersOfMonth(previousRows);

    const isActive = (s: CustomerMonthStat | undefined): s is CustomerMonthStat =>
        s !== undefined && s.amount > 0;

    const newCustomers: CustomerMonthStat[] = [];
    const retainedCustomers: RetainedCustomer[] = [];
    const churnedCustomers: CustomerMonthStat[] = [];

    const codes = new Set<string>([...cur.keys(), ...prev.keys()]);
    for (const code of codes) {
        const c = cur.get(code);
        const p = prev.get(code);
        const cActive = isActive(c);
        const pActive = isActive(p);

        if (cActive && !pActive) {
            newCustomers.push(c);
        } else if (!cActive && pActive) {
            churnedCustomers.push(p);
        } else if (cActive && pActive) {
            const amountDelta = c.amount - p.amount;
            const countDelta = c.count - p.count;
            retainedCustomers.push({
                ...c,
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
            };
            map.set(r.productName, s);
        }
        s.count += r.count;
        s.amount += r.amount;
        s.costAmount += r.costAmount;
        s.profit += r.profit;
        s.customers.add(r.customerCode);
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
