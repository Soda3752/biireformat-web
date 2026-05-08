/**
 * B1 客戶分群表：新增 / 零售新增 / 流失 / 零售流失 / 留存 / 客戶流動 六個分頁切換，
 * 共用同一張表骨架，欄位略有差異：
 *   - 新增 / 流失 / 零售新增 / 零售流失：客戶 / 線別 / 金額 / 數量 / 毛利 / 動作
 *   - 留存：客戶 / 線別 / 上月金額 / 本月金額 / Δ金額 / Δ% / Δ數量 / Δ%
 *   - 客戶流動：以卡片呈現「新增 vs 流失 = 淨差額」（並列全部 + 零售兩組）
 *
 * 「零售」標記為全域 Set（localStorage 持久化），詳見 retail-store.ts。
 */

import type {
    CustomerMonthStat,
    CustomerSegmentation,
    MetricTriple,
    RetainedCustomer,
} from './month-aggregators';
import {isRetail, setRetail, subscribeRetail} from './retail-store';

type SegmentKind = 'new' | 'retail-new' | 'churned' | 'retail-churned' | 'retained' | 'flow';

type SortDir = 'asc' | 'desc';

/** 簡單分頁（新增 / 零售新增 / 流失 / 零售流失）的可排序欄位 */
type SimpleSortKey = 'name' | 'line' | 'amount' | 'count' | 'profit';

/** 留存分頁的可排序欄位 */
type RetainedSortKey =
    | 'name'
    | 'line'
    | 'prevAmount'
    | 'amount'
    | 'amountDelta'
    | 'amountDeltaPct'
    | 'countDelta'
    | 'countDeltaPct';

/** 字串/數字共用 sort（字串用 zh-TW localeCompare） */
function sortRows<T>(
    rows: ReadonlyArray<T>,
    valueOf: (r: T) => number | string,
    dir: SortDir
): T[] {
    const sign = dir === 'asc' ? 1 : -1;
    return [...rows].sort((a, b) => {
        const av = valueOf(a);
        const bv = valueOf(b);
        if (typeof av === 'string' && typeof bv === 'string') {
            return av.localeCompare(bv, 'zh-TW') * sign;
        }
        return (Number(av) - Number(bv)) * sign;
    });
}

/** 表頭點擊後的指示器 HTML（依當前排序鍵決定 ▲▼ 或預留空間） */
function sortIndicatorHtml(active: boolean, dir: SortDir): string {
    if (!active) return '<span class="sort-indicator"></span>';
    return `<span class="sort-indicator is-active">${dir === 'asc' ? '▲' : '▼'}</span>`;
}

export interface MonthTotalsRef {
    label: string;
    amount: number;
    profit: number;
    count: number;
    /** 該月跨客戶中位數（用於每位客戶的「vs 中位數」比較標記） */
    median: MetricTriple;
}

interface Controller {
    element: HTMLElement;

    setData(seg: CustomerSegmentation, current: MonthTotalsRef, previous: MonthTotalsRef): void;
}

export interface CustomerSegmentTableOptions {
    onCustomerClick?: (customerCode: string) => void;
}

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');
const fmtCount = (v: number) => v.toLocaleString('zh-TW');
const fmtPct = (v: number | null) => (v === null ? '—' : `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`);
const fmtDelta = (v: number) => `${v >= 0 ? '+' : ''}${fmtMoney(Math.round(v))}`;
const ratioPct = (numerator: number, denominator: number): string =>
    denominator > 0 ? `${((numerator / denominator) * 100).toFixed(1)}%` : '—';

/**
 * 「客戶值 vs 該月中位數」徽章。median ≤ 0 時不可比較（樣本退化），回空字串。
 * 顯示為小尺寸百分比 chip：>0 綠 / <0 紅 / =0 灰。
 */
function vsMedianBadge(value: number, median: number): string {
    if (median <= 0) return '';
    const delta = value - median;
    const pct = (delta / median) * 100;
    const tone = delta > 0 ? 'is-positive' : delta < 0 ? 'is-negative' : 'is-neutral';
    const sign = delta > 0 ? '+' : '';
    const display = Math.abs(pct) >= 1000 ? `${sign}${(pct / 100).toFixed(1)}×` : `${sign}${pct.toFixed(0)}%`;
    return `<span class="cross-month-vs-median ${tone}" title="vs 中位數 ${fmtMoney(Math.round(median))}">${display}</span>`;
}

/** 表頭旁的中位數提示徽章（解釋小百分比的含義） */
function medianHintCaption(label: string, median: MetricTriple): string {
    return `<div class="cross-month-segment-caption">
        ${escapeHtml(label)} 中位數
        <span>營收 <strong>${fmtMoney(Math.round(median.amount))}</strong></span>
        <span>數量 <strong>${fmtCount(Math.round(median.count))}</strong></span>
        <span>毛利 <strong>${fmtMoney(Math.round(median.profit))}</strong></span>
        <span class="cross-month-segment-caption-note">每筆右側徽章代表該客戶相對於中位數的偏離（+ 高於、− 低於）</span>
    </div>`;
}

export function createCustomerSegmentTable(opts: CustomerSegmentTableOptions = {}): Controller {
    const root = document.createElement('div');
    root.className = 'analytics-detail cross-month-segment-table';
    root.innerHTML = `
    <div class="cross-month-segment-tabs" data-role="tabs">
      <button type="button" class="cross-month-segment-tab is-active" data-kind="new">
        <span class="cross-month-segment-tab-label">新增客戶</span>
        <span class="cross-month-segment-tab-count" data-role="count-new">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="retail-new">
        <span class="cross-month-segment-tab-label">零售新增</span>
        <span class="cross-month-segment-tab-count" data-role="count-retail-new">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="churned">
        <span class="cross-month-segment-tab-label">流失客戶</span>
        <span class="cross-month-segment-tab-count" data-role="count-churned">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="retail-churned">
        <span class="cross-month-segment-tab-label">零售流失</span>
        <span class="cross-month-segment-tab-count" data-role="count-retail-churned">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="retained">
        <span class="cross-month-segment-tab-label">留存客戶</span>
        <span class="cross-month-segment-tab-count" data-role="count-retained">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="flow">
        <span class="cross-month-segment-tab-label">客戶流動</span>
      </button>
    </div>
    <div class="cross-month-segment-summary" data-role="summary"></div>
    <div class="cross-month-segment-bars" data-role="bars"></div>
    <div data-role="median-caption"></div>
    <div class="cross-month-flow-cards" data-role="flow-cards"></div>
    <div class="analytics-detail-table-wrap" data-role="table-wrap">
      <table class="analytics-detail-table-el cross-month-segment-tableel">
        <thead data-role="thead"></thead>
        <tbody data-role="tbody"></tbody>
      </table>
    </div>
  `;

    const thead = root.querySelector<HTMLElement>('[data-role="thead"]')!;
    const tbody = root.querySelector<HTMLElement>('[data-role="tbody"]')!;
    const tabHost = root.querySelector<HTMLElement>('[data-role="tabs"]')!;
    const summaryEl = root.querySelector<HTMLElement>('[data-role="summary"]')!;
    const barsEl = root.querySelector<HTMLElement>('[data-role="bars"]')!;
    const medianCaptionEl = root.querySelector<HTMLElement>('[data-role="median-caption"]')!;
    const flowCardsEl = root.querySelector<HTMLElement>('[data-role="flow-cards"]')!;
    const tableWrapEl = root.querySelector<HTMLElement>('[data-role="table-wrap"]')!;
    const countNew = root.querySelector<HTMLElement>('[data-role="count-new"]')!;
    const countRetailNew = root.querySelector<HTMLElement>('[data-role="count-retail-new"]')!;
    const countChurn = root.querySelector<HTMLElement>('[data-role="count-churned"]')!;
    const countRetailChurn = root.querySelector<HTMLElement>('[data-role="count-retail-churned"]')!;
    const countRetain = root.querySelector<HTMLElement>('[data-role="count-retained"]')!;

    let activeKind: SegmentKind = 'new';
    let lastSeg: CustomerSegmentation | null = null;
    let lastCurrent: MonthTotalsRef | null = null;
    let lastPrevious: MonthTotalsRef | null = null;

    // 排序狀態：簡單表 4 個 tab 共用一份；留存表獨立一份
    // 預設與 month-aggregators 既有的排序方向一致（amount/amountDelta desc）
    const simpleSort: {key: SimpleSortKey; dir: SortDir} = {key: 'amount', dir: 'desc'};
    const retainedSort: {key: RetainedSortKey; dir: SortDir} = {key: 'amountDelta', dir: 'desc'};

    const splitByRetail = <T extends { customerCode: string }>(rows: ReadonlyArray<T>) => {
        const regular: T[] = [];
        const retail: T[] = [];
        for (const r of rows) {
            (isRetail(r.customerCode) ? retail : regular).push(r);
        }
        return {regular, retail};
    };

    const updateCounts = () => {
        if (!lastSeg) return;
        const newSplit = splitByRetail(lastSeg.newCustomers);
        const churnSplit = splitByRetail(lastSeg.churnedCustomers);
        countNew.textContent = String(newSplit.regular.length);
        countRetailNew.textContent = String(newSplit.retail.length);
        countChurn.textContent = String(churnSplit.regular.length);
        countRetailChurn.textContent = String(churnSplit.retail.length);
        countRetain.textContent = String(lastSeg.retainedCustomers.length);
    };

    tabHost.addEventListener('click', (e) => {
        const target = (e.target as HTMLElement).closest<HTMLElement>('.cross-month-segment-tab');
        if (!target) return;
        const kind = target.dataset.kind as SegmentKind | undefined;
        if (!kind || kind === activeKind) return;
        activeKind = kind;
        tabHost.querySelectorAll('.cross-month-segment-tab').forEach((el) => {
            el.classList.toggle('is-active', (el as HTMLElement).dataset.kind === kind);
        });
        renderTable();
    });

    // thead 排序：事件代理（thead 的 innerHTML 會在每次 render 重置，但 thead 元素本身不變）
    thead.addEventListener('click', (e) => {
        const th = (e.target as HTMLElement).closest<HTMLElement>('th[data-sort-key]');
        if (!th) return;
        const key = th.dataset.sortKey;
        if (!key) return;
        if (activeKind === 'retained') {
            const k = key as RetainedSortKey;
            if (retainedSort.key === k) {
                retainedSort.dir = retainedSort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                retainedSort.key = k;
                // 文字欄預設升冪、數字欄預設降冪（更貼近使用者習慣）
                retainedSort.dir = (k === 'name' || k === 'line') ? 'asc' : 'desc';
            }
        } else {
            const k = key as SimpleSortKey;
            if (simpleSort.key === k) {
                simpleSort.dir = simpleSort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                simpleSort.key = k;
                simpleSort.dir = (k === 'name' || k === 'line') ? 'asc' : 'desc';
            }
        }
        renderTable();
    });

    const sumStats = (rows: ReadonlyArray<{ amount: number; profit: number; count: number }>): {
        amount: number;
        profit: number;
        count: number;
    } => {
        let amount = 0, profit = 0, count = 0;
        for (const r of rows) {
            amount += r.amount;
            profit += r.profit;
            count += r.count;
        }
        return {amount, profit, count};
    };

    const renderSummary = () => {
        if (!lastSeg || !lastCurrent || !lastPrevious) {
            summaryEl.innerHTML = '';
            return;
        }
        if (activeKind === 'retained') {
            const stats = sumStats(lastSeg.retainedCustomers);
            const prevStats = sumStats(lastSeg.retainedCustomers.map((r) => ({
                amount: r.prevAmount,
                profit: r.prevProfit,
                count: r.prevCount,
            })));
            summaryEl.innerHTML = renderSummaryHtml({
                count: lastSeg.retainedCustomers.length,
                rows: [
                    {
                        label: `${lastPrevious.label} 營收`,
                        value: fmtMoney(Math.round(prevStats.amount)),
                        ratio: ratioPct(prevStats.amount, lastPrevious.amount),
                        ratioLabel: `佔 ${lastPrevious.label}`,
                    },
                    {
                        label: `${lastCurrent.label} 營收`,
                        value: fmtMoney(Math.round(stats.amount)),
                        ratio: ratioPct(stats.amount, lastCurrent.amount),
                        ratioLabel: `佔 ${lastCurrent.label}`,
                    },
                    {
                        label: `${lastPrevious.label} 毛利`,
                        value: fmtMoney(Math.round(prevStats.profit)),
                        ratio: ratioPct(prevStats.profit, lastPrevious.profit),
                        ratioLabel: `佔 ${lastPrevious.label}`,
                    },
                    {
                        label: `${lastCurrent.label} 毛利`,
                        value: fmtMoney(Math.round(stats.profit)),
                        ratio: ratioPct(stats.profit, lastCurrent.profit),
                        ratioLabel: `佔 ${lastCurrent.label}`,
                    },
                ],
            });
            return;
        }

        const isNewSide = activeKind === 'new' || activeKind === 'retail-new';
        const isRetailTab = activeKind === 'retail-new' || activeKind === 'retail-churned';
        const source = isNewSide ? lastSeg.newCustomers : lastSeg.churnedCustomers;
        const split = splitByRetail(source);
        const tabRows = isRetailTab ? split.retail : split.regular;
        const denom = isNewSide ? lastCurrent : lastPrevious;
        const stats = sumStats(tabRows);
        summaryEl.innerHTML = renderSummaryHtml({
            count: tabRows.length,
            rows: [
                {
                    label: '營收',
                    value: fmtMoney(Math.round(stats.amount)),
                    ratio: ratioPct(stats.amount, denom.amount),
                    ratioLabel: `佔 ${denom.label}`,
                },
                {
                    label: '毛利',
                    value: fmtMoney(Math.round(stats.profit)),
                    ratio: ratioPct(stats.profit, denom.profit),
                    ratioLabel: `佔 ${denom.label}`,
                },
                {
                    label: '數量',
                    value: fmtCount(stats.count),
                    ratio: ratioPct(stats.count, denom.count),
                    ratioLabel: `佔 ${denom.label}`,
                },
            ],
        });
    };

    const TAB_COLORS: Record<SegmentKind, string> = {
        'new': '#5b8def',
        'retail-new': '#9b6ddf',
        'churned': '#e74c3c',
        'retail-churned': '#f0a07c',
        'retained': '#7ac74f',
        'flow': '#0ea5e9',
    };

    const renderBarRow = (label: string, num: number, den: number, fmt: (v: number) => string, color: string): string => {
        const pct = den > 0 ? (num / den) * 100 : 0;
        const fillW = Math.min(pct, 100);
        const pctStr = den > 0 ? `${pct.toFixed(1)}%` : '—';
        return `
            <div class="cross-month-bar-row">
                <span class="cross-month-bar-label">${escapeHtml(label)}</span>
                <div class="cross-month-bar-track">
                    <div class="cross-month-bar-fill" style="width: ${fillW}%; background: ${color}"></div>
                </div>
                <span class="cross-month-bar-value">${escapeHtml(fmt(num))} / ${escapeHtml(fmt(den))}</span>
                <span class="cross-month-bar-pct">${escapeHtml(pctStr)}</span>
            </div>
        `;
    };

    const renderBars = () => {
        if (!lastSeg || !lastCurrent || !lastPrevious) {
            barsEl.innerHTML = '';
            return;
        }
        const color = TAB_COLORS[activeKind];

        if (activeKind === 'retained') {
            // 留存：本月 3 條（vs 本月）+ 上月 3 條（vs 上月），合併在一起
            const curStats = sumStats(lastSeg.retainedCustomers);
            const prevStats = sumStats(lastSeg.retainedCustomers.map((r) => ({
                amount: r.prevAmount,
                profit: r.prevProfit,
                count: r.prevCount,
            })));
            barsEl.innerHTML = `
                <div class="cross-month-bars-group-label">${escapeHtml(lastPrevious.label)} 留存佔比</div>
                ${renderBarRow('營收', Math.round(prevStats.amount), Math.round(lastPrevious.amount), fmtMoney, color)}
                ${renderBarRow('毛利', Math.round(prevStats.profit), Math.round(lastPrevious.profit), fmtMoney, color)}
                ${renderBarRow('數量', prevStats.count, lastPrevious.count, fmtCount, color)}
                <div class="cross-month-bars-group-label">${escapeHtml(lastCurrent.label)} 留存佔比</div>
                ${renderBarRow('營收', Math.round(curStats.amount), Math.round(lastCurrent.amount), fmtMoney, color)}
                ${renderBarRow('毛利', Math.round(curStats.profit), Math.round(lastCurrent.profit), fmtMoney, color)}
                ${renderBarRow('數量', curStats.count, lastCurrent.count, fmtCount, color)}
            `;
            return;
        }

        const isNewSide = activeKind === 'new' || activeKind === 'retail-new';
        const isRetailTab = activeKind === 'retail-new' || activeKind === 'retail-churned';
        const source = isNewSide ? lastSeg.newCustomers : lastSeg.churnedCustomers;
        const split = splitByRetail(source);
        const tabRows = isRetailTab ? split.retail : split.regular;
        const denom = isNewSide ? lastCurrent : lastPrevious;
        const stats = sumStats(tabRows);

        barsEl.innerHTML = `
            ${renderBarRow('營收', Math.round(stats.amount), Math.round(denom.amount), fmtMoney, color)}
            ${renderBarRow('毛利', Math.round(stats.profit), Math.round(denom.profit), fmtMoney, color)}
            ${renderBarRow('數量', stats.count, denom.count, fmtCount, color)}
        `;
    };

    const renderTable = () => {
        if (!lastSeg) {
            thead.innerHTML = '';
            tbody.innerHTML = '';
            summaryEl.innerHTML = '';
            barsEl.innerHTML = '';
            medianCaptionEl.innerHTML = '';
            flowCardsEl.innerHTML = '';
            return;
        }
        if (activeKind === 'flow') {
            // 客戶流動分頁不使用 table，改以卡片呈現「新增 vs 流失 = 淨差額」
            summaryEl.innerHTML = '';
            barsEl.innerHTML = '';
            medianCaptionEl.innerHTML = '';
            thead.innerHTML = '';
            tbody.innerHTML = '';
            tableWrapEl.style.display = 'none';
            flowCardsEl.style.display = '';
            renderFlowCards();
            return;
        }
        tableWrapEl.style.display = '';
        flowCardsEl.style.display = 'none';
        flowCardsEl.innerHTML = '';
        renderSummary();
        renderBars();
        renderMedianCaption();
        if (activeKind === 'retained') {
            renderRetainedTable(lastSeg.retainedCustomers);
            return;
        }
        const isNewSide = activeKind === 'new' || activeKind === 'retail-new';
        const isRetailTab = activeKind === 'retail-new' || activeKind === 'retail-churned';
        const source = isNewSide ? lastSeg.newCustomers : lastSeg.churnedCustomers;
        const split = splitByRetail(source);
        const rows = isRetailTab ? split.retail : split.regular;
        renderSimpleTable(rows, isNewSide, isRetailTab);
    };

    /**
     * 客戶流動分頁：以卡片呈現「新增營收 / 流失營收 / 淨差額」與客戶數對比，
     * 並列「一般客戶（已扣除零售）」與「零售客戶」兩組。
     */
    const renderFlowCards = () => {
        if (!lastSeg) {
            flowCardsEl.innerHTML = '';
            return;
        }
        const newSplit = splitByRetail(lastSeg.newCustomers);
        const churnSplit = splitByRetail(lastSeg.churnedCustomers);

        flowCardsEl.innerHTML = `
            ${renderFlowGroup('一般客戶', newSplit.regular, churnSplit.regular)}
            ${renderFlowGroup('零售客戶', newSplit.retail, churnSplit.retail)}
        `;
    };

    const renderFlowGroup = (
        label: string,
        newRows: ReadonlyArray<CustomerMonthStat>,
        churnRows: ReadonlyArray<CustomerMonthStat>
    ): string => {
        const newAmount = newRows.reduce((s, r) => s + r.amount, 0);
        const churnAmount = churnRows.reduce((s, r) => s + r.amount, 0);
        const amountNet = newAmount - churnAmount;
        const countNet = newRows.length - churnRows.length;

        const netToneAmount = amountNet > 0 ? 'is-positive' : amountNet < 0 ? 'is-negative' : 'is-neutral';
        const netToneCount = countNet > 0 ? 'is-positive' : countNet < 0 ? 'is-negative' : 'is-neutral';

        return `
            <div class="cross-month-flow-group">
                <div class="cross-month-flow-group-label">${escapeHtml(label)}</div>
                <div class="cross-month-flow-row">
                    <div class="cross-month-flow-card is-new">
                        <div class="cross-month-flow-card-title">新增</div>
                        <div class="cross-month-flow-card-amount">${fmtMoney(Math.round(newAmount))}</div>
                        <div class="cross-month-flow-card-sub">${newRows.length} 位客戶</div>
                    </div>
                    <div class="cross-month-flow-op">−</div>
                    <div class="cross-month-flow-card is-churn">
                        <div class="cross-month-flow-card-title">流失</div>
                        <div class="cross-month-flow-card-amount">${fmtMoney(Math.round(churnAmount))}</div>
                        <div class="cross-month-flow-card-sub">${churnRows.length} 位客戶</div>
                    </div>
                    <div class="cross-month-flow-op">=</div>
                    <div class="cross-month-flow-card is-net ${netToneAmount}">
                        <div class="cross-month-flow-card-title">淨差額</div>
                        <div class="cross-month-flow-card-amount">${fmtDelta(amountNet)}</div>
                        <div class="cross-month-flow-card-sub ${netToneCount}">${countNet >= 0 ? '+' : ''}${countNet} 位客戶</div>
                    </div>
                </div>
            </div>
        `;
    };

    /** 依當前分頁決定中位數依據哪一個月並渲染提示 caption。 */
    const renderMedianCaption = () => {
        if (!lastCurrent || !lastPrevious) {
            medianCaptionEl.innerHTML = '';
            return;
        }
        if (activeKind === 'retained') {
            // 留存：兩月並列，雙月份中位數一起標示
            medianCaptionEl.innerHTML = `
                ${medianHintCaption(lastPrevious.label, lastPrevious.median)}
                ${medianHintCaption(lastCurrent.label, lastCurrent.median)}
            `;
            return;
        }
        const isNewSide = activeKind === 'new' || activeKind === 'retail-new';
        const ref = isNewSide ? lastCurrent : lastPrevious;
        medianCaptionEl.innerHTML = medianHintCaption(ref.label, ref.median);
    };

    const valueOfSimple = (r: CustomerMonthStat, k: SimpleSortKey): number | string => {
        switch (k) {
            case 'name': return r.customerName;
            case 'line': return r.line;
            case 'amount': return r.amount;
            case 'count': return r.count;
            case 'profit': return r.profit;
        }
    };

    const renderSimpleTable = (
        rows: ReadonlyArray<CustomerMonthStat>,
        isNewSide: boolean,
        isRetailTab: boolean
    ) => {
        const ref = isNewSide ? lastCurrent : lastPrevious;
        const monthLabel = ref?.label ?? '';
        const median = ref?.median;
        const ind = (k: SimpleSortKey) => sortIndicatorHtml(simpleSort.key === k, simpleSort.dir);
        thead.innerHTML = `
      <tr>
        <th data-sort-key="name" class="is-sortable">客戶${ind('name')}</th>
        <th data-sort-key="line" class="is-sortable">線別${ind('line')}</th>
        <th data-sort-key="amount" class="is-sortable is-right">${monthLabel}　金額${ind('amount')}</th>
        <th data-sort-key="count" class="is-sortable is-right">${monthLabel}　數量${ind('count')}</th>
        <th data-sort-key="profit" class="is-sortable is-right">${monthLabel}　毛利${ind('profit')}</th>
        <th class="is-center">操作</th>
      </tr>
    `;
        if (rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="6" class="cross-month-empty">無資料</td></tr>`;
            return;
        }
        const sortedRows = sortRows(rows, (r) => valueOfSimple(r, simpleSort.key), simpleSort.dir);
        const actionLabel = isRetailTab ? '取消零售' : '標為零售';
        const actionClass = isRetailTab ? 'is-unmark' : 'is-mark';
        tbody.innerHTML = sortedRows
            .map((r) => {
                const amountBadge = median ? vsMedianBadge(r.amount, median.amount) : '';
                const countBadge = median ? vsMedianBadge(r.count, median.count) : '';
                const profitBadge = median ? vsMedianBadge(r.profit, median.profit) : '';
                return `
        <tr class="${opts.onCustomerClick ? 'is-clickable' : ''}" data-code="${escapeHtml(r.customerCode)}">
          <td>${escapeHtml(r.customerName)}<span class="cross-month-cell-sub">(${escapeHtml(r.customerCode)})</span></td>
          <td>${escapeHtml(r.line)}</td>
          <td class="is-right cross-month-cell-with-badge">${fmtMoney(Math.round(r.amount))}${amountBadge}</td>
          <td class="is-right cross-month-cell-with-badge">${fmtCount(r.count)}${countBadge}</td>
          <td class="is-right cross-month-cell-with-badge ${r.profit < 0 ? 'cell-negative' : ''}">${fmtMoney(Math.round(r.profit))}${profitBadge}</td>
          <td class="is-center">
            <button type="button" class="cross-month-retail-btn ${actionClass}"
                    data-action="toggle-retail" data-code="${escapeHtml(r.customerCode)}">
              ${actionLabel}
            </button>
          </td>
        </tr>
      `;
            })
            .join('');
        attachRowClick();
        attachRetailToggle();
    };

    const valueOfRetained = (r: RetainedCustomer, k: RetainedSortKey): number | string => {
        switch (k) {
            case 'name': return r.customerName;
            case 'line': return r.line;
            case 'prevAmount': return r.prevAmount;
            case 'amount': return r.amount;
            case 'amountDelta': return r.amountDelta;
            case 'amountDeltaPct': return r.amountDeltaPct ?? Number.NEGATIVE_INFINITY;
            case 'countDelta': return r.countDelta;
            case 'countDeltaPct': return r.countDeltaPct ?? Number.NEGATIVE_INFINITY;
        }
    };

    const renderRetainedTable = (rows: ReadonlyArray<RetainedCustomer>) => {
        const prevLbl = lastPrevious?.label ?? '';
        const curLbl = lastCurrent?.label ?? '';
        const prevMedian = lastPrevious?.median;
        const curMedian = lastCurrent?.median;
        const ind = (k: RetainedSortKey) => sortIndicatorHtml(retainedSort.key === k, retainedSort.dir);
        thead.innerHTML = `
      <tr>
        <th data-sort-key="name" class="is-sortable">客戶${ind('name')}</th>
        <th data-sort-key="line" class="is-sortable">線別${ind('line')}</th>
        <th data-sort-key="prevAmount" class="is-sortable is-right">${prevLbl} 金額${ind('prevAmount')}</th>
        <th data-sort-key="amount" class="is-sortable is-right">${curLbl} 金額${ind('amount')}</th>
        <th data-sort-key="amountDelta" class="is-sortable is-right" title="本月 - 上月">Δ金額${ind('amountDelta')}</th>
        <th data-sort-key="amountDeltaPct" class="is-sortable is-right">Δ%${ind('amountDeltaPct')}</th>
        <th data-sort-key="countDelta" class="is-sortable is-right" title="本月 - 上月">Δ數量${ind('countDelta')}</th>
        <th data-sort-key="countDeltaPct" class="is-sortable is-right">Δ數量%${ind('countDeltaPct')}</th>
      </tr>
    `;
        if (rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="cross-month-empty">無資料</td></tr>`;
            return;
        }
        const sortedRows = sortRows(rows, (r) => valueOfRetained(r, retainedSort.key), retainedSort.dir);
        tbody.innerHTML = sortedRows
            .map((r) => {
                const upClass = r.amountDelta >= 0 ? 'cell-positive' : 'cell-negative';
                const cntClass = r.countDelta >= 0 ? 'cell-positive' : 'cell-negative';
                const prevAmountBadge = prevMedian ? vsMedianBadge(r.prevAmount, prevMedian.amount) : '';
                const curAmountBadge = curMedian ? vsMedianBadge(r.amount, curMedian.amount) : '';
                return `
        <tr class="${opts.onCustomerClick ? 'is-clickable' : ''}" data-code="${escapeHtml(r.customerCode)}">
          <td>${escapeHtml(r.customerName)}<span class="cross-month-cell-sub">(${escapeHtml(r.customerCode)})</span></td>
          <td>${escapeHtml(r.line)}</td>
          <td class="is-right cross-month-cell-with-badge">${fmtMoney(Math.round(r.prevAmount))}${prevAmountBadge}</td>
          <td class="is-right cross-month-cell-with-badge">${fmtMoney(Math.round(r.amount))}${curAmountBadge}</td>
          <td class="is-right ${upClass}">${fmtDelta(r.amountDelta)}</td>
          <td class="is-right ${upClass}">${fmtPct(r.amountDeltaPct)}</td>
          <td class="is-right ${cntClass}">${r.countDelta >= 0 ? '+' : ''}${fmtCount(r.countDelta)}</td>
          <td class="is-right ${cntClass}">${fmtPct(r.countDeltaPct)}</td>
        </tr>
      `;
            })
            .join('');
        attachRowClick();
    };

    const attachRowClick = () => {
        if (!opts.onCustomerClick) return;
        tbody.querySelectorAll<HTMLElement>('tr[data-code]').forEach((tr) => {
            tr.addEventListener('click', (e) => {
                if ((e.target as HTMLElement).closest('[data-action="toggle-retail"]')) return;
                const code = tr.dataset.code;
                if (code) opts.onCustomerClick!(code);
            });
        });
    };

    const attachRetailToggle = () => {
        tbody.querySelectorAll<HTMLButtonElement>('[data-action="toggle-retail"]').forEach((btn) => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const code = btn.dataset.code;
                if (!code) return;
                setRetail(code, !isRetail(code));
            });
        });
    };

    const unsubscribe = subscribeRetail(() => {
        if (!root.isConnected) {
            unsubscribe();
            return;
        }
        if (!lastSeg) return;
        updateCounts();
        renderTable();
    });

    return {
        element: root,
        setData(seg, current, previous) {
            lastSeg = seg;
            lastCurrent = current;
            lastPrevious = previous;
            updateCounts();
            renderTable();
        },
    };
}

interface SummaryRow {
    label: string;
    value: string;
    ratio: string;
    ratioLabel: string;
}

function renderSummaryHtml(data: { count: number; rows: ReadonlyArray<SummaryRow> }): string {
    const itemHtml = data.rows
        .map((r) => `
            <span class="cross-month-segment-summary-item">
                <span class="cross-month-segment-summary-key">${escapeHtml(r.label)}</span>
                <span class="cross-month-segment-summary-val">${escapeHtml(r.value)}</span>
                <span class="cross-month-segment-summary-ratio" title="${escapeHtml(r.ratioLabel)}">${escapeHtml(r.ratioLabel)} ${escapeHtml(r.ratio)}</span>
            </span>
        `)
        .join('');
    return `
        <span class="cross-month-segment-summary-count">共 ${data.count} 間客戶</span>
        ${itemHtml}
    `;
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
