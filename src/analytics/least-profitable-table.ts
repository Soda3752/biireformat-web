/**
 * 客戶低價值分析表：以「平均售價（單價）」與「平均單品毛利」為主要訊號，
 * 找出「買很多但平均單價偏低 / 每單位賺得少」的客戶。
 *
 * 預設排序：平均售價 升冪（最低單價 = 第 1 名），搭配「最低數量門檻」可篩掉
 * 量很少卻碰巧買到便宜品項的客戶，聚焦在真正有量但拉低均價的客戶。
 *
 * 與 detail-table.ts 為兄弟元件，但聚合粒度為「客戶」而非「明細列」。
 */

import type {AnalyticsRow} from './dataset-builder';
import {marginPct} from './aggregators';

type SortKey =
    | 'rank'
    | 'customer'
    | 'count'
    | 'avgPrice'
    | 'avgUnitProfit'
    | 'amount'
    | 'costAmount'
    | 'profit'
    | 'margin';
type SortDir = 'asc' | 'desc';

const TOP_N_OPTIONS: Array<{ value: number; label: string }> = [
    {value: 10, label: '前 10'},
    {value: 20, label: '前 20'},
    {value: 50, label: '前 50'},
    {value: 0, label: '全部'},
];

interface CustomerAgg {
    customerCode: string;
    customerName: string;
    lines: Set<string>;
    count: number;
    amount: number;
    costAmount: number;
    profit: number;
    avgPrice: number;       // amount / count
    avgUnitProfit: number;  // profit / count
    allCostUnset: boolean;
}

export interface LeastProfitableTableOptions {
    /** 點擊任一列時觸發，傳入該列的客戶代碼，呼叫端可據此套用篩選器。 */
    onCustomerClick?: (customerCode: string) => void;
}

export interface LeastProfitableTableController {
    element: HTMLElement;

    setRows(rows: ReadonlyArray<AnalyticsRow>): void;
}

export function createLeastProfitableTable(
    options: LeastProfitableTableOptions = {}
): LeastProfitableTableController {
    const root = document.createElement('div');
    root.className = 'analytics-detail analytics-least-profit';
    root.innerHTML = `
    <div class="analytics-detail-toolbar analytics-least-profit-toolbar">
      <label class="analytics-least-profit-field">
        <span>最低數量</span>
        <input type="number" min="0" step="1" value="0" data-role="min-count" placeholder="0">
      </label>
      <label class="analytics-least-profit-field">
        <span>顯示</span>
        <select data-role="topn">
          ${TOP_N_OPTIONS.map((o) => `<option value="${o.value}">${o.label}</option>`).join('')}
        </select>
      </label>
      <span class="analytics-detail-stats" data-role="stats"></span>
    </div>
    <div class="analytics-detail-table-wrap">
      <table class="analytics-detail-table-el">
        <thead>
          <tr>
            <th data-sort-key="rank" class="is-right">排名<span class="sort-indicator"></span></th>
            <th data-sort-key="customer">客戶<span class="sort-indicator"></span></th>
            <th>線別</th>
            <th data-sort-key="count" class="is-right">銷售數量<span class="sort-indicator"></span></th>
            <th data-sort-key="avgPrice" class="is-right">平均售價<span class="sort-indicator"></span></th>
            <th data-sort-key="avgUnitProfit" class="is-right">平均單品毛利<span class="sort-indicator"></span></th>
            <th data-sort-key="amount" class="is-right">營收<span class="sort-indicator"></span></th>
            <th data-sort-key="costAmount" class="is-right">成本<span class="sort-indicator"></span></th>
            <th data-sort-key="profit" class="is-right">毛利<span class="sort-indicator"></span></th>
            <th data-sort-key="margin" class="is-right">毛利率<span class="sort-indicator"></span></th>
          </tr>
        </thead>
        <tbody data-role="tbody"></tbody>
      </table>
    </div>
  `;

    const minCountInput = root.querySelector<HTMLInputElement>('[data-role="min-count"]')!;
    const topnSelect = root.querySelector<HTMLSelectElement>('[data-role="topn"]')!;
    const statsEl = root.querySelector<HTMLElement>('[data-role="stats"]')!;
    const tbody = root.querySelector<HTMLTableSectionElement>('[data-role="tbody"]')!;

    let allRows: ReadonlyArray<AnalyticsRow> = [];
    let sortKey: SortKey = 'avgPrice';
    let sortDir: SortDir = 'asc';
    let topN = 20;
    let minCount = 0;

    const fmt = (v: number) => v.toLocaleString('zh-TW');
    const fmtMoney = (v: number) => Math.round(v).toLocaleString('zh-TW');
    const fmtPct = (v: number | null) => (v === null ? '—' : `${v.toFixed(1)}%`);
    const fmtAvgPrice = (v: number) => (v > 0 ? v.toFixed(1) : '—');

    const aggregate = (): CustomerAgg[] => {
        const map = new Map<string, CustomerAgg>();
        for (const r of allRows) {
            let g = map.get(r.customerCode);
            if (!g) {
                g = {
                    customerCode: r.customerCode,
                    customerName: r.customerName,
                    lines: new Set(),
                    count: 0,
                    amount: 0,
                    costAmount: 0,
                    profit: 0,
                    avgPrice: 0,
                    avgUnitProfit: 0,
                    allCostUnset: true,
                };
                map.set(r.customerCode, g);
            }
            g.lines.add(r.line);
            g.count += r.count;
            g.amount += r.amount;
            g.costAmount += r.costAmount;
            g.profit += r.profit;
            if (!r.isCostUnset) g.allCostUnset = false;
        }
        // 計算 avgPrice / avgUnitProfit
        for (const g of map.values()) {
            g.avgPrice = g.count > 0 ? g.amount / g.count : 0;
            g.avgUnitProfit = g.count > 0 ? g.profit / g.count : 0;
        }
        return [...map.values()];
    };

    const compareBy = (a: CustomerAgg, b: CustomerAgg): number => {
        if (sortKey === 'customer') return a.customerName.localeCompare(b.customerName);
        if (sortKey === 'count') return a.count - b.count;
        if (sortKey === 'avgPrice') return a.avgPrice - b.avgPrice;
        if (sortKey === 'avgUnitProfit') return a.avgUnitProfit - b.avgUnitProfit;
        if (sortKey === 'amount') return a.amount - b.amount;
        if (sortKey === 'costAmount') return a.costAmount - b.costAmount;
        if (sortKey === 'profit') return a.profit - b.profit;
        if (sortKey === 'margin') {
            const am = marginPct(a.profit, a.amount);
            const bm = marginPct(b.profit, b.amount);
            // 把 null（amount=0，無毛利率可言）一律排到最後
            if (am === null && bm === null) return 0;
            if (am === null) return 1;
            if (bm === null) return -1;
            return am - bm;
        }
        // rank：fallback，沿用 avgPrice 升冪
        return a.avgPrice - b.avgPrice;
    };

    const render = () => {
        const aggsAll = aggregate();
        const aggs = aggsAll.filter((g) => g.count >= minCount);

        // 整體基準（套最低數量門檻後計算，比較有可比性）
        const totalAmount = aggs.reduce((s, g) => s + g.amount, 0);
        const totalCount = aggs.reduce((s, g) => s + g.count, 0);
        const overallAvgPrice = totalCount > 0 ? totalAmount / totalCount : 0;

        // 用「升冪 by avgPrice」決定 rank（最低均價 = 第 1 名）
        const ranked = [...aggs].sort((a, b) => a.avgPrice - b.avgPrice);
        const rankMap = new Map<string, number>();
        ranked.forEach((g, i) => rankMap.set(g.customerCode, i + 1));

        // 套使用者選擇的排序
        const sorted = [...aggs].sort((a, b) => {
            if (sortKey === 'rank') {
                const cmp = (rankMap.get(a.customerCode) ?? 0) - (rankMap.get(b.customerCode) ?? 0);
                return sortDir === 'asc' ? cmp : -cmp;
            }
            const cmp = compareBy(a, b);
            return sortDir === 'asc' ? cmp : -cmp;
        });

        const limited = topN > 0 ? sorted.slice(0, topN) : sorted;
        const filterHint = minCount > 0 ? `（已過濾數量≥${minCount}）` : '';
        const benchmark = overallAvgPrice > 0 ? ` / 整體平均售價 ${overallAvgPrice.toFixed(1)}` : '';
        statsEl.textContent = `符合條件 ${aggs.length.toLocaleString()} 位 / 全部 ${aggsAll.length.toLocaleString()} 位${filterHint} / 顯示 ${limited.length.toLocaleString()} 位${benchmark}`;

        if (limited.length === 0) {
            tbody.innerHTML = `<tr><td colspan="10" style="text-align:center;color:var(--color-text-muted);padding:24px">無符合條件的客戶</td></tr>`;
        } else {
            tbody.innerHTML = limited
                .map((g) => {
                    const rank = rankMap.get(g.customerCode) ?? 0;
                    const margin = marginPct(g.profit, g.amount);
                    const isLoss = g.profit < 0;
                    // 平均售價低於整體均價 → 視覺提示
                    const isLowPrice = overallAvgPrice > 0 && g.avgPrice > 0 && g.avgPrice < overallAvgPrice;
                    const lines = [...g.lines].sort().join('、');
                    const warn = g.allCostUnset && g.amount > 0
                        ? `<span class="analytics-warn-pill" title="該客戶所有商品未填成本，毛利視為 0">⚠</span>`
                        : '';
                    const rowClass = [
                        isLoss ? 'is-loss' : '',
                        isLowPrice ? 'is-low-price' : '',
                        options.onCustomerClick ? 'is-clickable' : '',
                    ].filter(Boolean).join(' ');
                    return `
              <tr class="${rowClass}" data-customer-code="${escapeHtml(g.customerCode)}" ${options.onCustomerClick ? 'title="點擊以篩選此客戶"' : ''}>
                <td class="is-right">${rank}</td>
                <td>${escapeHtml(g.customerName)}（${escapeHtml(g.customerCode)}）${warn}</td>
                <td>${escapeHtml(lines)}</td>
                <td class="is-right">${fmt(g.count)}</td>
                <td class="is-right">${fmtAvgPrice(g.avgPrice)}</td>
                <td class="is-right">${fmtAvgPrice(g.avgUnitProfit)}</td>
                <td class="is-right">${fmtMoney(g.amount)}</td>
                <td class="is-right">${fmtMoney(g.costAmount)}</td>
                <td class="is-right">${fmtMoney(g.profit)}</td>
                <td class="is-right">${fmtPct(margin)}</td>
              </tr>
            `;
                })
                .join('');
        }

        // 排序指示器
        root.querySelectorAll<HTMLElement>('th[data-sort-key]').forEach((th) => {
            const ind = th.querySelector<HTMLElement>('.sort-indicator')!;
            if (th.dataset.sortKey === sortKey) {
                ind.textContent = sortDir === 'asc' ? ' ▲' : ' ▼';
            } else {
                ind.textContent = '';
            }
        });

        // 列點擊 → 觸發 onCustomerClick（每次 render 都重綁，因為 innerHTML 重建）
        if (options.onCustomerClick) {
            tbody.querySelectorAll<HTMLTableRowElement>('tr[data-customer-code]').forEach((tr) => {
                tr.addEventListener('click', () => {
                    const code = tr.dataset.customerCode ?? '';
                    if (code) options.onCustomerClick!(code);
                });
            });
        }
    };

    // 排序事件
    root.querySelectorAll<HTMLElement>('th[data-sort-key]').forEach((th) => {
        th.addEventListener('click', () => {
            const k = th.dataset.sortKey as SortKey;
            if (sortKey === k) {
                sortDir = sortDir === 'asc' ? 'desc' : 'asc';
            } else {
                sortKey = k;
                // 數值欄預設升冪（看最低），文字欄升冪
                sortDir = 'asc';
            }
            render();
        });
    });

    topnSelect.value = String(topN);
    topnSelect.addEventListener('change', () => {
        topN = Number(topnSelect.value);
        render();
    });

    let minCountTimer: number | null = null;
    minCountInput.addEventListener('input', () => {
        if (minCountTimer) window.clearTimeout(minCountTimer);
        minCountTimer = window.setTimeout(() => {
            const v = Number(minCountInput.value);
            minCount = Number.isFinite(v) && v > 0 ? Math.floor(v) : 0;
            render();
        }, 200);
    });

    return {
        element: root,
        setRows(rows) {
            allRows = rows;
            render();
        },
    };
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
