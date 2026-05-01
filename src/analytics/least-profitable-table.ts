/**
 * 客戶低價值分析表：以「相對均價損益」為主要訊號，
 * 找出「賣愈多虧愈多 / 該調漲價格」的客戶。
 *
 * 相對均價損益 = Σ (客戶該商品單價 − 該商品在資料集的整體加權均價) × 數量
 *   - 正值 = 該客戶整體高於均價（營利、優質客戶）
 *   - 負值 = 該客戶整體低於均價（虧損、該漲價）
 * 每單位損益 = 相對均價損益 ÷ 該客戶總銷售數量，輔助判斷該漲多少。
 *
 * 預設排序：相對均價損益 升冪（最虧 = 第 1 名 = 最該漲價），搭配「最低數量門檻」
 * 可篩掉量很少的客戶，聚焦在真正有量但拉低我方收入的客戶。
 *
 * 與 detail-table.ts 為兄弟元件，但聚合粒度為「客戶」而非「明細列」。
 */

import type {AnalyticsRow} from './dataset-builder';
import {marginPct} from './aggregators';

type SortKey =
    | 'rank'
    | 'customer'
    | 'count'
    | 'relativePnL'
    | 'perUnitPnL'
    | 'avgUnitProfit'
    | 'amount'
    | 'costAmount'
    | 'profit'
    | 'margin';
type SortDir = 'asc' | 'desc';

interface CustomerAgg {
    customerCode: string;
    customerName: string;
    lines: Set<string>;
    count: number;
    amount: number;
    costAmount: number;
    profit: number;
    avgUnitProfit: number;  // profit / count
    relativePnL: number;    // Σ (客戶單價 − 商品均價) × 數量；正 = 營利、負 = 虧損（該漲價）
    perUnitPnL: number;     // relativePnL / count
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
            <th data-sort-key="relativePnL" class="is-right" title="Σ (客戶單價 − 該商品整體均價) × 數量。正值 = 高於均價（營利、優質客戶）；負值 = 低於均價（虧損、該漲價）。">相對均價損益<span class="sort-indicator"></span></th>
            <th data-sort-key="perUnitPnL" class="is-right" title="相對均價損益 ÷ 總銷售數量。平均每件相對均價多收（+）或少收（−）多少元，輔助判斷該漲多少。">每單位損益<span class="sort-indicator"></span></th>
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
    const statsEl = root.querySelector<HTMLElement>('[data-role="stats"]')!;
    const tbody = root.querySelector<HTMLTableSectionElement>('[data-role="tbody"]')!;

    let allRows: ReadonlyArray<AnalyticsRow> = [];
    let sortKey: SortKey = 'relativePnL';
    let sortDir: SortDir = 'asc';
    let minCount = 0;

    const fmt = (v: number) => v.toLocaleString('zh-TW');
    const fmtMoney = (v: number) => Math.round(v).toLocaleString('zh-TW');
    const fmtPct = (v: number | null) => (v === null ? '—' : `${v.toFixed(1)}%`);
    const fmtAvgPrice = (v: number) => (v > 0 ? v.toFixed(1) : '—');
    const fmtSignedMoney = (v: number) => {
        const r = Math.round(v);
        if (r === 0) return '0';
        return (r > 0 ? '+' : '') + r.toLocaleString('zh-TW');
    };
    const fmtSignedDecimal = (v: number) => {
        if (Math.abs(v) < 0.05) return '0';
        return (v > 0 ? '+' : '') + v.toFixed(1);
    };

    const aggregate = (): CustomerAgg[] => {
        // 第一輪：算每項商品在「目前資料集」的整體均價（跨所有客戶）
        // 作為「相對均價損失」的基準線
        const productStats = new Map<string, { amount: number; count: number }>();
        for (const r of allRows) {
            let s = productStats.get(r.productName);
            if (!s) {
                s = {amount: 0, count: 0};
                productStats.set(r.productName, s);
            }
            s.amount += r.amount;
            s.count += r.count;
        }

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
                    avgUnitProfit: 0,
                    relativePnL: 0,
                    perUnitPnL: 0,
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

            const ps = productStats.get(r.productName);
            const productAvg = ps && ps.count > 0 ? ps.amount / ps.count : 0;
            g.relativePnL += (r.price - productAvg) * r.count;
        }
        // 計算 avgUnitProfit / perUnitPnL
        for (const g of map.values()) {
            g.avgUnitProfit = g.count > 0 ? g.profit / g.count : 0;
            g.perUnitPnL = g.count > 0 ? g.relativePnL / g.count : 0;
        }
        return [...map.values()];
    };

    const compareBy = (a: CustomerAgg, b: CustomerAgg): number => {
        if (sortKey === 'customer') return a.customerName.localeCompare(b.customerName);
        if (sortKey === 'count') return a.count - b.count;
        if (sortKey === 'relativePnL') return a.relativePnL - b.relativePnL;
        if (sortKey === 'perUnitPnL') return a.perUnitPnL - b.perUnitPnL;
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
        // rank：fallback，沿用 relativePnL 升冪（最虧的 = 最該漲價 = 第 1 名）
        return a.relativePnL - b.relativePnL;
    };

    const render = () => {
        const aggsAll = aggregate();
        const aggs = aggsAll.filter((g) => g.count >= minCount);

        // 用「升冪 by relativePnL」決定 rank（最虧 = 第 1 名 = 最該漲價）
        const ranked = [...aggs].sort((a, b) => a.relativePnL - b.relativePnL);
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

        const limited = sorted;
        const filterHint = minCount > 0 ? `（已過濾數量≥${minCount}）` : '';
        statsEl.textContent = `符合條件 ${aggs.length.toLocaleString()} 位 / 全部 ${aggsAll.length.toLocaleString()} 位${filterHint} / 顯示 ${limited.length.toLocaleString()} 位`;

        if (limited.length === 0) {
            tbody.innerHTML = `<tr><td colspan="11" style="text-align:center;color:var(--color-text-muted);padding:24px">無符合條件的客戶</td></tr>`;
        } else {
            tbody.innerHTML = limited
                .map((g) => {
                    const rank = rankMap.get(g.customerCode) ?? 0;
                    const margin = marginPct(g.profit, g.amount);
                    const isLoss = g.profit < 0;
                    const lines = [...g.lines].sort().join('、');
                    const warn = g.allCostUnset && g.amount > 0
                        ? `<span class="analytics-warn-pill" title="該客戶所有商品未填成本，毛利視為 0">⚠</span>`
                        : '';
                    const rowClass = [
                        isLoss ? 'is-loss' : '',
                        options.onCustomerClick ? 'is-clickable' : '',
                    ].filter(Boolean).join(' ');
                    const pnlClass = g.relativePnL > 0
                        ? 'cell-premium'
                        : g.relativePnL < 0
                            ? 'cell-raise-price'
                            : '';
                    return `
              <tr class="${rowClass}" data-customer-code="${escapeHtml(g.customerCode)}" ${options.onCustomerClick ? 'title="點擊以篩選此客戶"' : ''}>
                <td class="is-right">${rank}</td>
                <td>${escapeHtml(g.customerName)}（${escapeHtml(g.customerCode)}）${warn}</td>
                <td>${escapeHtml(lines)}</td>
                <td class="is-right">${fmt(g.count)}</td>
                <td class="is-right ${pnlClass}">${fmtSignedMoney(g.relativePnL)}</td>
                <td class="is-right ${pnlClass}">${fmtSignedDecimal(g.perUnitPnL)}</td>
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
