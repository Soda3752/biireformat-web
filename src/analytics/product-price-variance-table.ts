/**
 * 商品價差分析表：以「商品」為單位，計算整體加權平均售價，
 * 並比對該商品在每個客戶手中的售價偏離 % ，找出「同一商品在不同客戶售價落差大」的漏血點。
 *
 * 主表一列一商品，點擊可展開該商品的客戶層級明細（客戶售價 vs 整體均價）。
 *
 * 與 least-profitable-table 互補：那張從「客戶」端看，這張從「商品」端看。
 */

import type {AnalyticsRow} from './dataset-builder';

type SortKey =
    | 'rank'
    | 'product'
    | 'customerCount'
    | 'count'
    | 'avgPrice'
    | 'priceRange'
    | 'cv'
    | 'lowCount'
    | 'priceGap'
    | 'amount';
type SortDir = 'asc' | 'desc';

type CustomerSortKey =
    | 'customer'
    | 'line'
    | 'count'
    | 'avgPrice'
    | 'deviation'
    | 'amount';

const DEVIATION_OPTIONS: Array<{ value: number; label: string }> = [
    {value: 5, label: '低於均價 5%'},
    {value: 10, label: '低於均價 10%'},
    {value: 20, label: '低於均價 20%'},
];

interface CustomerStat {
    customerCode: string;
    customerName: string;
    line: string;
    count: number;
    amount: number;
    avgPrice: number;       // amount / count
    deviationPct: number;   // (avgPrice - overallAvg) / overallAvg * 100
}

interface ProductAgg {
    productName: string;
    category: string;
    count: number;          // Σcount
    amount: number;         // Σamount
    overallAvg: number;     // amount / count
    minPrice: number;       // 客戶層級均價的最小值
    maxPrice: number;       // 客戶層級均價的最大值
    cv: number;             // 變異係數（%）：weighted std / mean
    customers: CustomerStat[];
    lowCount: number;       // 偏離門檻以下的客戶數
    priceGap: number;       // Σ (overallAvg - cust.avgPrice) × cust.count，僅算低於門檻的客戶
    allCostUnset: boolean;
}

export interface ProductPriceVarianceTableOptions {
    /** 點擊商品列展開外的「篩選此商品」連動，呼叫端可據此套用 productNames 篩選 */
    onProductFilter?: (productName: string) => void;
    /** 子表內點擊客戶 → 套用客戶篩選 */
    onCustomerClick?: (customerCode: string) => void;
}

export interface ProductPriceVarianceTableController {
    element: HTMLElement;

    setRows(rows: ReadonlyArray<AnalyticsRow>): void;
}

export function createProductPriceVarianceTable(
    options: ProductPriceVarianceTableOptions = {}
): ProductPriceVarianceTableController {
    const root = document.createElement('div');
    root.className = 'analytics-detail analytics-price-variance';
    root.innerHTML = `
    <div class="analytics-detail-toolbar analytics-least-profit-toolbar">
      <label class="analytics-least-profit-field">
        <span>最低數量</span>
        <input type="number" min="0" step="1" value="0" data-role="min-count" placeholder="0">
      </label>
      <label class="analytics-least-profit-field">
        <span>偏離門檻</span>
        <select data-role="threshold">
          ${DEVIATION_OPTIONS.map((o) => `<option value="${o.value}">${o.label}</option>`).join('')}
        </select>
      </label>
      <span class="analytics-detail-stats" data-role="stats"></span>
    </div>
    <div class="analytics-detail-table-wrap">
      <table class="analytics-detail-table-el analytics-price-variance-table">
        <thead>
          <tr>
            <th data-sort-key="rank" class="is-right">排名<span class="sort-indicator"></span></th>
            <th data-sort-key="product">商品<span class="sort-indicator"></span></th>
            <th>分類</th>
            <th data-sort-key="customerCount" class="is-right">客戶數<span class="sort-indicator"></span></th>
            <th data-sort-key="count" class="is-right">銷售數量<span class="sort-indicator"></span></th>
            <th data-sort-key="avgPrice" class="is-right">整體均價<span class="sort-indicator"></span></th>
            <th data-sort-key="priceRange" class="is-right">最低～最高<span class="sort-indicator"></span></th>
            <th data-sort-key="cv" class="is-right">變異 CV<span class="sort-indicator"></span></th>
            <th data-sort-key="lowCount" class="is-right">低價客戶數<span class="sort-indicator"></span></th>
            <th data-sort-key="priceGap" class="is-right">均價缺口<span class="sort-indicator"></span></th>
            <th data-sort-key="amount" class="is-right">營收<span class="sort-indicator"></span></th>
          </tr>
        </thead>
        <tbody data-role="tbody"></tbody>
      </table>
    </div>
  `;

    const minCountInput = root.querySelector<HTMLInputElement>('[data-role="min-count"]')!;
    const thresholdSelect = root.querySelector<HTMLSelectElement>('[data-role="threshold"]')!;
    const statsEl = root.querySelector<HTMLElement>('[data-role="stats"]')!;
    const tbody = root.querySelector<HTMLTableSectionElement>('[data-role="tbody"]')!;

    let allRows: ReadonlyArray<AnalyticsRow> = [];
    let sortKey: SortKey = 'priceGap';
    let sortDir: SortDir = 'desc';
    let customerSortKey: CustomerSortKey = 'deviation';
    let customerSortDir: SortDir = 'asc';
    let minCount = 0;
    let threshold = 10;
    const expanded = new Set<string>();

    const fmt = (v: number) => v.toLocaleString('zh-TW');
    const fmtMoney = (v: number) => Math.round(v).toLocaleString('zh-TW');
    const fmtAvgPrice = (v: number) => (v > 0 ? v.toFixed(1) : '—');
    const fmtPct = (v: number) => `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`;

    const aggregate = (): ProductAgg[] => {
        // 第一層：依商品分群
        const productMap = new Map<string, {
            productName: string;
            category: string;
            count: number;
            amount: number;
            allCostUnset: boolean;
            customerMap: Map<string, CustomerStat>;
        }>();

        for (const r of allRows) {
            let p = productMap.get(r.productName);
            if (!p) {
                p = {
                    productName: r.productName,
                    category: r.category,
                    count: 0,
                    amount: 0,
                    allCostUnset: true,
                    customerMap: new Map(),
                };
                productMap.set(r.productName, p);
            }
            p.count += r.count;
            p.amount += r.amount;
            if (!r.isCostUnset) p.allCostUnset = false;

            let c = p.customerMap.get(r.customerCode);
            if (!c) {
                c = {
                    customerCode: r.customerCode,
                    customerName: r.customerName,
                    line: r.line,
                    count: 0,
                    amount: 0,
                    avgPrice: 0,
                    deviationPct: 0,
                };
                p.customerMap.set(r.customerCode, c);
            }
            c.count += r.count;
            c.amount += r.amount;
        }

        const out: ProductAgg[] = [];
        for (const p of productMap.values()) {
            if (p.count <= 0) continue;
            const overallAvg = p.amount / p.count;

            const customers: CustomerStat[] = [];
            for (const c of p.customerMap.values()) {
                if (c.count <= 0) continue;
                c.avgPrice = c.amount / c.count;
                c.deviationPct = overallAvg > 0
                    ? ((c.avgPrice - overallAvg) / overallAvg) * 100
                    : 0;
                customers.push(c);
            }

            // 變異係數：客戶層級均價的「以數量加權」標準差 / 整體均價
            let weightedSqDiff = 0;
            for (const c of customers) {
                const d = c.avgPrice - overallAvg;
                weightedSqDiff += c.count * d * d;
            }
            const variance = p.count > 0 ? weightedSqDiff / p.count : 0;
            const cv = overallAvg > 0 ? (Math.sqrt(variance) / overallAvg) * 100 : 0;

            const minPrice = customers.length > 0
                ? Math.min(...customers.map((c) => c.avgPrice))
                : 0;
            const maxPrice = customers.length > 0
                ? Math.max(...customers.map((c) => c.avgPrice))
                : 0;

            // 低價客戶（偏離 ≤ -threshold%）
            let lowCount = 0;
            let priceGap = 0;
            for (const c of customers) {
                if (c.deviationPct <= -threshold) {
                    lowCount += 1;
                    priceGap += (overallAvg - c.avgPrice) * c.count;
                }
            }

            // 客戶子表預設按偏離 % 升冪（最低價在前）
            customers.sort((a, b) => a.deviationPct - b.deviationPct);

            out.push({
                productName: p.productName,
                category: p.category,
                count: p.count,
                amount: p.amount,
                overallAvg,
                minPrice,
                maxPrice,
                cv,
                customers,
                lowCount,
                priceGap,
                allCostUnset: p.allCostUnset,
            });
        }
        return out;
    };

    const compareBy = (a: ProductAgg, b: ProductAgg): number => {
        if (sortKey === 'product') return a.productName.localeCompare(b.productName);
        if (sortKey === 'customerCount') return a.customers.length - b.customers.length;
        if (sortKey === 'count') return a.count - b.count;
        if (sortKey === 'avgPrice') return a.overallAvg - b.overallAvg;
        if (sortKey === 'priceRange') return (a.maxPrice - a.minPrice) - (b.maxPrice - b.minPrice);
        if (sortKey === 'cv') return a.cv - b.cv;
        if (sortKey === 'lowCount') return a.lowCount - b.lowCount;
        if (sortKey === 'priceGap') return a.priceGap - b.priceGap;
        if (sortKey === 'amount') return a.amount - b.amount;
        // rank：fallback，沿用 priceGap 降冪
        return a.priceGap - b.priceGap;
    };

    const compareCustomerBy = (a: CustomerStat, b: CustomerStat): number => {
        if (customerSortKey === 'customer') return a.customerName.localeCompare(b.customerName);
        if (customerSortKey === 'line') return a.line.localeCompare(b.line);
        if (customerSortKey === 'count') return a.count - b.count;
        if (customerSortKey === 'avgPrice') return a.avgPrice - b.avgPrice;
        if (customerSortKey === 'deviation') return a.deviationPct - b.deviationPct;
        if (customerSortKey === 'amount') return a.amount - b.amount;
        return 0;
    };

    const customerSortIndicator = (k: CustomerSortKey): string => {
        if (k !== customerSortKey) return '';
        return customerSortDir === 'asc' ? ' ▲' : ' ▼';
    };

    const renderCustomerSubRows = (p: ProductAgg): string => {
        if (p.customers.length === 0) {
            return `<tr><td colspan="11" class="analytics-price-variance-empty">無客戶資料</td></tr>`;
        }
        const sortedCustomers = [...p.customers].sort((a, b) => {
            const cmp = compareCustomerBy(a, b);
            return customerSortDir === 'asc' ? cmp : -cmp;
        });
        const rows = sortedCustomers.map((c) => {
            const isLow = c.deviationPct <= -threshold;
            const isPremium = c.deviationPct >= threshold;
            const cls = isLow ? 'is-low-deviation' : (isPremium ? 'is-high-deviation' : '');
            const clickable = options.onCustomerClick ? 'is-clickable' : '';
            return `
        <tr class="analytics-price-variance-subrow ${cls} ${clickable}"
            data-customer-code="${escapeHtml(c.customerCode)}"
            ${options.onCustomerClick ? 'title="點擊以篩選此客戶"' : ''}>
          <td>${escapeHtml(c.customerName)}（${escapeHtml(c.customerCode)}）</td>
          <td>${escapeHtml(c.line)}</td>
          <td class="is-right">${fmt(c.count)}</td>
          <td class="is-right">${fmtAvgPrice(c.avgPrice)}</td>
          <td class="is-right analytics-price-variance-dev">${fmtPct(c.deviationPct)}</td>
          <td class="is-right">${fmtMoney(c.amount)}</td>
        </tr>
      `;
        }).join('');

        return `
      <tr class="analytics-price-variance-expand-row">
        <td colspan="11">
          <table class="analytics-price-variance-subtable">
            <thead>
              <tr>
                <th data-customer-sort-key="customer">客戶${customerSortIndicator('customer')}</th>
                <th data-customer-sort-key="line">線別${customerSortIndicator('line')}</th>
                <th data-customer-sort-key="count" class="is-right">數量${customerSortIndicator('count')}</th>
                <th data-customer-sort-key="avgPrice" class="is-right">客戶均價${customerSortIndicator('avgPrice')}</th>
                <th data-customer-sort-key="deviation" class="is-right">vs 整體均價${customerSortIndicator('deviation')}</th>
                <th data-customer-sort-key="amount" class="is-right">營收${customerSortIndicator('amount')}</th>
              </tr>
            </thead>
            <tbody>${rows}</tbody>
          </table>
        </td>
      </tr>
    `;
    };

    const render = () => {
        const aggsAll = aggregate();
        const aggs = aggsAll.filter((g) => g.count >= minCount);

        // rank：固定用「均價缺口」降冪（缺口最大 = 第 1 名）
        const ranked = [...aggs].sort((a, b) => b.priceGap - a.priceGap);
        const rankMap = new Map<string, number>();
        ranked.forEach((g, i) => rankMap.set(g.productName, i + 1));

        const sorted = [...aggs].sort((a, b) => {
            if (sortKey === 'rank') {
                const cmp = (rankMap.get(a.productName) ?? 0) - (rankMap.get(b.productName) ?? 0);
                return sortDir === 'asc' ? cmp : -cmp;
            }
            const cmp = compareBy(a, b);
            return sortDir === 'asc' ? cmp : -cmp;
        });

        const limited = sorted;
        const filterHint = minCount > 0 ? `（已過濾數量≥${minCount}）` : '';
        statsEl.textContent = `符合條件 ${aggs.length.toLocaleString()} 項商品 / 全部 ${aggsAll.length.toLocaleString()} 項${filterHint} / 顯示 ${limited.length.toLocaleString()} 項`;

        if (limited.length === 0) {
            tbody.innerHTML = `<tr><td colspan="11" style="text-align:center;color:var(--color-text-muted);padding:24px">無符合條件的商品</td></tr>`;
        } else {
            tbody.innerHTML = limited.map((p) => {
                const rank = rankMap.get(p.productName) ?? 0;
                const isExpanded = expanded.has(p.productName);
                const chevron = isExpanded ? '▼' : '▶';
                const warn = p.allCostUnset && p.amount > 0
                    ? `<span class="analytics-warn-pill" title="該商品未填成本，毛利視為 0">⚠</span>`
                    : '';
                const rangeText = p.minPrice > 0 && p.maxPrice > 0
                    ? `${p.minPrice.toFixed(1)} ～ ${p.maxPrice.toFixed(1)}`
                    : '—';
                const cvText = p.cv > 0 ? `${p.cv.toFixed(1)}%` : '—';
                const gapClass = p.priceGap > 0 ? 'analytics-price-variance-gap' : '';
                const mainRow = `
            <tr class="analytics-price-variance-mainrow ${isExpanded ? 'is-expanded' : ''}"
                data-product-name="${escapeHtml(p.productName)}">
              <td class="is-right">${rank}</td>
              <td>
                <button type="button" class="analytics-price-variance-toggle"
                        data-role="toggle"
                        aria-expanded="${isExpanded}"
                        title="${isExpanded ? '收合' : '展開客戶明細'}">
                  <span class="analytics-price-variance-chevron">${chevron}</span>
                  ${escapeHtml(p.productName)}${warn}
                </button>
              </td>
              <td>${escapeHtml(p.category)}</td>
              <td class="is-right">${fmt(p.customers.length)}</td>
              <td class="is-right">${fmt(p.count)}</td>
              <td class="is-right">${fmtAvgPrice(p.overallAvg)}</td>
              <td class="is-right">${rangeText}</td>
              <td class="is-right">${cvText}</td>
              <td class="is-right">${p.lowCount > 0 ? fmt(p.lowCount) : '—'}</td>
              <td class="is-right ${gapClass}">${p.priceGap > 0 ? fmtMoney(p.priceGap) : '—'}</td>
              <td class="is-right">${fmtMoney(p.amount)}</td>
            </tr>
          `;
                const subRows = isExpanded ? renderCustomerSubRows(p) : '';
                return mainRow + subRows;
            }).join('');
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

        // 主列展開 / 篩選
        tbody.querySelectorAll<HTMLTableRowElement>('tr.analytics-price-variance-mainrow').forEach((tr) => {
            const productName = tr.dataset.productName ?? '';
            if (!productName) return;
            const toggleBtn = tr.querySelector<HTMLButtonElement>('[data-role="toggle"]');
            toggleBtn?.addEventListener('click', (e) => {
                e.stopPropagation();
                if (expanded.has(productName)) expanded.delete(productName);
                else expanded.add(productName);
                render();
            });
            if (options.onProductFilter) {
                tr.addEventListener('click', (e) => {
                    if ((e.target as HTMLElement).closest('[data-role="toggle"]')) return;
                    options.onProductFilter!(productName);
                });
                tr.classList.add('is-clickable');
                tr.title = '點擊以篩選此商品';
            }
        });

        // 子列點擊客戶
        if (options.onCustomerClick) {
            tbody.querySelectorAll<HTMLTableRowElement>('tr.analytics-price-variance-subrow').forEach((tr) => {
                tr.addEventListener('click', () => {
                    const code = tr.dataset.customerCode ?? '';
                    if (code) options.onCustomerClick!(code);
                });
            });
        }

        // 子表標頭點擊 → 切換子表排序
        tbody.querySelectorAll<HTMLElement>('th[data-customer-sort-key]').forEach((th) => {
            th.addEventListener('click', (e) => {
                e.stopPropagation();
                const k = th.dataset.customerSortKey as CustomerSortKey;
                if (customerSortKey === k) {
                    customerSortDir = customerSortDir === 'asc' ? 'desc' : 'asc';
                } else {
                    customerSortKey = k;
                    // 文字欄升冪、數值欄降冪；deviation 預設升冪（最低價在前）
                    customerSortDir = (k === 'customer' || k === 'line' || k === 'deviation') ? 'asc' : 'desc';
                }
                render();
            });
        });
    };

    root.querySelectorAll<HTMLElement>('th[data-sort-key]').forEach((th) => {
        th.addEventListener('click', () => {
            const k = th.dataset.sortKey as SortKey;
            if (sortKey === k) {
                sortDir = sortDir === 'asc' ? 'desc' : 'asc';
            } else {
                sortKey = k;
                // 數值欄預設降冪（看最高），文字欄升冪
                sortDir = (k === 'product') ? 'asc' : 'desc';
            }
            render();
        });
    });

    thresholdSelect.value = String(threshold);
    thresholdSelect.addEventListener('change', () => {
        threshold = Number(thresholdSelect.value);
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
