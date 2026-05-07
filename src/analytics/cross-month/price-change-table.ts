/**
 * C1 商品漲價影響表：列出兩月都出現、加權均價有變動的商品。
 *
 * 預設依「Δ均價%」絕對值由大到小排序，可點欄位切換排序。
 * 漲價列以紅色 chip 標示、降價以綠色 chip 標示。
 */

import type {ProductPriceChange} from './month-aggregators';

type SortKey =
    | 'product'
    | 'priceChangePct'
    | 'priceChange'
    | 'currentAvgPrice'
    | 'countChangePct'
    | 'amountChange'
    | 'amountChangePct'
    | 'profitChange'
    | 'profitChangePct';
type SortDir = 'asc' | 'desc';

export interface PriceChangeTableOptions {
    onProductClick?: (productName: string) => void;
}

interface Controller {
    element: HTMLElement;

    setData(rows: ReadonlyArray<ProductPriceChange>, currentLabel: string, previousLabel: string): void;
}

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');
const fmtPrice = (v: number) => v.toFixed(2);
const fmtPct = (v: number | null) => (v === null ? '—' : `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`);
const fmtMoneyDelta = (v: number) => `${v >= 0 ? '+' : ''}${fmtMoney(Math.round(v))}`;

export function createPriceChangeTable(opts: PriceChangeTableOptions = {}): Controller {
    const root = document.createElement('div');
    root.className = 'analytics-detail cross-month-price-change';
    root.innerHTML = `
    <div class="analytics-detail-toolbar cross-month-price-change-toolbar">
      <label class="analytics-least-profit-field">
        <span>最小變動 %</span>
        <input type="number" min="0" step="0.5" value="1" data-role="min-pct" placeholder="1">
      </label>
      <label class="analytics-least-profit-field">
        <span>方向</span>
        <select data-role="direction">
          <option value="all">全部</option>
          <option value="up">僅漲價</option>
          <option value="down">僅降價</option>
        </select>
      </label>
      <span class="analytics-detail-stats" data-role="stats"></span>
    </div>
    <div class="analytics-detail-table-wrap">
      <table class="analytics-detail-table-el">
        <thead>
          <tr>
            <th data-sort-key="product">商品<span class="sort-indicator"></span></th>
            <th class="is-right" data-role="th-prev-price"></th>
            <th class="is-right" data-role="th-curr-price"></th>
            <th data-sort-key="priceChange" class="is-right">Δ均價<span class="sort-indicator"></span></th>
            <th data-sort-key="priceChangePct" class="is-right">Δ均價%<span class="sort-indicator"></span></th>
            <th data-sort-key="countChangePct" class="is-right" title="(本月數量 - 上月數量) / 上月數量">Δ數量%<span class="sort-indicator"></span></th>
            <th data-sort-key="amountChange" class="is-right">Δ營收<span class="sort-indicator"></span></th>
            <th data-sort-key="amountChangePct" class="is-right">Δ營收%<span class="sort-indicator"></span></th>
            <th data-sort-key="profitChange" class="is-right">Δ毛利<span class="sort-indicator"></span></th>
            <th data-sort-key="profitChangePct" class="is-right">Δ毛利%<span class="sort-indicator"></span></th>
          </tr>
        </thead>
        <tbody data-role="tbody"></tbody>
      </table>
    </div>
  `;

    const tbody = root.querySelector<HTMLElement>('[data-role="tbody"]')!;
    const stats = root.querySelector<HTMLElement>('[data-role="stats"]')!;
    const minPctInput = root.querySelector<HTMLInputElement>('[data-role="min-pct"]')!;
    const directionSelect = root.querySelector<HTMLSelectElement>('[data-role="direction"]')!;
    const thPrev = root.querySelector<HTMLElement>('[data-role="th-prev-price"]')!;
    const thCurr = root.querySelector<HTMLElement>('[data-role="th-curr-price"]')!;

    let raw: ReadonlyArray<ProductPriceChange> = [];
    let currentLabel = '';
    let previousLabel = '';
    let sortKey: SortKey = 'priceChangePct';
    let sortDir: SortDir = 'desc';

    const valueOf = (r: ProductPriceChange, k: SortKey): number | string => {
        switch (k) {
            case 'product':
                return r.productName;
            case 'priceChange':
                return r.priceChange;
            case 'priceChangePct':
                return Math.abs(r.priceChangePct ?? 0);
            case 'currentAvgPrice':
                return r.currentAvgPrice;
            case 'countChangePct':
                return r.countChangePct ?? 0;
            case 'amountChange':
                return r.amountChange;
            case 'amountChangePct':
                return r.amountChangePct ?? 0;
            case 'profitChange':
                return r.profitChange;
            case 'profitChangePct':
                return r.profitChangePct ?? 0;
        }
    };

    const filterAndSort = (): ProductPriceChange[] => {
        const minPct = Math.max(0, Number(minPctInput.value) || 0);
        const direction = directionSelect.value as 'all' | 'up' | 'down';
        const filtered = raw.filter((r) => {
            const pct = r.priceChangePct ?? 0;
            if (Math.abs(pct) < minPct) return false;
            if (direction === 'up' && pct <= 0) return false;
            if (direction === 'down' && pct >= 0) return false;
            return true;
        });
        const dir = sortDir === 'asc' ? 1 : -1;
        filtered.sort((a, b) => {
            const av = valueOf(a, sortKey);
            const bv = valueOf(b, sortKey);
            if (typeof av === 'string' && typeof bv === 'string') {
                return av.localeCompare(bv) * dir;
            }
            return ((Number(av) - Number(bv))) * dir;
        });
        return filtered;
    };

    const renderRows = () => {
        const rows = filterAndSort();
        stats.textContent = rows.length === 0
            ? `共 ${raw.length} 項商品兩月都有，當前條件下沒有符合的`
            : `共 ${rows.length} / ${raw.length} 項`;
        if (rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="10" class="cross-month-empty">無資料</td></tr>`;
            return;
        }
        tbody.innerHTML = rows
            .map((r) => {
                const priceUp = r.priceChange > 0;
                const priceDown = r.priceChange < 0;
                const priceCellClass = priceUp ? 'cell-price-up' : priceDown ? 'cell-price-down' : '';
                const priceTagClass = priceUp ? 'is-up' : priceDown ? 'is-down' : '';
                const priceArrow = priceUp ? '↑' : priceDown ? '↓' : '·';
                const distinctNotice = r.currentDistinctPrices > 1 || r.prevDistinctPrices > 1
                    ? `<span class="cross-month-mixed-prices" title="該商品在不同客戶間有 ${r.prevDistinctPrices}→${r.currentDistinctPrices} 種不同單價（顯示為加權均價）">混價</span>`
                    : '';
                return `
        <tr class="${opts.onProductClick ? 'is-clickable' : ''}" data-name="${escapeHtml(r.productName)}">
          <td>
            <span class="cross-month-price-arrow ${priceTagClass}">${priceArrow}</span>
            ${escapeHtml(r.productName)}
            ${distinctNotice}
          </td>
          <td class="is-right">${fmtPrice(r.prevAvgPrice)}</td>
          <td class="is-right ${priceCellClass}">${fmtPrice(r.currentAvgPrice)}</td>
          <td class="is-right ${priceCellClass}">${fmtMoneyDelta(r.priceChange)}</td>
          <td class="is-right ${priceCellClass}">${fmtPct(r.priceChangePct)}</td>
          <td class="is-right ${classOfDelta(r.countChange)}">${fmtPct(r.countChangePct)}</td>
          <td class="is-right ${classOfDelta(r.amountChange)}">${fmtMoneyDelta(r.amountChange)}</td>
          <td class="is-right ${classOfDelta(r.amountChange)}">${fmtPct(r.amountChangePct)}</td>
          <td class="is-right ${classOfDelta(r.profitChange)}">${fmtMoneyDelta(r.profitChange)}</td>
          <td class="is-right ${classOfDelta(r.profitChange)}">${fmtPct(r.profitChangePct)}</td>
        </tr>
      `;
            })
            .join('');
        attachRowClick();
    };

    const updateHeaderLabels = () => {
        thPrev.textContent = `${previousLabel} 均價`;
        thCurr.textContent = `${currentLabel} 均價`;
    };

    const updateSortIndicators = () => {
        root.querySelectorAll<HTMLElement>('th[data-sort-key]').forEach((th) => {
            const k = th.dataset.sortKey;
            const indicator = th.querySelector<HTMLElement>('.sort-indicator');
            if (!indicator) return;
            indicator.textContent = k === sortKey ? (sortDir === 'asc' ? ' ▲' : ' ▼') : '';
        });
    };

    root.querySelectorAll<HTMLElement>('th[data-sort-key]').forEach((th) => {
        th.addEventListener('click', () => {
            const k = th.dataset.sortKey as SortKey | undefined;
            if (!k) return;
            if (k === sortKey) {
                sortDir = sortDir === 'asc' ? 'desc' : 'asc';
            } else {
                sortKey = k;
                sortDir = (k === 'product') ? 'asc' : 'desc';
            }
            updateSortIndicators();
            renderRows();
        });
    });

    const attachRowClick = () => {
        if (!opts.onProductClick) return;
        tbody.querySelectorAll<HTMLElement>('tr[data-name]').forEach((tr) => {
            tr.addEventListener('click', () => {
                const name = tr.dataset.name;
                if (name) opts.onProductClick!(name);
            });
        });
    };

    minPctInput.addEventListener('input', renderRows);
    directionSelect.addEventListener('change', renderRows);

    updateSortIndicators();

    return {
        element: root,
        setData(rows, current, previous) {
            raw = rows;
            currentLabel = current;
            previousLabel = previous;
            updateHeaderLabels();
            renderRows();
        },
    };
}

function classOfDelta(v: number): string {
    if (v > 0) return 'cell-positive';
    if (v < 0) return 'cell-negative';
    return '';
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
