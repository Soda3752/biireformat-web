/**
 * 各路線營收排行：以「線別」聚合營收（amount），用水平長條由高到低排名呈現，
 * 每條標示金額與佔比，最上方顯示總營收（含稅）。可切換升冪／降冪排序，
 * 並提供「只列印本區塊」的按鈕。與上方全域篩選器連動（setRows 傳入篩選後 rows）。
 *
 * 名次固定以營收高低計算（最高 = 第 1 名），排序切換只改變列的呈現順序，
 * 故名次徽章的意義不隨排序方向改變。
 */

import type {AnalyticsRow} from './dataset-builder';
import {groupBy, sumAmountWithTax} from './aggregators';
import {icon} from '@/ui/icons';

export interface RouteRevenueController {
    element: HTMLElement;

    setRows(rows: ReadonlyArray<AnalyticsRow>): void;
}

export function createRouteRevenueSection(): RouteRevenueController {
    let rows: ReadonlyArray<AnalyticsRow> = [];
    let sortDesc = true;

    const root = document.createElement('section');
    root.className = 'analytics-route-revenue';
    root.innerHTML = `
      <div class="analytics-route-revenue-head">
        <div class="analytics-route-revenue-titles">
          <div class="analytics-section-label">各路線營收排行</div>
          <div class="analytics-route-revenue-total">
            <span class="analytics-route-revenue-total-label">總營收</span>
            <span class="analytics-route-revenue-total-value" data-role="total-value">0</span>
            <span class="analytics-route-revenue-total-hint" data-role="total-hint"></span>
          </div>
        </div>
        <div class="analytics-route-revenue-actions">
          <div class="analytics-route-revenue-sort">
            <button type="button" class="is-active" data-role="sort-desc">由高到低</button>
            <button type="button" data-role="sort-asc">由低到高</button>
          </div>
          <button type="button" class="btn btn-secondary" data-role="print">
            ${icon('printer', 16)} 列印
          </button>
        </div>
      </div>
      <div class="analytics-route-revenue-list" data-role="list"></div>
    `;

    const totalValueEl = root.querySelector<HTMLElement>('[data-role="total-value"]')!;
    const totalHintEl = root.querySelector<HTMLElement>('[data-role="total-hint"]')!;
    const listEl = root.querySelector<HTMLElement>('[data-role="list"]')!;
    const sortDescBtn = root.querySelector<HTMLButtonElement>('[data-role="sort-desc"]')!;
    const sortAscBtn = root.querySelector<HTMLButtonElement>('[data-role="sort-asc"]')!;
    const printBtn = root.querySelector<HTMLButtonElement>('[data-role="print"]')!;

    const fmt = (v: number) => Math.round(v).toLocaleString('zh-TW');

    const render = () => {
        // 一律以營收高到低分組，用於計算名次與最大值（長條相對長度基準）。
        const groups = groupBy(rows, 'line', 'amount', true);
        const total = groups.reduce((s, g) => s + g.amount, 0);

        totalValueEl.textContent = fmt(total);
        totalHintEl.textContent = groups.length > 0 ? `含稅 ${fmt(sumAmountWithTax(rows))}` : '';

        if (groups.length === 0) {
            listEl.innerHTML = '<div class="analytics-route-revenue-empty">無資料</div>';
            return;
        }

        const maxAmount = groups[0].amount || 1;
        const rankOf = new Map<string, number>();
        groups.forEach((g, i) => rankOf.set(g.key, i + 1));

        const display = sortDesc ? groups : [...groups].reverse();
        listEl.innerHTML = display
            .map((g) => {
                const pct = total > 0 ? (g.amount / total) * 100 : 0;
                const width = (g.amount / maxAmount) * 100;
                return `
                  <div class="analytics-route-bar-row">
                    <div class="analytics-route-bar-rank">${rankOf.get(g.key)}</div>
                    <div class="analytics-route-bar-name">${escapeHtml(g.key)}</div>
                    <div class="analytics-route-bar-track">
                      <div class="analytics-route-bar-fill" style="width:${width.toFixed(1)}%"></div>
                    </div>
                    <div class="analytics-route-bar-amount">${fmt(g.amount)}</div>
                    <div class="analytics-route-bar-pct">${pct.toFixed(1)}%</div>
                  </div>`;
            })
            .join('');
    };

    const setSort = (desc: boolean) => {
        if (desc === sortDesc) return;
        sortDesc = desc;
        sortDescBtn.classList.toggle('is-active', desc);
        sortAscBtn.classList.toggle('is-active', !desc);
        render();
    };
    sortDescBtn.addEventListener('click', () => setSort(true));
    sortAscBtn.addEventListener('click', () => setSort(false));

    // afterprint 監聽只註冊一次負責復原（列印結束或取消都會觸發），
    // 避免每次點擊都新增監聽器，也確保 body class 不會卡住讓整頁空白。
    window.addEventListener('afterprint', () => {
        document.body.classList.remove('printing-route-revenue');
    });
    printBtn.addEventListener('click', () => {
        document.body.classList.add('printing-route-revenue');
        window.print();
    });

    return {
        element: root,
        setRows(next) {
            rows = next;
            render();
        },
    };
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;');
}
