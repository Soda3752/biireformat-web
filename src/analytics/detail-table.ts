/**
 * 明細表元件：可排序、可文字搜尋、可分頁、可匯出 CSV。
 * 純 DOM 實作，不引入 grid library。
 */

import type {AnalyticsRow} from './dataset-builder';

type SortKey = 'day' | 'customerName' | 'line' | 'productName' | 'category' | 'count' | 'price' | 'amount';
type SortDir = 'asc' | 'desc';

const PAGE_SIZE = 50;

const COLUMNS: Array<{ key: SortKey; label: string; align?: 'right' }> = [
    {key: 'day', label: '日'},
    {key: 'customerName', label: '客戶'},
    {key: 'line', label: '線別'},
    {key: 'productName', label: '商品'},
    {key: 'category', label: '分類'},
    {key: 'count', label: '數量', align: 'right'},
    {key: 'price', label: '單價', align: 'right'},
    {key: 'amount', label: '金額', align: 'right'},
];

export interface DetailTableController {
    element: HTMLElement;

    setRows(rows: ReadonlyArray<AnalyticsRow>): void;

    /** 取得目前篩選+排序後可見的列（給 CSV 匯出用） */
    getCurrentRows(): AnalyticsRow[];

    scrollIntoView(): void;
}

export function createDetailTable(): DetailTableController {
    const root = document.createElement('div');
    root.className = 'analytics-detail';
    root.innerHTML = `
    <div class="analytics-detail-toolbar">
      <input type="search" class="analytics-detail-search" placeholder="搜尋客戶或商品名稱…" data-role="search">
      <span class="analytics-detail-stats" data-role="stats"></span>
    </div>
    <div class="analytics-detail-table-wrap">
      <table class="analytics-detail-table-el">
        <thead>
          <tr>
            ${COLUMNS.map(
        (c) =>
            `<th data-sort-key="${c.key}" class="${c.align === 'right' ? 'is-right' : ''}">${c.label}<span class="sort-indicator"></span></th>`
    ).join('')}
          </tr>
        </thead>
        <tbody data-role="tbody"></tbody>
      </table>
    </div>
    <div class="analytics-detail-pager" data-role="pager"></div>
  `;

    const searchInput = root.querySelector<HTMLInputElement>('[data-role="search"]')!;
    const statsEl = root.querySelector<HTMLElement>('[data-role="stats"]')!;
    const tbody = root.querySelector<HTMLTableSectionElement>('[data-role="tbody"]')!;
    const pager = root.querySelector<HTMLElement>('[data-role="pager"]')!;

    let allRows: ReadonlyArray<AnalyticsRow> = [];
    let sortKey: SortKey = 'day';
    let sortDir: SortDir = 'asc';
    let page = 1;
    let searchText = '';

    const computeFiltered = (): AnalyticsRow[] => {
        let result = allRows.filter((r) => {
            if (!searchText) return true;
            return r.customerName.includes(searchText) || r.productName.includes(searchText);
        });
        result = [...result].sort((a, b) => {
            const av = a[sortKey];
            const bv = b[sortKey];
            let cmp = 0;
            if (typeof av === 'number' && typeof bv === 'number') cmp = av - bv;
            else cmp = String(av).localeCompare(String(bv));
            return sortDir === 'asc' ? cmp : -cmp;
        });
        return result;
    };

    const fmt = (v: number) => v.toLocaleString('zh-TW');

    const render = () => {
        const filtered = computeFiltered();
        const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
        if (page > totalPages) page = totalPages;
        const start = (page - 1) * PAGE_SIZE;
        const visible = filtered.slice(start, start + PAGE_SIZE);

        statsEl.textContent = `共 ${filtered.length.toLocaleString()} 筆 / 第 ${page} / ${totalPages} 頁`;

        tbody.innerHTML = visible
            .map(
                (r) => `
          <tr>
            <td>${r.day}</td>
            <td>${escapeHtml(r.customerName)}</td>
            <td>${escapeHtml(r.line)}</td>
            <td>${escapeHtml(r.productName)}</td>
            <td>${escapeHtml(r.category)}</td>
            <td class="is-right">${fmt(r.count)}</td>
            <td class="is-right">${fmt(r.price)}</td>
            <td class="is-right">${fmt(r.amount)}</td>
          </tr>
        `
            )
            .join('');

        // 排序指示器
        root.querySelectorAll<HTMLElement>('th[data-sort-key]').forEach((th) => {
            const ind = th.querySelector<HTMLElement>('.sort-indicator')!;
            if (th.dataset.sortKey === sortKey) {
                ind.textContent = sortDir === 'asc' ? ' ▲' : ' ▼';
            } else {
                ind.textContent = '';
            }
        });

        // 分頁按鈕
        pager.innerHTML = '';
        if (totalPages > 1) {
            const mkBtn = (label: string, target: number, disabled = false) => {
                const b = document.createElement('button');
                b.type = 'button';
                b.textContent = label;
                b.disabled = disabled;
                b.className = 'analytics-detail-pager-btn';
                b.addEventListener('click', () => {
                    page = target;
                    render();
                });
                pager.appendChild(b);
            };
            mkBtn('« 首頁', 1, page === 1);
            mkBtn('上一頁', Math.max(1, page - 1), page === 1);
            const span = document.createElement('span');
            span.className = 'analytics-detail-pager-info';
            span.textContent = ` ${page} / ${totalPages} `;
            pager.appendChild(span);
            mkBtn('下一頁', Math.min(totalPages, page + 1), page === totalPages);
            mkBtn('末頁 »', totalPages, page === totalPages);
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
                sortDir = 'asc';
            }
            render();
        });
    });

    // 搜尋事件
    let searchTimer: number | null = null;
    searchInput.addEventListener('input', () => {
        if (searchTimer) window.clearTimeout(searchTimer);
        searchTimer = window.setTimeout(() => {
            searchText = searchInput.value.trim();
            page = 1;
            render();
        }, 200);
    });

    return {
        element: root,
        setRows(rows) {
            allRows = rows;
            page = 1;
            render();
        },
        getCurrentRows() {
            return computeFiltered();
        },
        scrollIntoView() {
            root.scrollIntoView({behavior: 'smooth', block: 'start'});
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

/* ================ CSV 匯出 ================ */

export function rowsToCsv(rows: ReadonlyArray<AnalyticsRow>): string {
    const headers = ['檔案', '民國年', '月', '日', '客戶代碼', '客戶名稱', '線別', '商品', '分類', '數量', '單價', '金額', '月結', '含稅', '現金'];
    const escape = (v: string | number | boolean) => {
        const s = String(v);
        if (/[",\n]/.test(s)) return `"${s.replaceAll('"', '""')}"`;
        return s;
    };
    const lines = [headers.join(',')];
    for (const r of rows) {
        lines.push(
            [
                r.fileName,
                r.year,
                r.month,
                r.day,
                r.customerCode,
                r.customerName,
                r.line,
                r.productName,
                r.category,
                r.count,
                r.price,
                r.amount,
                r.isMonthly ? 'Y' : '',
                r.isNeedTex ? 'Y' : '',
                r.isCashUser ? 'Y' : '',
            ]
                .map(escape)
                .join(',')
        );
    }
    // UTF-8 BOM 確保 Excel 開啟中文不亂碼
    return '﻿' + lines.join('\n');
}
