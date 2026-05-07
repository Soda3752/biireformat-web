/**
 * B1 客戶分群表：新增 / 流失 / 留存 三個分頁切換，
 * 共用同一張表骨架，欄位略有差異：
 *   - 新增 / 流失：客戶 / 線別 / 金額 / 數量 / 毛利
 *   - 留存：客戶 / 線別 / 上月金額 / 本月金額 / Δ金額 / Δ% / Δ數量 / Δ%
 */

import type {CustomerMonthStat, CustomerSegmentation, RetainedCustomer,} from './month-aggregators';

type SegmentKind = 'new' | 'churned' | 'retained';

interface Controller {
    element: HTMLElement;

    setData(seg: CustomerSegmentation, currentLabel: string, previousLabel: string): void;
}

export interface CustomerSegmentTableOptions {
    onCustomerClick?: (customerCode: string) => void;
}

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');
const fmtCount = (v: number) => v.toLocaleString('zh-TW');
const fmtPct = (v: number | null) => (v === null ? '—' : `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`);
const fmtDelta = (v: number) => `${v >= 0 ? '+' : ''}${fmtMoney(Math.round(v))}`;

export function createCustomerSegmentTable(opts: CustomerSegmentTableOptions = {}): Controller {
    const root = document.createElement('div');
    root.className = 'analytics-detail cross-month-segment-table';
    root.innerHTML = `
    <div class="cross-month-segment-tabs" data-role="tabs">
      <button type="button" class="cross-month-segment-tab is-active" data-kind="new">
        <span class="cross-month-segment-tab-label">新增客戶</span>
        <span class="cross-month-segment-tab-count" data-role="count-new">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="churned">
        <span class="cross-month-segment-tab-label">流失客戶</span>
        <span class="cross-month-segment-tab-count" data-role="count-churned">0</span>
      </button>
      <button type="button" class="cross-month-segment-tab" data-kind="retained">
        <span class="cross-month-segment-tab-label">留存客戶</span>
        <span class="cross-month-segment-tab-count" data-role="count-retained">0</span>
      </button>
    </div>
    <div class="analytics-detail-table-wrap">
      <table class="analytics-detail-table-el cross-month-segment-tableel">
        <thead data-role="thead"></thead>
        <tbody data-role="tbody"></tbody>
      </table>
    </div>
  `;

    const thead = root.querySelector<HTMLElement>('[data-role="thead"]')!;
    const tbody = root.querySelector<HTMLElement>('[data-role="tbody"]')!;
    const tabHost = root.querySelector<HTMLElement>('[data-role="tabs"]')!;
    const countNew = root.querySelector<HTMLElement>('[data-role="count-new"]')!;
    const countChurn = root.querySelector<HTMLElement>('[data-role="count-churned"]')!;
    const countRetain = root.querySelector<HTMLElement>('[data-role="count-retained"]')!;

    let activeKind: SegmentKind = 'new';
    let lastSeg: CustomerSegmentation | null = null;
    let lastCurrentLabel = '';
    let lastPreviousLabel = '';

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

    const renderTable = () => {
        if (!lastSeg) {
            thead.innerHTML = '';
            tbody.innerHTML = '';
            return;
        }
        if (activeKind === 'new' || activeKind === 'churned') {
            renderSimpleTable(activeKind === 'new' ? lastSeg.newCustomers : lastSeg.churnedCustomers);
        } else {
            renderRetainedTable(lastSeg.retainedCustomers);
        }
    };

    const renderSimpleTable = (rows: ReadonlyArray<CustomerMonthStat>) => {
        const monthLabel = activeKind === 'new' ? lastCurrentLabel : lastPreviousLabel;
        thead.innerHTML = `
      <tr>
        <th>客戶</th>
        <th>線別</th>
        <th class="is-right">${monthLabel}　金額</th>
        <th class="is-right">${monthLabel}　數量</th>
        <th class="is-right">${monthLabel}　毛利</th>
      </tr>
    `;
        if (rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" class="cross-month-empty">無資料</td></tr>`;
            return;
        }
        tbody.innerHTML = rows
            .map((r) => `
        <tr class="${opts.onCustomerClick ? 'is-clickable' : ''}" data-code="${escapeHtml(r.customerCode)}">
          <td>${escapeHtml(r.customerName)}<span class="cross-month-cell-sub">(${escapeHtml(r.customerCode)})</span></td>
          <td>${escapeHtml(r.line)}</td>
          <td class="is-right">${fmtMoney(Math.round(r.amount))}</td>
          <td class="is-right">${fmtCount(r.count)}</td>
          <td class="is-right ${r.profit < 0 ? 'cell-negative' : ''}">${fmtMoney(Math.round(r.profit))}</td>
        </tr>
      `)
            .join('');
        attachRowClick();
    };

    const renderRetainedTable = (rows: ReadonlyArray<RetainedCustomer>) => {
        thead.innerHTML = `
      <tr>
        <th>客戶</th>
        <th>線別</th>
        <th class="is-right">${lastPreviousLabel} 金額</th>
        <th class="is-right">${lastCurrentLabel} 金額</th>
        <th class="is-right" title="本月 - 上月">Δ金額</th>
        <th class="is-right">Δ%</th>
        <th class="is-right" title="本月 - 上月">Δ數量</th>
        <th class="is-right">Δ數量%</th>
      </tr>
    `;
        if (rows.length === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="cross-month-empty">無資料</td></tr>`;
            return;
        }
        tbody.innerHTML = rows
            .map((r) => {
                const upClass = r.amountDelta >= 0 ? 'cell-positive' : 'cell-negative';
                const cntClass = r.countDelta >= 0 ? 'cell-positive' : 'cell-negative';
                return `
        <tr class="${opts.onCustomerClick ? 'is-clickable' : ''}" data-code="${escapeHtml(r.customerCode)}">
          <td>${escapeHtml(r.customerName)}<span class="cross-month-cell-sub">(${escapeHtml(r.customerCode)})</span></td>
          <td>${escapeHtml(r.line)}</td>
          <td class="is-right">${fmtMoney(Math.round(r.prevAmount))}</td>
          <td class="is-right">${fmtMoney(Math.round(r.amount))}</td>
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
            tr.addEventListener('click', () => {
                const code = tr.dataset.code;
                if (code) opts.onCustomerClick!(code);
            });
        });
    };

    return {
        element: root,
        setData(seg, currentLabel, previousLabel) {
            lastSeg = seg;
            lastCurrentLabel = currentLabel;
            lastPreviousLabel = previousLabel;
            countNew.textContent = String(seg.newCustomers.length);
            countChurn.textContent = String(seg.churnedCustomers.length);
            countRetain.textContent = String(seg.retainedCustomers.length);
            renderTable();
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
