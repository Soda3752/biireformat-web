/**
 * 對帳 2.0 預覽元件
 *
 * 由兩個區塊組成：
 *  1. 主表（客戶 / 線別 / 模式 / 應收 / 已收 / 差額 / 狀態）
 *  2. 「需人工複核」警示區塊（多重匹配 / 未配對 / 未設 storeCode），未設 storeCode 列提供「補建」按鈕
 *
 * 統計列（總筆數 / 已配對 / 未收 / 部分 / 超收 / 待覆核）由呼叫端組裝，本元件僅負責資料表渲染。
 */

import type {
    CustomerReconcileRow,
    ManualReviewCandidate,
    ManualReviewItem,
    ReconcileResult,
    ReconcileStatus,
} from '@/domain/bank-reconcile-service';
import type {BankRowMatch} from '@/domain/bank-match-service';

export interface BankReconcileTableOptions {
    onAddClick: (row: BankRowMatch) => void;
    onCandidateClick: (candidate: ManualReviewCandidate, item: ManualReviewItem) => void;
}

export interface BankReconcileTableHandle {
    readonly element: HTMLElement;

    setData(result: ReconcileResult): void;

    clear(): void;
}

type TabKey = 'matched' | 'unmatched' | 'manual';

const TAB_DEFS: ReadonlyArray<{ key: TabKey; label: string }> = [
    {key: 'matched', label: '已匹配'},
    {key: 'unmatched', label: '無匹配'},
    {key: 'manual', label: '需人工複核'},
];

const DEFAULT_TAB: TabKey = 'matched';

/**
 * 銀行 CSV 預期欄位順序（與 trans-record-reader / bank-name-writer 對齊）。
 * 當原始 CSV 缺 header 時做為 tooltip 欄位名稱備援。
 */
const DEFAULT_BANK_HEADER: ReadonlyArray<string> = [
    '日期', '摘要', '幣別', '支出金額', '存入金額', '餘額', '備註', '轉出入帳號', '註記',
];

export function createBankReconcileTable(options: BankReconcileTableOptions): BankReconcileTableHandle {
    const element = document.createElement('section');
    element.className = 'reconcile-preview';
    element.hidden = true;

    element.innerHTML = `
      <div class="reconcile-summary" data-role="summary"></div>

      <div class="reconcile-tabs" role="tablist" data-role="tabs">
        ${TAB_DEFS.map((t) => `
          <button type="button" class="reconcile-tab" role="tab" data-tab="${t.key}" aria-selected="false">
            <span class="reconcile-tab-label">${t.label}</span>
            <span class="reconcile-tab-badge" data-role="badge-${t.key}">0</span>
          </button>
        `).join('')}
      </div>

      <div class="reconcile-tab-panel" data-role="panel-customer" role="tabpanel">
        <div class="reconcile-search-bar">
          <input type="search" class="reconcile-search-input" data-role="search-customer"
                 placeholder="搜尋客戶編號或名稱..." autocomplete="off">
          <button type="button" class="reconcile-search-clear" data-role="search-customer-clear" aria-label="清除搜尋" hidden>✕</button>
        </div>
        <div class="reconcile-scroll">
          <table class="reconcile-table">
            <thead>
              <tr>
                <th class="col-code">編號</th>
                <th class="col-name">客戶</th>
                <th class="col-line">線別</th>
                <th class="col-mode">模式</th>
                <th class="col-money">應收</th>
                <th class="col-money">已收</th>
                <th class="col-money">差額</th>
                <th class="col-status">狀態</th>
                <th class="col-receipts">匯款詳情</th>
              </tr>
            </thead>
            <tbody data-role="customer-tbody"></tbody>
          </table>
        </div>
      </div>

      <div class="reconcile-tab-panel" data-role="panel-manual" role="tabpanel" hidden>
        <div class="reconcile-search-bar">
          <input type="search" class="reconcile-search-input" data-role="search-manual"
                 placeholder="搜尋客戶編號或名稱..." autocomplete="off">
          <button type="button" class="reconcile-search-clear" data-role="search-manual-clear" aria-label="清除搜尋" hidden>✕</button>
        </div>
        <div class="reconcile-scroll">
          <table class="reconcile-table reconcile-manual-table">
            <thead>
              <tr>
                <th class="col-type">類型</th>
                <th class="col-receipts">匯款詳情</th>
                <th class="col-summary">摘要</th>
                <th class="col-money">存入</th>
                <th class="col-candidate">配對候選</th>
                <th class="col-action">操作</th>
              </tr>
            </thead>
            <tbody data-role="manual-tbody"></tbody>
          </table>
        </div>
      </div>
    `;

    const summaryEl = element.querySelector<HTMLElement>('[data-role="summary"]')!;
    const tabsEl = element.querySelector<HTMLElement>('[data-role="tabs"]')!;
    const customerPanel = element.querySelector<HTMLElement>('[data-role="panel-customer"]')!;
    const manualPanel = element.querySelector<HTMLElement>('[data-role="panel-manual"]')!;
    const customerTbody = element.querySelector<HTMLElement>('[data-role="customer-tbody"]')!;
    const manualTbody = element.querySelector<HTMLElement>('[data-role="manual-tbody"]')!;
    const customerSearchInput = element.querySelector<HTMLInputElement>('[data-role="search-customer"]')!;
    const manualSearchInput = element.querySelector<HTMLInputElement>('[data-role="search-manual"]')!;
    const customerSearchClear = element.querySelector<HTMLButtonElement>('[data-role="search-customer-clear"]')!;
    const manualSearchClear = element.querySelector<HTMLButtonElement>('[data-role="search-manual-clear"]')!;

    let activeTab: TabKey = DEFAULT_TAB;
    let lastResult: ReconcileResult | null = null;
    let lastHeader: ReadonlyArray<string> = DEFAULT_BANK_HEADER;
    const queries: Record<TabKey, string> = {matched: '', unmatched: '', manual: ''};

    const setActiveTab = (next: TabKey) => {
        activeTab = next;
        refreshTabState();
        refreshPanels();
    };

    const refreshTabState = () => {
        for (const btn of tabsEl.querySelectorAll<HTMLButtonElement>('.reconcile-tab')) {
            const key = btn.dataset.tab as TabKey;
            const isActive = key === activeTab;
            btn.classList.toggle('is-active', isActive);
            btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
        }
    };

    const refreshBadges = () => {
        if (!lastResult) return;
        const matchedRows = lastResult.customers.filter((c) => c.received > 0);
        const unmatchedRows = lastResult.customers.filter((c) => c.received === 0);
        const counts: Record<TabKey, number> = {
            matched: matchedRows.length,
            unmatched: unmatchedRows.length,
            manual: lastResult.manualReviewItems.length,
        };
        for (const t of TAB_DEFS) {
            const badge = element.querySelector<HTMLElement>(`[data-role="badge-${t.key}"]`);
            if (badge) {
                badge.textContent = String(counts[t.key]);
                badge.classList.toggle('is-zero', counts[t.key] === 0);
                badge.classList.toggle('is-warn', t.key === 'manual' && counts[t.key] > 0);
                badge.classList.toggle('is-danger', t.key === 'unmatched' && counts[t.key] > 0);
            }
        }
    };

    const refreshPanels = () => {
        const showCustomer = activeTab === 'matched' || activeTab === 'unmatched';
        customerPanel.hidden = !showCustomer;
        manualPanel.hidden = activeTab !== 'manual';

        if ((activeTab === 'matched' || activeTab === 'unmatched') && lastResult) {
            const baseRows = activeTab === 'matched'
                ? lastResult.customers.filter((c) => c.received > 0)
                : lastResult.customers.filter((c) => c.received === 0);
            const query = queries[activeTab];
            const rows = filterCustomers(baseRows, query);
            const emptyText = buildCustomerEmptyText(activeTab, baseRows.length, query);
            renderCustomers(customerTbody, rows, lastHeader, emptyText);

            customerSearchInput.value = query;
            customerSearchClear.hidden = query === '';
        }

        if (activeTab === 'manual' && lastResult) {
            const query = queries.manual;
            const items = filterManual(lastResult.manualReviewItems, query);
            renderManual(
                manualTbody, items, lastHeader,
                options.onAddClick, options.onCandidateClick,
                buildManualEmptyText(lastResult.manualReviewItems.length, query),
            );

            manualSearchInput.value = query;
            manualSearchClear.hidden = query === '';
        }
    };

    for (const btn of tabsEl.querySelectorAll<HTMLButtonElement>('.reconcile-tab')) {
        btn.addEventListener('click', () => setActiveTab(btn.dataset.tab as TabKey));
    }

    customerSearchInput.addEventListener('input', () => {
        if (activeTab !== 'matched' && activeTab !== 'unmatched') return;
        queries[activeTab] = customerSearchInput.value;
        refreshPanels();
    });

    manualSearchInput.addEventListener('input', () => {
        queries.manual = manualSearchInput.value;
        refreshPanels();
    });

    customerSearchClear.addEventListener('click', () => {
        if (activeTab !== 'matched' && activeTab !== 'unmatched') return;
        queries[activeTab] = '';
        customerSearchInput.value = '';
        customerSearchInput.focus();
        refreshPanels();
    });

    manualSearchClear.addEventListener('click', () => {
        queries.manual = '';
        manualSearchInput.value = '';
        manualSearchInput.focus();
        refreshPanels();
    });

    const setData = (result: ReconcileResult) => {
        lastResult = result;
        lastHeader = pickHeader(result.bankHeader);
        renderSummary(summaryEl, result);
        refreshBadges();
        refreshTabState();
        refreshPanels();
        element.hidden = false;
    };

    const clear = () => {
        lastResult = null;
        lastHeader = DEFAULT_BANK_HEADER;
        hideRawRowPopover();
        summaryEl.innerHTML = '';
        customerTbody.innerHTML = '';
        manualTbody.innerHTML = '';
        queries.matched = '';
        queries.unmatched = '';
        queries.manual = '';
        customerSearchInput.value = '';
        manualSearchInput.value = '';
        customerSearchClear.hidden = true;
        manualSearchClear.hidden = true;
        for (const t of TAB_DEFS) {
            const badge = element.querySelector<HTMLElement>(`[data-role="badge-${t.key}"]`);
            if (badge) badge.textContent = '0';
        }
        element.hidden = true;
    };

    return {element, setData, clear};
}

function renderSummary(host: HTMLElement, result: ReconcileResult): void {
    const s = result.summary;
    host.innerHTML = `
      <span class="reconcile-chip">客戶 <strong>${s.customerCount}</strong> 間</span>
      <span class="reconcile-chip">應收 <strong>${formatMoney(s.totalReceivable)}</strong></span>
      <span class="reconcile-chip">已收 <strong>${formatMoney(s.totalReceived)}</strong></span>
      <span class="reconcile-chip ${diffChipClass(s.totalDiff)}">差額 <strong>${formatMoney(s.totalDiff)}</strong></span>
    `;
}

function renderCustomers(
    tbody: HTMLElement,
    rows: ReadonlyArray<CustomerReconcileRow>,
    header: ReadonlyArray<string>,
    emptyText = '無資料',
): void {
    tbody.innerHTML = '';
    if (rows.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="reconcile-empty">${emptyText}</td></tr>`;
        return;
    }

    const fragment = document.createDocumentFragment();
    for (const r of rows) {
        const tr = document.createElement('tr');
        tr.className = `reconcile-row status-${r.status}`;
        if (r.isCashUser) tr.classList.add('is-cash');

        tr.appendChild(td('col-code', r.customerCode));
        tr.appendChild(td('col-name', r.customerName));
        tr.appendChild(td('col-line', r.customerLine));
        tr.appendChild(td('col-mode', formatMode(r)));
        tr.appendChild(td('col-money', formatMoney(r.receivable)));
        tr.appendChild(td('col-money', formatMoney(r.received)));

        const diffTd = td('col-money', signedMoney(r.diff));
        if (r.diff < 0) diffTd.classList.add('is-negative');
        else if (r.diff > 0) diffTd.classList.add('is-positive');
        tr.appendChild(diffTd);

        const statusTd = document.createElement('td');
        statusTd.className = 'col-status';
        statusTd.appendChild(buildStatusPill(r.status));
        tr.appendChild(statusTd);

        tr.appendChild(buildReceiptsCell(r.matchedRows, header));

        fragment.appendChild(tr);
    }
    tbody.appendChild(fragment);
}

function buildReceiptsCell(rows: ReadonlyArray<BankRowMatch>, header: ReadonlyArray<string>): HTMLElement {
    const cell = document.createElement('td');
    cell.className = 'col-receipts';
    if (rows.length === 0) {
        const empty = document.createElement('span');
        empty.className = 'reconcile-receipt-empty';
        empty.textContent = '—';
        cell.appendChild(empty);
        return cell;
    }

    const list = document.createElement('ul');
    list.className = 'reconcile-receipt-list';
    for (const row of rows) {
        const item = document.createElement('li');
        item.className = 'reconcile-receipt-item';

        const date = document.createElement('span');
        date.className = 'reconcile-receipt-date';
        date.textContent = (row.date ?? '').trim() || '—';
        item.appendChild(date);

        const account = (row.account ?? '').trim();
        if (account) {
            const acc = document.createElement('span');
            acc.className = 'reconcile-receipt-account';
            acc.textContent = `#${account}`;
            item.appendChild(acc);
        }

        item.appendChild(buildRawRowInfoIcon(row, header));

        list.appendChild(item);
    }
    cell.appendChild(list);
    return cell;
}

function renderManual(
    tbody: HTMLElement,
    items: ReadonlyArray<ManualReviewItem>,
    header: ReadonlyArray<string>,
    onAddClick: (row: BankRowMatch) => void,
    onCandidateClick: (candidate: ManualReviewCandidate, item: ManualReviewItem) => void,
    emptyText = '沒有需要人工處理的交易',
): void {
    tbody.innerHTML = '';
    if (items.length === 0) {
        tbody.innerHTML = `<tr><td colspan="6" class="reconcile-empty">${emptyText}</td></tr>`;
        return;
    }

    const fragment = document.createDocumentFragment();
    for (const item of items) {
        const tr = document.createElement('tr');
        tr.className = `manual-row reason-${item.reason}`;

        const typeTd = document.createElement('td');
        typeTd.className = 'col-type';
        typeTd.textContent = reasonLabel(item.reason);
        tr.appendChild(typeTd);

        tr.appendChild(buildReceiptsCell([item.row], header));
        tr.appendChild(td('col-summary', item.row.summary));
        tr.appendChild(td('col-money', item.row.deposit || '—'));

        const candTd = document.createElement('td');
        candTd.className = 'col-candidate';
        if (item.candidates.length > 0) {
            for (const c of item.candidates) {
                const chip = document.createElement('button');
                chip.type = 'button';
                chip.className = 'reconcile-cand-chip is-clickable';
                chip.title = '點擊編輯此末五碼設定';

                const name = document.createElement('strong');
                name.textContent = c.customerName;
                chip.appendChild(name);
                if (c.customerLine) {
                    chip.appendChild(document.createTextNode(` / ${c.customerLine}`));
                }
                if (c.storeCode) {
                    const code = document.createElement('span');
                    code.className = 'reconcile-cand-code';
                    code.textContent = ` #${c.storeCode}`;
                    chip.appendChild(code);
                } else {
                    const codeMissing = document.createElement('span');
                    codeMissing.className = 'reconcile-cand-code is-missing';
                    codeMissing.textContent = ' #未設';
                    chip.appendChild(codeMissing);
                }
                chip.addEventListener('click', () => onCandidateClick(c, item));
                candTd.appendChild(chip);
            }
        } else {
            const empty = document.createElement('span');
            empty.className = 'reconcile-cand-account';
            empty.textContent = '(無)';
            candTd.appendChild(empty);
        }
        tr.appendChild(candTd);

        const actionTd = document.createElement('td');
        actionTd.className = 'col-action';
        if (item.reason === 'unmatched') {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-secondary btn-xs';
            btn.textContent = '補建';
            btn.addEventListener('click', () => onAddClick(item.row));
            actionTd.appendChild(btn);
        } else if (item.reason === 'no-store-code') {
            const hint = document.createElement('span');
            hint.className = 'reconcile-action-hint';
            hint.textContent = '請至設定補編號';
            actionTd.appendChild(hint);
        } else {
            actionTd.textContent = '—';
        }
        tr.appendChild(actionTd);

        fragment.appendChild(tr);
    }
    tbody.appendChild(fragment);
}

function buildStatusPill(s: ReconcileStatus): HTMLElement {
    const span = document.createElement('span');
    span.className = `reconcile-status-pill is-${s}`;
    span.textContent = statusLabel(s);
    return span;
}

function statusLabel(s: ReconcileStatus): string {
    switch (s) {
        case 'matched':
            return '✓ 已收';
        case 'unpaid':
            return '✗ 未收';
        case 'partial':
            return '⚠ 部分';
        case 'overpaid':
            return '⚠ 超收';
        case 'na':
            return '—';
    }
}

function reasonLabel(reason: ManualReviewItem['reason']): string {
    switch (reason) {
        case 'multi-match':
            return '多重匹配';
        case 'no-store-code':
            return '未設編號';
        case 'unmatched':
            return '未配對';
    }
}

function diffChipClass(diff: number): string {
    if (diff === 0) return 'is-ok';
    return diff < 0 ? 'is-danger' : 'is-warn';
}

function formatMode(r: CustomerReconcileRow): string {
    let s = '';
    if (r.isCashUser) s += '現';
    if (r.isMonthly) s += '月';
    if (r.isNeedTex) s += '稅';
    return s || '—';
}

function formatMoney(n: number): string {
    return n.toLocaleString('zh-TW');
}

function signedMoney(n: number): string {
    if (n === 0) return '0';
    const sign = n > 0 ? '+' : '';
    return `${sign}${n.toLocaleString('zh-TW')}`;
}

function td(className: string, text: string): HTMLElement {
    const el = document.createElement('td');
    el.className = className;
    el.textContent = text;
    return el;
}

function filterCustomers(
    rows: ReadonlyArray<CustomerReconcileRow>,
    query: string,
): CustomerReconcileRow[] {
    const q = query.trim().toLowerCase();
    if (q === '') return [...rows];
    return rows.filter((c) =>
        c.customerCode.toLowerCase().includes(q) ||
        c.customerName.toLowerCase().includes(q),
    );
}

function filterManual(
    items: ReadonlyArray<ManualReviewItem>,
    query: string,
): ManualReviewItem[] {
    const q = query.trim().toLowerCase();
    if (q === '') return [...items];
    return items.filter((item) =>
        item.candidates.some((c) =>
            c.storeCode.toLowerCase().includes(q) ||
            c.customerName.toLowerCase().includes(q),
        ),
    );
}

function buildCustomerEmptyText(tab: 'matched' | 'unmatched', baseCount: number, query: string): string {
    const trimmed = query.trim();
    if (trimmed !== '' && baseCount > 0) return `找不到符合「${trimmed}」的客戶`;
    return tab === 'matched' ? '尚無已匹配客戶' : '所有客戶都有匯款';
}

function buildManualEmptyText(baseCount: number, query: string): string {
    const trimmed = query.trim();
    if (trimmed !== '' && baseCount > 0) return `找不到符合「${trimmed}」的交易`;
    return '沒有需要人工處理的交易';
}

function pickHeader(headerFromResult: ReadonlyArray<string> | null | undefined): ReadonlyArray<string> {
    if (!headerFromResult || headerFromResult.length === 0) return DEFAULT_BANK_HEADER;
    return headerFromResult;
}

// --- Raw row tooltip (滑過 ⓘ 顯示銀行對帳單原始欄位) -------------------------

let rawRowPopoverEl: HTMLElement | null = null;
let rawRowPopoverAnchor: HTMLElement | null = null;
let rawRowScrollHandlerBound = false;

function ensureRawRowPopover(): HTMLElement {
    if (rawRowPopoverEl) return rawRowPopoverEl;
    const el = document.createElement('div');
    el.className = 'reconcile-raw-popover';
    el.setAttribute('role', 'tooltip');
    el.hidden = true;
    document.body.appendChild(el);
    rawRowPopoverEl = el;

    if (!rawRowScrollHandlerBound) {
        // 任何 scroll/resize 都直接收起，避免 popover 與 anchor 錯位。
        const dismiss = () => hideRawRowPopover();
        window.addEventListener('scroll', dismiss, true);
        window.addEventListener('resize', dismiss);
        rawRowScrollHandlerBound = true;
    }
    return el;
}

function hideRawRowPopover(): void {
    rawRowPopoverAnchor = null;
    if (rawRowPopoverEl) {
        rawRowPopoverEl.hidden = true;
        rawRowPopoverEl.innerHTML = '';
    }
}

function showRawRowPopover(anchor: HTMLElement, row: BankRowMatch, header: ReadonlyArray<string>): void {
    const pop = ensureRawRowPopover();
    rawRowPopoverAnchor = anchor;
    pop.innerHTML = '';

    const titleEl = document.createElement('div');
    titleEl.className = 'reconcile-raw-popover-title';
    const fileLine = row.fileLineNumber > 0 ? row.fileLineNumber : (row.rowIndex + 2);
    const source = row.sourceFile ? `${row.sourceFile}　` : '';
    titleEl.textContent = `${source}銀行對帳單 第 ${fileLine} 行`;
    pop.appendChild(titleEl);

    const table = document.createElement('table');
    table.className = 'reconcile-raw-popover-table';
    const tbody = document.createElement('tbody');
    const total = Math.max(header.length, row.raw.length, DEFAULT_BANK_HEADER.length);
    for (let i = 0; i < total; i++) {
        const labelText = header[i] ?? DEFAULT_BANK_HEADER[i] ?? `欄${i + 1}`;
        const valueText = (row.raw[i] ?? '').toString();
        const trEl = document.createElement('tr');
        const th = document.createElement('th');
        th.textContent = labelText;
        const valTd = document.createElement('td');
        const trimmed = valueText.trim();
        if (trimmed === '') {
            valTd.className = 'is-empty';
            valTd.textContent = '—';
        } else {
            valTd.textContent = valueText;
        }
        trEl.appendChild(th);
        trEl.appendChild(valTd);
        tbody.appendChild(trEl);
    }
    table.appendChild(tbody);
    pop.appendChild(table);

    pop.hidden = false;
    positionRawRowPopover(anchor, pop);
}

function positionRawRowPopover(anchor: HTMLElement, pop: HTMLElement): void {
    const rect = anchor.getBoundingClientRect();
    const margin = 8;
    pop.style.left = '0px';
    pop.style.top = '0px';
    const popRect = pop.getBoundingClientRect();
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    // 預設貼右側，超出右邊界則改貼左側
    let left = rect.right + margin;
    if (left + popRect.width > vw - margin) {
        left = rect.left - margin - popRect.width;
    }
    if (left < margin) left = margin;

    let top = rect.top;
    if (top + popRect.height > vh - margin) {
        top = vh - margin - popRect.height;
    }
    if (top < margin) top = margin;

    pop.style.left = `${Math.round(left)}px`;
    pop.style.top = `${Math.round(top)}px`;
}

function buildRawRowInfoIcon(row: BankRowMatch, header: ReadonlyArray<string>): HTMLElement {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'reconcile-raw-info';
    btn.setAttribute('aria-label', '查看原始銀行對帳單資料');
    btn.textContent = 'ⓘ';

    const open = () => showRawRowPopover(btn, row, header);
    const close = () => {
        if (rawRowPopoverAnchor === btn) hideRawRowPopover();
    };

    btn.addEventListener('mouseenter', open);
    btn.addEventListener('mouseleave', close);
    btn.addEventListener('focus', open);
    btn.addEventListener('blur', close);
    // 點擊不觸發父層連動，僅做切換顯示
    btn.addEventListener('click', (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (rawRowPopoverAnchor === btn && rawRowPopoverEl && !rawRowPopoverEl.hidden) {
            hideRawRowPopover();
        } else {
            open();
        }
    });
    return btn;
}
