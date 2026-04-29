/**
 * 全域篩選器 UI：日期區間 + 線別 + 客戶 + 商品 + 分類 + 結帳模式 + 檔案來源。
 *
 * 採「點開下拉、勾選 chip」的形式，避免一次塞太多控制項。
 * 篩選變更會觸發 onChange callback，由 panel 統一重繪所有圖表。
 */

import type {AnalyticsDataset} from './dataset-builder';
import type {FilterState} from './filter-engine';
import {EMPTY_FILTER} from './filter-engine';

export interface FilterUiController {
    element: HTMLElement;

    setDataset(dataset: AnalyticsDataset): void;

    /** 由外部（圖表鑽取）程式化套用篩選條件 */
    applyPatch(patch: Partial<FilterState>): void;

    reset(): void;

    getState(): FilterState;
}

export interface FilterUiOptions {
    onChange: (state: FilterState) => void;
}

interface OptionItem {
    value: string;
    label: string;
}

export function createFilterUi(options: FilterUiOptions): FilterUiController {
    const root = document.createElement('div');
    root.className = 'analytics-filters-bar';
    root.innerHTML = `
    <div class="analytics-filter-group">
      <label class="analytics-filter-label">日期</label>
      <div class="analytics-filter-day-range">
        <input type="number" min="1" max="31" placeholder="起" data-role="day-min" class="analytics-filter-day-input">
        <span>~</span>
        <input type="number" min="1" max="31" placeholder="迄" data-role="day-max" class="analytics-filter-day-input">
      </div>
    </div>
    <div class="analytics-filter-group" data-role="grp-line">
      <label class="analytics-filter-label">線別</label>
      <div class="analytics-filter-chips" data-role="line-chips"></div>
    </div>
    <div class="analytics-filter-group" data-role="grp-category">
      <label class="analytics-filter-label">分類</label>
      <div class="analytics-filter-chips" data-role="category-chips"></div>
    </div>
    <div class="analytics-filter-group">
      <label class="analytics-filter-label">客戶</label>
      <select multiple size="1" data-role="customer-select" class="analytics-filter-select"></select>
    </div>
    <div class="analytics-filter-group">
      <label class="analytics-filter-label">排除客戶</label>
      <select multiple size="1" data-role="excluded-customer-select" class="analytics-filter-select"></select>
    </div>
    <div class="analytics-filter-group">
      <label class="analytics-filter-label">商品</label>
      <select multiple size="1" data-role="product-select" class="analytics-filter-select"></select>
    </div>
    <div class="analytics-filter-group">
      <label class="analytics-filter-label">結帳</label>
      <div class="analytics-filter-chips" data-role="pay-chips">
        <button type="button" class="analytics-filter-chip" data-pay="monthly">月結</button>
        <button type="button" class="analytics-filter-chip" data-pay="needTex">含稅</button>
        <button type="button" class="analytics-filter-chip" data-pay="cash">現金</button>
      </div>
    </div>
    <div class="analytics-filter-group" data-role="grp-files" hidden>
      <label class="analytics-filter-label">檔案</label>
      <div class="analytics-filter-chips" data-role="file-chips"></div>
    </div>
    <div class="analytics-filter-group analytics-filter-actions">
      <button type="button" class="btn btn-secondary" data-role="reset">重置</button>
    </div>
  `;

    const dayMin = root.querySelector<HTMLInputElement>('[data-role="day-min"]')!;
    const dayMax = root.querySelector<HTMLInputElement>('[data-role="day-max"]')!;
    const lineChipsHost = root.querySelector<HTMLElement>('[data-role="line-chips"]')!;
    const categoryChipsHost = root.querySelector<HTMLElement>('[data-role="category-chips"]')!;
    const customerSelect = root.querySelector<HTMLSelectElement>('[data-role="customer-select"]')!;
    const excludedCustomerSelect = root.querySelector<HTMLSelectElement>('[data-role="excluded-customer-select"]')!;
    const productSelect = root.querySelector<HTMLSelectElement>('[data-role="product-select"]')!;
    const payChips = root.querySelectorAll<HTMLButtonElement>('[data-pay]');
    const fileChipsHost = root.querySelector<HTMLElement>('[data-role="file-chips"]')!;
    const fileGroup = root.querySelector<HTMLElement>('[data-role="grp-files"]')!;
    const resetBtn = root.querySelector<HTMLButtonElement>('[data-role="reset"]')!;

    let state: FilterState = {...EMPTY_FILTER};
    let lineOptions: OptionItem[] = [];
    let categoryOptions: OptionItem[] = [];
    let fileOptions: OptionItem[] = [];

    const renderChips = (
        host: HTMLElement,
        items: OptionItem[],
        selected: ReadonlySet<string> | null,
        onToggle: (value: string) => void
    ) => {
        host.innerHTML = items
            .map((it) => {
                const active = selected ? selected.has(it.value) : false;
                return `<button type="button" class="analytics-filter-chip ${active ? 'is-active' : ''}" data-value="${escapeHtml(it.value)}">${escapeHtml(it.label)}</button>`;
            })
            .join('');
        host.querySelectorAll<HTMLButtonElement>('button[data-value]').forEach((btn) => {
            btn.addEventListener('click', () => {
                onToggle(btn.dataset.value!);
            });
        });
    };

    const refreshChips = () => {
        renderChips(lineChipsHost, lineOptions, state.lines, (v) => toggleSet('lines', v));
        renderChips(categoryChipsHost, categoryOptions, state.categories, (v) => toggleSet('categories', v));
        renderChips(fileChipsHost, fileOptions, state.fileIds, (v) => toggleSet('fileIds', v));
    };

    const toggleSet = (
        key: 'lines' | 'categories' | 'fileIds' | 'customerCodes' | 'excludedCustomerCodes' | 'productNames',
        value: string
    ) => {
        const cur = state[key];
        let next: Set<string> | null;
        if (!cur) {
            next = new Set([value]);
        } else {
            next = new Set(cur);
            if (next.has(value)) next.delete(value);
            else next.add(value);
            if (next.size === 0) next = null;
        }
        state = {...state, [key]: next};
        refreshChips();
        refreshPayChips();
        syncSelectMulti(customerSelect, state.customerCodes);
        syncSelectMulti(excludedCustomerSelect, state.excludedCustomerCodes);
        syncSelectMulti(productSelect, state.productNames);
        options.onChange(state);
    };

    const togglePayment = (key: 'paymentMonthly' | 'paymentNeedTex' | 'paymentCash') => {
        // 三段式：null（不限） → true → null（再點關閉）
        const cur = state[key];
        const next = cur === null ? true : null;
        state = {...state, [key]: next};
        refreshPayChips();
        options.onChange(state);
    };

    const refreshPayChips = () => {
        payChips.forEach((btn) => {
            const k = btn.dataset.pay!;
            const stateKey = k === 'monthly' ? 'paymentMonthly' : k === 'needTex' ? 'paymentNeedTex' : 'paymentCash';
            const v = state[stateKey as keyof FilterState];
            btn.classList.toggle('is-active', v === true);
        });
    };

    payChips.forEach((btn) => {
        btn.addEventListener('click', () => {
            const k = btn.dataset.pay!;
            togglePayment(
                k === 'monthly' ? 'paymentMonthly' : k === 'needTex' ? 'paymentNeedTex' : 'paymentCash'
            );
        });
    });

    // day range
    const onDayChange = () => {
        const min = parseDay(dayMin.value);
        const max = parseDay(dayMax.value);
        state = {...state, dayMin: min, dayMax: max};
        options.onChange(state);
    };
    dayMin.addEventListener('change', onDayChange);
    dayMax.addEventListener('change', onDayChange);

    // multi-select customer / product（用瀏覽器原生多選）
    customerSelect.addEventListener('change', () => {
        const set = new Set<string>();
        for (const opt of Array.from(customerSelect.selectedOptions)) set.add(opt.value);
        state = {...state, customerCodes: set.size === 0 ? null : set};
        options.onChange(state);
    });
    excludedCustomerSelect.addEventListener('change', () => {
        const set = new Set<string>();
        for (const opt of Array.from(excludedCustomerSelect.selectedOptions)) set.add(opt.value);
        state = {...state, excludedCustomerCodes: set.size === 0 ? null : set};
        options.onChange(state);
    });
    productSelect.addEventListener('change', () => {
        const set = new Set<string>();
        for (const opt of Array.from(productSelect.selectedOptions)) set.add(opt.value);
        state = {...state, productNames: set.size === 0 ? null : set};
        options.onChange(state);
    });

    resetBtn.addEventListener('click', () => {
        state = {...EMPTY_FILTER};
        dayMin.value = '';
        dayMax.value = '';
        refreshChips();
        refreshPayChips();
        syncSelectMulti(customerSelect, null);
        syncSelectMulti(excludedCustomerSelect, null);
        syncSelectMulti(productSelect, null);
        options.onChange(state);
    });

    return {
        element: root,
        setDataset(dataset) {
            // 收集選項
            const lineSet = new Set<string>();
            const categorySet = new Set<string>();
            const customerMap = new Map<string, string>();
            const productSet = new Set<string>();
            for (const r of dataset.rows) {
                lineSet.add(r.line);
                categorySet.add(r.category);
                if (!customerMap.has(r.customerCode)) customerMap.set(r.customerCode, `${r.customerName}(${r.customerCode})`);
                productSet.add(r.productName);
            }
            lineOptions = [...lineSet].sort().map((v) => ({value: v, label: v}));
            categoryOptions = [...categorySet].sort().map((v) => ({value: v, label: v}));
            fileOptions = dataset.files.map((f) => ({value: f.id, label: `${f.year}/${f.month} ${f.name}`}));

            const customerOptionsHtml = [...customerMap]
                .sort((a, b) => compareCustomerCode(a[0], b[0]))
                .map(([code, label]) => `<option value="${escapeHtml(code)}">${escapeHtml(label)}</option>`)
                .join('');
            customerSelect.innerHTML = customerOptionsHtml;
            excludedCustomerSelect.innerHTML = customerOptionsHtml;
            productSelect.innerHTML = [...productSet]
                .sort()
                .map((p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`)
                .join('');

            fileGroup.hidden = dataset.files.length < 2;
            refreshChips();
            refreshPayChips();
            syncSelectMulti(customerSelect, state.customerCodes);
            syncSelectMulti(excludedCustomerSelect, state.excludedCustomerCodes);
            syncSelectMulti(productSelect, state.productNames);
        },
        applyPatch(patch) {
            state = {...state, ...patch};
            // 同步 UI
            if ('dayMin' in patch) dayMin.value = state.dayMin == null ? '' : String(state.dayMin);
            if ('dayMax' in patch) dayMax.value = state.dayMax == null ? '' : String(state.dayMax);
            refreshChips();
            refreshPayChips();
            syncSelectMulti(customerSelect, state.customerCodes);
            syncSelectMulti(excludedCustomerSelect, state.excludedCustomerCodes);
            syncSelectMulti(productSelect, state.productNames);
            options.onChange(state);
        },
        reset() {
            state = {...EMPTY_FILTER};
            dayMin.value = '';
            dayMax.value = '';
            refreshChips();
            refreshPayChips();
            syncSelectMulti(customerSelect, null);
            syncSelectMulti(excludedCustomerSelect, null);
            syncSelectMulti(productSelect, null);
            options.onChange(state);
        },
        getState() {
            return state;
        },
    };
}

function parseDay(v: string): number | null {
    const n = Number(v);
    if (!Number.isFinite(n) || n < 1 || n > 31) return null;
    return Math.floor(n);
}

function syncSelectMulti(select: HTMLSelectElement, set: ReadonlySet<string> | null): void {
    Array.from(select.options).forEach((opt) => {
        opt.selected = set ? set.has(opt.value) : false;
    });
}

function compareCustomerCode(a: string, b: string): number {
    const numA = Number(a);
    const numB = Number(b);
    const aIsNum = Number.isFinite(numA);
    const bIsNum = Number.isFinite(numB);
    if (aIsNum && bIsNum) return numA - numB;
    if (aIsNum) return -1;
    if (bIsNum) return 1;
    return a.localeCompare(b);
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}

