/**
 * 全域篩選器引擎：把 AnalyticsDataset 套用 FilterState 後產出 filteredDataset。
 */

import type {AnalyticsDataset, AnalyticsRow} from './dataset-builder';

export interface FilterState {
    fileIds: ReadonlySet<string> | null;
    lines: ReadonlySet<string> | null;
    customerCodes: ReadonlySet<string> | null;
    excludedCustomerCodes: ReadonlySet<string> | null;
    productNames: ReadonlySet<string> | null;
    categories: ReadonlySet<string> | null;
    dayMin: number | null; // 1~31，null 表示不限
    dayMax: number | null;
    paymentMonthly: boolean | null; // null 表示不過濾，true 限月結，false 限非月結
    paymentNeedTex: boolean | null;
    paymentCash: boolean | null;
}

export const EMPTY_FILTER: FilterState = {
    fileIds: null,
    lines: null,
    customerCodes: null,
    excludedCustomerCodes: null,
    productNames: null,
    categories: null,
    dayMin: null,
    dayMax: null,
    paymentMonthly: null,
    paymentNeedTex: null,
    paymentCash: null,
};

export function applyFilter(
    dataset: AnalyticsDataset,
    filter: FilterState
): AnalyticsDataset {
    const rows = dataset.rows.filter((r) => matchRow(r, filter));
    return {
        rows,
        files: dataset.files,
        unmatchedProducts: dataset.unmatchedProducts,
        unsetCostProducts: dataset.unsetCostProducts,
    };
}

function matchRow(r: AnalyticsRow, f: FilterState): boolean {
    if (f.fileIds && !f.fileIds.has(r.fileId)) return false;
    if (f.lines && !f.lines.has(r.line)) return false;
    if (f.excludedCustomerCodes && f.excludedCustomerCodes.has(r.customerCode)) return false;
    if (f.customerCodes && !f.customerCodes.has(r.customerCode)) return false;
    if (f.productNames && !f.productNames.has(r.productName)) return false;
    if (f.categories && !f.categories.has(r.category)) return false;
    if (f.dayMin != null && r.day < f.dayMin) return false;
    if (f.dayMax != null && r.day > f.dayMax) return false;
    if (f.paymentMonthly != null && r.isMonthly !== f.paymentMonthly) return false;
    if (f.paymentNeedTex != null && r.isNeedTex !== f.paymentNeedTex) return false;
    if (f.paymentCash != null && r.isCashUser !== f.paymentCash) return false;
    return true;
}

export function isFilterActive(f: FilterState): boolean {
    return (
        f.fileIds !== null ||
        f.lines !== null ||
        f.customerCodes !== null ||
        f.excludedCustomerCodes !== null ||
        f.productNames !== null ||
        f.categories !== null ||
        f.dayMin !== null ||
        f.dayMax !== null ||
        f.paymentMonthly !== null ||
        f.paymentNeedTex !== null ||
        f.paymentCash !== null
    );
}
