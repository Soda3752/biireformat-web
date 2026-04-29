/**
 * 客戶代碼排序的載入器：來源為 localStorage 中的 CSV 文字（雙欄，由設定頁管理）。
 *
 * 兩欄：客戶代碼 / 客戶名稱
 * - 排序語意只看「客戶代碼」(對應桌面版 OrderReader.kt)，「客戶名稱」僅作編輯時辨識用
 * - 帳單分頁與明細分頁各有一份排序，互不影響
 */

import Papa from 'papaparse';

import {LEGACY_CUSTOMER_ORDER_KEY, localSettings} from '@/infra/local-settings-store';

export interface CustomerOrderEntry {
    code: string;
    name: string;
}

export const CUSTOMER_ORDER_HEADER = ['客戶代碼', '客戶名稱'] as const;

export type CustomerOrderVariant = 'bill' | 'overview';

let cacheBill: CustomerOrderEntry[] | null = null;
let cacheOverview: CustomerOrderEntry[] | null = null;

migrateLegacyKeyOnce();

// ============================================================================
// 帳單客戶排序
// ============================================================================

export function loadCustomerOrderBill(): CustomerOrderEntry[] {
    if (cacheBill) return cacheBill;
    const csv = localSettings.getCustomerOrderBill();
    cacheBill = csv === null ? [] : parseCustomerOrderCsv(csv);
    return cacheBill;
}

export function getCustomerOrderBill(): CustomerOrderEntry[] {
    return loadCustomerOrderBill();
}

export function invalidateCustomerOrderBill(): void {
    cacheBill = null;
}

// ============================================================================
// 明細客戶排序
// ============================================================================

export function loadCustomerOrderOverview(): CustomerOrderEntry[] {
    if (cacheOverview) return cacheOverview;
    const csv = localSettings.getCustomerOrderOverview();
    cacheOverview = csv === null ? [] : parseCustomerOrderCsv(csv);
    return cacheOverview;
}

export function getCustomerOrderOverview(): CustomerOrderEntry[] {
    return loadCustomerOrderOverview();
}

export function invalidateCustomerOrderOverview(): void {
    cacheOverview = null;
}

// ============================================================================
// 共用 helper
// ============================================================================

export function parseCustomerOrderCsv(text: string): CustomerOrderEntry[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: true});
    return collectCustomerEntries(parsed.data as string[][]);
}

/**
 * 解析雙欄客戶排序資料：
 * - 第一列若是 header（值等於「客戶代碼」）則跳過，否則保留為實際資料
 *   → 同時相容自家匯出的雙欄 header、舊版單欄 header、以及桌面版無 header 的 .xlsx
 * - 以「客戶代碼」為 key 去重；空白代碼略過
 * - 缺第二欄時 name 為空字串
 */
export function collectCustomerEntries(rows: string[][]): CustomerOrderEntry[] {
    const out: CustomerOrderEntry[] = [];
    const seen = new Set<string>();
    const startIdx =
        rows.length > 0 && (rows[0][0] ?? '').trim() === CUSTOMER_ORDER_HEADER[0] ? 1 : 0;
    for (let i = startIdx; i < rows.length; i++) {
        const row = rows[i];
        if (!row || row.length === 0) continue;
        const code = String(row[0] ?? '').trim();
        if (code.length === 0) continue;
        if (seen.has(code)) continue;
        seen.add(code);
        const name = String(row[1] ?? '').trim();
        out.push({code, name});
    }
    return out;
}

export function serializeCustomerOrderCsv(entries: CustomerOrderEntry[]): string {
    const data: string[][] = [
        Array.from(CUSTOMER_ORDER_HEADER),
        ...entries.map((e) => [e.code, e.name]),
    ];
    return Papa.unparse(data, {newline: '\n'});
}

/**
 * 一次性把舊的單一 customerOrder key 遷移到 bill / overview 兩份。
 * 遷移後刪除舊 key；之後啟動就直接走新 key。
 *
 * 規則：只有當新 key 都尚未設定時才複製（避免覆蓋使用者已分別調整過的內容）。
 */
function migrateLegacyKeyOnce(): void {
    let legacy: string | null = null;
    try {
        legacy = localStorage.getItem(LEGACY_CUSTOMER_ORDER_KEY);
    } catch {
        return;
    }
    if (legacy === null) return;

    if (!localSettings.hasCustomerOrderBill()) localSettings.setCustomerOrderBill(legacy);
    if (!localSettings.hasCustomerOrderOverview()) localSettings.setCustomerOrderOverview(legacy);
    try {
        localStorage.removeItem(LEGACY_CUSTOMER_ORDER_KEY);
    } catch {
        // ignore
    }
}
