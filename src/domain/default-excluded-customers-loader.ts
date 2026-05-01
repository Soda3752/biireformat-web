/**
 * 預設排除店家載入器：來源為 localStorage 中的 CSV 文字（雙欄，由設定頁管理）。
 *
 * 兩欄：客戶編號 / 備註
 * - 數據分析開啟時會自動把這些客戶編號預先勾入「排除客戶」篩選器
 * - 重置篩選器後仍會回到此預設狀態
 */

import Papa from 'papaparse';

import {localSettings} from '@/infra/local-settings-store';

export interface DefaultExcludedCustomerEntry {
    code: string;
    name: string;
}

export const DEFAULT_EXCLUDED_CUSTOMERS_HEADER = ['客戶編號', '備註'] as const;

let cache: DefaultExcludedCustomerEntry[] | null = null;

export function loadDefaultExcludedCustomers(): DefaultExcludedCustomerEntry[] {
    if (cache) return cache;
    const csv = localSettings.getDefaultExcludedCustomers();
    cache = csv === null ? [] : parseDefaultExcludedCustomersCsv(csv);
    return cache;
}

export function getDefaultExcludedCustomerCodeSet(): Set<string> {
    const entries = loadDefaultExcludedCustomers();
    const out = new Set<string>();
    for (const e of entries) {
        if (e.code.length > 0) out.add(e.code);
    }
    return out;
}

export function invalidateDefaultExcludedCustomers(): void {
    cache = null;
}

export function parseDefaultExcludedCustomersCsv(text: string): DefaultExcludedCustomerEntry[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: true});
    return collectDefaultExcludedEntries(parsed.data as string[][]);
}

/**
 * 解析雙欄資料：第一欄為客戶編號（必填、去重），第二欄為備註（可選）。
 * 若第一列是 header（值等於「客戶編號」）則跳過。
 */
export function collectDefaultExcludedEntries(rows: string[][]): DefaultExcludedCustomerEntry[] {
    const out: DefaultExcludedCustomerEntry[] = [];
    const seen = new Set<string>();
    const startIdx =
        rows.length > 0 && (rows[0][0] ?? '').trim() === DEFAULT_EXCLUDED_CUSTOMERS_HEADER[0] ? 1 : 0;
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

export function serializeDefaultExcludedCustomersCsv(entries: DefaultExcludedCustomerEntry[]): string {
    const data: string[][] = [
        Array.from(DEFAULT_EXCLUDED_CUSTOMERS_HEADER),
        ...entries.map((e) => [e.code, e.name]),
    ];
    return Papa.unparse(data, {newline: '\n'});
}
