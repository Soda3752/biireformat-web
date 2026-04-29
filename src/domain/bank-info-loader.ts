/**
 * 末五碼對照資料的載入器：來源為 localStorage 中的 CSV 文字（由設定頁管理）。
 *
 * - 三欄：客戶名稱 / 線別 / 末五碼
 * - 與桌面版 BankInfoReader 對齊：lastFiveDigit 空白略過、整列重複去重
 * - 設定頁編輯後呼叫 invalidateBankInfos() 清快取
 */

import Papa from 'papaparse';

import type {BankInfo} from '@/domain/models/bank-info';
import {equalsBankInfo} from '@/domain/models/bank-info';
import {localSettings} from '@/infra/local-settings-store';

export const BANK_INFO_HEADER = ['客戶名稱', '線別', '末五碼'] as const;

let cache: BankInfo[] | null = null;

/**
 * 確保載入。沒有設定資料時回傳空陣列（不丟錯，由 UI 端自行決定要不要提示）。
 */
export function loadBankInfos(): BankInfo[] {
    if (cache) return cache;
    const csv = localSettings.getLastFiveDigit();
    cache = csv === null ? [] : parseBankInfoCsv(csv);
    return cache;
}

export function getBankInfos(): BankInfo[] {
    return loadBankInfos();
}

export function invalidateBankInfos(): void {
    cache = null;
}

export function parseBankInfoCsv(text: string): BankInfo[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: true});
    const result: BankInfo[] = [];
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        const customerName = String(row[0] ?? '').trim();
        const customerLine = String(row[1] ?? '').trim();
        const lastFiveDigit = String(row[2] ?? '').trim();
        if (lastFiveDigit.length === 0) continue;
        const info: BankInfo = {customerName, customerLine, lastFiveDigit};
        if (!result.some((existing) => equalsBankInfo(existing, info))) {
            result.push(info);
        }
    }
    return result;
}

export function serializeBankInfoCsv(rows: BankInfo[]): string {
    const data: string[][] = [Array.from(BANK_INFO_HEADER)];
    for (const r of rows) {
        data.push([r.customerName, r.customerLine, r.lastFiveDigit]);
    }
    return Papa.unparse(data, {newline: '\n'});
}
