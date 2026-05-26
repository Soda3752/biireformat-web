/**
 * 末五碼對照資料的載入器：來源為 localStorage 中的 CSV 文字（由設定頁管理）。
 *
 * - 四欄：客戶名稱 / 店家編號 / 線別 / 末五碼
 * - 向後相容舊版 3 欄（客戶名稱 / 線別 / 末五碼），讀取時自動補 storeCode=''
 * - 與桌面版 BankInfoReader 對齊：lastFiveDigit 空白略過、整列重複去重
 * - 設定頁編輯後呼叫 invalidateBankInfos() 清快取
 */

import Papa from 'papaparse';

import type {BankInfo} from '@/domain/models/bank-info';
import {equalsBankInfo} from '@/domain/models/bank-info';
import {localSettings} from '@/infra/local-settings-store';

export const BANK_INFO_HEADER = ['客戶名稱', '店家編號', '線別', '末五碼'] as const;
const LEGACY_HEADER_3COL = ['客戶名稱', '線別', '末五碼'] as const;

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
    const isLegacy = detectLegacyHeader(parsed.data[0]);
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        const info = isLegacy ? parseLegacyRow(row) : parseRow(row);
        if (info.lastFiveDigit.length === 0) continue;
        if (!result.some((existing) => equalsBankInfo(existing, info))) {
            result.push(info);
        }
    }
    return result;
}

function detectLegacyHeader(headerRow: string[] | undefined): boolean {
    if (!headerRow) return false;
    const trimmed = headerRow.map((c) => String(c ?? '').trim());
    if (trimmed[1] === BANK_INFO_HEADER[1]) return false; // 第二欄是「店家編號」→ 新格式
    if (
        trimmed[0] === LEGACY_HEADER_3COL[0] &&
        trimmed[1] === LEGACY_HEADER_3COL[1] &&
        trimmed[2] === LEGACY_HEADER_3COL[2]
    ) {
        return true;
    }
    return false;
}

function parseRow(row: string[]): BankInfo {
    return {
        customerName: String(row[0] ?? '').trim(),
        storeCode: String(row[1] ?? '').trim(),
        customerLine: String(row[2] ?? '').trim(),
        lastFiveDigit: String(row[3] ?? '').trim(),
    };
}

function parseLegacyRow(row: string[]): BankInfo {
    return {
        customerName: String(row[0] ?? '').trim(),
        storeCode: '',
        customerLine: String(row[1] ?? '').trim(),
        lastFiveDigit: String(row[2] ?? '').trim(),
    };
}

export function serializeBankInfoCsv(rows: BankInfo[]): string {
    const data: string[][] = [Array.from(BANK_INFO_HEADER)];
    for (const r of rows) {
        data.push([r.customerName, r.storeCode, r.customerLine, r.lastFiveDigit]);
    }
    return Papa.unparse(data, {newline: '\n'});
}
