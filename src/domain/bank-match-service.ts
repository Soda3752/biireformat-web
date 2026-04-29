import type {BankInfo} from '@/domain/models/bank-info';
import {isSameAccount} from '@/domain/models/bank-info';

/**
 * 將銀行對帳單原始資料 (CSV 解析後 string[][]) 與末五碼對照表比對，
 * 產出 UI 預覽所需的結構化結果。
 *
 * 比對邏輯與 bank-name-writer.ts 對齊：
 * - 對每列丟掉尾端空白欄，最後一欄即為比對目標（中文摘要或隱碼帳號）
 * - 用 isSameAccount 找出所有匹配的 BankInfo
 *
 * 桌面版 CSV 第一列為標題（日期, 摘要, 幣別, 支出金額, 存入金額, 餘額, 備註, 轉出入帳號, 註記）
 * 標題列在輸出時是原樣輸出，不需參與配對；UI 預覽會把它過濾掉。
 */

export interface BankRowMatch {
    rowIndex: number;
    raw: string[];
    trimmed: string[];
    date: string;
    summary: string;
    deposit: string;
    account: string;
    matches: BankInfo[];
}

export interface BankMatchResult {
    header: string[] | null;
    rows: BankRowMatch[];
    matchedCount: number;
    unmatchedCount: number;
}

const HEADER_TOKENS = ['日期', '摘要'] as const;

const isHeaderRow = (row: ReadonlyArray<string>): boolean =>
    HEADER_TOKENS.every((token, i) => (row[i] ?? '').trim() === token);

const dropTrailingBlanks = (row: ReadonlyArray<string>): string[] => {
    const out = [...row];
    while (out.length > 0 && (out[out.length - 1] ?? '').trim() === '') out.pop();
    return out;
};

const findMatches = (
    account: string,
    bankInfos: ReadonlyArray<BankInfo>
): BankInfo[] => {
    const trimmed = account.trim();
    if (trimmed.length === 0) return [];
    return bankInfos.filter((info) => isSameAccount(info, trimmed));
};

export const matchTransRecord = (
    originData: ReadonlyArray<ReadonlyArray<string>>,
    bankInfos: ReadonlyArray<BankInfo>
): BankMatchResult => {
    let header: string[] | null = null;
    let startIdx = 0;
    if (originData.length > 0 && isHeaderRow(originData[0])) {
        header = [...originData[0]];
        startIdx = 1;
    }

    const rows: BankRowMatch[] = [];
    let matchedCount = 0;
    for (let i = startIdx; i < originData.length; i++) {
        const raw = [...originData[i]];
        const trimmed = dropTrailingBlanks(raw);
        const account = trimmed[trimmed.length - 1] ?? '';
        const matches = findMatches(account, bankInfos);
        if (matches.length > 0) matchedCount++;
        rows.push({
            rowIndex: i - startIdx,
            raw,
            trimmed,
            date: (raw[0] ?? '').trim(),
            summary: (raw[1] ?? '').trim(),
            deposit: (raw[4] ?? '').trim(),
            account,
            matches,
        });
    }

    return {
        header,
        rows,
        matchedCount,
        unmatchedCount: rows.length - matchedCount,
    };
};
