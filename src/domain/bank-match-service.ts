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
    /** UI 顯示用：多候選以「、」串接。配對邏輯實際使用 accounts。 */
    account: string;
    /** 配對候選清單：標準銀行 CSV 只有 1 個（最末欄）；郵局合併後可能多個。 */
    accounts: string[];
    matches: BankInfo[];
}

export interface BankMatchResult {
    header: string[] | null;
    rows: BankRowMatch[];
    matchedCount: number;
    unmatchedCount: number;
}

const HEADER_TOKENS = ['日期', '摘要'] as const;

/**
 * 配對候選欄收集起始 index（對齊標準 9 欄格式的「註記」欄）。
 * - 標準銀行 CSV：trimmed 最多 9 欄，僅 index 8 為候選 → 行為與舊版相同
 * - 郵局合併後：主行 index 8 為主帳號，index 9+ 為補充行帳號/銀行名，全數成為候選
 */
const ACCOUNT_CANDIDATE_START = 8;

const isHeaderRow = (row: ReadonlyArray<string>): boolean =>
    HEADER_TOKENS.every((token, i) => (row[i] ?? '').trim() === token);

const dropTrailingBlanks = (row: ReadonlyArray<string>): string[] => {
    const out = [...row];
    while (out.length > 0 && (out[out.length - 1] ?? '').trim() === '') out.pop();
    return out;
};

export const collectAccounts = (trimmed: ReadonlyArray<string>): string[] => {
    const out: string[] = [];
    for (let i = ACCOUNT_CANDIDATE_START; i < trimmed.length; i++) {
        const v = (trimmed[i] ?? '').trim();
        if (v !== '') out.push(v);
    }
    // 後備：trimmed 不到 ACCOUNT_CANDIDATE_START 欄時，取最末非空欄
    if (out.length === 0 && trimmed.length > 0) {
        const v = (trimmed[trimmed.length - 1] ?? '').trim();
        if (v !== '') out.push(v);
    }
    return out;
};

export const findMatchesFromCandidates = (
    accounts: ReadonlyArray<string>,
    bankInfos: ReadonlyArray<BankInfo>
): BankInfo[] => {
    const seen = new Set<string>();
    const out: BankInfo[] = [];
    for (const acc of accounts) {
        const trimmed = acc.trim();
        if (trimmed.length === 0) continue;
        for (const info of bankInfos) {
            if (!isSameAccount(info, trimmed)) continue;
            const key = `${info.customerName}|${info.customerLine}|${info.lastFiveDigit}`;
            if (seen.has(key)) continue;
            seen.add(key);
            out.push(info);
        }
    }
    return out;
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
        const accounts = collectAccounts(trimmed);
        const matches = findMatchesFromCandidates(accounts, bankInfos);
        if (matches.length > 0) matchedCount++;
        rows.push({
            rowIndex: i - startIdx,
            raw,
            trimmed,
            date: (raw[0] ?? '').trim(),
            summary: (raw[1] ?? '').trim(),
            deposit: (raw[4] ?? '').trim(),
            account: accounts.join('、'),
            accounts,
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
