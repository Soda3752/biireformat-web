/**
 * 對帳 2.0 核心：以「客戶」為主軸聚合應收與已收。
 *
 * 輸入：
 *  - bill：帳單 .xlsx 解析後的客戶清單（提供應收）
 *  - bankResult：銀行對帳單 .csv/.xlsx 解析後的逐筆配對結果（提供已收）
 *  - bankInfos：末五碼對照表（提供 storeCode → CustomerModel.code 連結；目前 service 內僅作型別語義紀錄，
 *               實際配對在 bankResult.rows[i].matches 已由 matchTransRecord 完成）
 *
 * 配對策略（嚴格 storeCode）：
 *  - 一筆交易 matches.length === 1 且 matches[0].storeCode 非空 → 計入該 storeCode 對應客戶的「已收」
 *  - 其餘（無匹配 / 多重匹配 / 匹配到但 storeCode 為空）→ 不計入任何客戶，落入 manualReviewRows
 *
 * 輸出：以 CustomerModel.customerModels 順序為主表順序（與帳單原順序一致）。
 */

import type {Bill} from '@/domain/models/bill';
import type {CustomerModel} from '@/domain/models/customer-model';
import type {BankInfo} from '@/domain/models/bank-info';
import type {BankMatchResult, BankRowMatch} from '@/domain/bank-match-service';

export type ReconcileStatus = 'matched' | 'unpaid' | 'partial' | 'overpaid' | 'na';

export type ManualReviewReason = 'multi-match' | 'no-store-code' | 'unmatched';

export interface ReconcileMatchedRow extends BankRowMatch {
    /** 跨月旗標：true 表示該列被「金額容差篩選」排除、不計入 received，但仍保留供顯示。 */
    crossMonth: boolean;
}

export interface CustomerReconcileRow {
    customerCode: string;
    customerName: string;
    customerLine: string;
    isCashUser: boolean;
    isMonthly: boolean;
    isNeedTex: boolean;
    receivable: number;
    received: number;
    diff: number;
    status: ReconcileStatus;
    matchedRows: ReconcileMatchedRow[];
}

/**
 * 銀行手續費容差預設值（元）。
 *
 * 場景：銀行對帳單可能包含跨月入帳（同 storeCode 上月匯款），導致應為 matched 的客戶被累加成 overpaid。
 * 處理規則：對每位客戶的 matchedRows
 *  1. 若存在某筆 (receivable - deposit) 在 [0, feeTolerance] 內 → 取「|receivable - deposit| 最小」那筆當本月款；
 *     其餘列標記 crossMonth=true，不計入 received。
 *  2. 若無任何列落在容差內 → 退回原本「全部累加」邏輯（保留分期付款、部分付款的語意）。
 */
export const DEFAULT_FEE_TOLERANCE = 15;

export interface ReconcileOptions {
    /** 銀行手續費容差（元）。預設 15。 */
    feeTolerance: number;
}

export interface ManualReviewCandidate {
    customerName: string;
    customerLine: string;
    storeCode: string;
    /** BankInfo 原始末五碼，編輯/定位該筆設定用（與 row.account 不一定相等，例如 wildcard 隱碼）。 */
    lastFiveDigit: string;
}

export interface ManualReviewItem {
    row: BankRowMatch;
    reason: ManualReviewReason;
    /** 多重匹配時候選客戶（顯示用），其他情況為空陣列。 */
    candidates: ReadonlyArray<ManualReviewCandidate>;
}

export interface ReconcileSummary {
    totalReceivable: number;
    totalReceived: number;
    totalDiff: number;
    customerCount: number;
    manualReviewCount: number;
    unpaidCount: number;
    partialCount: number;
    overpaidCount: number;
    matchedCount: number;
}

export interface ReconcileResult {
    bill: Bill;
    customers: CustomerReconcileRow[];
    manualReviewItems: ManualReviewItem[];
    rawRows: BankRowMatch[];
    /** 銀行 CSV 標題列（如有），用於 UI tooltip 對照欄位名。 */
    bankHeader: string[] | null;
    summary: ReconcileSummary;
}

export function reconcileByCustomer(
    bill: Bill,
    bankResult: BankMatchResult,
    _bankInfos: ReadonlyArray<BankInfo>,
    options: ReconcileOptions = {feeTolerance: DEFAULT_FEE_TOLERANCE},
): ReconcileResult {
    // 以 storeCode 為鍵，收集所有對應到該客戶的銀行列（先不累加，留給 buildCustomerRow 做跨月篩選）
    const matchedRowsByCode = new Map<string, BankRowMatch[]>();
    const manualReviewItems: ManualReviewItem[] = [];

    for (const row of bankResult.rows) {
        if (row.matches.length === 0) {
            manualReviewItems.push({row, reason: 'unmatched', candidates: []});
            continue;
        }

        if (row.matches.length > 1) {
            manualReviewItems.push({
                row,
                reason: 'multi-match',
                candidates: row.matches.map((m) => ({
                    customerName: m.customerName,
                    customerLine: m.customerLine,
                    storeCode: m.storeCode,
                    lastFiveDigit: m.lastFiveDigit,
                })),
            });
            continue;
        }

        // matches.length === 1
        const info = row.matches[0];
        const code = info.storeCode.trim();
        if (code === '') {
            manualReviewItems.push({
                row,
                reason: 'no-store-code',
                candidates: [{
                    customerName: info.customerName,
                    customerLine: info.customerLine,
                    storeCode: info.storeCode,
                    lastFiveDigit: info.lastFiveDigit,
                }],
            });
            continue;
        }

        const list = matchedRowsByCode.get(code);
        if (list) list.push(row);
        else matchedRowsByCode.set(code, [row]);
    }

    const feeTolerance = Math.max(0, options.feeTolerance);
    const customers: CustomerReconcileRow[] = bill.customerModels.map((c) =>
        buildCustomerRow(c, matchedRowsByCode.get(c.code) ?? [], feeTolerance),
    );

    return {
        bill,
        customers,
        manualReviewItems,
        rawRows: [...bankResult.rows],
        bankHeader: bankResult.header,
        summary: buildSummary(customers, manualReviewItems),
    };
}

function buildCustomerRow(
    c: CustomerModel,
    bankRows: ReadonlyArray<BankRowMatch>,
    feeTolerance: number,
): CustomerReconcileRow {
    const receivable = c.isNeedTex ? c.getAfterTexSum() : c.getTotalPrice();
    const {received, matchedRows} = applyFeeToleranceFilter(receivable, bankRows, feeTolerance);
    const diff = received - receivable;
    return {
        customerCode: c.code,
        customerName: c.name,
        customerLine: deriveLine(c.code),
        isCashUser: c.isCashUser,
        isMonthly: c.isMonthly,
        isNeedTex: c.isNeedTex,
        receivable,
        received,
        diff,
        status: computeStatus(receivable, received),
        matchedRows,
    };
}

/**
 * 用「金額最接近應收 + 容差內」規則挑出當月匹配款，其餘列標跨月。
 *
 * - bestIdx：rows 中 (receivable - deposit) ∈ [0, feeTolerance] 且 |receivable - deposit| 最小的索引。
 *   tie-breaker：金額相同時取索引較後者（在合併多份對帳單時通常是日期較新的）。
 * - 若 bestIdx 存在：received = rows[bestIdx].deposit，其他列 crossMonth=true。
 * - 若不存在（如分期付款 500+500、無人匹配等情境）：received = 全部 deposit 加總，所有列 crossMonth=false。
 */
function applyFeeToleranceFilter(
    receivable: number,
    rows: ReadonlyArray<BankRowMatch>,
    feeTolerance: number,
): { received: number; matchedRows: ReconcileMatchedRow[] } {
    if (rows.length === 0) {
        return {received: 0, matchedRows: []};
    }

    let bestIdx = -1;
    let bestDelta = Number.POSITIVE_INFINITY;
    for (let i = 0; i < rows.length; i++) {
        const deposit = parseDeposit(rows[i].deposit);
        const delta = receivable - deposit;
        if (delta < 0 || delta > feeTolerance) continue;
        if (delta <= bestDelta) {
            bestDelta = delta;
            bestIdx = i;
        }
    }

    if (bestIdx === -1) {
        const received = rows.reduce((sum, r) => sum + parseDeposit(r.deposit), 0);
        return {
            received,
            matchedRows: rows.map((r) => ({...r, crossMonth: false})),
        };
    }

    const received = parseDeposit(rows[bestIdx].deposit);
    return {
        received,
        matchedRows: rows.map((r, i) => ({...r, crossMonth: i !== bestIdx})),
    };
}

function buildSummary(
    customers: ReadonlyArray<CustomerReconcileRow>,
    manualReviewItems: ReadonlyArray<ManualReviewItem>,
): ReconcileSummary {
    let totalReceivable = 0;
    let totalReceived = 0;
    let matchedCount = 0;
    let unpaidCount = 0;
    let partialCount = 0;
    let overpaidCount = 0;
    for (const r of customers) {
        totalReceivable += r.receivable;
        totalReceived += r.received;
        switch (r.status) {
            case 'matched':
                matchedCount++;
                break;
            case 'unpaid':
                unpaidCount++;
                break;
            case 'partial':
                partialCount++;
                break;
            case 'overpaid':
                overpaidCount++;
                break;
            case 'na':
                break;
        }
    }
    return {
        totalReceivable,
        totalReceived,
        totalDiff: totalReceived - totalReceivable,
        customerCount: customers.length,
        manualReviewCount: manualReviewItems.length,
        unpaidCount,
        partialCount,
        overpaidCount,
        matchedCount,
    };
}

function computeStatus(receivable: number, received: number): ReconcileStatus {
    if (receivable === 0 && received === 0) return 'na';
    if (received === 0) return 'unpaid';
    if (received === receivable) return 'matched';
    if (received > receivable) return 'overpaid';
    return 'partial';
}

function deriveLine(code: string): string {
    if (!code) return '';
    return `第${code.charAt(0)}線`;
}

function parseDeposit(raw: string): number {
    const s = String(raw ?? '').replace(/,/g, '').trim();
    if (s === '') return 0;
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
}
