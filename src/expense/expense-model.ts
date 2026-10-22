/**
 * 「支出」分頁的資料模型。
 *
 * 一個月份 = 一張收支表月工作表解析後的結果（MonthExpense），
 * 以 sortKey（`${民國年}-${MM}`，與跨月分析的 MonthKey 同規約）為唯一鍵，重新上傳即覆蓋。
 */

/** 收支表的 6 大支出類別（順序即顯示順序，與 Excel 年度總表一致）。 */
export const EXPENSE_CATEGORIES = ['原物料', '營運', '車輛', '人事', '其他', '銀行'] as const;

export type ExpenseCategory = (typeof EXPENSE_CATEGORIES)[number];

export function isExpenseCategory(text: string): text is ExpenseCategory {
    return (EXPENSE_CATEGORIES as readonly string[]).includes(text);
}

export interface ExpenseItem {
    category: ExpenseCategory;
    /** 品項／項目名稱（人事區塊無名稱時會退回用人名） */
    name: string;
    amount: number;
    /** 同列其他欄位的文字（廠商、月份、人名、備註…），以「・」串接 */
    note: string;
    /** 來源儲存格（例如 E4），方便對照 Excel */
    cell: string;
}

export interface IncomeItem {
    name: string;
    amount: number;
}

export interface MonthExpense {
    /** 民國年 */
    year: number;
    month: number;
    /** `${year}-${MM}` */
    sortKey: string;
    sheetName: string;
    /** 各大類小計（以 Excel 小計格為準；找不到小計時用明細加總） */
    totals: Record<ExpenseCategory, number>;
    items: ExpenseItem[];
    /** Excel 左側「收入」欄（僅供參考對照，營收以網站帳單為準） */
    income: IncomeItem[];
    incomeTotal: number;
    /** 解析時的提醒（小計與明細加總不符等） */
    warnings: string[];
    sourceFile: string;
    uploadedAt: number;
}

export function monthSortKey(year: number, month: number): string {
    return `${year}-${String(month).padStart(2, '0')}`;
}

export function emptyTotals(): Record<ExpenseCategory, number> {
    return {原物料: 0, 營運: 0, 車輛: 0, 人事: 0, 其他: 0, 銀行: 0};
}

export function sumTotals(totals: Record<ExpenseCategory, number>): number {
    return EXPENSE_CATEGORIES.reduce((s, c) => s + totals[c], 0);
}
