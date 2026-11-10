/**
 * 把「月支出」與「網站營收」對齊成每月一列的損益摘要，並提供年度累計。
 *
 * 結餘分三層（「銀行」是還貸款本金，不算費用，畫面上稱「還款」）：
 * - 營業費用 = 原物料 + 營運 + 車輛 + 人事 + 其他（店務部分，不含家用）
 * - 營業結餘 = 網站營收 − 營業費用
 * - 還款後結餘 = 營業結餘 − 還款
 * - 實際剩下的錢 = 還款後結餘 − 家用（＝ 網站營收 − 全部支出）
 */

import {EXPENSE_CATEGORIES, type ExpenseCategory, type MonthExpense, sumTotals} from './expense-model';
import {isHouseholdName} from './expense-store';
import type {MonthRevenue} from './revenue-snapshot';

/** 收支表大類名稱（資料結構沿用）→ 畫面顯示名稱。 */
export function categoryLabel(cat: ExpenseCategory): string {
    return cat === '銀行' ? '還款' : cat;
}

export interface MonthSummary {
    sortKey: string;
    year: number;
    month: number;
    expense: MonthExpense;
    /** 沒有載入該月帳單時為 null */
    revenue: number | null;
    /** 原物料：Excel 實付 */
    materialPaid: number;
    totals: Record<ExpenseCategory, number>;
    /** 「其他」中算家用的金額（預設全部，扣掉被取消勾選＝店務的項目） */
    household: number;
    /** 「其他」中算店務的金額 */
    otherShop: number;
    /** 還款（收支表「銀行」大類，還貸款本金，不算費用） */
    repayment: number;
    /** 營業費用：原物料＋營運＋車輛＋人事＋其他店務（不含還款、家用） */
    operatingExpense: number;
    /** 店務支出（全部支出 − 家用；含還款） */
    shopExpense: number;
    /** 全部支出（含還款、家用，等於 Excel 月支出合計） */
    totalExpense: number;
    /** 營收 − Excel 原物料實付 */
    grossProfit: number | null;
    /** 營收 − 營業費用 */
    operatingBalance: number | null;
    /** 營業結餘 − 還款 */
    afterRepayBalance: number | null;
    /** 還款後結餘 − 家用（＝ 營收 − 全部支出） */
    netBalance: number | null;
}

/** @param shopNames 被取消家用勾選（=店務）的「其他」項目名稱；其餘「其他」項目一律算家用 */
export function summarizeMonth(expense: MonthExpense, revenue: MonthRevenue | undefined, shopNames: ReadonlySet<string>): MonthSummary {
    const householdAmount = expense.items
        .filter((it) => it.category === '其他' && isHouseholdName(it.name, shopNames))
        .reduce((s, it) => s + it.amount, 0);
    const totalExpense = sumTotals(expense.totals);
    const repayment = expense.totals['銀行'];
    const operatingExpense = totalExpense - repayment - householdAmount;
    const rev = revenue ? revenue.amount : null;
    const operatingBalance = rev === null ? null : rev - operatingExpense;
    const afterRepayBalance = operatingBalance === null ? null : operatingBalance - repayment;
    return {
        sortKey: expense.sortKey,
        year: expense.year,
        month: expense.month,
        expense,
        revenue: rev,
        materialPaid: expense.totals['原物料'],
        totals: expense.totals,
        household: householdAmount,
        otherShop: expense.totals['其他'] - householdAmount,
        repayment,
        operatingExpense,
        shopExpense: totalExpense - householdAmount,
        totalExpense,
        grossProfit: rev === null ? null : rev - expense.totals['原物料'],
        operatingBalance,
        afterRepayBalance,
        netBalance: afterRepayBalance === null ? null : afterRepayBalance - householdAmount,
    };
}

export interface YearSummary {
    months: number;
    /** 有網站營收的月份數 */
    revenueMonths: number;
    revenue: number;
    materialPaid: number;
    totals: Record<ExpenseCategory, number>;
    household: number;
    otherShop: number;
    repayment: number;
    operatingExpense: number;
    shopExpense: number;
    totalExpense: number;
    /** 以下四項只累計有網站營收的月份 */
    grossProfit: number;
    operatingBalance: number;
    afterRepayBalance: number;
    netBalance: number;
}

export function summarizeYear(rows: ReadonlyArray<MonthSummary>): YearSummary {
    const totals = Object.fromEntries(EXPENSE_CATEGORIES.map((c) => [c, 0])) as Record<ExpenseCategory, number>;
    const y: YearSummary = {
        months: rows.length,
        revenueMonths: 0,
        revenue: 0,
        materialPaid: 0,
        totals,
        household: 0,
        otherShop: 0,
        repayment: 0,
        operatingExpense: 0,
        shopExpense: 0,
        totalExpense: 0,
        grossProfit: 0,
        operatingBalance: 0,
        afterRepayBalance: 0,
        netBalance: 0,
    };
    for (const r of rows) {
        y.materialPaid += r.materialPaid;
        for (const c of EXPENSE_CATEGORIES) totals[c] += r.totals[c];
        y.household += r.household;
        y.otherShop += r.otherShop;
        y.repayment += r.repayment;
        y.operatingExpense += r.operatingExpense;
        y.shopExpense += r.shopExpense;
        y.totalExpense += r.totalExpense;
        if (r.revenue !== null) {
            y.revenueMonths += 1;
            y.revenue += r.revenue;
            y.grossProfit += r.grossProfit ?? 0;
            y.operatingBalance += r.operatingBalance ?? 0;
            y.afterRepayBalance += r.afterRepayBalance ?? 0;
            y.netBalance += r.netBalance ?? 0;
        }
    }
    return y;
}
