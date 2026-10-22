/**
 * 把「月支出」與「網站營收」對齊成每月一列的損益摘要，並提供年度累計。
 */

import {EXPENSE_CATEGORIES, type ExpenseCategory, type MonthExpense, sumTotals} from './expense-model';
import {householdKey} from './expense-store';
import type {MonthRevenue} from './revenue-snapshot';

export interface MonthSummary {
    sortKey: string;
    year: number;
    month: number;
    expense: MonthExpense;
    /** 沒有載入該月帳單時為 null */
    revenue: number | null;
    /** 原物料：Excel 實付 */
    materialPaid: number;
    /** 原物料：網站推算（數量 × 商品成本） */
    materialEst: number | null;
    unsetCostProducts: string[];
    totals: Record<ExpenseCategory, number>;
    /** 「其他」中被勾為家用的金額 */
    household: number;
    /** 店務支出（全部支出 − 家用） */
    shopExpense: number;
    /** 全部支出（含家用，等於 Excel 月支出合計） */
    totalExpense: number;
    /** 營收 − Excel 原物料實付 */
    grossProfit: number | null;
    /** 營收 − 店務支出 */
    shopBalance: number | null;
    /** 營收 − 全部支出 */
    allBalance: number | null;
}

export function summarizeMonth(expense: MonthExpense, revenue: MonthRevenue | undefined, household: ReadonlySet<string>): MonthSummary {
    const householdAmount = expense.items
        .filter((it) => it.category === '其他' && household.has(householdKey(it.name)))
        .reduce((s, it) => s + it.amount, 0);
    const totalExpense = sumTotals(expense.totals);
    const shopExpense = totalExpense - householdAmount;
    const rev = revenue ? revenue.amount : null;
    return {
        sortKey: expense.sortKey,
        year: expense.year,
        month: expense.month,
        expense,
        revenue: rev,
        materialPaid: expense.totals['原物料'],
        materialEst: revenue ? revenue.estCost : null,
        unsetCostProducts: revenue?.unsetCostProducts ?? [],
        totals: expense.totals,
        household: householdAmount,
        shopExpense,
        totalExpense,
        grossProfit: rev === null ? null : rev - expense.totals['原物料'],
        shopBalance: rev === null ? null : rev - shopExpense,
        allBalance: rev === null ? null : rev - totalExpense,
    };
}

export interface YearSummary {
    months: number;
    /** 有網站營收的月份數 */
    revenueMonths: number;
    revenue: number;
    materialPaid: number;
    /** 只累計有網站營收的月份，方便與 materialEst 對照 */
    materialPaidWithRevenue: number;
    materialEst: number;
    totals: Record<ExpenseCategory, number>;
    household: number;
    shopExpense: number;
    totalExpense: number;
    /** 以下三項只累計有網站營收的月份 */
    grossProfit: number;
    shopBalance: number;
    allBalance: number;
}

export function summarizeYear(rows: ReadonlyArray<MonthSummary>): YearSummary {
    const totals = Object.fromEntries(EXPENSE_CATEGORIES.map((c) => [c, 0])) as Record<ExpenseCategory, number>;
    const y: YearSummary = {
        months: rows.length,
        revenueMonths: 0,
        revenue: 0,
        materialPaid: 0,
        materialPaidWithRevenue: 0,
        materialEst: 0,
        totals,
        household: 0,
        shopExpense: 0,
        totalExpense: 0,
        grossProfit: 0,
        shopBalance: 0,
        allBalance: 0,
    };
    for (const r of rows) {
        y.materialPaid += r.materialPaid;
        for (const c of EXPENSE_CATEGORIES) totals[c] += r.totals[c];
        y.household += r.household;
        y.shopExpense += r.shopExpense;
        y.totalExpense += r.totalExpense;
        if (r.revenue !== null) {
            y.revenueMonths += 1;
            y.revenue += r.revenue;
            y.materialPaidWithRevenue += r.materialPaid;
            y.materialEst += r.materialEst ?? 0;
            y.grossProfit += r.grossProfit ?? 0;
            y.shopBalance += r.shopBalance ?? 0;
            y.allBalance += r.allBalance ?? 0;
        }
    }
    return y;
}
