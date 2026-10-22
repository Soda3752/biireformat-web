/**
 * 「支出」分頁的 ECharts option builder。
 */

import type {BarSeriesOption, EChartsOption} from 'echarts';
import {CATEGORY_PALETTE, PALETTE} from '@/analytics/chart-colors';
import {EXPENSE_CATEGORIES} from './expense-model';
import type {MonthSummary} from './expense-summary';

const fmtMoney = (v: number) => Math.round(v).toLocaleString('zh-TW');
const moneyTooltip = (v: unknown) => (typeof v === 'number' ? fmtMoney(v) : '—');

const HOUSEHOLD_COLOR = '#bdc3c7';

/** 每月 網站營收 / 店務支出 長條 ＋ 店務結餘、含家用結餘 折線 */
export function balanceOption(rows: ReadonlyArray<MonthSummary>): EChartsOption {
    return {
        tooltip: {trigger: 'axis', valueFormatter: moneyTooltip},
        legend: {top: 0},
        grid: {left: 8, right: 8, top: 36, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => `${r.month}月`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {name: '網站營收', type: 'bar', data: rows.map((r) => r.revenue), itemStyle: {color: PALETTE.revenue}},
            {name: '店務支出', type: 'bar', data: rows.map((r) => r.shopExpense), itemStyle: {color: PALETTE.cost}},
            {
                name: '店務結餘',
                type: 'line',
                data: rows.map((r) => r.shopBalance),
                itemStyle: {color: PALETTE.profit},
                lineStyle: {width: 3},
                connectNulls: false,
            },
            {
                name: '含家用結餘',
                type: 'line',
                data: rows.map((r) => r.allBalance),
                itemStyle: {color: PALETTE.accent},
                lineStyle: {type: 'dashed'},
                connectNulls: false,
            },
        ],
    };
}

/** 每月支出組成（堆疊）：其他拆成「其他（店務）」與「家用」 */
export function compositionOption(rows: ReadonlyArray<MonthSummary>): EChartsOption {
    const series: BarSeriesOption[] = EXPENSE_CATEGORIES.map((cat, i) => ({
        name: cat === '其他' ? '其他（店務）' : cat,
        type: 'bar' as const,
        stack: 'expense',
        data: rows.map((r) => (cat === '其他' ? r.totals[cat] - r.household : r.totals[cat])),
        itemStyle: {color: CATEGORY_PALETTE[i % CATEGORY_PALETTE.length]},
    }));
    series.push({
        name: '家用',
        type: 'bar',
        stack: 'expense',
        data: rows.map((r) => r.household),
        itemStyle: {color: HOUSEHOLD_COLOR},
    });
    return {
        tooltip: {trigger: 'axis', axisPointer: {type: 'shadow'}, valueFormatter: moneyTooltip},
        legend: {top: 0, type: 'scroll'},
        grid: {left: 8, right: 8, top: 36, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => `${r.month}月`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series,
    };
}

/** 原物料：Excel 實付 vs 網站推算 */
export function materialOption(rows: ReadonlyArray<MonthSummary>): EChartsOption {
    return {
        tooltip: {trigger: 'axis', valueFormatter: moneyTooltip},
        legend: {top: 0},
        grid: {left: 8, right: 8, top: 36, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => `${r.month}月`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {name: 'Excel 實付', type: 'bar', data: rows.map((r) => r.materialPaid), itemStyle: {color: PALETTE.cost}},
            {name: '網站推算', type: 'bar', data: rows.map((r) => r.materialEst), itemStyle: {color: CATEGORY_PALETTE[4]}},
        ],
    };
}
