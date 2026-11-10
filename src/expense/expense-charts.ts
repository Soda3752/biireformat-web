/**
 * 「支出」分頁的 ECharts option builder。
 */

import type {BarSeriesOption, EChartsOption} from 'echarts';
import {CATEGORY_PALETTE, PALETTE} from '@/analytics/chart-colors';
import {EXPENSE_CATEGORIES} from './expense-model';
import {categoryLabel, type MonthSummary} from './expense-summary';
import {HEALTHY_MATERIAL_LABOR_RATIO, type RollingMargin, SHARE_KEYS, type ShareRow} from './expense-analysis';

const fmtMoney = (v: number) => Math.round(v).toLocaleString('zh-TW');
const moneyTooltip = (v: unknown) => (typeof v === 'number' ? fmtMoney(v) : '—');

const HOUSEHOLD_COLOR = '#bdc3c7';

/** 每月 網站營收 / 營業費用 長條 ＋ 營業結餘、還款後結餘、實際剩下的錢 折線 */
export function balanceOption(rows: ReadonlyArray<MonthSummary>): EChartsOption {
    return {
        tooltip: {trigger: 'axis', valueFormatter: moneyTooltip},
        legend: {top: 0, type: 'scroll'},
        grid: {left: 8, right: 8, top: 36, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => `${r.month}月`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {name: '網站營收', type: 'bar', data: rows.map((r) => r.revenue), itemStyle: {color: PALETTE.revenue}},
            {name: '營業費用', type: 'bar', data: rows.map((r) => r.operatingExpense), itemStyle: {color: PALETTE.cost}},
            {
                name: '營業結餘',
                type: 'line',
                data: rows.map((r) => r.operatingBalance),
                itemStyle: {color: PALETTE.profit},
                lineStyle: {width: 3},
                connectNulls: false,
            },
            {
                name: '還款後結餘',
                type: 'line',
                data: rows.map((r) => r.afterRepayBalance),
                itemStyle: {color: PALETTE.accent},
                connectNulls: false,
            },
            {
                name: '實際剩下的錢',
                type: 'line',
                data: rows.map((r) => r.netBalance),
                itemStyle: {color: '#e74c3c'},
                lineStyle: {type: 'dashed'},
                connectNulls: false,
            },
        ],
    };
}

/** 每月支出組成（堆疊）：其他拆成「其他（店務）」與「家用」 */
export function compositionOption(rows: ReadonlyArray<MonthSummary>): EChartsOption {
    const series: BarSeriesOption[] = EXPENSE_CATEGORIES.map((cat, i) => ({
        name: cat === '其他' ? '其他（店務）' : categoryLabel(cat),
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

/** 每月原物料（Excel 實付） */
export function materialOption(rows: ReadonlyArray<MonthSummary>): EChartsOption {
    return {
        tooltip: {trigger: 'axis', valueFormatter: moneyTooltip},
        grid: {left: 8, right: 8, top: 16, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => `${r.month}月`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {name: 'Excel 實付', type: 'bar', data: rows.map((r) => r.materialPaid), itemStyle: {color: PALETTE.cost}},
        ],
    };
}

const pctLabel = (v: number) => `${Math.round(v)}%`;
const pctTooltip = (v: unknown) => (typeof v === 'number' ? `${v.toFixed(1)}%` : '—');
const toPct = (v: number | null) => (v === null ? null : Math.round(v * 1000) / 10);

/** 各項占營收 %（堆疊）＋「原物料＋人事」折線與一般參考線 */
export function shareOption(rows: ReadonlyArray<ShareRow>): EChartsOption {
    const colors: Record<string, string> = {
        原物料: CATEGORY_PALETTE[0],
        人事: CATEGORY_PALETTE[3],
        營運: CATEGORY_PALETTE[1],
        車輛: CATEGORY_PALETTE[2],
        其他店務: CATEGORY_PALETTE[4],
        還款: CATEGORY_PALETTE[5],
        家用: HOUSEHOLD_COLOR,
        營業結餘: PALETTE.profit,
    };
    const stackKeys = SHARE_KEYS.filter((k) => k !== '營業結餘');
    const series: EChartsOption['series'] = [
        ...stackKeys.map((k) => ({
            name: k,
            type: 'bar' as const,
            stack: 'share',
            data: rows.map((r) => toPct(r.pct[k])),
            itemStyle: {color: colors[k]},
        })),
        {
            name: '原物料＋人事',
            type: 'line' as const,
            data: rows.map((r) => toPct(r.materialLabor)),
            itemStyle: {color: '#2c3e50'},
            lineStyle: {width: 2},
            markLine: {
                silent: true,
                symbol: 'none',
                lineStyle: {type: 'dashed', color: '#e74c3c'},
                label: {formatter: `一般參考 ${Math.round(HEALTHY_MATERIAL_LABOR_RATIO * 100)}%`, position: 'insideEndTop'},
                data: [{yAxis: HEALTHY_MATERIAL_LABOR_RATIO * 100}],
            },
        },
    ];
    return {
        tooltip: {trigger: 'axis', axisPointer: {type: 'shadow'}, valueFormatter: pctTooltip},
        legend: {top: 0, type: 'scroll'},
        grid: {left: 8, right: 8, top: 36, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => (r.month === null ? '年度' : `${r.month}月`)), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: pctLabel}},
        series,
    };
}

/** 單月毛利率 vs 近 3 個月平均毛利率 */
export function marginOption(rows: ReadonlyArray<RollingMargin>): EChartsOption {
    return {
        tooltip: {trigger: 'axis', valueFormatter: pctTooltip},
        legend: {top: 0},
        grid: {left: 8, right: 8, top: 36, bottom: 8, containLabel: true},
        xAxis: {type: 'category', data: rows.map((r) => `${r.month}月`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', scale: true, axisLabel: {formatter: pctLabel}},
        series: [
            {
                name: '單月毛利率',
                type: 'line',
                data: rows.map((r) => toPct(r.margin)),
                itemStyle: {color: '#bdc3c7'},
                lineStyle: {type: 'dashed'},
            },
            {
                name: '近 3 個月平均',
                type: 'line',
                data: rows.map((r) => toPct(r.average)),
                itemStyle: {color: PALETTE.profit},
                lineStyle: {width: 3},
            },
        ],
    };
}
