/**
 * A1 月度趨勢折線圖 option builder。
 * 可切換指標：營收 / 毛利 / 數量 / 客戶數。
 */

import type {EChartsOption} from 'echarts';
import type {MonthlyTotals} from './month-aggregators';
import {PALETTE} from '@/analytics/chart-colors';

export type TrendMetric = 'amount' | 'profit' | 'count' | 'customerCount';

export interface TrendMetricSpec {
    key: TrendMetric;
    label: string;
    color: string;
    /** 是否錢類（影響 axis label 與 tooltip 格式） */
    isMoney: boolean;
}

export const TREND_METRICS: ReadonlyArray<TrendMetricSpec> = [
    {key: 'amount', label: '營收', color: PALETTE.revenue, isMoney: true},
    {key: 'profit', label: '毛利', color: PALETTE.profit, isMoney: true},
    {key: 'count', label: '數量', color: PALETTE.accent, isMoney: false},
    {key: 'customerCount', label: '客戶數', color: PALETTE.cost, isMoney: false},
];

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');

function metricValue(t: MonthlyTotals, m: TrendMetric): number {
    if (m === 'amount') return t.amount;
    if (m === 'profit') return Math.round(t.profit);
    if (m === 'count') return t.count;
    return t.customerCount;
}

export function monthlyTrendOption(
    monthly: ReadonlyArray<MonthlyTotals>,
    metric: TrendMetric
): EChartsOption {
    const spec = TREND_METRICS.find((m) => m.key === metric)!;
    const xData = monthly.map((m) => m.key.label);
    const values = monthly.map((m) => metricValue(m, metric));
    const valueFmt = spec.isMoney ? fmtMoney : (v: number) => String(v);

    return {
        grid: {top: 30, right: 30, bottom: 50, left: 70},
        tooltip: {
            trigger: 'axis',
            formatter: (raw) => {
                const params = Array.isArray(raw) ? raw : [raw];
                const p = params[0] as { name?: string; value?: number; marker?: string; dataIndex?: number };
                const idx = p.dataIndex ?? 0;
                const t = monthly[idx];
                if (!t) return '';
                const head = `<div style="font-weight:600">${p.name ?? ''}</div>`
                    + `${p.marker ?? ''}${spec.label}　${valueFmt(Number(p.value ?? 0))}`;
                const meta = `<div style="margin-top:6px;border-top:1px dashed #ddd;padding-top:4px">`
                    + `<div>營收　　${fmtMoney(t.amount)}</div>`
                    + `<div>毛利　　${fmtMoney(Math.round(t.profit))}</div>`
                    + `<div>數量　　${t.count.toLocaleString('zh-TW')}</div>`
                    + `<div>客戶數　${t.customerCount}</div>`
                    + `</div>`;
                return head + meta;
            },
        },
        xAxis: {
            type: 'category',
            data: xData,
            axisLabel: {fontSize: 11},
        },
        yAxis: {
            type: 'value',
            axisLabel: {formatter: (v: number) => valueFmt(v)},
        },
        series: [
            {
                name: spec.label,
                type: 'line',
                smooth: true,
                symbol: 'circle',
                symbolSize: 8,
                data: values,
                areaStyle: {opacity: 0.15, color: spec.color},
                itemStyle: {color: spec.color},
                lineStyle: {width: 2.5, color: spec.color},
                label: {
                    show: true,
                    position: 'top',
                    fontSize: 11,
                    formatter: (p) => valueFmt(Number(p.value ?? 0)),
                },
            },
        ],
    };
}
