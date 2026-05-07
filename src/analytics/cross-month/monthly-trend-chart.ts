/**
 * 跨月趨勢折線圖 option builder。
 * 輸入為時間升冪的「趨勢點」陣列，可為月度或每 5 日為一段的細粒度資料。
 * 可切換指標：營收 / 毛利 / 數量 / 客戶數。
 */

import type {EChartsOption} from 'echarts';
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

/** 趨勢圖最少結構：x 軸 label + 四個指標。月度與 5 日段皆相容。 */
export interface TrendPoint {
    key: {
        /** x 軸顯示用緊湊標籤 */
        label: string;
        /** tooltip 顯示用完整標籤；缺省時退回 label */
        fullLabel?: string;
    };
    amount: number;
    count: number;
    profit: number;
    customerCount: number;
}

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');

function metricValue(t: TrendPoint, m: TrendMetric): number {
    if (m === 'amount') return t.amount;
    if (m === 'profit') return Math.round(t.profit);
    if (m === 'count') return t.count;
    return t.customerCount;
}

export function monthlyTrendOption(
    points: ReadonlyArray<TrendPoint>,
    metric: TrendMetric
): EChartsOption {
    const spec = TREND_METRICS.find((m) => m.key === metric)!;
    const xData = points.map((m) => m.key.label);
    const values = points.map((m) => metricValue(m, metric));
    const valueFmt = spec.isMoney ? fmtMoney : (v: number) => String(v);
    // 點數越多，標籤越擠 → 自動旋轉、隱藏值標籤避免互相覆蓋
    const dense = points.length > 8;

    return {
        grid: {top: 30, right: 30, bottom: dense ? 70 : 50, left: 70},
        tooltip: {
            trigger: 'axis',
            formatter: (raw) => {
                const params = Array.isArray(raw) ? raw : [raw];
                const p = params[0] as { name?: string; value?: number; marker?: string; dataIndex?: number };
                const idx = p.dataIndex ?? 0;
                const t = points[idx];
                if (!t) return '';
                const title = t.key.fullLabel ?? p.name ?? t.key.label;
                const head = `<div style="font-weight:600">${title}</div>`
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
            axisLabel: {
                fontSize: 11,
                interval: 0,
                rotate: dense ? 30 : 0,
            },
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
                symbolSize: dense ? 6 : 8,
                data: values,
                areaStyle: {opacity: 0.15, color: spec.color},
                itemStyle: {color: spec.color},
                lineStyle: {width: 2.5, color: spec.color},
                label: {
                    show: !dense,
                    position: 'top',
                    fontSize: 11,
                    formatter: (p) => valueFmt(Number(p.value ?? 0)),
                },
            },
        ],
    };
}
