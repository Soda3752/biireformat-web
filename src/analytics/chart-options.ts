/**
 * 12 張 ECharts 圖的 option builder。
 * 每個函式接 AnalyticsRow[]（已套篩選）+ 視需要的設定，回傳 EChartsOption。
 *
 * 統一規約：
 * - 不引用 echarts 模組（型別來自 EChartsOption），讓本檔可單元測試
 * - tooltip 走 axis trigger 為主，符合中文閱讀習慣
 * - 顏色集中於 chart-colors.ts，避免 hex 散落
 * - 所有圖 tooltip 共用 formatProfitFooter() 輸出成本／毛利／毛利率
 */

import type {EChartsOption} from 'echarts';
import type {AnalyticsRow} from './dataset-builder';
import type {GroupSum} from './aggregators';
import {dailySeries, groupBy, groupByCustomer, marginPct, topN, weekdaySeries} from './aggregators';
import {WEEKDAY_NAMES} from '@/domain/date-utility';
import {CATEGORY_PALETTE, lightenHex, PALETTE} from './chart-colors';

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');
const fmtPct = (v: number) => `${v.toFixed(1)}%`;

export type MetricKind = 'amount' | 'count' | 'profit';

/** 取得指定 metric 對應的數值 */
function metricValueFromGroup(g: GroupSum, metric: MetricKind): number {
    if (metric === 'count') return g.count;
    if (metric === 'profit') return Math.round(g.profit);
    return g.amount;
}

function metricValueFromDaily(p: { amount: number; count: number; profit: number }, metric: MetricKind): number {
    if (metric === 'count') return p.count;
    if (metric === 'profit') return Math.round(p.profit);
    return p.amount;
}

function metricValueFromRow(r: AnalyticsRow, metric: MetricKind): number {
    if (metric === 'count') return r.count;
    if (metric === 'profit') return r.profit;
    return r.amount;
}

function metricSeriesName(metric: MetricKind): string {
    if (metric === 'count') return '銷售數量';
    if (metric === 'profit') return '毛利';
    return '營收';
}

function metricColor(metric: MetricKind): string {
    if (metric === 'profit') return PALETTE.profit;
    return PALETTE.revenue;
}

function metricValueFmt(metric: MetricKind): (v: number) => string {
    return metric === 'count' ? (v: number) => String(v) : fmtMoney;
}

interface ProfitMeta {
    amount: number;
    costAmount: number;
    profit: number;
    allCostUnset: boolean;
}

function formatProfitFooter(meta: ProfitMeta | undefined | null): string {
    if (!meta) return '';
    const pct = marginPct(meta.profit, meta.amount);
    const pctText = pct === null ? '—' : fmtPct(pct);
    const warn = meta.allCostUnset && meta.amount > 0
        ? `<div style="margin-top:4px;font-size:11px;color:#999">⚠ 部分商品未填成本，計算視為 0</div>`
        : '';
    return `<div style="margin-top:6px;border-top:1px dashed #ddd;padding-top:4px">`
        + `<div>成本　　${fmtMoney(Math.round(meta.costAmount))}</div>`
        + `<div>毛利　　${fmtMoney(Math.round(meta.profit))}</div>`
        + `<div>毛利率　${pctText}</div>`
        + warn
        + `</div>`;
}

interface AxisTooltipParam {
    marker?: string;
    seriesName?: string;
    name?: string;
    axisValueLabel?: string;
    dataIndex?: number;
    value?: number | string;
}

function buildAxisTooltip(
    raw: AxisTooltipParam | AxisTooltipParam[],
    metas: ReadonlyArray<ProfitMeta>,
    valueFmt: (v: number) => string = fmtMoney,
): string {
    const params = Array.isArray(raw) ? raw : [raw];
    const first = params[0];
    const idx = first.dataIndex ?? 0;
    const meta = metas[idx];
    const head = params
        .map((p) => `${p.marker ?? ''}${p.seriesName ?? ''}　${valueFmt(Number(p.value ?? 0))}`)
        .join('<br/>');
    const title = `<div style="font-weight:600">${first.axisValueLabel ?? first.name ?? ''}</div>`;
    return title + head + formatProfitFooter(meta);
}

/* ================ T1.2 每日營收趨勢 ================ */
export function dailyTrendOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind): EChartsOption {
    const points = dailySeries(rows);
    const xData = points.map((p) => `${p.day}日`);
    const isAmount = metric === 'amount';
    const valueFmt = metricValueFmt(metric);

    const series: NonNullable<EChartsOption['series']> = isAmount
        ? [
            {
                name: '營收',
                type: 'line',
                smooth: true,
                symbol: 'circle',
                symbolSize: 6,
                data: points.map((p) => p.amount),
                areaStyle: {opacity: 0.15, color: PALETTE.revenue},
                itemStyle: {color: PALETTE.revenue},
                lineStyle: {width: 2, color: PALETTE.revenue},
            },
            {
                name: '毛利',
                type: 'line',
                smooth: true,
                symbol: 'circle',
                symbolSize: 5,
                data: points.map((p) => Math.round(p.profit)),
                itemStyle: {color: PALETTE.profit},
                lineStyle: {width: 2, color: PALETTE.profit},
            },
        ]
        : [
            {
                name: metricSeriesName(metric),
                type: 'line',
                smooth: true,
                symbol: 'circle',
                symbolSize: 6,
                data: points.map((p) => metricValueFromDaily(p, metric)),
                areaStyle: {opacity: 0.15, color: metricColor(metric)},
                itemStyle: {color: metricColor(metric)},
                lineStyle: {width: 2, color: metricColor(metric)},
            },
        ];

    return {
        grid: {top: isAmount ? 40 : 30, right: 20, bottom: 40, left: 60},
        legend: isAmount ? {top: 0, left: 'center', textStyle: {fontSize: 11}} : {show: false},
        tooltip: {
            trigger: 'axis',
            formatter: (raw) => buildAxisTooltip(raw as AxisTooltipParam | AxisTooltipParam[], points, valueFmt),
        },
        xAxis: {type: 'category', data: xData, axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series,
    };
}

/* ================ T1.3 線別佔比（金額/數量/利潤切換） ================ */
export function linePieOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind = 'amount'): EChartsOption {
    const groups = groupBy(rows, 'line', metric, true);
    const groupMap = new Map(groups.map((g) => [g.key, g]));
    return {
        tooltip: {
            trigger: 'item',
            formatter: (raw) => {
                const p = Array.isArray(raw) ? raw[0] : raw;
                const name = String(p.name ?? '');
                const g = groupMap.get(name);
                const pctText = typeof p.percent === 'number' ? `（${p.percent.toFixed(1)}%）` : '';
                const head = `<div style="font-weight:600">${name}</div>${p.marker ?? ''}${fmtMoney(Number(p.value ?? 0))}${pctText}`;
                return head + formatProfitFooter(g);
            },
        },
        legend: {
            bottom: 0,
            left: 'center',
            textStyle: {fontSize: 11},
            itemWidth: 12,
            itemHeight: 8,
            itemGap: 8,
        },
        series: [
            {
                type: 'pie',
                radius: ['38%', '58%'],
                center: ['50%', '42%'],
                avoidLabelOverlap: true,
                minAngle: 2,
                label: {
                    formatter: '{b} {d}%',
                    fontSize: 11,
                    overflow: 'truncate',
                    width: 80,
                },
                labelLine: {
                    length: 8,
                    length2: 8,
                    smooth: true,
                },
                labelLayout: {
                    hideOverlap: true,
                },
                data: groups.map((g) => ({name: g.key, value: metricValueFromGroup(g, metric)})),
            },
        ],
    };
}

/* ================ 路線分析：每日折線（每條線一條線） ================ */
export function routeDailyLineOption(
    rows: ReadonlyArray<AnalyticsRow>,
    metric: MetricKind,
    lines: ReadonlyArray<string>,
    colorOf: (line: string) => string
): EChartsOption {
    const sel = new Set(lines);
    const dayset = new Set<number>();
    const perLine = new Map<string, Map<number, number>>();
    for (const r of rows) {
        if (!sel.has(r.line)) continue;
        dayset.add(r.day);
        let m = perLine.get(r.line);
        if (!m) {
            m = new Map();
            perLine.set(r.line, m);
        }
        m.set(r.day, (m.get(r.day) ?? 0) + metricValueFromRow(r, metric));
    }
    const days = [...dayset].sort((a, b) => a - b);
    const valueFmt = metricValueFmt(metric);
    const series: NonNullable<EChartsOption['series']> = lines.map((line) => ({
        name: line,
        type: 'line',
        smooth: true,
        symbol: 'circle',
        symbolSize: 5,
        data: days.map((d) => Math.round(perLine.get(line)?.get(d) ?? 0)),
        itemStyle: {color: colorOf(line)},
        lineStyle: {width: 2, color: colorOf(line)},
    }));
    return {
        grid: {top: 20, right: 20, bottom: 50, left: 60},
        legend: {bottom: 0, left: 'center', textStyle: {fontSize: 11}, itemWidth: 14, itemHeight: 8, itemGap: 8},
        tooltip: {trigger: 'axis', valueFormatter: (v) => valueFmt(Number(v))},
        xAxis: {type: 'category', data: days.map((d) => `${d}日`), axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series,
    };
}

/* ================ 路線分析：佔比圓餅 ================ */
export function routePieOption(
    rows: ReadonlyArray<AnalyticsRow>,
    metric: MetricKind,
    lines: ReadonlyArray<string>,
    colorOf: (line: string) => string
): EChartsOption {
    const gmap = new Map(groupBy(rows, 'line', metric, true).map((g) => [g.key, g]));
    const valueFmt = metricValueFmt(metric);
    // 圓餅面積不可為負，利潤虧損取 0（tooltip 仍顯示真實毛利）。
    const data = lines.map((line) => {
        const g = gmap.get(line);
        return {name: line, value: g ? Math.max(0, metricValueFromGroup(g, metric)) : 0, itemStyle: {color: colorOf(line)}};
    });
    return {
        tooltip: {
            trigger: 'item',
            formatter: (raw) => {
                const p = Array.isArray(raw) ? raw[0] : raw;
                const name = String(p.name ?? '');
                const g = gmap.get(name);
                const pctText = typeof p.percent === 'number' ? `（${p.percent.toFixed(1)}%）` : '';
                const head = `<div style="font-weight:600">${name}</div>${p.marker ?? ''}${valueFmt(Number(p.value ?? 0))}${pctText}`;
                return head + formatProfitFooter(g);
            },
        },
        legend: {bottom: 0, left: 'center', textStyle: {fontSize: 11}, itemWidth: 12, itemHeight: 8, itemGap: 8},
        series: [
            {
                type: 'pie',
                radius: ['38%', '58%'],
                center: ['50%', '42%'],
                avoidLabelOverlap: true,
                minAngle: 2,
                label: {formatter: '{b} {d}%', fontSize: 11, overflow: 'truncate', width: 80},
                labelLine: {length: 8, length2: 8, smooth: true},
                labelLayout: {hideOverlap: true},
                data,
            },
        ],
    };
}

/* ================ 路線分析：總計條狀（各線比較） ================ */
export function routeBarOption(
    rows: ReadonlyArray<AnalyticsRow>,
    metric: MetricKind,
    lines: ReadonlyArray<string>,
    colorOf: (line: string) => string
): EChartsOption {
    const sel = new Set(lines);
    // groupBy 已依 metric 由大到小排序，過濾出勾選的線即為排名順序。
    const ordered = groupBy(rows, 'line', metric, true).filter((g) => sel.has(g.key));
    const valueFmt = metricValueFmt(metric);
    return {
        grid: {top: 20, right: 20, bottom: 30, left: 60},
        tooltip: {
            trigger: 'axis',
            axisPointer: {type: 'shadow'},
            formatter: (raw) => buildAxisTooltip(raw as AxisTooltipParam | AxisTooltipParam[], ordered, valueFmt),
        },
        xAxis: {type: 'category', data: ordered.map((g) => g.key), axisLabel: {fontSize: 11, interval: 0}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {
                name: metricSeriesName(metric),
                type: 'bar',
                data: ordered.map((g) => ({
                    value: metricValueFromGroup(g, metric),
                    itemStyle: {color: colorOf(g.key), borderRadius: [4, 4, 0, 0]},
                })),
                barMaxWidth: 40,
                label: {show: true, position: 'top', fontSize: 11, formatter: (p) => valueFmt(Number(p.value))},
            },
        ],
    };
}

/* ================ T1.4 商品 Top 10 ================ */
export function productTopOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind): EChartsOption {
    const groups = topN(groupBy(rows, 'productName', metric, true), 10).reverse();
    const isAmount = metric === 'amount';
    const valueFmt = metricValueFmt(metric);

    const series: NonNullable<EChartsOption['series']> = [
        {
            name: metricSeriesName(metric),
            type: 'bar',
            data: groups.map((g) => metricValueFromGroup(g, metric)),
            itemStyle: {borderRadius: [0, 4, 4, 0], color: metricColor(metric)},
            barMaxWidth: 22,
            label: {show: true, position: 'right', fontSize: 11, formatter: (p) => valueFmt(Number(p.value))},
        },
    ];
    if (isAmount) {
        series.push({
            name: '毛利',
            type: 'bar',
            data: groups.map((g) => Math.round(g.profit)),
            itemStyle: {borderRadius: [0, 4, 4, 0], color: PALETTE.profit},
            barMaxWidth: 22,
            label: {show: true, position: 'right', fontSize: 11, formatter: (p) => fmtMoney(Number(p.value))},
        });
    }

    return {
        grid: {top: isAmount ? 40 : 20, right: 60, bottom: 30, left: 100},
        legend: isAmount ? {top: 0, left: 'center', textStyle: {fontSize: 11}} : {show: false},
        tooltip: {
            trigger: 'axis',
            axisPointer: {type: 'shadow'},
            formatter: (raw) => buildAxisTooltip(raw as AxisTooltipParam | AxisTooltipParam[], groups, valueFmt),
        },
        xAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        yAxis: {type: 'category', data: groups.map((g) => g.key), axisLabel: {fontSize: 11}},
        series,
    };
}

/* ================ T2.1 客戶 Top 10（金額/數量/利潤切換） ================ */
export function customerTopOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind = 'amount'): EChartsOption {
    const groups = topN(groupByCustomer(rows, metric, true), 10).reverse();
    const isAmount = metric === 'amount';
    const valueFmt = metricValueFmt(metric);

    const series: NonNullable<EChartsOption['series']> = [
        {
            name: metricSeriesName(metric),
            type: 'bar',
            data: groups.map((g) => metricValueFromGroup(g, metric)),
            itemStyle: {borderRadius: [0, 4, 4, 0], color: metricColor(metric)},
            barMaxWidth: 22,
            label: {show: true, position: 'right', fontSize: 11, formatter: (p) => valueFmt(Number(p.value))},
        },
    ];
    if (isAmount) {
        series.push({
            name: '毛利',
            type: 'bar',
            data: groups.map((g) => Math.round(g.profit)),
            itemStyle: {borderRadius: [0, 4, 4, 0], color: PALETTE.profit},
            barMaxWidth: 22,
            label: {show: true, position: 'right', fontSize: 11, formatter: (p) => fmtMoney(Number(p.value))},
        });
    }

    return {
        grid: {top: isAmount ? 40 : 20, right: 60, bottom: 30, left: 140},
        legend: isAmount ? {top: 0, left: 'center', textStyle: {fontSize: 11}} : {show: false},
        tooltip: {
            trigger: 'axis',
            axisPointer: {type: 'shadow'},
            formatter: (raw) => buildAxisTooltip(raw as AxisTooltipParam | AxisTooltipParam[], groups, valueFmt),
        },
        xAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        yAxis: {type: 'category', data: groups.map((g) => g.key), axisLabel: {fontSize: 11}},
        series,
    };
}

/* ================ T2.3 商品分類 Treemap（金額/數量/利潤切換 + 毛利率色階） ================ */
export function categoryTreemapOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind = 'amount'): EChartsOption {
    const byCategory = groupBy(rows, 'category', metric, true);
    const byCategoryMap = new Map(byCategory.map((g) => [g.key, g]));

    const productByCategory = new Map<string, Map<string, number>>();
    for (const r of rows) {
        let m = productByCategory.get(r.category);
        if (!m) {
            m = new Map();
            productByCategory.set(r.category, m);
        }
        const v = metric === 'count' ? r.count : metric === 'profit' ? r.profit : r.amount;
        m.set(r.productName, (m.get(r.productName) ?? 0) + v);
    }

    // treemap value 必須非負，對 profit 模式下虧損商品取 0（tooltip 仍顯示真實值）
    const sizeOf = (v: number) => Math.max(0, v);

    const data = byCategory.map((cat, i) => {
        const parentColor = CATEGORY_PALETTE[i % CATEGORY_PALETTE.length];
        const entries = [...(productByCategory.get(cat.key) ?? new Map())] as Array<[string, number]>;
        const values = entries.map(([, v]) => v);
        const max = values.length > 0 ? Math.max(...values) : 1;
        const min = values.length > 0 ? Math.min(...values) : 0;
        const range = max - min;
        const products = entries.map(([name, value]) => {
            // 大 value → 0 (原色)、小 value → 0.65 (淡)
            const ratio = range > 0 ? 1 - (value - min) / range : 0;
            const lighten = ratio * 0.65;
            return {
                name,
                value: sizeOf(value),
                itemStyle: {
                    color: lightenHex(parentColor, lighten),
                    borderWidth: 0,
                },
            };
        });
        return {
            name: cat.key,
            value: sizeOf(metricValueFromGroup(cat, metric)),
            itemStyle: {
                color: parentColor,
                borderWidth: 0,
                gapWidth: 0,
            },
            children: products,
        };
    });

    const metricLabel = metric === 'count' ? '數量' : metric === 'profit' ? '利潤' : '金額';

    return {
        tooltip: {
            formatter: (raw) => {
                const p = Array.isArray(raw) ? raw[0] : raw;
                const name = String(p.name ?? '');
                const head = `<div style="font-weight:600">${name}</div>${metricLabel}　${fmtMoney(Number(p.value ?? 0))}`;
                const path = (p as unknown as { treePathInfo?: Array<{ name: string }> }).treePathInfo ?? [];
                let cat: GroupSum | undefined;
                if (path.length >= 2) cat = byCategoryMap.get(path[1].name);
                else cat = byCategoryMap.get(name);
                return head + formatProfitFooter(cat);
            },
        },
        series: [
            {
                type: 'treemap',
                roam: false,
                nodeClick: false,
                breadcrumb: {show: false},
                label: {show: true, formatter: '{b}', fontSize: 11},
                upperLabel: {show: true, height: 20, fontSize: 12, fontWeight: 'bold'},
                levels: [
                    {},
                    {
                        itemStyle: {borderWidth: 0, gapWidth: 0},
                    },
                    {
                        itemStyle: {borderWidth: 0, gapWidth: 0},
                        upperLabel: {show: false},
                    },
                ],
                data,
            },
        ],
    };
}

/* ================ T2.4 客戶帕累托（金額/數量/利潤切換） ================ */
export function customerParetoOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind = 'amount'): EChartsOption {
    const groups = topN(groupByCustomer(rows, metric, true), 30); // 取前 30 客戶
    const values = groups.map((g) => metricValueFromGroup(g, metric));
    const total = values.reduce((s, v) => s + v, 0) || 1;
    let acc = 0;
    const cumPct = values.map((v) => {
        acc += v;
        return Number(((acc / total) * 100).toFixed(2));
    });
    const seriesName = metricSeriesName(metric);
    const valueFmt = metricValueFmt(metric);
    return {
        grid: {top: 40, right: 60, bottom: 80, left: 60},
        tooltip: {
            trigger: 'axis',
            formatter: (raw) => {
                const params = Array.isArray(raw) ? raw : [raw];
                const idx = (params[0]?.dataIndex ?? 0) as number;
                const g = groups[idx];
                const head = params
                    .map((p) => {
                        const name = p.seriesName ?? '';
                        const isPct = name === '累積佔比';
                        const v = Number(p.value ?? 0);
                        return `${p.marker ?? ''}${name}　${isPct ? `${v}%` : valueFmt(v)}`;
                    })
                    .join('<br/>');
                const title = `<div style="font-weight:600">${g?.key ?? ''}</div>`;
                return title + head + formatProfitFooter(g);
            },
        },
        legend: {top: 0, left: 'center'},
        xAxis: {
            type: 'category',
            data: groups.map((g) => g.key),
            axisLabel: {rotate: 35, fontSize: 10, interval: 0},
        },
        yAxis: [
            {type: 'value', name: seriesName, axisLabel: {formatter: (v: number) => fmtMoney(v)}},
            {type: 'value', name: '累積%', max: 100, axisLabel: {formatter: '{value}%'}},
        ],
        series: [
            {
                name: seriesName,
                type: 'bar',
                data: values,
                itemStyle: {color: metricColor(metric)},
            },
            {
                name: '累積佔比',
                type: 'line',
                yAxisIndex: 1,
                data: cumPct,
                smooth: false,
                symbol: 'circle',
                symbolSize: 5,
                itemStyle: {color: PALETTE.cost},
                markLine: {
                    silent: true,
                    symbol: 'none',
                    label: {formatter: '80%'},
                    lineStyle: {color: PALETTE.cost, type: 'dashed'},
                    data: [{yAxis: 80}],
                },
            },
        ],
    };
}

/* ================ T2.5 星期銷售熱度 ================ */
export function weekdayOption(rows: ReadonlyArray<AnalyticsRow>, metric: MetricKind): EChartsOption {
    const series = weekdaySeries(rows);
    // 重排為 一二三四五六日
    const order = [1, 2, 3, 4, 5, 6, 0];
    const reordered = order.map((i) => series[i]);
    const valueFmt = metricValueFmt(metric);
    const barColor = metric === 'profit' ? PALETTE.profit : PALETTE.accent;
    return {
        grid: {top: 20, right: 20, bottom: 30, left: 60},
        tooltip: {
            trigger: 'axis',
            axisPointer: {type: 'shadow'},
            formatter: (raw) => buildAxisTooltip(raw as AxisTooltipParam | AxisTooltipParam[], reordered, valueFmt),
        },
        xAxis: {type: 'category', data: order.map((i) => `週${WEEKDAY_NAMES[i]}`)},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {
                name: metricSeriesName(metric),
                type: 'bar',
                data: reordered.map((p) => metricValueFromDaily(p, metric)),
                itemStyle: {color: barColor, borderRadius: [4, 4, 0, 0]},
                label: {show: true, position: 'top', fontSize: 11, formatter: (p) => valueFmt(Number(p.value))},
            },
        ],
    };
}

