/**
 * 12 張 ECharts 圖的 option builder。
 * 每個函式接 AnalyticsRow[]（已套篩選）+ 視需要的設定，回傳 EChartsOption。
 *
 * 統一規約：
 * - 不引用 echarts 模組（型別來自 EChartsOption），讓本檔可單元測試
 * - tooltip 走 axis trigger 為主，符合中文閱讀習慣
 * - 顏色走預設 palette，後續可注入 theme
 */

import type {EChartsOption} from 'echarts';
import type {AnalyticsDataset, AnalyticsRow} from './dataset-builder';
import {
    crossAggregate,
    dailySeries,
    groupBy,
    groupByCustomer,
    paymentModeCounts,
    topN,
    weekdaySeries,
} from './aggregators';
import {WEEKDAY_NAMES} from '@/domain/date-utility';

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');

/* ================ T1.2 每日營收趨勢 ================ */
export function dailyTrendOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count'): EChartsOption {
    const series = dailySeries(rows);
    const xData = series.map((p) => `${p.day}日`);
    const yData = series.map((p) => (metric === 'amount' ? p.amount : p.count));
    return {
        grid: {top: 30, right: 20, bottom: 40, left: 60},
        tooltip: {
            trigger: 'axis',
            valueFormatter: (v) => (metric === 'amount' ? fmtMoney(Number(v)) : String(v)),
        },
        xAxis: {type: 'category', data: xData, axisLabel: {fontSize: 11}},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {
                type: 'line',
                smooth: true,
                symbol: 'circle',
                symbolSize: 6,
                data: yData,
                areaStyle: {opacity: 0.15},
                lineStyle: {width: 2},
            },
        ],
    };
}

/* ================ T1.3 線別營收佔比 ================ */
export function linePieOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const groups = groupBy(rows, 'line', 'amount', true);
    return {
        tooltip: {
            trigger: 'item',
            valueFormatter: (v) => fmtMoney(Number(v)),
        },
        legend: {bottom: 0, left: 'center', textStyle: {fontSize: 11}},
        series: [
            {
                type: 'pie',
                radius: ['45%', '70%'],
                center: ['50%', '45%'],
                avoidLabelOverlap: true,
                label: {formatter: '{b}\n{d}%', fontSize: 11},
                data: groups.map((g) => ({name: g.key, value: g.amount})),
            },
        ],
    };
}

/* ================ T1.4 商品 Top 10 ================ */
export function productTopOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count'): EChartsOption {
    const groups = topN(groupBy(rows, 'productName', metric, true), 10).reverse();
    return {
        grid: {top: 20, right: 40, bottom: 30, left: 100},
        tooltip: {trigger: 'axis', valueFormatter: (v) => fmtMoney(Number(v))},
        xAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        yAxis: {type: 'category', data: groups.map((g) => g.key), axisLabel: {fontSize: 11}},
        series: [
            {
                type: 'bar',
                data: groups.map((g) => (metric === 'amount' ? g.amount : g.count)),
                itemStyle: {borderRadius: [0, 4, 4, 0]},
                label: {show: true, position: 'right', fontSize: 11, formatter: (p) => fmtMoney(Number(p.value))},
            },
        ],
    };
}

/* ================ T2.1 客戶 Top 10 ================ */
export function customerTopOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const groups = topN(groupByCustomer(rows, true), 10).reverse();
    return {
        grid: {top: 20, right: 40, bottom: 30, left: 140},
        tooltip: {trigger: 'axis', valueFormatter: (v) => fmtMoney(Number(v))},
        xAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        yAxis: {type: 'category', data: groups.map((g) => g.key), axisLabel: {fontSize: 11}},
        series: [
            {
                type: 'bar',
                data: groups.map((g) => g.amount),
                itemStyle: {borderRadius: [0, 4, 4, 0], color: '#5b8def'},
                label: {show: true, position: 'right', fontSize: 11, formatter: (p) => fmtMoney(Number(p.value))},
            },
        ],
    };
}

/* ================ T2.2 結帳模式（重疊計數） ================ */
export function paymentModeOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const c = paymentModeCounts(rows);
    const total = c.total || 1;
    const data = [
        {name: '月結', value: c.monthly},
        {name: '含稅', value: c.needTex},
        {name: '現金', value: c.cash},
    ];
    return {
        grid: {top: 30, right: 20, bottom: 40, left: 50},
        tooltip: {
            trigger: 'axis',
            formatter: (params) => {
                const arr = Array.isArray(params) ? params : [params];
                const p = arr[0] as { value: number; name: string; axisValue?: string };
                const v = Number(p.value);
                return `${p.axisValue ?? p.name}: ${v} 人（${((v / total) * 100).toFixed(1)}%）`;
            },
        },
        xAxis: {type: 'category', data: data.map((d) => d.name)},
        yAxis: {type: 'value', name: '客戶數', minInterval: 1},
        title: {
            text: '同一客戶可同時屬於多類，總和可能 > 100%',
            left: 'center',
            top: 'bottom',
            textStyle: {fontSize: 11, fontWeight: 'normal', color: '#888'},
        },
        series: [
            {
                type: 'bar',
                data: data.map((d) => d.value),
                itemStyle: {color: '#7ac74f', borderRadius: [4, 4, 0, 0]},
                label: {
                    show: true,
                    position: 'top',
                    fontSize: 12,
                    formatter: (p) => {
                        const v = Number(p.value);
                        return `${v}\n(${((v / total) * 100).toFixed(0)}%)`;
                    },
                },
            },
        ],
    };
}

/* ================ T2.3 商品分類 Treemap ================ */
export function categoryTreemapOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const byCategory = groupBy(rows, 'category', 'amount', true);
    const productByCategory = new Map<string, Map<string, number>>();
    for (const r of rows) {
        let m = productByCategory.get(r.category);
        if (!m) {
            m = new Map();
            productByCategory.set(r.category, m);
        }
        m.set(r.productName, (m.get(r.productName) ?? 0) + r.amount);
    }
    const data = byCategory.map((cat) => {
        const products = [...(productByCategory.get(cat.key) ?? new Map())].map(([name, amount]) => ({
            name,
            value: amount,
        }));
        return {name: cat.key, value: cat.amount, children: products};
    });
    return {
        tooltip: {
            formatter: (params) => {
                const p = Array.isArray(params) ? params[0] : params;
                return `${p.name}<br/>金額: ${fmtMoney(Number(p.value))}`;
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
                    {
                        itemStyle: {borderColor: '#fff', borderWidth: 2, gapWidth: 2},
                    },
                    {
                        colorSaturation: [0.35, 0.55],
                        itemStyle: {borderColorSaturation: 0.7, gapWidth: 1, borderWidth: 1},
                    },
                ],
                data,
            },
        ],
    };
}

/* ================ T2.4 客戶帕累托 ================ */
export function customerParetoOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const groups = topN(groupByCustomer(rows, true), 30); // 取前 30 客戶
    const total = groups.reduce((s, g) => s + g.amount, 0) || 1;
    let acc = 0;
    const cumPct = groups.map((g) => {
        acc += g.amount;
        return Number(((acc / total) * 100).toFixed(2));
    });
    return {
        grid: {top: 40, right: 60, bottom: 80, left: 60},
        tooltip: {trigger: 'axis'},
        legend: {top: 0, left: 'center'},
        xAxis: {
            type: 'category',
            data: groups.map((g) => g.key),
            axisLabel: {rotate: 35, fontSize: 10, interval: 0},
        },
        yAxis: [
            {type: 'value', name: '營收', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
            {type: 'value', name: '累積%', max: 100, axisLabel: {formatter: '{value}%'}},
        ],
        series: [
            {
                name: '營收',
                type: 'bar',
                data: groups.map((g) => g.amount),
                itemStyle: {color: '#5b8def'},
            },
            {
                name: '累積佔比',
                type: 'line',
                yAxisIndex: 1,
                data: cumPct,
                smooth: false,
                symbol: 'circle',
                symbolSize: 5,
                itemStyle: {color: '#f5a623'},
                markLine: {
                    silent: true,
                    symbol: 'none',
                    label: {formatter: '80%'},
                    lineStyle: {color: '#f5a623', type: 'dashed'},
                    data: [{yAxis: 80}],
                },
            },
        ],
    };
}

/* ================ T2.5 星期銷售熱度 ================ */
export function weekdayOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count'): EChartsOption {
    const series = weekdaySeries(rows);
    // 重排為 一二三四五六日
    const order = [1, 2, 3, 4, 5, 6, 0];
    const reordered = order.map((i) => series[i]);
    return {
        grid: {top: 20, right: 20, bottom: 30, left: 60},
        tooltip: {trigger: 'axis', valueFormatter: (v) => fmtMoney(Number(v))},
        xAxis: {type: 'category', data: order.map((i) => `週${WEEKDAY_NAMES[i]}`)},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {
                type: 'bar',
                data: reordered.map((p) => (metric === 'amount' ? p.amount : p.count)),
                itemStyle: {color: '#9b59b6', borderRadius: [4, 4, 0, 0]},
                label: {show: true, position: 'top', fontSize: 11, formatter: (p) => fmtMoney(Number(p.value))},
            },
        ],
    };
}

/* ================ T3.1 客戶 × 商品熱力圖 ================ */
export function customerProductHeatmapOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const m = crossAggregate(rows, 'productName', 'customerCode', 'amount', 15, 15);
    const customerNameMap = new Map<string, string>();
    for (const r of rows) customerNameMap.set(r.customerCode, r.customerName);
    const yLabels = m.yKeys.map((c) => customerNameMap.get(c) ?? c);
    return {
        tooltip: {
            formatter: (params) => {
                const p = Array.isArray(params) ? params[0] : params;
                const v = (p.value as [number, number, number])[2];
                return `${p.name}<br/>${yLabels[(p.value as [number, number, number])[1]]}<br/>金額: ${fmtMoney(v)}`;
            },
        },
        grid: {top: 30, right: 20, bottom: 80, left: 130},
        xAxis: {
            type: 'category',
            data: m.xKeys,
            axisLabel: {rotate: 35, fontSize: 10, interval: 0},
            splitArea: {show: true}
        },
        yAxis: {type: 'category', data: yLabels, axisLabel: {fontSize: 10}, splitArea: {show: true}},
        visualMap: {
            min: 0,
            max: m.max || 1,
            calculable: true,
            orient: 'horizontal',
            left: 'center',
            bottom: 0,
            inRange: {color: ['#f6f7fb', '#5b8def', '#1a4fb8']},
        },
        series: [
            {
                type: 'heatmap',
                data: m.values,
                emphasis: {itemStyle: {shadowBlur: 8, shadowColor: 'rgba(0,0,0,0.3)'}},
            },
        ],
    };
}

/* ================ T3.2 日期 × 商品熱力圖 ================ */
export function dateProductHeatmapOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    // y = Top 15 商品（依數量）
    const productTop = topN(groupBy(rows, 'productName', 'count', true), 15).map((g) => g.key);
    const productSet = new Set(productTop);
    const productIdx = new Map(productTop.map((p, i) => [p, i]));
    const days = [...new Set(rows.map((r) => r.day))].sort((a, b) => a - b);
    const dayIdx = new Map(days.map((d, i) => [d, i]));

    const matrix = new Map<string, number>();
    for (const r of rows) {
        if (!productSet.has(r.productName)) continue;
        const k = `${r.day}|${r.productName}`;
        matrix.set(k, (matrix.get(k) ?? 0) + r.count);
    }
    const values: Array<[number, number, number]> = [];
    let max = 0;
    for (const [k, v] of matrix) {
        const [d, p] = k.split('|');
        values.push([dayIdx.get(Number(d))!, productIdx.get(p)!, v]);
        if (v > max) max = v;
    }

    return {
        tooltip: {
            formatter: (params) => {
                const p = Array.isArray(params) ? params[0] : params;
                const v = (p.value as [number, number, number])[2];
                return `${days[(p.value as [number, number, number])[0]]}日<br/>${productTop[(p.value as [number, number, number])[1]]}<br/>數量: ${v}`;
            },
        },
        grid: {top: 30, right: 20, bottom: 80, left: 110},
        xAxis: {type: 'category', data: days.map((d) => `${d}日`), axisLabel: {fontSize: 10}, splitArea: {show: true}},
        yAxis: {type: 'category', data: productTop, axisLabel: {fontSize: 10}, splitArea: {show: true}},
        visualMap: {
            min: 0,
            max: max || 1,
            calculable: true,
            orient: 'horizontal',
            left: 'center',
            bottom: 0,
            inRange: {color: ['#fef9e7', '#f5a623', '#c0392b']},
        },
        series: [
            {
                type: 'heatmap',
                data: values,
            },
        ],
    };
}

/* ================ T3.3 異常偵測 ================ */
export interface AnomalyPoint {
    day: number;
    amount: number;
    zscore: number;
}

export function detectAnomalies(rows: ReadonlyArray<AnalyticsRow>): {
    series: ReturnType<typeof dailySeries>;
    anomalies: AnomalyPoint[];
} {
    const series = dailySeries(rows);
    if (series.length < 3) return {series, anomalies: []};
    const amounts = series.map((p) => p.amount);
    const mean = amounts.reduce((a, b) => a + b, 0) / amounts.length;
    const variance = amounts.reduce((a, b) => a + (b - mean) ** 2, 0) / amounts.length;
    const std = Math.sqrt(variance);
    const anomalies: AnomalyPoint[] = [];
    if (std > 0) {
        for (const p of series) {
            const z = (p.amount - mean) / std;
            if (Math.abs(z) > 2) anomalies.push({day: p.day, amount: p.amount, zscore: Number(z.toFixed(2))});
        }
    }
    return {series, anomalies};
}

export function anomalyOption(rows: ReadonlyArray<AnalyticsRow>): EChartsOption {
    const {series, anomalies} = detectAnomalies(rows);
    return {
        grid: {top: 30, right: 20, bottom: 30, left: 60},
        tooltip: {trigger: 'axis', valueFormatter: (v) => fmtMoney(Number(v))},
        xAxis: {type: 'category', data: series.map((p) => `${p.day}日`)},
        yAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        series: [
            {
                type: 'line',
                smooth: true,
                data: series.map((p) => p.amount),
                lineStyle: {width: 2},
                markPoint: {
                    symbol: 'pin',
                    symbolSize: 50,
                    itemStyle: {color: '#e74c3c'},
                    label: {formatter: '異常', fontSize: 10, color: '#fff'},
                    data: anomalies.map((a) => ({xAxis: `${a.day}日`, yAxis: a.amount, name: `z=${a.zscore}`})),
                },
            },
        ],
    };
}

/* ================ T3.4 月對月成長 ================ */
export function monthOverMonthOption(dataset: AnalyticsDataset): EChartsOption {
    // 一個 file = 一個月份
    const fileLabels = dataset.files.map((f) => `${f.year}年${f.month}月`);
    const fileAmount: number[] = [];
    const fileCount: number[] = [];
    const fileCustomers: number[] = [];
    for (const f of dataset.files) {
        const fileRows = dataset.rows.filter((r) => r.fileId === f.id);
        let amount = 0;
        let count = 0;
        const customers = new Set<string>();
        for (const r of fileRows) {
            amount += r.amount;
            count += r.count;
            customers.add(r.customerCode);
        }
        fileAmount.push(amount);
        fileCount.push(count);
        fileCustomers.push(customers.size);
    }

    return {
        grid: {top: 50, right: 20, bottom: 40, left: 70},
        tooltip: {trigger: 'axis'},
        legend: {top: 0, left: 'center'},
        xAxis: {type: 'category', data: fileLabels},
        yAxis: [
            {type: 'value', name: '金額/數量', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
            {type: 'value', name: '客戶數', minInterval: 1},
        ],
        series: [
            {name: '營收', type: 'bar', data: fileAmount, itemStyle: {color: '#5b8def'}},
            {name: '銷售數量', type: 'bar', data: fileCount, itemStyle: {color: '#7ac74f'}},
            {
                name: '客戶數',
                type: 'line',
                yAxisIndex: 1,
                data: fileCustomers,
                itemStyle: {color: '#f5a623'},
                lineStyle: {width: 2},
                symbolSize: 8
            },
        ],
    };
}
