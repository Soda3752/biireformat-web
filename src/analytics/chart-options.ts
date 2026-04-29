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
import {dailySeries, groupBy, groupByCustomer, topN, weekdaySeries,} from './aggregators';
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

/* ================ T1.3 線別佔比（金額/數量切換） ================ */
export function linePieOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count' = 'amount'): EChartsOption {
    const groups = groupBy(rows, 'line', metric, true);
    return {
        tooltip: {
            trigger: 'item',
            valueFormatter: (v) => fmtMoney(Number(v)),
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
                data: groups.map((g) => ({name: g.key, value: metric === 'amount' ? g.amount : g.count})),
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

/* ================ T2.1 客戶 Top 10（金額/數量切換） ================ */
export function customerTopOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count' = 'amount'): EChartsOption {
    const groups = topN(groupByCustomer(rows, metric, true), 10).reverse();
    return {
        grid: {top: 20, right: 40, bottom: 30, left: 140},
        tooltip: {trigger: 'axis', valueFormatter: (v) => fmtMoney(Number(v))},
        xAxis: {type: 'value', axisLabel: {formatter: (v: number) => fmtMoney(v)}},
        yAxis: {type: 'category', data: groups.map((g) => g.key), axisLabel: {fontSize: 11}},
        series: [
            {
                type: 'bar',
                data: groups.map((g) => (metric === 'amount' ? g.amount : g.count)),
                itemStyle: {borderRadius: [0, 4, 4, 0], color: '#5b8def'},
                label: {show: true, position: 'right', fontSize: 11, formatter: (p) => fmtMoney(Number(p.value))},
            },
        ],
    };
}

/* ================ T2.3 商品分類 Treemap（金額/數量切換） ================ */
export function categoryTreemapOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count' = 'amount'): EChartsOption {
    const byCategory = groupBy(rows, 'category', metric, true);
    const productByCategory = new Map<string, Map<string, number>>();
    for (const r of rows) {
        let m = productByCategory.get(r.category);
        if (!m) {
            m = new Map();
            productByCategory.set(r.category, m);
        }
        const v = metric === 'amount' ? r.amount : r.count;
        m.set(r.productName, (m.get(r.productName) ?? 0) + v);
    }
    const data = byCategory.map((cat) => {
        const products = [...(productByCategory.get(cat.key) ?? new Map())].map(([name, value]) => ({
            name,
            value,
        }));
        return {name: cat.key, value: metric === 'amount' ? cat.amount : cat.count, children: products};
    });
    const metricLabel = metric === 'amount' ? '金額' : '數量';
    return {
        tooltip: {
            formatter: (params) => {
                const p = Array.isArray(params) ? params[0] : params;
                return `${p.name}<br/>${metricLabel}: ${fmtMoney(Number(p.value))}`;
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

/* ================ T2.4 客戶帕累托（金額/數量切換） ================ */
export function customerParetoOption(rows: ReadonlyArray<AnalyticsRow>, metric: 'amount' | 'count' = 'amount'): EChartsOption {
    const groups = topN(groupByCustomer(rows, metric, true), 30); // 取前 30 客戶
    const values = groups.map((g) => (metric === 'amount' ? g.amount : g.count));
    const total = values.reduce((s, v) => s + v, 0) || 1;
    let acc = 0;
    const cumPct = values.map((v) => {
        acc += v;
        return Number(((acc / total) * 100).toFixed(2));
    });
    const seriesName = metric === 'amount' ? '營收' : '銷售數量';
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
            {type: 'value', name: seriesName, axisLabel: {formatter: (v: number) => fmtMoney(v)}},
            {type: 'value', name: '累積%', max: 100, axisLabel: {formatter: '{value}%'}},
        ],
        series: [
            {
                name: seriesName,
                type: 'bar',
                data: values,
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
