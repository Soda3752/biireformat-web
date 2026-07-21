/**
 * KPI 卡片計算與渲染（純 HTML，不用 ECharts）。
 */

import {distinctCount, sumAmount, sumAmountWithTax, sumCostAmount, sumCount, sumProfit} from './aggregators';

export interface KpiData {
    totalAmount: number;
    totalAmountWithTax: number;
    totalCount: number;
    orderRowCount: number;
    customerCount: number;
    productCount: number;
    avgPerCustomer: number;
    avgPerDay: number;
    avgCountPerCustomer: number;
    avgCountPerDay: number;
    totalCost: number;
    totalProfit: number;
    grossMarginPct: number; // 0~100
    unsetCostProductCount: number;
}

export function computeKpi(dataset: {
    rows: ReadonlyArray<import('./dataset-builder').AnalyticsRow>;
    unsetCostProducts?: ReadonlyArray<string>;
}): KpiData {
    const rows = dataset.rows;
    const totalAmount = sumAmount(rows);
    const totalCount = sumCount(rows);
    const customerCount = distinctCount(rows, 'customerCode');
    const productCount = distinctCount(rows, 'productName');
    const days = distinctCount(rows, 'day');
    const totalCost = sumCostAmount(rows);
    const totalProfit = sumProfit(rows);
    const grossMarginPct = totalAmount > 0 ? (totalProfit / totalAmount) * 100 : 0;

    return {
        totalAmount,
        totalAmountWithTax: Math.round(sumAmountWithTax(rows)),
        totalCount,
        orderRowCount: rows.length,
        customerCount,
        productCount,
        avgPerCustomer: customerCount > 0 ? Math.round(totalAmount / customerCount) : 0,
        avgPerDay: days > 0 ? Math.round(totalAmount / days) : 0,
        avgCountPerCustomer: customerCount > 0 ? Math.round(totalCount / customerCount) : 0,
        avgCountPerDay: days > 0 ? Math.round(totalCount / days) : 0,
        totalCost: Math.round(totalCost),
        totalProfit: Math.round(totalProfit),
        grossMarginPct,
        unsetCostProductCount: dataset.unsetCostProducts?.length ?? 0,
    };
}

export function renderKpiCards(host: HTMLElement, kpi: KpiData): void {
    const fmt = (v: number) => v.toLocaleString('zh-TW');
    const fmtPct = (v: number) => `${v.toFixed(1)}%`;
    const costHint = kpi.unsetCostProductCount > 0
        ? `${kpi.unsetCostProductCount} 項商品未填成本（視為 0）`
        : undefined;
    const items: Array<{ label: string; value: string; hint?: string }> = [
        {label: '總營收', value: fmt(kpi.totalAmount), hint: `含稅 ${fmt(kpi.totalAmountWithTax)}`},
        {label: '總銷售數量', value: fmt(kpi.totalCount)},
        {label: '明細列數', value: fmt(kpi.orderRowCount)},
        {label: '客戶數', value: fmt(kpi.customerCount)},
        {label: '商品種類', value: fmt(kpi.productCount)},
        {label: '平均客單價', value: fmt(kpi.avgPerCustomer), hint: `平均客戶數量 ${fmt(kpi.avgCountPerCustomer)}`},
        {label: '平均日營收', value: fmt(kpi.avgPerDay), hint: `平均日數量 ${fmt(kpi.avgCountPerDay)}`},
        {label: '總成本', value: fmt(kpi.totalCost), hint: costHint},
        {label: '總毛利', value: fmt(kpi.totalProfit)},
        {label: '毛利率', value: fmtPct(kpi.grossMarginPct)},
    ];
    host.innerHTML = items
        .map(
            (it) => `
        <div class="analytics-kpi-card">
          <div class="analytics-kpi-label">${it.label}</div>
          <div class="analytics-kpi-value">${it.value}</div>
          ${it.hint ? `<div class="analytics-kpi-hint">${it.hint}</div>` : ''}
        </div>
      `
        )
        .join('');
}
