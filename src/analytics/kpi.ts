/**
 * KPI 卡片計算與渲染（純 HTML，不用 ECharts）。
 */

import {distinctCount, sumAmount, sumCount} from './aggregators';

export interface KpiData {
    totalAmount: number;
    totalAmountWithTax: number;
    orderRowCount: number;
    customerCount: number;
    productCount: number;
    avgPerCustomer: number;
    avgPerDay: number;
}

export function computeKpi(dataset: { rows: ReadonlyArray<import('./dataset-builder').AnalyticsRow> }): KpiData {
    const rows = dataset.rows;
    const totalAmount = sumAmount(rows);
    const totalCount = sumCount(rows);
    const customerCount = distinctCount(rows, 'customerCode');
    const productCount = distinctCount(rows, 'productName');
    const days = distinctCount(rows, 'day');

    // 含稅金額：有 isNeedTex 旗標的客戶銷貨額外加 5%
    let amountWithTax = 0;
    for (const r of rows) {
        amountWithTax += r.isNeedTex ? r.amount * 1.05 : r.amount;
    }

    return {
        totalAmount,
        totalAmountWithTax: Math.round(amountWithTax),
        orderRowCount: rows.length,
        customerCount,
        productCount,
        avgPerCustomer: customerCount > 0 ? Math.round(totalAmount / customerCount) : 0,
        avgPerDay: days > 0 ? Math.round(totalAmount / days) : 0,
        // totalCount 雖未顯示在 KPI 卡，但保留欄位以利未來使用
        ...({_totalCount: totalCount} as Record<string, number>),
    };
}

export function renderKpiCards(host: HTMLElement, kpi: KpiData): void {
    const fmt = (v: number) => v.toLocaleString('zh-TW');
    const items: Array<{ label: string; value: string; hint?: string }> = [
        {label: '總營收', value: fmt(kpi.totalAmount), hint: `含稅 ${fmt(kpi.totalAmountWithTax)}`},
        {label: '明細列數', value: fmt(kpi.orderRowCount)},
        {label: '客戶數', value: fmt(kpi.customerCount)},
        {label: '商品種類', value: fmt(kpi.productCount)},
        {label: '平均客單價', value: fmt(kpi.avgPerCustomer)},
        {label: '平均日營收', value: fmt(kpi.avgPerDay)},
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
