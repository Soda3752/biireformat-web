/**
 * 帳單檔 → 每月商品銷售快照，並彙總成每月網站營收。
 *
 * 營收計算沿用數據分析的 buildDataset（數量 × 單價），月份以 row.year + row.month 切分，
 * 與跨月數據分析的 rowMonthSortKey 規約一致。
 */

import {parseBillFile} from '@/readers/bill-reader';
import {Bill} from '@/domain/models/bill';
import {ExcelRowType} from '@/domain/excel-row-data';
import {loadCategoryMap} from '@/analytics/category-loader';
import {loadCostMap} from '@/analytics/cost-loader';
import {buildDataset} from '@/analytics/dataset-builder';
import {rowMonthSortKey} from '@/analytics/cross-month/month-aggregators';
import type {ProductSale, RevenueFile} from './expense-store';

export function revenueFileId(file: File): string {
    return `${file.name}__${file.size}`;
}

export async function parseBillToRevenueFile(file: File): Promise<RevenueFile> {
    const bill = new Bill();
    await parseBillFile(file, ({type, values}) => {
        if (type === ExcelRowType.BillSettingInfo) bill.setBillDateInfo(values);
        else if (type === ExcelRowType.CustomerData) bill.newCustomerModel(values);
        else if (type === ExcelRowType.CustomerSetting) bill.configCustomer(values);
        else if (type === ExcelRowType.ProductRowSetting) bill.setProductSetting(values);
        else if (type === ExcelRowType.ProductSellData) bill.addProduct(values);
    });
    bill.appendCustomer();

    const id = revenueFileId(file);
    const [categoryMap, costMap] = await Promise.all([loadCategoryMap(), loadCostMap()]);
    const dataset = buildDataset([{id, name: file.name, bill}], categoryMap, costMap);
    if (dataset.rows.length === 0) throw new Error('沒有讀到任何銷售資料，請確認是「明細編號順序」帳單檔');

    const months = new Map<string, Map<string, ProductSale>>();
    for (const row of dataset.rows) {
        const key = rowMonthSortKey(row);
        let products = months.get(key);
        if (!products) {
            products = new Map();
            months.set(key, products);
        }
        const p = products.get(row.productName) ?? {name: row.productName, count: 0, amount: 0};
        p.count += row.count;
        p.amount += row.amount;
        products.set(row.productName, p);
    }

    const out: Record<string, ProductSale[]> = {};
    for (const [key, products] of months) out[key] = [...products.values()];
    return {id, fileName: file.name, months: out, savedAt: Date.now()};
}

export interface MonthRevenue {
    /** 網站營收（數量 × 單價） */
    amount: number;
    fileNames: string[];
}

export function computeMonthRevenue(files: ReadonlyArray<RevenueFile>): Map<string, MonthRevenue> {
    const out = new Map<string, MonthRevenue>();
    for (const f of files) {
        for (const [key, products] of Object.entries(f.months)) {
            const acc = out.get(key) ?? {amount: 0, fileNames: []};
            for (const p of products) acc.amount += p.amount;
            acc.fileNames.push(f.fileName);
            out.set(key, acc);
        }
    }
    return out;
}
