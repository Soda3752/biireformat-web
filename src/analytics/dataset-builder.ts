/**
 * 把多份 Bill 領域物件展平成 AnalyticsRow 陣列，供圖表 adapter 聚合。
 *
 * AnalyticsRow 故意設計成「扁平、純值」的形式，讓 filter-engine 與 chart-adapter
 * 可以用簡單的 Array.filter / reduce 處理，不必再回頭呼叫 Bill 物件方法。
 */

import type {Bill} from '@/domain/models/bill';
import {getWeekday} from '@/domain/date-utility';
import type {CategoryMap} from './category-loader';
import {UNCATEGORIZED} from './category-loader';

export interface LoadedFileMeta {
    id: string;
    name: string;
    bill: Bill;
}

export interface AnalyticsRow {
    fileId: string;
    fileName: string;
    year: string;        // 民國年
    month: number;
    day: number;
    weekday: number;     // 0=日 ~ 6=六
    customerCode: string;
    customerName: string;
    line: string;        // 第N線
    isMonthly: boolean;
    isNeedTex: boolean;
    isCashUser: boolean;
    productName: string;
    category: string;
    count: number;
    price: number;
    amount: number;      // count × price
}

export interface AnalyticsFileEntry {
    id: string;
    name: string;
    year: string;
    month: number;
    customerCount: number;
    rowCount: number;
}

export interface AnalyticsDataset {
    rows: AnalyticsRow[];
    files: AnalyticsFileEntry[];
    unmatchedProducts: string[];
}

export function buildDataset(
    files: ReadonlyArray<LoadedFileMeta>,
    categoryMap: CategoryMap
): AnalyticsDataset {
    const rows: AnalyticsRow[] = [];
    const fileEntries: AnalyticsFileEntry[] = [];
    const unmatched = new Set<string>();

    for (const file of files) {
        const {bill} = file;
        const dateInfo = bill.billDateInfo;
        if (!dateInfo) continue;

        const year = dateInfo.year;
        const month = Number(dateInfo.month);

        let fileRowCount = 0;
        for (const customer of bill.customerModels) {
            const line = `第${customer.code.substring(0, 1)}線`;

            for (const product of customer.productList) {
                const category = categoryMap.get(product.name);
                if (!category) unmatched.add(product.name);

                for (const order of product.orderList) {
                    rows.push({
                        fileId: file.id,
                        fileName: file.name,
                        year,
                        month,
                        day: order.day,
                        weekday: getWeekday(year, month, order.day),
                        customerCode: customer.code,
                        customerName: customer.name,
                        line,
                        isMonthly: customer.isMonthly,
                        isNeedTex: customer.isNeedTex,
                        isCashUser: customer.isCashUser,
                        productName: product.name,
                        category: category ?? UNCATEGORIZED,
                        count: order.count,
                        price: product.price,
                        amount: order.count * product.price,
                    });
                    fileRowCount += 1;
                }
            }
        }

        fileEntries.push({
            id: file.id,
            name: file.name,
            year,
            month,
            customerCount: bill.customerModels.length,
            rowCount: fileRowCount,
        });
    }

    return {
        rows,
        files: fileEntries,
        unmatchedProducts: [...unmatched].sort(),
    };
}
