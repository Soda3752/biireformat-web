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
import type {CostMap} from './cost-loader';

/** 各線別代表地區名稱（以客戶代碼首碼對應）。 */
const LINE_NAMES: Record<string, string> = {
    '1': '彰化',
    '2': '和美',
    '3': '溪湖',
    '4': '台中',
    '5': '市區',
    '6': '海線',
    '7': '員林+石',
    '8': '逢甲',
    '9': '貨運',
};

/** 線別標籤：第X線 + 地區名（無對應名稱時只顯示第X線）。 */
function lineLabel(code: string): string {
    const key = code.substring(0, 1);
    return `第${key}線${LINE_NAMES[key] ?? ''}`;
}

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
    line: string;        // 第N線+地區名（如「第1線彰化」）
    isMonthly: boolean;
    isNeedTex: boolean;
    isCashUser: boolean;
    productName: string;
    category: string;
    count: number;
    price: number;
    amount: number;      // count × price
    cost: number;        // 單品成本（未填視為 0）
    costAmount: number;  // count × cost
    profit: number;      // amount - costAmount
    isCostUnset: boolean; // 該商品於 daily_report_list 未填成本
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
    unsetCostProducts: string[];
}

export function buildDataset(
    files: ReadonlyArray<LoadedFileMeta>,
    categoryMap: CategoryMap,
    costMap: CostMap
): AnalyticsDataset {
    const rows: AnalyticsRow[] = [];
    const fileEntries: AnalyticsFileEntry[] = [];
    const unmatched = new Set<string>();
    const unsetCost = new Set<string>();

    for (const file of files) {
        const {bill} = file;
        const dateInfo = bill.billDateInfo;
        if (!dateInfo) continue;

        const year = dateInfo.year;
        const month = Number(dateInfo.month);

        let fileRowCount = 0;
        for (const customer of bill.customerModels) {
            const line = lineLabel(customer.code);

            for (const product of customer.productList) {
                const category = categoryMap.get(product.name);
                if (!category) unmatched.add(product.name);

                const costLookup = costMap.get(product.name);
                const isCostUnset = costLookup === undefined || costLookup === null;
                const cost = isCostUnset ? 0 : costLookup;
                if (isCostUnset) unsetCost.add(product.name);

                for (const order of product.orderList) {
                    const amount = order.count * product.price;
                    const costAmount = order.count * cost;
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
                        amount,
                        cost,
                        costAmount,
                        profit: amount - costAmount,
                        isCostUnset,
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
        unsetCostProducts: [...unsetCost].sort(),
    };
}
