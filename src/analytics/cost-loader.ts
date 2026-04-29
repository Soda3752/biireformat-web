/**
 * 數據分析分頁的「商品名稱 → 成本」查找表。
 *
 * 重用 `domain/daily-report-loader` 的 CSV 解析（已含設定頁 localStorage 覆寫支援）。
 * 此處把 `Map<groupName, DailyProduct[]>` 反轉為 `Map<productName, cost | null>`：
 * - 成本為空字串：對應到 null（代表「未填」，UI 會在下方提示補填）
 * - 成本為數字：對應到 number
 * - 對於同名商品採「先到先得」原則，與 category-loader 行為一致
 */

import {getCostNumber, hasCost} from '@/domain/models/daily-product';
import {loadDailyReportTemplate} from '@/domain/daily-report-loader';

export type CostMap = ReadonlyMap<string, number | null>;

let cached: Promise<CostMap> | null = null;

export const loadCostMap = (): Promise<CostMap> => {
    if (!cached) cached = build();
    return cached;
};

export const invalidateCostMap = (): void => {
    cached = null;
};

const build = async (): Promise<CostMap> => {
    const grouped = await loadDailyReportTemplate();
    const map = new Map<string, number | null>();
    for (const products of grouped.values()) {
        for (const p of products) {
            if (map.has(p.name)) continue;
            map.set(p.name, hasCost(p) ? getCostNumber(p) : null);
        }
    }
    return map;
};
