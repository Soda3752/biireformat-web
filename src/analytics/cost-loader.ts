/**
 * 數據分析分頁的「商品名稱 → 成本」查找表。
 *
 * 成本資料已從 daily_report_list 搬移到 cargo_sort（帳單排序），
 * 故此處改讀 `domain/sorting-list`（已含設定頁 localStorage 覆寫支援）。
 * 把 cargoItems 反轉為 `Map<productName, cost | null>`：
 * - 成本為空白：對應到 null（代表「未填」，UI 會在下方提示補填）
 * - 成本為數字：對應到 number
 * - 對於同名商品採「先到先得」原則，與 category-loader 行為一致
 */

import {loadSortingList} from '@/domain/sorting-list';

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
    const list = await loadSortingList();
    const map = new Map<string, number | null>();
    for (const item of list.cargoItems) {
        if (map.has(item.name)) continue;
        map.set(item.name, item.cost);
    }
    return map;
};
