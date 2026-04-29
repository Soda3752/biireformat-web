/**
 * 數據分析分頁的「商品名稱 → 分類」查找表。
 *
 * 重用 `domain/daily-report-loader` 的 CSV 解析（已含設定頁 localStorage 覆寫支援）。
 * 此處只負責把 `Map<groupName, DailyProduct[]>` 反轉為 `Map<productName, groupName>`，
 * 並對重複名稱採「先到先得」原則（極少發生，但 csv 容許重名）。
 */

import {loadDailyReportTemplate} from '@/domain/daily-report-loader';

export type CategoryMap = ReadonlyMap<string, string>;

let cached: Promise<CategoryMap> | null = null;

export const loadCategoryMap = (): Promise<CategoryMap> => {
    if (!cached) cached = build();
    return cached;
};

export const invalidateCategoryMap = (): void => {
    cached = null;
};

const build = async (): Promise<CategoryMap> => {
    const grouped = await loadDailyReportTemplate();
    const map = new Map<string, string>();
    for (const [groupName, products] of grouped) {
        for (const p of products) {
            if (!map.has(p.name)) map.set(p.name, groupName);
        }
    }
    return map;
};

/** 在 CSV 找不到對應分類時的預設值。 */
export const UNCATEGORIZED = '其他';
