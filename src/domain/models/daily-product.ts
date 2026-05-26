/**
 * 對應桌面版 src/main/kotlin/dailyCountStatistics/model/Product.kt
 *
 * `costRaw` 為設定頁「單日數量」第 4 欄的原始字串。
 * 空字串代表使用者尚未填入成本（數據分析頁會在下方提示）。
 */
export interface DailyProduct {
  readonly code: string;
  readonly name: string;
  readonly groupName: string;
    readonly costRaw: string;
  count: number;
}

/** 解析 costRaw 為數字；空值或無法 parse 視為 0。 */
export const getCostNumber = (product: DailyProduct): number => {
    const n = Number(product.costRaw);
    return Number.isFinite(n) ? n : 0;
};

/** 是否已填入成本（非空白字串）。 */
export const hasCost = (product: DailyProduct): boolean =>
    product.costRaw.trim().length > 0;

/** 對應 Product.getTotalCount()：包含「(半)」名稱者算半個（向上取整） */
export const getTotalCount = (product: DailyProduct): number => {
  if (product.name.includes('(半)')) {
    return Math.floor((product.count + 1) / 2);
  }
  return product.count;
};

export type DailyProductMap = Map<string, DailyProduct[]>;

/** 深拷貝一份初始 product map（避免使用者多次解析時 count 累積汙染） */
export const cloneProductMap = (source: DailyProductMap): DailyProductMap => {
  const cloned: DailyProductMap = new Map();
  for (const [groupName, products] of source) {
    cloned.set(
      groupName,
      products.map((p) => ({ ...p }))
    );
  }
  return cloned;
};
