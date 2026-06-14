/**
 * 對應桌面版 src/main/kotlin/dailyCountStatistics/model/Product.kt
 *
 * `costRaw` 為「單日數量」CSV 第 4 欄的原始字串（已封存）。
 * 成本已搬移到「帳單排序」管理，數據分析改讀帳單排序的成本；
 * 此欄僅在解析舊資料時保留，目前不再參與任何成本計算。
 */
export interface DailyProduct {
  readonly code: string;
  readonly name: string;
  readonly groupName: string;
    readonly costRaw: string;
  count: number;
}

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
