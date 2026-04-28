/**
 * 對應桌面版 `models/ProductSetting.kt`。
 * 該設定來自帳單表頭一列 `日期 單號 ... 品名規格 數量 單位 單價 ...`，
 * 用 `indexOf` 找出每欄索引，後續銷貨列依此索引取值。
 */

export interface ProductSetting {
  orderDateIndex: number;
  productNameIndex: number;
  productCountIndex: number;
  productPriceIndex: number;
}

export function parseProductSetting(rowData: ReadonlyArray<string>): ProductSetting {
  return {
    orderDateIndex: rowData.indexOf('日期'),
    productNameIndex: rowData.indexOf('品名規格'),
    productCountIndex: rowData.indexOf('數量'),
    productPriceIndex: rowData.indexOf('單價'),
  };
}

export const DEFAULT_PRODUCT_SETTING: ProductSetting = {
  orderDateIndex: 0,
  productNameIndex: 3,
  productCountIndex: 4,
  productPriceIndex: 6,
};
