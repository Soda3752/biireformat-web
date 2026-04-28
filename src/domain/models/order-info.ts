/**
 * 對應桌面版 `models/OrderInfo.kt`。
 * 一筆銷貨資料：日期、月份、品名、數量、單價。
 * 桌面版 dateString 會去除 "銷" 前綴後再交給 DateUtility 解析。
 */

import { parseDayOfMonth, parseMonth } from '../date-utility';
import { DEFAULT_PRODUCT_SETTING, type ProductSetting } from './product-setting';

export interface OrderInfo {
  orderDate: number;
  month: number;
  productName: string;
  productCount: number;
  productPrice: number;
}

export function parseOrderInfo(
  rowData: ReadonlyArray<string>,
  productSetting?: ProductSetting | null
): OrderInfo {
  const indexSetting = productSetting ?? DEFAULT_PRODUCT_SETTING;
  const dateString = rowData[indexSetting.orderDateIndex].replace('銷', '');
  return {
    orderDate: parseDayOfMonth(dateString),
    month: parseMonth(dateString),
    productName: rowData[indexSetting.productNameIndex],
    productCount: Number(rowData[indexSetting.productCountIndex]),
    productPrice: Number(rowData[indexSetting.productPriceIndex]),
  };
}
