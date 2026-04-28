/**
 * 對應桌面版 `billReformat/core/ExcelRowData.kt`。
 * 依「列首文字」判斷該列屬於哪個區塊類型。
 */

export const ExcelRowType = {
  BillSettingInfo: 'BillSettingInfo',
  CustomerData: 'CustomerData',
  CustomerSetting: 'CustomerSetting',
  ProductRowSetting: 'ProductRowSetting',
  ProductSellData: 'ProductSellData',
} as const;

export type ExcelRowType = (typeof ExcelRowType)[keyof typeof ExcelRowType];

const ROW_TYPE_PREFIX: ReadonlyArray<{ type: ExcelRowType; startsWith: string }> = [
  { type: ExcelRowType.BillSettingInfo, startsWith: '客戶區間' },
  { type: ExcelRowType.CustomerData, startsWith: '客戶名稱' },
  { type: ExcelRowType.CustomerSetting, startsWith: '客戶地址' },
  { type: ExcelRowType.ProductRowSetting, startsWith: '日期' },
  { type: ExcelRowType.ProductSellData, startsWith: '銷' },
];

export function getRowType(rowData: ReadonlyArray<string>): ExcelRowType | null {
  const firstString = rowData[0];
  if (firstString == null) return null;
  for (const def of ROW_TYPE_PREFIX) {
    if (firstString.startsWith(def.startsWith)) return def.type;
  }
  return null;
}
