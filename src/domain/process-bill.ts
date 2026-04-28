/**
 * 對應桌面版 `billReformat/BillReformatTab.kt:124-141` 的 `processExcel`。
 * 把 Excel 列分派到對應的 Bill 方法上，最後 commit 最後一位客戶。
 *
 * 桌面版預留排除月份功能（`EXCLUDE_MONTH = -1` 即不排除）。
 */

import { ExcelRowType } from './excel-row-data';
import { Bill } from './models/bill';
import { parseBillFile } from '@/readers/bill-reader';

export const EXCLUDE_MONTH: number = -1;

export async function processBillFile(source: File | ArrayBuffer | Blob): Promise<Bill> {
  const bill = new Bill();
  await parseBillFile(source, ({ type, values }) => {
    switch (type) {
      case ExcelRowType.BillSettingInfo:
        bill.setBillDateInfo(values);
        break;
      case ExcelRowType.CustomerData:
        bill.newCustomerModel(values);
        break;
      case ExcelRowType.CustomerSetting:
        bill.configCustomer(values);
        break;
      case ExcelRowType.ProductRowSetting:
        bill.setProductSetting(values);
        break;
      case ExcelRowType.ProductSellData:
        bill.addProduct(values, EXCLUDE_MONTH);
        break;
    }
  });
  bill.appendCustomer();
  return bill;
}
