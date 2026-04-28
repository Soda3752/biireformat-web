/**
 * 對應桌面版 `models/Bill.kt`。
 * 整份帳單的容器：日期區間 + 多位客戶 + 「目前游標客戶」。
 * Reader 一邊掃 sheet，一邊呼叫這裡的方法把每一列塞進對的位置。
 */

import { CustomerModel } from './customer-model';
import { type BillDateInfo, parseBillDateInfo } from './bill-date-info';
import { parseOrderInfo } from './order-info';
import { parseProductSetting } from './product-setting';

export class Bill {
  /** 帳單日期資訊；要求在 setBillDateInfo 後才會被使用。 */
  billDateInfo!: BillDateInfo;
  readonly customerModels: CustomerModel[] = [];

  /** 目前正在組裝的客戶（reader 游標）。null 代表尚未開始或已 commit。 */
  currentCustomer: CustomerModel | null = null;

  setBillDateInfo(rowData: ReadonlyArray<string>): void {
    this.billDateInfo = parseBillDateInfo(rowData);
  }

  /**
   * 依客戶代碼前綴把客戶分到不同線別（例如「第1線」「第2線」）。
   */
  groupedCustomerByLine(): Map<string, CustomerModel[]> {
    const map = new Map<string, CustomerModel[]>();
    for (const c of this.customerModels) {
      const name = `第${c.code.substring(0, 1)}線`;
      const list = map.get(name);
      if (list) {
        list.push(c);
      } else {
        map.set(name, [c]);
      }
    }
    return map;
  }

  /**
   * 開始一位新客戶；若先前游標客戶尚未 commit，先 commit 進列表。
   */
  newCustomerModel(rowData: ReadonlyArray<string>): void {
    this.appendCustomer();
    this.currentCustomer = CustomerModel.newInstanceWithRowData(rowData);
  }

  /** 把游標客戶 commit 進列表並清空游標。 */
  appendCustomer(): void {
    if (this.currentCustomer) {
      this.customerModels.push(this.currentCustomer);
      this.currentCustomer = null;
    }
  }

  configCustomer(rowData: ReadonlyArray<string>): void {
    this.currentCustomer?.setUpConfig(rowData);
  }

  setProductSetting(rowData: ReadonlyArray<string>): void {
    if (this.currentCustomer) {
      this.currentCustomer.productSetting = parseProductSetting(rowData);
    }
  }

  /**
   * 加入一筆銷貨資料。傳入 excludeMonth 可排除特定月份（桌面版預留功能，預設 -1 即不排除）。
   */
  addProduct(rowData: ReadonlyArray<string>, excludeMonth?: number | null): void {
    if (!this.currentCustomer) return;
    const orderInfo = parseOrderInfo(rowData, this.currentCustomer.productSetting);
    if (excludeMonth != null && orderInfo.month === excludeMonth) return;
    this.currentCustomer.appendOrderInfo(orderInfo);
  }
}
