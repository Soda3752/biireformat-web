/**
 * 對應桌面版 `models/CustomerModel.kt`。
 * 一位客戶（含結帳模式旗標、產品列表、商品設定）。
 */

import type {OrderInfo} from './order-info';
import {Product} from './product';
import type {ProductSetting} from './product-setting';

const PARAM_NAME = '客戶名稱';
const PARAM_CODE = '客戶編號';
const PARAM_SETTING = '傳真電話';

export class CustomerModel {
  readonly name: string;
  readonly code: string;

  productSetting: ProductSetting | null = null;
  readonly productList: Product[] = [];

  /** 是否為月結 */
  isMonthly = false;
  /** 是否要加稅 */
  isNeedTex = false;
  /** 是否為現金結 */
  isCashUser = false;

  constructor(name: string, code: string) {
    this.name = name;
    this.code = code;
  }

  /**
   * 範例 rowData：[客戶名稱, 辭修弘爺, 客戶編號, 1001, 統一編號, , 電話號碼, ]
   */
  static newInstanceWithRowData(rowData: ReadonlyArray<string>): CustomerModel {
    const name = rowData[rowData.indexOf(PARAM_NAME) + 1];
    const code = rowData[rowData.indexOf(PARAM_CODE) + 1];
    return new CustomerModel(name, code);
  }

  appendOrderInfo(orderInfo: OrderInfo): void {
    const found = this.productList.find((p) => p.name === orderInfo.productName);
    if (found) {
      found.appendOrderInfo(orderInfo);
    } else {
      const np = new Product(orderInfo.productName, orderInfo.productPrice);
      np.appendOrderInfo(orderInfo);
      this.productList.push(np);
    }
  }

  /**
   * 從「客戶地址」列上的 PARAM_SETTING 欄位讀取結帳模式設定。
   * 桌面版規則：包含「月」 → 月結；包含「稅」 → 含稅；包含「現」 → 現金結。
   */
  setUpConfig(rowData: ReadonlyArray<string>): void {
    const idx = rowData.indexOf(PARAM_SETTING);
    const configInfo = idx >= 0 ? rowData[idx + 1] ?? '' : '';
    this.isMonthly = configInfo.includes('月');
    this.isNeedTex = configInfo.includes('稅');
    this.isCashUser = configInfo.includes('現');
  }

  getTotalPrice(): number {
    let sum = 0;
    for (const p of this.productList) sum += p.getTotalPrice();
    return Math.round(sum);
  }

  getTex(): number {
    return Math.round(this.getTotalPrice() * 0.05);
  }

  getAfterTexSum(): number {
    return this.getTotalPrice() + this.getTex();
  }

    /**
     * 以「(月, 日) slot 序列」加總總金額。
     * 用於帳單分頁：跨月或日期校正時，僅計入確實會輸出到表格中的訂單。
     */
    getTotalPriceForDates(sources: ReadonlyArray<{ month: number; day: number }>): number {
        let sum = 0;
        for (const p of this.productList) sum += p.getPriceForDates(sources);
        return Math.round(sum);
    }

    getTexForDates(sources: ReadonlyArray<{ month: number; day: number }>): number {
        return Math.round(this.getTotalPriceForDates(sources) * 0.05);
    }

    getAfterTexSumForDates(sources: ReadonlyArray<{ month: number; day: number }>): number {
        return this.getTotalPriceForDates(sources) + this.getTexForDates(sources);
    }
}
