/**
 * 對應桌面版 `models/Product.kt`。
 * 一個商品（品名 + 單價），底下持有多筆 (day, count) 訂單。
 */

import type {OrderInfo} from './order-info';

export interface ProductOrder {
  day: number;
    /** 訂單對應的月份。跨月日期區間時用於精準配對。 */
    month: number;
  count: number;
}

export class Product {
  readonly name: string;
  readonly price: number;
  readonly orderList: ProductOrder[] = [];

  constructor(name: string, price: number) {
    this.name = name;
    this.price = price;
  }

  appendOrderInfo(info: OrderInfo): void {
      this.orderList.push({day: info.orderDate, month: info.month, count: info.productCount});
  }

  getTotalCount(): number {
    let sum = 0;
    for (const o of this.orderList) sum += o.count;
    return sum;
  }

  /**
   * 在指定日期範圍內的數量加總。
   * 為維持其他寫入器（overview/delivery-fee/daily-count）的呼叫，僅以 day 比對，不檢查 month。
   * 帳單分頁的跨月/校正情境改走 getCountForDates。
   */
  getDateRangeCount(dateRange: ReadonlyArray<number>): number {
    let sum = 0;
    for (const day of dateRange) {
      const order = this.orderList.find((o) => o.day === day);
      if (order) sum += order.count;
    }
    return sum;
  }

  /**
   * 在指定日期範圍內的金額加總（數量 × 單價）。
   * 對齊桌面版同名 helper：先加總數量，再乘 price。
   */
  getDateRangePrice(dateRange: ReadonlyArray<number>): number {
    return this.getDateRangeCount(dateRange) * this.price;
  }

    /** 以「(月, 日)」配對加總數量；用於跨月或日期校正情境。 */
    getCountForDates(slots: ReadonlyArray<{ month: number; day: number }>): number {
        let sum = 0;
        for (const slot of slots) {
            const order = this.orderList.find((o) => o.day === slot.day && o.month === slot.month);
            if (order) sum += order.count;
        }
        return sum;
    }

    getPriceForDates(slots: ReadonlyArray<{ month: number; day: number }>): number {
        return this.getCountForDates(slots) * this.price;
    }

  getTotalPrice(): number {
    return this.getTotalCount() * this.price;
  }
}
