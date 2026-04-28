/**
 * 對應桌面版 `models/Product.kt`。
 * 一個商品（品名 + 單價），底下持有多筆 (day, count) 訂單。
 */

import type { OrderInfo } from './order-info';

export interface ProductOrder {
  day: number;
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
    this.orderList.push({ day: info.orderDate, count: info.productCount });
  }

  getTotalCount(): number {
    let sum = 0;
    for (const o of this.orderList) sum += o.count;
    return sum;
  }

  /**
   * 在指定日期範圍內的數量加總。
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

  getTotalPrice(): number {
    return this.getTotalCount() * this.price;
  }
}
