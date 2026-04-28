/**
 * 對應桌面版 `billOverView/overViewWriter/OverViewWriter.kt`。
 *
 * 結構：單一 .xlsx，依 `bill.groupedCustomerByLine()` 每個線別一張 sheet。
 * 每張 sheet 的內容：
 *   表頭列：編號 / 客戶名稱 / 金額 / 類型
 *   客戶列：依排序檔（orderList）順序排序，找不到排在最後
 *     - 金額：含稅客戶用 `getAfterTexSum()`，否則 `getTotalPrice()`
 *     - 類型：依 isCashUser / isMonthly / isNeedTex 組合「現 / 月 / 稅」
 *
 * 檔名：`${month}月_明細總覽.xlsx`
 */

import ExcelJS from 'exceljs';

import { buildStyle, createWorkbook, workbookToBlob } from '@/infra/excel-service';

import type { Bill } from '@/domain/models/bill';
import type { CustomerModel } from '@/domain/models/customer-model';

export interface OverviewFile {
  filename: string;
  blob: Blob;
}

const FILE_END_FIX = '.xlsx';

const HEADER = ['編號', '客戶名稱', '金額', '類型'] as const;

export class OverViewWriter {
  constructor(
    private readonly bill: Bill,
    private readonly orderList: ReadonlyArray<string>
  ) {}

  async write(): Promise<OverviewFile> {
    const wb = createWorkbook();
    const grouped = this.bill.groupedCustomerByLine();

    // 依線別名稱排序，讓 sheet 順序固定（第1線 → 第2線 → ...）
    const sortedEntries = [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b));

    for (const [linePrefix, customers] of sortedEntries) {
      const ws = wb.addWorksheet(linePrefix);
      this.renderSheet(ws, customers);
    }

    return {
      filename: this.getFilename(),
      blob: await workbookToBlob(wb),
    };
  }

  private renderSheet(ws: ExcelJS.Worksheet, customers: ReadonlyArray<CustomerModel>): void {
    const borderStyle = buildStyle({ border: true });

    // 表頭
    const headerRow = ws.addRow([...HEADER]);
    headerRow.eachCell({ includeEmpty: true }, (cell) => {
      cell.style = borderStyle;
    });

    // 依排序檔排序，找不到的放最後
    const sorted = [...customers].sort((a, b) => {
      const ia = this.orderList.indexOf(a.code);
      const ib = this.orderList.indexOf(b.code);
      const sa = ia < 0 ? Number.MAX_SAFE_INTEGER : ia;
      const sb = ib < 0 ? Number.MAX_SAFE_INTEGER : ib;
      return sa - sb;
    });

    for (const c of sorted) {
      const amount = c.isNeedTex ? c.getAfterTexSum() : c.getTotalPrice();
      const row = ws.addRow([c.code, c.name, amount, this.formatType(c)]);
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.style = borderStyle;
      });
    }
  }

  private formatType(c: CustomerModel): string {
    let s = '';
    if (c.isCashUser) s += '現';
    if (c.isMonthly) s += '月';
    if (c.isNeedTex) s += '稅';
    return s;
  }

  private getFilename(): string {
    return `${this.bill.billDateInfo.month}月_明細總覽${FILE_END_FIX}`;
  }
}
