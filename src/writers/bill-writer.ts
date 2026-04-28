/**
 * 對應桌面版 `billReformat/core/Writer.kt`。
 *
 * 三種輸出模式：
 *   - 半月結：依「線別」分檔，每檔一張 sheet（${prefix}_上 / ${prefix}_下）。
 *   - 全月結：所有月結客戶共一檔（sheet 名「月結」），每位客戶含上半/下半兩段。
 *   - 現金：永遠輸出（不論 isFullMonth），所有現金客戶共一檔（sheet 名「現金」）。
 *
 * 客戶分頁：每位客戶之後加一空白列 → 設定下一頁分頁符 → 再加一空白列。
 * 對應桌面版 `row {} ; setRowBreak(lastRowNum) ; row {}`。
 *
 * 檔名格式（與桌面版完全一致）：
 *   `${month}月_上半_${prefix}.xlsx`（dateRange 不含 18）
 *   `${month}月_下半_${prefix}.xlsx`（dateRange 包含 18）
 */

import ExcelJS from 'exceljs';

import {
  setColumnWidthPoi,
  setupPrintSetting,
  workbookToBlob,
  createWorkbook,
} from '@/infra/excel-service';

import {
  CENTER_LINE_WIDTH_POI,
  FIRST_LINE_WIDTH_POI,
  RIGHT_LINE_WIDTH_POI,
  createCustInfo,
  createDateRangeRow,
  createFooter,
  createHeader,
  createProductRow,
  createTotalRow,
  getCashCustomer,
  getHalfMonthlyCustomer,
  getMonthlyCustomer,
} from './sheet-extension';

import type { Bill } from '@/domain/models/bill';
import type { CustomerModel } from '@/domain/models/customer-model';

const FULL_MONTH_MAX_ROW = 19;
const FULL_MONTH_FIRST_RANGE: ReadonlyArray<number> = Array.from({ length: 15 }, (_, i) => i + 1);
const FULL_MONTH_SECOND_RANGE: ReadonlyArray<number> = Array.from({ length: 16 }, (_, i) => i + 16);

export interface BillFile {
  /** 輸出檔名（含 .xlsx），對應桌面版 `getFilePath` */
  filename: string;
  blob: Blob;
}

export class BillWriter {
  private readonly halfMonthMaxRow: number;

  constructor(
    private readonly bill: Bill,
    private readonly orderList: ReadonlyArray<string>
  ) {
    this.halfMonthMaxRow = bill.billDateInfo.dateRange.length + 3;
  }

  /** 產出所有檔案的 Blob 列表，呼叫端負責下載/打包 ZIP。 */
  async write(isFullMonth: boolean): Promise<BillFile[]> {
    const files: BillFile[] = [];

    if (isFullMonth) {
      const file = await this.writeFullMonth();
      if (file) files.push(file);
    } else {
      files.push(...(await this.writeHalfMonth()));
    }

    const cashFile = await this.writeCash();
    if (cashFile) files.push(cashFile);

    return files;
  }

  /* ------------------------------------------------------------
     P2.11 半月結：依「線別」分檔
     ------------------------------------------------------------ */
  private async writeHalfMonth(): Promise<BillFile[]> {
    const grouped = this.bill.groupedCustomerByLine();
    const files: BillFile[] = [];

    for (const [linePrefix, customers] of grouped) {
      const halfNonCash = this.sortByOrderList(getHalfMonthlyCustomer(customers));
      if (halfNonCash.length === 0) continue;

      const wb = createWorkbook();
      const ws = wb.addWorksheet(this.getSheetName(linePrefix));
      setupPrintSetting(ws);
      this.setupHalfMonthColumns(ws);

      this.writeCustomers(ws, halfNonCash, 'half');

      files.push({
        filename: this.getFilename(linePrefix),
        blob: await workbookToBlob(wb),
      });
    }
    return files;
  }

  /* ------------------------------------------------------------
     P2.12 全月結：月結客戶共一檔，sheet 名「月結」
     ------------------------------------------------------------ */
  private async writeFullMonth(): Promise<BillFile | null> {
    const monthly = this.sortByOrderList(getMonthlyCustomer(this.bill.customerModels));
    if (monthly.length === 0) return null;

    const wb = createWorkbook();
    const ws = wb.addWorksheet('月結');
    setupPrintSetting(ws);
    this.setupFullMonthColumns(ws);

    this.writeCustomers(ws, monthly, 'full');

    return {
      filename: this.getFilename('月結'),
      blob: await workbookToBlob(wb),
    };
  }

  /* ------------------------------------------------------------
     現金：所有現金客戶共一檔（與 isFullMonth 無關）
     ------------------------------------------------------------ */
  private async writeCash(): Promise<BillFile | null> {
    const cash = this.sortByOrderList(getCashCustomer(this.bill.customerModels));
    if (cash.length === 0) return null;

    const wb = createWorkbook();
    const ws = wb.addWorksheet('現金');
    setupPrintSetting(ws);
    this.setupHalfMonthColumns(ws);

    this.writeCustomers(ws, cash, 'half');

    return {
      filename: this.getFilename('現金'),
      blob: await workbookToBlob(wb),
    };
  }

  /* ------------------------------------------------------------
     單客戶內容（half / full month）
     ------------------------------------------------------------ */
  private writeCustomers(
    ws: ExcelJS.Worksheet,
    customers: ReadonlyArray<CustomerModel>,
    mode: 'half' | 'full'
  ): void {
    customers.forEach((customer, idx) => {
      if (mode === 'full') {
        this.writeFullMonthBlock(ws, customer);
      } else {
        this.writeHalfMonthBlock(ws, customer);
      }

      // 對應桌面版： row {} ; setRowBreak(lastRowNum) ; row {}
      ws.addRow([]); // (a)
      const nextRowNumber = (ws.lastRow?.number ?? 0) + 1;
      // 只要不是最後一位客戶，就把分頁設在 (c) 那一列
      if (idx < customers.length - 1) {
        ws.getRow(nextRowNumber).addPageBreak();
      }
      ws.addRow([]); // (c)
    });
  }

  private writeHalfMonthBlock(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    const max = this.halfMonthMaxRow;
    createHeader(ws, max);
    createCustInfo(ws, customer, this.bill.billDateInfo.year, this.bill.billDateInfo.month, max);
    createDateRangeRow(ws, this.bill.billDateInfo.dateRange);
    createProductRow(ws, customer.productList, this.bill.billDateInfo.dateRange);
    createTotalRow(ws, customer, max);
    createFooter(ws, max);
  }

  private writeFullMonthBlock(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    const max = FULL_MONTH_MAX_ROW;
    createHeader(ws, max);
    createCustInfo(ws, customer, this.bill.billDateInfo.year, this.bill.billDateInfo.month, max);

    // 上半月：日期 1..15，maxDate=16
    createDateRangeRow(ws, FULL_MONTH_FIRST_RANGE, 16);
    createProductRow(ws, customer.productList, FULL_MONTH_FIRST_RANGE, 16);

    // 中間空白列
    ws.addRow([]);

    // 下半月：日期 16..31，maxDate=16
    createDateRangeRow(ws, FULL_MONTH_SECOND_RANGE, 16);
    createProductRow(ws, customer.productList, FULL_MONTH_SECOND_RANGE, 16);

    createTotalRow(ws, customer, max);
    createFooter(ws, max);
  }

  /* ------------------------------------------------------------
     欄寬設定
     ------------------------------------------------------------ */

  private setupHalfMonthColumns(ws: ExcelJS.Worksheet): void {
    const max = this.halfMonthMaxRow;
    setColumnWidthPoi(ws, 1, FIRST_LINE_WIDTH_POI);
    for (let i = 2; i < max - 1; i++) {
      setColumnWidthPoi(ws, i, CENTER_LINE_WIDTH_POI);
    }
    for (let i = max - 1; i <= max + 1; i++) {
      setColumnWidthPoi(ws, i, RIGHT_LINE_WIDTH_POI);
    }
  }

  private setupFullMonthColumns(ws: ExcelJS.Worksheet): void {
    const max = FULL_MONTH_MAX_ROW;
    setColumnWidthPoi(ws, 1, FIRST_LINE_WIDTH_POI);
    for (let i = 2; i < max - 1; i++) {
      setColumnWidthPoi(ws, i, CENTER_LINE_WIDTH_POI);
    }
    for (let i = max - 1; i <= max + 1; i++) {
      setColumnWidthPoi(ws, i, RIGHT_LINE_WIDTH_POI);
    }
  }

  /* ------------------------------------------------------------
     工具
     ------------------------------------------------------------ */

  /** 客戶依排序檔的順序排列，找不到的排在最後（索引 MAX_SAFE_INTEGER）。 */
  private sortByOrderList(customers: ReadonlyArray<CustomerModel>): CustomerModel[] {
    return [...customers].sort((a, b) => {
      const ia = this.orderList.indexOf(a.code);
      const ib = this.orderList.indexOf(b.code);
      const sa = ia < 0 ? Number.MAX_SAFE_INTEGER : ia;
      const sb = ib < 0 ? Number.MAX_SAFE_INTEGER : ib;
      return sa - sb;
    });
  }

  private getFilename(prefix: string): string {
    const half = this.bill.billDateInfo.dateRange.includes(18) ? '下半' : '上半';
    return `${this.bill.billDateInfo.month}月_${half}_${prefix}.xlsx`;
  }

  private getSheetName(prefix: string): string {
    return `${prefix}${this.bill.billDateInfo.dateRange.includes(18) ? '_下' : '_上'}`;
  }
}
