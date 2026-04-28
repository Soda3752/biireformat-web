/**
 * 對應桌面版 `deliverFee/DeliveryFeeWriter.kt`。
 *
 * 結構：單一 .xlsx，單一 sheet「代送費」。
 *  1) 排除客戶 EXCLUDE_LIST = ['7001']
 *  2) 欄寬：第 1 欄寬欄、中段窄欄、最右 3 欄寬欄
 *  3) 每位客戶輸出：
 *      - createSingleLineCustInfo（單行客戶資訊）
 *      - createDateRangeRow（日期表頭）
 *      - 商品列：依 SortingList 排序，計算 deliveryFee × count
 *      - 客戶代送費總計列（紅字無，純黑框）
 *      - 空白列分隔
 *  4) 結尾總統計表：標題行（合併 4 欄）+ 每店家行（合併 4 欄）+ 總計行（合併 4 欄）
 *
 * 列印設定：A4、橫向、fitToPage、fitWidth=1
 * 檔名：`${month}月_(上半|下半)_代送費.xlsx`（dateRange 含 18 → 下半）
 */

import ExcelJS from 'exceljs';

import {
  PAPER_A4,
  buildStyle,
  createWorkbook,
  mergeRange,
  setColumnWidthPoi,
  setRowHeightPoi,
  setupPrintSetting,
  workbookToBlob,
  writeMixedRow,
  type StyledCell,
} from '@/infra/excel-service';

import {
  CENTER_LINE_WIDTH_POI,
  FIRST_LINE_WIDTH_POI,
  FONT_SIZE_POI,
  RIGHT_LINE_WIDTH_POI,
  ROW_HEIGHT_POI,
  createDateRangeRow,
  createSingleLineCustInfo,
} from './sheet-extension';

import { poiFontSize, KAI_FONT } from '@/infra/excel-service';

import type { Bill } from '@/domain/models/bill';
import type { CustomerModel } from '@/domain/models/customer-model';
import { getDeliveryFee, getItemIndex } from '@/domain/sorting-list';

const FILE_END_FIX = '.xlsx';
const EXCLUDE_LIST: ReadonlyArray<string> = ['7001'];

export interface DeliveryFeeFile {
  filename: string;
  blob: Blob;
}

export class DeliveryFeeWriter {
  /** 對應桌面版 `halfMonthMaxRowSize`：日期欄數 + 3（數量/單價/合計） */
  private readonly halfMonthMaxRow: number;

  constructor(private readonly bill: Bill) {
    this.halfMonthMaxRow = bill.billDateInfo.dateRange.length + 3;
  }

  async write(): Promise<DeliveryFeeFile> {
    const wb = createWorkbook();
    const ws = wb.addWorksheet('代送費');

    setupPrintSetting(ws, {
      paperSize: PAPER_A4,
      landscape: true,
      fitToPage: true,
      fitWidth: 1,
      fitHeight: 0,
    });

    this.setupColumns(ws);

    // 輸出每位客戶（排除 EXCLUDE_LIST）
    const customers = this.bill.customerModels.filter((c) => !EXCLUDE_LIST.includes(c.code));
    for (const customer of customers) {
      this.writeCustomer(ws, customer);
      ws.addRow([]); // 客戶分隔
    }

    // 結尾總統計表
    this.writeTotalStatistics(ws, customers);

    return {
      filename: this.getFilename(),
      blob: await workbookToBlob(wb),
    };
  }

  /* ------------------------------------------------------------
     欄寬：第 1 欄寬、中段窄、最右 3 欄寬
     ------------------------------------------------------------ */
  private setupColumns(ws: ExcelJS.Worksheet): void {
    const max = this.halfMonthMaxRow;
    setColumnWidthPoi(ws, 1, FIRST_LINE_WIDTH_POI);
    for (let i = 2; i < max - 1; i++) {
      setColumnWidthPoi(ws, i, CENTER_LINE_WIDTH_POI);
    }
    for (let i = max - 1; i <= max + 1; i++) {
      setColumnWidthPoi(ws, i, RIGHT_LINE_WIDTH_POI);
    }
  }

  /* ------------------------------------------------------------
     單客戶區塊
     ------------------------------------------------------------ */
  private writeCustomer(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    createSingleLineCustInfo(
      ws,
      customer,
      this.bill.billDateInfo.year,
      this.bill.billDateInfo.month,
      this.halfMonthMaxRow
    );
    createDateRangeRow(ws, this.bill.billDateInfo.dateRange);
    this.writeProductRows(ws, customer);
    this.writeCustomerTotalRow(ws, customer);
  }

  /** 商品列：依 SortingList 排序，每列輸出代送費明細。 */
  private writeProductRows(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    const cellStyle = buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI) },
      border: true,
      align: 'center',
    });

    const dateRange = this.bill.billDateInfo.dateRange;
    const sorted = [...customer.productList].sort(
      (a, b) => getItemIndex(a.name) - getItemIndex(b.name)
    );

    for (const product of sorted) {
      const deliveryFee = getDeliveryFee(product.name) ?? 0;
      const totalCount = product.getTotalCount();
      const totalPrice = totalCount * deliveryFee;

      const cells: StyledCell[] = [[product.name, cellStyle]];
      for (const day of dateRange) {
        const order = product.orderList.find((o) => o.day === day);
        cells.push([order ? order.count : '', cellStyle]);
      }
      cells.push([totalCount, cellStyle], [deliveryFee, cellStyle], [totalPrice, cellStyle]);

      const row = writeMixedRow(ws, cells, cellStyle);
      setRowHeightPoi(row, ROW_HEIGHT_POI);
    }
  }

  /**
   * 客戶代送費總計列：右側兩格「總計 / 金額」，前面留 (max - 1) 個空白。
   * 對應 `createDeliveryFeeTotalRow`。
   */
  private writeCustomerTotalRow(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    const borderStyle = buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI) },
      border: true,
      align: 'center',
    });

    const total = Math.round(
      customer.productList.reduce((sum, p) => {
        const fee = getDeliveryFee(p.name) ?? 0;
        return sum + p.getTotalCount() * fee;
      }, 0)
    );

    const cells: StyledCell[] = [];
    for (let i = 0; i < this.halfMonthMaxRow - 1; i++) {
      cells.push(['', borderStyle]);
    }
    cells.push(['總計', borderStyle], [total, borderStyle]);

    const row = writeMixedRow(ws, cells, borderStyle);
    setRowHeightPoi(row, ROW_HEIGHT_POI);
  }

  /* ------------------------------------------------------------
     總統計表（每店家逐列 + 總計，最右 4 欄合併儲存格）
     ------------------------------------------------------------ */
  private writeTotalStatistics(
    ws: ExcelJS.Worksheet,
    customers: ReadonlyArray<CustomerModel>
  ): void {
    const borderStyle = buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI) },
      border: true,
      align: 'center',
    });
    const emptyStyle = buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI) },
    });

    const max = this.halfMonthMaxRow;
    // 桌面版的合併範圍：col (max - 4) .. (max - 1)（POI 0-indexed）
    // → ExcelJS 1-indexed: col (max - 3) .. max
    const mergeStartCol = max - 3;
    const mergeEndCol = max;
    // 標題、店家、總計三類列的「右側 4 欄」前面要留多少空白：
    // 桌面版：repeat(halfMonthMaxRowSize - 4) cell("", emptyStyle)
    const leftPadCount = max - 4;

    // 標題行：「代送費」
    this.writeStatisticsRow(ws, '代送費', '', leftPadCount, borderStyle, emptyStyle, true);
    {
      const rowIndex = ws.lastRow!.number;
      mergeRange(ws, rowIndex, mergeStartCol, rowIndex, mergeEndCol);
    }

    // 每店家逐列
    const customerTotals = customers.map((c) => ({
      name: c.name,
      total: Math.round(
        c.productList.reduce((sum, p) => {
          const fee = getDeliveryFee(p.name) ?? 0;
          return sum + p.getTotalCount() * fee;
        }, 0)
      ),
    }));

    for (const { name, total } of customerTotals) {
      this.writeStatisticsRow(ws, name, total, leftPadCount, borderStyle, emptyStyle, false);
      const rowIndex = ws.lastRow!.number;
      mergeRange(ws, rowIndex, mergeStartCol, rowIndex, mergeEndCol);
    }

    // 總計行
    const grandTotal = customerTotals.reduce((sum, it) => sum + it.total, 0);
    this.writeStatisticsRow(ws, '總計', grandTotal, leftPadCount, borderStyle, emptyStyle, false);
    {
      const rowIndex = ws.lastRow!.number;
      mergeRange(ws, rowIndex, mergeStartCol, rowIndex, mergeEndCol);
    }
  }

  /**
   * 共用列輸出：左側 leftPadCount 個空白（emptyStyle）+
   * 標籤 cell（borderStyle）+ 3 個空白 cell（borderStyle）+
   * 數值 cell（borderStyle）。
   *
   * 標題行 `isHeader=true` 時數值欄輸出空字串（與桌面版一致）。
   */
  private writeStatisticsRow(
    ws: ExcelJS.Worksheet,
    label: string,
    value: number | string,
    leftPadCount: number,
    borderStyle: Partial<ExcelJS.Style>,
    emptyStyle: Partial<ExcelJS.Style>,
    isHeader: boolean
  ): void {
    const cells: StyledCell[] = [];
    for (let i = 0; i < leftPadCount; i++) cells.push(['', emptyStyle]);
    cells.push([label, borderStyle]);
    for (let i = 0; i < 3; i++) cells.push(['', borderStyle]);
    cells.push([isHeader ? '' : value, borderStyle]);

    const row = writeMixedRow(ws, cells);
    setRowHeightPoi(row, ROW_HEIGHT_POI);
  }

  /* ------------------------------------------------------------
     檔名
     ------------------------------------------------------------ */
  private getFilename(): string {
    const half = this.bill.billDateInfo.dateRange.includes(18) ? '下半' : '上半';
    return `${this.bill.billDateInfo.month}月_${half}_代送費${FILE_END_FIX}`;
  }
}
