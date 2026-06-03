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
  createWorkbook,
  type PrintSettingOptions,
  setColumnWidthPoi,
  setupPrintSetting,
  workbookToBlob,
} from '@/infra/excel-service';

import {
  BILL_DATE_ROW_HEIGHT_POI,
  CENTER_LINE_WIDTH_POI,
  createCustInfo,
  createDateRangeRow,
  createFooter,
  createHeader,
  createProductRow,
  createProductRowBySlots,
  createTotalRow,
  createTotalRowBySlots,
  type DateSlot,
  FIRST_LINE_WIDTH_POI,
  getCashCustomer,
  getHalfMonthlyCustomer,
  getMonthlyCustomer,
  RIGHT_LINE_WIDTH_POI,
} from './sheet-extension';

import type {Bill} from '@/domain/models/bill';
import type {CustomerModel} from '@/domain/models/customer-model';

const BILL_PRINT_SETTINGS: PrintSettingOptions = {
    landscape: true,
    scale: 80,
    horizontalCentered: true,
    verticalCentered: true,
    margins: {
        left: 0, right: 0,
        top: 0, bottom: 0,
        header: 0, footer: 0,
    },
};

const FULL_MONTH_MAX_ROW = 19;
const FULL_MONTH_FIRST_RANGE: ReadonlyArray<number> = Array.from({ length: 15 }, (_, i) => i + 1);
const FULL_MONTH_SECOND_RANGE: ReadonlyArray<number> = Array.from({ length: 16 }, (_, i) => i + 16);
const ONE_DAY_MS = 24 * 60 * 60 * 1000;

/* ------------------------------------------------------------
   壓進一頁：品項過多的請款單自動等比例縮小，避免破版
   ------------------------------------------------------------ */

/**
 * 單一客戶請款單區塊（標題～總計～頁尾，含其後 1 列空白）的「列高總和」上限（pt）。
 *
 * 為實測校正值，非理論推導：列印縮放（80%）實際上並不會放大每頁可用高度，內容仍須
 * 以自然高度塞進實體 A5 橫向頁面。由實際輸出檔觀察：某客戶區塊約 370pt（含上下半月
 * 兩段明細）會被擠到次頁破版，而 ~336pt（含）以下的客戶均可正常列印於一頁。
 *
 * 另外每頁起始尚有約 30pt 的分頁結構空白列（不計入此區塊），因此本上限 336pt 對應到
 * 「整頁約 366pt」——即目前確認可正常列印的最高頁面。如此一來：原本就放得下（≤336pt）
 * 的客戶完全不動，過高者等比例縮小至此上限，使每頁高度都不超過已知可正常列印的範圍。
 */
const PAGE_BUDGET_PT = 336;
/** 未設定列高的空白列，ExcelJS 預設約 15pt。 */
const DEFAULT_ROW_PT = 15;

/**
 * 將 [startRow, endRow]（單一客戶的請款單，含其後屬於本頁的空白列）壓進一頁。
 * 估算整段自然列高總和，超過單頁上限時，等比例縮小每一列的「列高」與「字級」，
 * 讓內容塞回一頁且字體不被裁切。原本就放得下的客戶不會被改動。
 */
function fitBlockToOnePage(ws: ExcelJS.Worksheet, startRow: number, endRow: number): void {
    let total = 0;
    for (let r = startRow; r <= endRow; r++) {
        total += ws.getRow(r).height ?? DEFAULT_ROW_PT;
    }
    if (total <= PAGE_BUDGET_PT) return;

    const scale = PAGE_BUDGET_PT / total;
    for (let r = startRow; r <= endRow; r++) {
        const row = ws.getRow(r);
        row.height = (row.height ?? DEFAULT_ROW_PT) * scale;
        row.eachCell({ includeEmpty: true }, (cell) => {
            const font = cell.font;
            if (font?.size) {
                cell.font = { ...font, size: font.size * scale };
            }
        });
    }
}

export interface BillFile {
  /** 輸出檔名（含 .xlsx），對應桌面版 `getFilePath` */
  filename: string;
  blob: Blob;
}

export class BillWriter {
  private readonly halfMonthMaxRow: number;
    /** 顯示用 day 序列，永遠等於原 dateRange（不隨校正變動） */
    private readonly displayDateRange: number[];
    /** 顯示用月份，取自原 dates[0] */
    private readonly displayMonth: string;
    /** 顯示用民國年 */
    private readonly displayYear: string;
    /**
     * DateSlot 序列：每格的 displayDay 是原日期，sourceMonth/Day 是「校正後要去原資料找的日期」。
     * shift=+1 → display 4/1 格的 source=3/31，達成「3/31 訂單併入 4/1 格」。
     */
    private readonly halfMonthSlots: DateSlot[];

  constructor(
    private readonly bill: Bill,
    private readonly orderList: ReadonlyArray<string>,
    dateShiftDays: number = 0
  ) {
    this.halfMonthMaxRow = bill.billDateInfo.dateRange.length + 3;

      const dates = bill.billDateInfo.dates;
      this.displayDateRange = dates.map((d) => d.getDate());
      const head = dates[0];
      this.displayMonth = String(head.getMonth() + 1);
      this.displayYear = String(head.getFullYear() - 1911);

      this.halfMonthSlots = dates.map((d) => {
          const src = new Date(d.getTime() - dateShiftDays * ONE_DAY_MS);
          return {
              displayDay: d.getDate(),
              sourceMonth: src.getMonth() + 1,
              sourceDay: src.getDate(),
          };
      });
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
        setupPrintSetting(ws, BILL_PRINT_SETTINGS);
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
      setupPrintSetting(ws, BILL_PRINT_SETTINGS);
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
      setupPrintSetting(ws, BILL_PRINT_SETTINGS);
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
      const blockStart = (ws.lastRow?.number ?? 0) + 1;
      if (mode === 'full') {
        this.writeFullMonthBlock(ws, customer);
      } else {
        this.writeHalfMonthBlock(ws, customer);
      }

      // 對應桌面版： row {} ; setRowBreak(lastRowNum) ; row {}
      ws.addRow([]); // (a)，仍屬本頁，一併納入「壓進一頁」估算
      const blockEnd = ws.lastRow?.number ?? blockStart;

      // 品項過多會破版的客戶：等比例縮小列高與字級，壓回一頁
      fitBlockToOnePage(ws, blockStart, blockEnd);

      const nextRowNumber = blockEnd + 1;
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
      createCustInfo(ws, customer, this.displayYear, this.displayMonth, max);
      createDateRangeRow(ws, this.displayDateRange, 0, BILL_DATE_ROW_HEIGHT_POI);
    const range = createProductRowBySlots(ws, customer.productList, this.halfMonthSlots);
    createTotalRowBySlots(ws, customer, max, this.halfMonthSlots, range ? [range] : []);
    createFooter(ws, max);
  }

  private writeFullMonthBlock(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    const max = FULL_MONTH_MAX_ROW;
      const {year, month} = this.bill.billDateInfo;
    createHeader(ws, max);
      createCustInfo(ws, customer, year, month, max);

    // 上半月：日期 1..15，maxDate=16
    createDateRangeRow(ws, FULL_MONTH_FIRST_RANGE, 16, BILL_DATE_ROW_HEIGHT_POI);
    const upperRange = createProductRow(ws, customer.productList, FULL_MONTH_FIRST_RANGE, 16);

    // 中間空白列
    ws.addRow([]);

    // 下半月：日期 16..31，maxDate=16
    createDateRangeRow(ws, FULL_MONTH_SECOND_RANGE, 16, BILL_DATE_ROW_HEIGHT_POI);
    const lowerRange = createProductRow(ws, customer.productList, FULL_MONTH_SECOND_RANGE, 16);

    const ranges = [upperRange, lowerRange].filter(
        (r): r is NonNullable<typeof r> => r !== null
    );
    createTotalRow(ws, customer, max, ranges);
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
      const half = this.displayDateRange.includes(18) ? '下半' : '上半';
      return `${this.displayMonth}月_${half}_${prefix}.xlsx`;
  }

  private getSheetName(prefix: string): string {
      return `${prefix}${this.displayDateRange.includes(18) ? '_下' : '_上'}`;
  }
}
