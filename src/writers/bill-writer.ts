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
  BILL_ROW_HEIGHT_POI,
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
   壓進一頁：品項過多的請款單自動套用緊湊版型，避免破版
   ------------------------------------------------------------ */

/**
 * 單一客戶請款單區塊（標題～總計～頁尾，含其後 1 列空白）的「列高總和」上限（pt）。
 * 實測校正值：整頁超過約 370pt 會破版。blank(a)/(c) 已壓到 1pt，故頁面總高
 * ≈ block + 2pt。本上限設為 360pt → 整頁約 362pt，保留約 8pt 安全餘裕。
 */
const PAGE_BUDGET_PT = 360;
/** 未設定列高的空白列，ExcelJS 預設約 15pt。 */
const DEFAULT_ROW_PT = 15;

/** 四捨五入到小數點後兩位，避免寫出過長小數／科學記號（如 1.2e-8）導致 Excel 無法開啟。 */
const round2 = (n: number): number => Math.round(n * 100) / 100;

/**
 * 緊湊版型行高（pt）與字級（pt）——依使用者指定（px × 0.75 換算）：
 *   頭4列（青坊食品行/請款單/客戶編號/名稱）→ 行高 21.75pt（29px）、字級 20pt
 *   中間內容（日期表頭/明細/總計）          → 行高 19.5pt（26px）、字級 15pt
 *   尾3列（訂貨專線/銀行/匯款提醒）         → 行高 12.75pt（17px）、字級 14pt
 */
const COMPACT_HEADER_H    = 21.75;  // 29px × 0.75
const COMPACT_HEADER_FONT = 20;
const COMPACT_CONTENT_H    = 19.5;  // 26px × 0.75
const COMPACT_CONTENT_FONT = 15;
const COMPACT_FOOTER_H    = 12.75;  // 17px × 0.75
const COMPACT_FOOTER_FONT = 14;
/** 結構性尾部空白列（blank a）的緊湊列高（pt）——壓到 1pt，不佔版面。 */
const COMPACT_BLANK_H = 1;
/** 視覺間距列（上半／下半分隔、頁尾前留白）的緊湊列高（pt）——保留足夠高度讓段落清楚可辨。 */
const COMPACT_GAP_H = 8;

/** 依列在區塊中的位置判斷緊湊版型字級；endRow（尾部空白）回傳 null。 */
function compactFontForRow(r: number, startRow: number, endRow: number): number | null {
    if (r === endRow) return null;                                           // 尾部空白列（a）
    if (r >= endRow - 3 && r <= endRow - 1) return COMPACT_FOOTER_FONT;     // 尾3列 footer（先判，防止極短區塊重疊）
    if (r <= startRow + 3) return COMPACT_HEADER_FONT;                      // 頭4列
    return COMPACT_CONTENT_FONT;
}

/**
 * 依列在區塊中的位置判斷緊湊版型列高（pt）。行高與字級獨立設定。
 * 有高度的內容列 → COMPACT_CONTENT_H；無高度的視覺間距列 → COMPACT_GAP_H；
 * 尾部結構空白列 (a) → COMPACT_BLANK_H。
 */
function compactHeightForRow(ws: ExcelJS.Worksheet, r: number, startRow: number, endRow: number): number {
    if (r === endRow) return COMPACT_BLANK_H;
    if (r >= endRow - 3 && r <= endRow - 1) return COMPACT_FOOTER_H;
    if (r <= startRow + 3) return COMPACT_HEADER_H;
    return ws.getRow(r).height != null ? COMPACT_CONTENT_H : COMPACT_GAP_H;
}

/**
 * 將 [startRow, endRow]（單一客戶的請款單，含其後屬於本頁的空白列）壓進一頁。
 *
 * 放得下的客戶（自然高度 ≤ PAGE_BUDGET_PT）完全不動。
 * 超出上限的客戶套用「緊湊版型」：
 *   - 頭4列 21.75pt/20pt、中間內容 19.5pt/15pt、尾3列 12.75pt/14pt（依使用者指定）。
 *   - 列高以字級為基準計算緊湊高度；若緊湊高度仍超過上限，再等比例縮小列高
 *     （字級維持不動，確保文字可讀）。
 *
 * 注意：ExcelJS 合併儲存格附屬格共用主格樣式，因此採兩階段（先快照、再寫入）
 * 以避免共用樣式被重複縮放而壓成趨近 0 / 科學記號，導致 Excel 無法開啟。
 */
function fitBlockToOnePage(ws: ExcelJS.Worksheet, startRow: number, endRow: number): void {
    let total = 0;
    for (let r = startRow; r <= endRow; r++) {
        total += ws.getRow(r).height ?? DEFAULT_ROW_PT;
    }
    if (total <= PAGE_BUDGET_PT) return;

    // 計算緊湊版型基礎列高總和，必要時等比例縮小列高（字級保持不動）
    let compactTotal = 0;
    for (let r = startRow; r <= endRow; r++) {
        compactTotal += compactHeightForRow(ws, r, startRow, endRow);
    }
    const scale = compactTotal > PAGE_BUDGET_PT ? PAGE_BUDGET_PT / compactTotal : 1;

    // 第一階段：快照目標列高與（主格的）目標字級，避免共用樣式重複縮放
    const targetHeights: number[] = [];
    const fontTargets: Array<{ cell: ExcelJS.Cell; font: Partial<ExcelJS.Font>; size: number }> = [];
    for (let r = startRow; r <= endRow; r++) {
        targetHeights.push(compactHeightForRow(ws, r, startRow, endRow) * scale);
        const targetSize = compactFontForRow(r, startRow, endRow);
        if (targetSize !== null) {
            ws.getRow(r).eachCell({ includeEmpty: true }, (cell) => {
                if (cell.master !== cell) return; // 跳過合併附屬格
                if (cell.font?.size) fontTargets.push({ cell, font: cell.font, size: targetSize });
            });
        }
    }

    // 第二階段：以絕對值寫回（不累乘，不受共用樣式影響）
    let i = 0;
    for (let r = startRow; r <= endRow; r++) {
        ws.getRow(r).height = round2(targetHeights[i++]);
    }
    for (const { cell, font, size } of fontTargets) {
        cell.font = { ...font, size };
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
      // 空白列 (a)/(c) 均設為 1pt，避免佔用版面空間
      const rowA = ws.addRow([]); // (a)：頁尾緩衝列，納入壓縮估算
      rowA.height = 1;
      const blockEnd = rowA.number;

      // 品項過多會破版的客戶：套用緊湊版型，壓回一頁；完成後確保 (a) 仍為 1pt
      fitBlockToOnePage(ws, blockStart, blockEnd);
      rowA.height = 1;

      // 分頁起始列 (c)：先建立列再標記分頁，避免預先建立列導致高度設定失效
      const rowC = ws.addRow([]);
      rowC.height = 1;
      if (idx < customers.length - 1) {
        rowC.addPageBreak();
      }
    });
  }

  private writeHalfMonthBlock(ws: ExcelJS.Worksheet, customer: CustomerModel): void {
    const max = this.halfMonthMaxRow;
    createHeader(ws, max);
      createCustInfo(ws, customer, this.displayYear, this.displayMonth, max);
      createDateRangeRow(ws, this.displayDateRange, 0, BILL_ROW_HEIGHT_POI);
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
    createDateRangeRow(ws, FULL_MONTH_FIRST_RANGE, 16, BILL_ROW_HEIGHT_POI);
    const upperRange = createProductRow(ws, customer.productList, FULL_MONTH_FIRST_RANGE, 16);

    // 中間空白列
    ws.addRow([]);

    // 下半月：日期 16..31，maxDate=16
    createDateRangeRow(ws, FULL_MONTH_SECOND_RANGE, 16, BILL_ROW_HEIGHT_POI);
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
