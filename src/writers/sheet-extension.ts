/**
 * 對應桌面版 `billReformat/core/SheetExtension.kt`，是帳單輸出的樣式 helper 全集。
 * 含：
 *  - 客戶分組：getMonthlyCustomer / getHalfMonthlyCustomer / getCashCustomer
 *  - 樣式 helper：createHeader / createCustInfo / createSingleLineCustInfo /
 *               createDateRangeRow / createProductRow / createTotalRow / createFooter
 *
 * 所有「POI 單位」均透過 `infra/excel-service` 的 `poiFontSize` / `setRowHeightPoi` 等換算。
 */

import ExcelJS from 'exceljs';

import {
  ARGB_RED,
  KAI_FONT,
  buildStyle,
  mergeRange,
  poiFontSize,
  setRowHeightPoi,
  writeMixedRow,
  type StyledCell,
} from '@/infra/excel-service';

import type { CustomerModel } from '@/domain/models/customer-model';
import type { Product } from '@/domain/models/product';
import { getItemIndex } from '@/domain/sorting-list';

/* ============================================================
   桌面版 Writer.companion 常數（POI 單位）
   ============================================================ */
export const TITLE_SIZE_POI = 400;
export const FONT_SIZE_POI = 300;
export const CUSTINFO_FONT_POI = FONT_SIZE_POI + 40;
export const ROW_HEIGHT_POI = 500;

export const FIRST_LINE_WIDTH_POI = 15 * 256;
export const CENTER_LINE_WIDTH_POI = 6 * 256;
export const RIGHT_LINE_WIDTH_POI = 10 * 256;

/* ============================================================
   客戶分組（對應 SheetExtension 的 list extension）
   ============================================================ */

/** 月結客戶（且非現金結） */
export const getMonthlyCustomer = (list: ReadonlyArray<CustomerModel>): CustomerModel[] =>
  list.filter((c) => c.isMonthly && !c.isCashUser);

/** 半月結客戶（非月結且非現金） */
export const getHalfMonthlyCustomer = (list: ReadonlyArray<CustomerModel>): CustomerModel[] =>
  list.filter((c) => !c.isMonthly && !c.isCashUser);

/** 現金結客戶 */
export const getCashCustomer = (list: ReadonlyArray<CustomerModel>): CustomerModel[] =>
  list.filter((c) => c.isCashUser);

/* ============================================================
   樣式建構器（取代桌面版每個 helper 內各自 createCellStyle）
   ============================================================ */

const styles = {
  /** 標題（標楷體、粗體、字級 20pt） */
  title: () =>
    buildStyle({
      font: { name: KAI_FONT, bold: true, size: poiFontSize(TITLE_SIZE_POI) },
      align: 'center',
    }),
  /** 客戶資訊欄（標楷體、粗體、字級 17pt） */
  custInfo: () =>
    buildStyle({
      font: { name: KAI_FONT, bold: true, size: poiFontSize(CUSTINFO_FONT_POI) },
      align: 'center',
    }),
  /** 一般置中文字（無框、無粗體、字級 15pt） */
  centerText: () =>
    buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI) },
      align: 'center',
    }),
  /** 商品列基本格式（含框） */
  cellWithBorder: () =>
    buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI) },
      border: true,
      align: 'center',
    }),
  /** 訂貨日期粗體 */
  dateBold: () =>
    buildStyle({
      font: { name: KAI_FONT, bold: true, size: poiFontSize(FONT_SIZE_POI) },
      border: true,
      align: 'center',
    }),
  /** 紅字總計列（含框） */
  redTotal: () =>
    buildStyle({
      font: { name: KAI_FONT, size: poiFontSize(FONT_SIZE_POI), color: ARGB_RED },
      border: true,
      align: 'center',
    }),
};

/* ============================================================
   區塊建構：對應 SheetExtension.* helpers
   ============================================================ */

/**
 * 標題區（青坊食品行 / 請款單），整列合併到 `maxCols + 1`。
 * 對應 `Sheet.createHeader(rowMaxSize)`.
 */
export function createHeader(sheet: ExcelJS.Worksheet, rowMaxSize: number): void {
  const headerStyle = styles.title();
  const r1 = sheet.addRow(['青坊食品行']);
  r1.getCell(1).style = headerStyle;
  mergeRange(sheet, r1.number, 1, r1.number, rowMaxSize + 1);

  const r2 = sheet.addRow(['請款單']);
  r2.getCell(1).style = headerStyle;
  mergeRange(sheet, r2.number, 1, r2.number, rowMaxSize + 1);
}

/**
 * 客戶資訊：兩列（客戶編號 / 客戶名稱 + 帳單年月）。
 * 對應 `Sheet.createCustInfo(...)`.
 */
export function createCustInfo(
  sheet: ExcelJS.Worksheet,
  customer: CustomerModel,
  billYear: string,
  billMonth: string,
  rowMaxSize: number
): void {
  const custFont = styles.custInfo();

  // 客戶編號列：標籤 + 值（值合併欄 2..5）
  const codeRow = sheet.addRow(['客戶編號', customer.code]);
  codeRow.getCell(1).style = custFont;
  codeRow.getCell(2).style = custFont;
  mergeRange(sheet, codeRow.number, 2, codeRow.number, 5);

  // 客戶名稱列：repeat(rowMaxSize - 6) 空白 + 帳單年月
  const nameValues: Array<string | number> = ['客戶名稱', customer.name];
  const padCount = rowMaxSize - 6;
  for (let i = 0; i < padCount; i++) nameValues.push('');
  nameValues.push(billYear, '年', billMonth, '月');

  const nameRow = sheet.addRow(nameValues);
  setRowHeightPoi(nameRow, ROW_HEIGHT_POI);
  nameRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.style = custFont;
  });
  mergeRange(sheet, nameRow.number, 2, nameRow.number, 5);
}

/**
 * 單行客戶資訊（代碼 + 名稱合併顯示），用於明細/代送費分頁。
 * 對應 `Sheet.createSingleLineCustInfo(...)`.
 */
export function createSingleLineCustInfo(
  sheet: ExcelJS.Worksheet,
  customer: CustomerModel,
  billYear: string,
  billMonth: string,
  rowMaxSize: number
): void {
  const custFont = styles.custInfo();

  const values: Array<string | number> = ['客戶名稱', `${customer.code} ${customer.name}`];
  const padCount = rowMaxSize - 5;
  for (let i = 0; i < padCount; i++) values.push('');
  values.push(billYear, '年', billMonth, '月');

  const r = sheet.addRow(values);
  setRowHeightPoi(r, ROW_HEIGHT_POI);
  r.eachCell({ includeEmpty: true }, (cell) => {
    cell.style = custFont;
  });
  mergeRange(sheet, r.number, 2, r.number, 7);
}

/**
 * 訂貨日期表頭：訂貨日期 / 1 2 3 ... / 數量 / 單價 / 合計。
 * 對應 `Sheet.createDateRangeRow(dateRange, maxDate)`.
 */
export function createDateRangeRow(
  sheet: ExcelJS.Worksheet,
  dateRange: ReadonlyArray<number>,
  maxDate = 0
): void {
  const base = styles.cellWithBorder();
  const dateBold = styles.dateBold();

  const cells: StyledCell[] = [['訂貨日期', base]];
  for (const day of dateRange) cells.push([day, dateBold]);

  if (dateRange.length < maxDate) {
    const pad = maxDate - dateRange.length;
    for (let i = 0; i < pad; i++) cells.push(['', base]);
  }

  cells.push(['數量', base], ['單價', base], ['合計', base]);
  const row = writeMixedRow(sheet, cells, base);
  setRowHeightPoi(row, ROW_HEIGHT_POI);
}

/**
 * 商品資料列：依 SortingList 排序後逐列輸出。
 * 對應 `Sheet.createProductRow(productList, dateRange, maxDate)`.
 */
export function createProductRow(
  sheet: ExcelJS.Worksheet,
  productList: ReadonlyArray<Product>,
  dateRange: ReadonlyArray<number>,
  maxDate = 0
): void {
  const cellStyle = styles.cellWithBorder();

  // P2.13：商品依 SortingList 排序
  const sorted = [...productList].sort((a, b) => getItemIndex(a.name) - getItemIndex(b.name));

  for (const product of sorted) {
    const cells: StyledCell[] = [[product.name, cellStyle]];

    for (const day of dateRange) {
      const order = product.orderList.find((o) => o.day === day);
      cells.push([order ? order.count : '', cellStyle]);
    }

    if (dateRange.length < maxDate) {
      const pad = maxDate - dateRange.length;
      for (let i = 0; i < pad; i++) cells.push(['', cellStyle]);
    }

    cells.push(
      [product.getDateRangeCount(dateRange), cellStyle],
      [product.price, cellStyle],
      [product.getDateRangePrice(dateRange), cellStyle]
    );

    const row = writeMixedRow(sheet, cells, cellStyle);
    setRowHeightPoi(row, ROW_HEIGHT_POI);
  }
}

/**
 * 總計列（紅字）。若客戶含稅，再多兩列：稅金 / 含稅小計。
 * 對應 `Sheet.createTotalRow(customer, rowMaxSize)`.
 */
export function createTotalRow(
  sheet: ExcelJS.Worksheet,
  customer: CustomerModel,
  rowMaxSize: number
): void {
  const redStyle = styles.redTotal();

  // 總計列：repeat(rowMaxSize - 1) 空白 + 「總計」+ 金額
  const totalCells: StyledCell[] = [];
  for (let i = 0; i < rowMaxSize - 1; i++) totalCells.push(['', redStyle]);
  totalCells.push(['總計', redStyle], [customer.getTotalPrice(), redStyle]);
  const totalRow = writeMixedRow(sheet, totalCells, redStyle);
  setRowHeightPoi(totalRow, ROW_HEIGHT_POI);

  // 含稅客戶：稅金行 + 含稅小計行（P2.14）
  if (customer.isNeedTex) {
    const taxCells: StyledCell[] = [];
    for (let i = 0; i < rowMaxSize - 1; i++) taxCells.push(['', redStyle]);
    taxCells.push(['稅金', redStyle], [customer.getTex(), redStyle]);
    const taxRow = writeMixedRow(sheet, taxCells, redStyle);
    setRowHeightPoi(taxRow, ROW_HEIGHT_POI);

    const sumCells: StyledCell[] = [];
    for (let i = 0; i < rowMaxSize; i++) sumCells.push(['', redStyle]);
    sumCells.push([customer.getAfterTexSum(), redStyle]);
    const sumRow = writeMixedRow(sheet, sumCells, redStyle);
    setRowHeightPoi(sumRow, ROW_HEIGHT_POI);
  }
}

/**
 * 帳單頁尾：訂貨專線 / 銀行資訊 / 提醒。
 * 對應 `Sheet.createFooter(rowMaxSize)`.
 */
export function createFooter(sheet: ExcelJS.Worksheet, rowMaxSize: number): void {
  const base = styles.centerText();

  // 對應 `row { }` 空白列
  sheet.addRow([]);

  const lines = [
    '訂貨專線:(04)7359885   7359886    ',
    '銀行代號:\t822 中國信託\t\t\t戶名：青坊食品行\t\t\t\t\t帳號：554540468218\t\t\t\t\t\t分行別：六家庄分行',
    '※ 匯款後請告知戶名或末5碼 ※',
  ];

  for (const text of lines) {
    const r = sheet.addRow([text]);
    r.getCell(1).style = base;
    mergeRange(sheet, r.number, 1, r.number, rowMaxSize + 1);
  }
}
