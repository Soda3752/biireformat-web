import ExcelJS from 'exceljs';

/**
 * ExcelJS Wrapper — 對應桌面版 SheetExtension.kt 的樣式 helper。
 * 同時包含 POI 單位（1/20 pt、1/256 char）與 ExcelJS 單位（pt、char）的換算。
 */

export const KAI_FONT = '楷體-簡';

/** POI 1/20 pt → ExcelJS pt */
export const poiFontSize = (poi: number): number => poi / 20;

/** POI 1/256 char → ExcelJS char */
export const poiColumnWidth = (poi: number): number => poi / 256;

/** POI 1/20 pt → ExcelJS pt（row height 同字級單位） */
export const poiRowHeight = (poi: number): number => poi / 20;

/** Color helper：對應 POI Font.COLOR_RED */
export const ARGB_RED = 'FFFF0000';
export const ARGB_BLACK = 'FF000000';

/** Paper sizes（ExcelJS 直接接受 POI 同義代碼） */
export const PAPER_A5 = 11;
export const PAPER_A4 = 9;

/* ========================================================================
   樣式建構器
   ======================================================================== */

export interface BorderOptions {
  style?: ExcelJS.BorderStyle;
}

export const thinBorder = (style: ExcelJS.BorderStyle = 'thin'): Partial<ExcelJS.Borders> => ({
  top: { style },
  bottom: { style },
  left: { style },
  right: { style },
});

export const centerAlign: Partial<ExcelJS.Alignment> = {
  horizontal: 'center',
  vertical: 'middle',
};

export interface FontConfig {
  size?: number;
  bold?: boolean;
  italic?: boolean;
  name?: string;
  color?: string;
}

export const buildFont = (cfg: FontConfig): Partial<ExcelJS.Font> => {
  const font: Partial<ExcelJS.Font> = {};
  if (cfg.size !== undefined) font.size = cfg.size;
  if (cfg.bold !== undefined) font.bold = cfg.bold;
  if (cfg.italic !== undefined) font.italic = cfg.italic;
  if (cfg.name !== undefined) font.name = cfg.name;
  if (cfg.color !== undefined) font.color = { argb: cfg.color };
  return font;
};

export interface CellStyleConfig {
  font?: FontConfig;
  border?: boolean | ExcelJS.BorderStyle;
  align?: 'center' | 'left' | 'right' | Partial<ExcelJS.Alignment>;
  fill?: string;
  numFmt?: string;
}

/** 建立 cell style（產出 Partial<Style>，可直接賦給 cell.style 或 row.style） */
export const buildStyle = (cfg: CellStyleConfig): Partial<ExcelJS.Style> => {
  const style: Partial<ExcelJS.Style> = {};

  if (cfg.font) style.font = buildFont(cfg.font);

  if (cfg.border) {
    const borderStyle = typeof cfg.border === 'string' ? cfg.border : 'thin';
    style.border = thinBorder(borderStyle);
  }

  if (cfg.align) {
    if (cfg.align === 'center') {
      style.alignment = centerAlign;
    } else if (cfg.align === 'left') {
      style.alignment = { horizontal: 'left', vertical: 'middle' };
    } else if (cfg.align === 'right') {
      style.alignment = { horizontal: 'right', vertical: 'middle' };
    } else {
      style.alignment = cfg.align;
    }
  }

  if (cfg.fill) {
    style.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: cfg.fill },
    };
  }

  if (cfg.numFmt) style.numFmt = cfg.numFmt;

  return style;
};

/* ========================================================================
   Sheet 操作 helper
   ======================================================================== */

export interface PrintMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

export interface PrintSettingOptions {
  paperSize?: number;
  landscape?: boolean;
  fitToPage?: boolean;
  fitWidth?: number;
  fitHeight?: number;
  scale?: number;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
  margins?: PrintMargins;
}

/** 對應 SheetExtension.setUpPrintSetting()：A5 橫向、fitToPage */
export const setupPrintSetting = (
  sheet: ExcelJS.Worksheet,
  options: PrintSettingOptions = {}
): void => {
  const useScale = typeof options.scale === 'number';
  const pageSetup: Partial<ExcelJS.PageSetup> = {
    ...sheet.pageSetup,
    paperSize: options.paperSize ?? PAPER_A5,
    orientation: options.landscape === false ? 'portrait' : 'landscape',
    fitToPage: useScale ? false : (options.fitToPage ?? true),
    fitToWidth: options.fitWidth ?? 1,
    fitToHeight: options.fitHeight ?? 0,
  };
  if (useScale) pageSetup.scale = options.scale;
  if (options.horizontalCentered !== undefined) pageSetup.horizontalCentered = options.horizontalCentered;
  if (options.verticalCentered !== undefined) pageSetup.verticalCentered = options.verticalCentered;
  if (options.margins) {
    pageSetup.margins = {
      left: options.margins.left ?? 0,
      right: options.margins.right ?? 0,
      top: options.margins.top ?? 0,
      bottom: options.margins.bottom ?? 0,
      header: options.margins.header ?? 0,
      footer: options.margins.footer ?? 0,
    };
  }
  sheet.pageSetup = pageSetup;
};

/**
 * 合併儲存格（對應 POI addMergedRegion）。
 * 注意：POI 是 0-indexed，ExcelJS mergeCells 是 1-indexed，這個 helper 讓使用者用 ExcelJS 1-indexed。
 */
export const mergeRange = (
  sheet: ExcelJS.Worksheet,
  startRow: number,
  startCol: number,
  endRow: number,
  endCol: number
): void => {
  sheet.mergeCells(startRow, startCol, endRow, endCol);
};

/** 整列合併（對應 mergeFullRow）——傳入 1-indexed 列號與欄位數 */
export const mergeFullRow = (
  sheet: ExcelJS.Worksheet,
  rowIndex: number,
  totalColumns: number,
  startCol = 1
): void => {
  sheet.mergeCells(rowIndex, startCol, rowIndex, totalColumns);
};

/** 設定欄寬（POI 1/256 char 換算） */
export const setColumnWidthPoi = (
  sheet: ExcelJS.Worksheet,
  columnIndex: number,
  poiWidth: number
): void => {
  sheet.getColumn(columnIndex).width = poiColumnWidth(poiWidth);
};

/** 設定列高（POI 1/20 pt 換算） */
export const setRowHeightPoi = (row: ExcelJS.Row, poiHeight: number): void => {
  row.height = poiRowHeight(poiHeight);
};

/** 在指定列前加入分頁符號（對應 setRowBreak） */
export const addPageBreakAfter = (sheet: ExcelJS.Worksheet, rowIndex: number): void => {
  const row = sheet.getRow(rowIndex);
  row.addPageBreak();
};

/* ========================================================================
   Row 寫入便利方法
   ======================================================================== */

/**
 * 在 worksheet 寫一列資料，並對所有 cell 套用相同樣式。
 * 對應 excelkt 的 row { ... } DSL。
 */
export const writeStyledRow = (
  sheet: ExcelJS.Worksheet,
  values: ReadonlyArray<string | number | null | undefined>,
  style?: Partial<ExcelJS.Style>
): ExcelJS.Row => {
  const row = sheet.addRow(values as Array<string | number | null | undefined>);
  if (style) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.style = { ...cell.style, ...style };
    });
  }
  return row;
};

/**
 * 在 worksheet 寫一列資料，每個 cell 各自帶不同樣式（傳入 [value, style] tuple）。
 * 沒指定樣式的 cell 會用 fallbackStyle。
 */
export type StyledCell = [
  value: string | number | null | undefined,
  style?: Partial<ExcelJS.Style>,
];

export const writeMixedRow = (
  sheet: ExcelJS.Worksheet,
  cells: ReadonlyArray<StyledCell>,
  fallbackStyle?: Partial<ExcelJS.Style>
): ExcelJS.Row => {
  const values = cells.map(([v]) => v);
  const row = sheet.addRow(values as Array<string | number | null | undefined>);
  cells.forEach(([, style], idx) => {
    const cell = row.getCell(idx + 1);
    const merged = { ...(fallbackStyle ?? {}), ...(style ?? {}) };
    cell.style = merged;
  });
  return row;
};

/**
 * 加一個空白列（對應 row {}），便於分隔區塊
 */
export const addBlankRow = (sheet: ExcelJS.Worksheet): ExcelJS.Row => sheet.addRow([]);

/* ========================================================================
   Workbook helper
   ======================================================================== */

/** 建立新 workbook，預設帶 BOM 與 UTF-8 設定 */
export const createWorkbook = (): ExcelJS.Workbook => {
  const wb = new ExcelJS.Workbook();
  wb.creator = '青坊食品行 帳單處理工具';
  wb.created = new Date();
  return wb;
};

/** 將 workbook 輸出為 Blob，用於前端下載 */
export const workbookToBlob = async (wb: ExcelJS.Workbook): Promise<Blob> => {
  const buffer = await wb.xlsx.writeBuffer();
  return new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
};
