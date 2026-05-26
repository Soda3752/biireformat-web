/**
 * 「生成手填本」.xlsx 輸出。
 *
 * 對應範本 `1.彰化.xls` 的版面：
 *   - 兩個分頁：上（日 1~15）、下（日 16~31）
 *   - 標楷體、合併、邊框、列高 17.25pt
 *   - A4 橫向列印
 *   - 客戶區塊 = max(8, 品名數) 列
 *   - 列印演算法：累積列數超過頁面上限時插入 page break，禁止客戶跨頁
 */

import ExcelJS from 'exceljs';

import {
    addPageBreakAfter,
    ARGB_BLACK,
    buildStyle,
    createWorkbook,
    KAI_FONT,
    mergeRange,
    PAPER_A4,
} from '@/infra/excel-service';
import {
    HANDFILL_MANIFEST_SHEET,
    type HandfillBook,
    type HandfillCustomer,
    lineFullName,
} from '@/domain/models/handfill-book';

/* ====================== 樣式常數 ======================= */
const TITLE_FONT_PT = 20;       // 範本：400 POI = 20pt
const CONTENT_FONT_PT = 12;     // 範本：240 POI = 12pt
const ROW_HEIGHT_PT = 15.8;     // 配合縮減邊距，讓 36 列填滿 A4 橫向可印區域

// 欄寬：A4 橫向、左右邊界 0.25" 時可印寬約 806pt。
// 在原本（左右 0.05"）的欄寬基礎上整體放大 ~8%，讓表格右側貼齊可印區。
const COL_WIDTHS_UPPER = [21, 15.5, 5, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 6.4, 8.3, 10.6];
const COL_WIDTHS_LOWER = [21, 15.5, 5, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 5.95, 8.3, 10.6];

const HEADER_ROW_COUNT = 2;             // 每頁欄位標頭 2 列（「客戶名稱/品名/單/...」+ 日期列）
const TITLE_ROW_COUNT_FIRST = 2;        // 第一頁多 2 列（線別 + 年月）
const TOP_BLANK_ROW_COUNT_FIRST = 0;    // 第一頁頂端不留空白，讓 4 客戶（32 列）能塞進 36 列預算
const TOP_BLANK_ROW_COUNT = 2;          // 第 2 頁起頂端預留 2 列空白
// A4 橫向可印列數（14.5pt × 36 = 522pt < 538pt 可印高）
const A4_LANDSCAPE_USABLE_ROWS = 36;

const STD_BLOCK_ROWS = 8;

// 分頁規劃：預設 4 客戶/頁；超過 budget 時自動退到 3 → 2 → 1 並均分剩餘列
const TARGET_CUSTOMERS_PER_PAGE = 4;

interface SheetSpec {
    name: '上' | '下';
    dayStart: number;   // 1 or 16
    dayCount: number;   // 15 or 16
    totalCols: number;  // 20 or 21
    qtyCol: number;     // 1-indexed
    totalCol: number;   // 1-indexed
    colWidths: number[];
}

const SHEET_UPPER: SheetSpec = {
    name: '上',
    dayStart: 1,
    dayCount: 15,
    totalCols: 20,
    qtyCol: 19,
    totalCol: 20,
    colWidths: COL_WIDTHS_UPPER,
};

const SHEET_LOWER: SheetSpec = {
    name: '下',
    dayStart: 16,
    dayCount: 16,
    totalCols: 21,
    qtyCol: 20,
    totalCol: 21,
    colWidths: COL_WIDTHS_LOWER,
};

/* ====================== 公開 API ======================= */

export async function buildHandfillWorkbook(book: HandfillBook): Promise<ExcelJS.Workbook> {
    const wb = createWorkbook();
    writeSheet(wb, book, SHEET_UPPER);
    writeSheet(wb, book, SHEET_LOWER);
    writeManifest(wb, book);
    return wb;
}

/**
 * 寫入隱藏 metadata 分頁，A1 存整本 book 的 JSON。
 * 設為 veryHidden 讓使用者無法透過 Excel「取消隱藏」選單看到，避免誤改造成資料/manifest 不同步。
 * Reader 讀檔時優先讀此分頁，繞過版面分析的不確定性。
 */
function writeManifest(wb: ExcelJS.Workbook, book: HandfillBook): void {
    const sheet = wb.addWorksheet(HANDFILL_MANIFEST_SHEET, {
        state: 'veryHidden',
    });
    sheet.getCell(1, 1).value = JSON.stringify(book);
}

/* ====================== Sheet 寫入 ======================= */

function writeSheet(wb: ExcelJS.Workbook, book: HandfillBook, spec: SheetSpec): void {
    const sheet = wb.addWorksheet(spec.name);

    // 預設列高（避免某些未顯式設定的列回到 Excel 預設高度）
    sheet.properties.defaultRowHeight = ROW_HEIGHT_PT;

    // 1. 欄寬
    spec.colWidths.forEach((w, i) => {
        sheet.getColumn(i + 1).width = w;
    });

    // 2. 列印設定（A4 橫向）
    // 不使用 fitToPage：欄寬加總 ≈ 758pt 已小於可印寬，自然能單頁顯示。
    // 啟用 fitToPage 反而會做隱式縮放，連帶把列高一起縮小，造成下方留白。
    // 上下邊界 0.15" 是讓 36 列塞進 A4 橫向可印高的必要值；
    // 左右改 0.25"（≈ 6.35mm）避開多數印表機的硬體不可印區。
    sheet.pageSetup = {
        paperSize: PAPER_A4,
        orientation: 'landscape',
        scale: 105,
        horizontalCentered: true,
        verticalCentered: true,
        margins: {
            left: 0.0, right: 0.0,
            top: 0.15, bottom: 0.0,
            header: 0.05, footer: 0.0,
        },
    };

    // 3. 寫入：頂端空白 + 標題 + 欄頭 + 客戶區塊（含 page break）
    let currentRow = 1;
    currentRow += TOP_BLANK_ROW_COUNT_FIRST;
    currentRow = writeTitle(sheet, book, spec, currentRow);
    currentRow = writeColumnHeader(sheet, spec, currentRow);

    const PAGE_1_BUDGET = A4_LANDSCAPE_USABLE_ROWS - TOP_BLANK_ROW_COUNT_FIRST - TITLE_ROW_COUNT_FIRST - HEADER_ROW_COUNT;
    const PAGE_OTHER_BUDGET = A4_LANDSCAPE_USABLE_ROWS - TOP_BLANK_ROW_COUNT - HEADER_ROW_COUNT;

    let customerIdx = 0;
    let pageIdx = 0;
    while (customerIdx < book.customers.length) {
        const budget = pageIdx === 0 ? PAGE_1_BUDGET : PAGE_OTHER_BUDGET;
        const remaining = book.customers.slice(customerIdx);
        const blockSizes = planPage(remaining, budget);

        for (let i = 0; i < blockSizes.length; i++) {
            currentRow = writeCustomerBlock(sheet, spec, currentRow, remaining[i], blockSizes[i]);
        }
        customerIdx += blockSizes.length;

        // 最後一頁若客戶不足 TARGET_CUSTOMERS_PER_PAGE，補上同樣格式的空白區塊，
        // 維持 4 個 8 列格子的版面（在剩餘 budget 範圍內）。
        if (customerIdx >= book.customers.length && blockSizes.length < TARGET_CUSTOMERS_PER_PAGE) {
            const usedRows = blockSizes.reduce((a, b) => a + b, 0);
            const slack = budget - usedRows;
            const emptyCount = Math.min(
                TARGET_CUSTOMERS_PER_PAGE - blockSizes.length,
                Math.floor(slack / STD_BLOCK_ROWS),
            );
            for (let i = 0; i < emptyCount; i++) {
                currentRow = writeEmptyBlock(sheet, spec, currentRow, STD_BLOCK_ROWS);
            }
        }

        if (customerIdx < book.customers.length) {
            addPageBreakAfter(sheet, currentRow - 1);
            currentRow += TOP_BLANK_ROW_COUNT;
            currentRow = writeColumnHeader(sheet, spec, currentRow);
        }
        pageIdx++;
    }

    // 4. 列高（使用自行追蹤的 lastRow，避免 actualRowCount 未涵蓋只有空白 cell 的列）
    const lastRow = currentRow - 1;
    for (let r = 1; r <= lastRow; r++) {
        sheet.getRow(r).height = ROW_HEIGHT_PT;
    }

    // 5. 凍結窗格：凍結頂端空白 + 標題 + 欄頭
    sheet.views = [{
        state: 'frozen',
        xSplit: 0,
        ySplit: TOP_BLANK_ROW_COUNT_FIRST + TITLE_ROW_COUNT_FIRST + HEADER_ROW_COUNT
    }];
}

/* ====================== 區塊 helper ======================= */

function writeTitle(
    sheet: ExcelJS.Worksheet,
    book: HandfillBook,
    spec: SheetSpec,
    startRow: number
): number {
    const titleStyle = buildStyle({
        font: {name: KAI_FONT, size: TITLE_FONT_PT, bold: true},
        align: 'center',
    });

    // (一)彰化 → col 1~2 合併
    setCell(sheet, startRow, 1, lineFullName(book), titleStyle);
    mergeRange(sheet, startRow, 1, startRow + 1, 2);

    // 「年」字 → 跟 col 15:17 合併（範例為 col 16 顯示年數）
    // 為了結構清晰，這裡分散在不同欄
    const halfSuffix = spec.name === '上' ? '月上' : '月下';

    if (spec.name === '上') {
        // 上 sheet: 年份 col 15:17 合併, 「年」col 17:18, 月份 col 18:19, 「月上」col 19:20
        setCell(sheet, startRow, 16, book.year, titleStyle);
        mergeRange(sheet, startRow, 16, startRow + 1, 17);
        setCell(sheet, startRow, 18, '年', titleStyle);
        mergeRange(sheet, startRow, 18, startRow + 1, 18);
        setCell(sheet, startRow, 19, book.month, titleStyle);
        mergeRange(sheet, startRow, 19, startRow + 1, 19);
        setCell(sheet, startRow, 20, halfSuffix, titleStyle);
        mergeRange(sheet, startRow, 20, startRow + 1, 20);
    } else {
        // 下 sheet: 對齊上 sheet，「月下」落在總計欄 (col 21)，整體向右一欄
        // 年份合併 col 17:18（單欄寬度 5.5 不足以容納 20pt 三位數年份，會顯示 ##）
        setCell(sheet, startRow, 17, book.year, titleStyle);
        mergeRange(sheet, startRow, 17, startRow + 1, 18);
        setCell(sheet, startRow, 19, '年', titleStyle);
        mergeRange(sheet, startRow, 19, startRow + 1, 19);
        setCell(sheet, startRow, 20, book.month, titleStyle);
        mergeRange(sheet, startRow, 20, startRow + 1, 20);
        setCell(sheet, startRow, 21, halfSuffix, titleStyle);
        mergeRange(sheet, startRow, 21, startRow + 1, 21);
    }

    return startRow + 2;
}

function writeColumnHeader(sheet: ExcelJS.Worksheet, spec: SheetSpec, startRow: number): number {
    const headerStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT, bold: true},
        align: 'center',
        border: 'thin',
    });

    const r1 = startRow;
    const r2 = startRow + 1;

    // 客戶名稱 (col 1, 直向合併)
    setCell(sheet, r1, 1, '客戶名稱', headerStyle);
    setCell(sheet, r2, 1, '', headerStyle);
    mergeRange(sheet, r1, 1, r2, 1);

    // 品名 (col 2, 直向合併)
    setCell(sheet, r1, 2, '品名', headerStyle);
    setCell(sheet, r2, 2, '', headerStyle);
    mergeRange(sheet, r1, 2, r2, 2);

    // 單價 (col 3, 上下各一字；移除兩格中間的水平框線)
    const unitTopStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT, bold: true},
        align: 'center',
    });
    unitTopStyle.border = {
        top: {style: 'thin', color: {argb: ARGB_BLACK}},
        left: {style: 'thin', color: {argb: ARGB_BLACK}},
        right: {style: 'thin', color: {argb: ARGB_BLACK}},
    };
    const unitBottomStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT, bold: true},
        align: 'center',
    });
    unitBottomStyle.border = {
        bottom: {style: 'thin', color: {argb: ARGB_BLACK}},
        left: {style: 'thin', color: {argb: ARGB_BLACK}},
        right: {style: 'thin', color: {argb: ARGB_BLACK}},
    };
    setCell(sheet, r1, 3, '單', unitTopStyle);
    setCell(sheet, r2, 3, '價', unitBottomStyle);

    // 訂購單日數量 標籤 (col 4~end-2，橫向合併 r1)
    setCell(sheet, r1, 4, '訂購單日數量', headerStyle);
    mergeRange(sheet, r1, 4, r1, 3 + spec.dayCount);

    // 日期列 (r2, col 4 起 dayCount 個)
    for (let d = 0; d < spec.dayCount; d++) {
        setCell(sheet, r2, 4 + d, spec.dayStart + d, headerStyle);
    }

    // 數量 (qtyCol, 直向合併)
    setCell(sheet, r1, spec.qtyCol, '數量', headerStyle);
    setCell(sheet, r2, spec.qtyCol, '', headerStyle);
    mergeRange(sheet, r1, spec.qtyCol, r2, spec.qtyCol);

    // 總計 (totalCol, 直向合併)
    setCell(sheet, r1, spec.totalCol, '總計', headerStyle);
    setCell(sheet, r2, spec.totalCol, '', headerStyle);
    mergeRange(sheet, r1, spec.totalCol, r2, spec.totalCol);

    return startRow + 2;
}

function writeCustomerBlock(
    sheet: ExcelJS.Worksheet,
    spec: SheetSpec,
    startRow: number,
    cust: HandfillCustomer,
    blockRows: number
): number {
    // 樣式
    const col0Style = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'left',
        // col 0 只有左右框（與下一列連通），上下無
    });
    // ↑ 範本：客戶代號 / 名稱 / 電話 cells 只有左右框
    col0Style.border = {
        left: {style: 'thin', color: {argb: ARGB_BLACK}},
        right: {style: 'thin', color: {argb: ARGB_BLACK}},
    };

    const productNameStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'left',
        border: 'thin',
    });

    const unitPriceStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'center',
        border: 'thin',
    });

    const dayStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'center',
        border: 'thin',
    });

    // 總計欄只有左右框，內部不畫水平格線
    const totalStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'center',
    });
    totalStyle.border = {
        left: {style: 'thin', color: {argb: ARGB_BLACK}},
        right: {style: 'thin', color: {argb: ARGB_BLACK}},
    };

    // === col 0 排版 ===
    // 標準: row+0=ID, row+1=名稱, row+2=空, row+3~5=休息日, row+6=電話1, row+7=電話2
    // 品名 > 8 時: 名稱頂、電話底，中間放休息日 + 留白
    const col0Layout = layoutCol0(cust, blockRows);
    for (let i = 0; i < blockRows; i++) {
        setCell(sheet, startRow + i, 1, col0Layout[i] ?? '', col0Style);
    }

    // === col 1 (品名) + col 2 (單價) ===
    for (let i = 0; i < blockRows; i++) {
        const prod = cust.products[i];
        setCell(sheet, startRow + i, 2, prod?.name ?? '', productNameStyle);
        setCell(sheet, startRow + i, 3, prod?.unitPrice ?? '', unitPriceStyle);
    }

    // === col 4 ~ 日數欄 + 數量（空白 + 邊框）===
    for (let i = 0; i < blockRows; i++) {
        for (let c = 4; c < spec.totalCol; c++) {
            setCell(sheet, startRow + i, c, '', dayStyle);
        }
        // 總計欄：只有左右框
        setCell(sheet, startRow + i, spec.totalCol, '', totalStyle);
    }

    // === 區塊底部粗框（店家分隔線）===
    const lastRow = startRow + blockRows - 1;
    for (let c = 1; c <= spec.totalCol; c++) {
        const cell = sheet.getCell(lastRow, c);
        cell.style = {
            ...cell.style,
            border: {
                ...(cell.style.border ?? {}),
                bottom: {style: 'medium', color: {argb: ARGB_BLACK}},
            },
        };
    }

    return startRow + blockRows;
}

/**
 * 計算 col 0 在客戶區塊內的逐列內容。
 *
 * 標準佈局 (blockRows >= 8)：
 *   row+0   ID
 *   row+1   名稱
 *   row+2   (空)
 *   row+3~5 休息日 1~3（若有）
 *   row+(N-2) 電話 1
 *   row+(N-1) 電話 2
 *
 * 休息日多於 3 個時，會延伸佔用中間更多列；
 * 電話多於 2 個時，會擠到電話 1 之前；
 * 整體保證代號頂、電話底、中間放休息日+留白。
 */
function layoutCol0(cust: HandfillCustomer, blockRows: number): string[] {
    const out: string[] = new Array(blockRows).fill('');
    out[0] = cust.customerId;
    if (blockRows >= 2) out[1] = cust.customerName;

    // 電話放在最後 phones.length 列（從底向上）
    const phonesNonEmpty = cust.phones.filter((p) => p.trim());
    const phoneCount = phonesNonEmpty.length;
    for (let i = 0; i < phoneCount; i++) {
        const targetRow = blockRows - phoneCount + i;
        if (targetRow >= 2) out[targetRow] = phonesNonEmpty[i];
    }

    // 休息日從 row+3 起放，最多到電話之前
    const restNonEmpty = cust.restNotes.filter((r) => r.trim());
    const restStartRow = 3;
    const restEndRowExclusive = Math.max(restStartRow, blockRows - phoneCount);
    for (let i = 0; i < restNonEmpty.length && i < restEndRowExclusive - restStartRow; i++) {
        out[restStartRow + i] = restNonEmpty[i];
    }
    // 若超出可放範圍，剩餘休息日捨棄（理論上不會發生：blockRows >= 8 時可容納大量備註）

    return out;
}

/**
 * 寫入一個沒有客戶資料的空白區塊，使用與 writeCustomerBlock 相同的格線與底部分隔線。
 * 用於最後一頁不足 4 個客戶時撐滿版面。
 */
function writeEmptyBlock(
    sheet: ExcelJS.Worksheet,
    spec: SheetSpec,
    startRow: number,
    blockRows: number
): number {
    const col0Style = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'left',
    });
    col0Style.border = {
        left: {style: 'thin', color: {argb: ARGB_BLACK}},
        right: {style: 'thin', color: {argb: ARGB_BLACK}},
    };

    const cellStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'center',
        border: 'thin',
    });

    const totalStyle = buildStyle({
        font: {name: KAI_FONT, size: CONTENT_FONT_PT},
        align: 'center',
    });
    totalStyle.border = {
        left: {style: 'thin', color: {argb: ARGB_BLACK}},
        right: {style: 'thin', color: {argb: ARGB_BLACK}},
    };

    for (let i = 0; i < blockRows; i++) {
        setCell(sheet, startRow + i, 1, '', col0Style);
        for (let c = 2; c < spec.totalCol; c++) {
            setCell(sheet, startRow + i, c, '', cellStyle);
        }
        setCell(sheet, startRow + i, spec.totalCol, '', totalStyle);
    }

    // 底部粗框
    const lastRow = startRow + blockRows - 1;
    for (let c = 1; c <= spec.totalCol; c++) {
        const cell = sheet.getCell(lastRow, c);
        cell.style = {
            ...cell.style,
            border: {
                ...(cell.style.border ?? {}),
                bottom: {style: 'medium', color: {argb: ARGB_BLACK}},
            },
        };
    }

    return startRow + blockRows;
}

function setCell(
    sheet: ExcelJS.Worksheet,
    row: number,
    col: number,
    value: string | number,
    style?: Partial<ExcelJS.Style>
): void {
    const cell = sheet.getCell(row, col);
    cell.value = value === '' ? null : value;
    if (style) {
        cell.style = {...cell.style, ...style};
    }
}

/* ====================== 分頁規劃 ======================= */

/**
 * 規劃單頁要放幾個客戶、各自佔多少列。
 *
 * 策略：
 *   1. 從目標 4 個客戶開始嘗試，計算自然列數總和 (max(8, 品名數))。
 *   2. 若 4 個塞不下 → 退到 3 → 2 → 1，直到放得下為止。
 *   3. 放得下後，把 budget 剩餘列均分給該頁所有客戶。
 *   4. 萬一單一客戶都超過 budget（品名數 >> budget），仍放 1 個，會自然溢出到下一頁。
 */
function planPage(remaining: HandfillCustomer[], budget: number): number[] {
    const naturalOf = (c: HandfillCustomer) => Math.max(STD_BLOCK_ROWS, c.products.length);

    const maxTry = Math.min(TARGET_CUSTOMERS_PER_PAGE, remaining.length);
    for (let take = maxTry; take >= 1; take--) {
        const natural = remaining.slice(0, take).map(naturalOf);
        const sum = natural.reduce((a, b) => a + b, 0);
        if (sum <= budget) {
            // 最後一頁（已涵蓋所有剩餘客戶）→ 保留自然 8 列，下方留白；
            // 中間頁 → 把多餘列均分撐滿，避免出現空白單列影響閱讀。
            if (take === remaining.length) return natural;
            return distributeSlack(natural, budget);
        }
    }

    // 連 1 個自然尺寸都超過 budget → 直接放，會自然跨頁
    return [naturalOf(remaining[0])];
}

/** 把 budget - 自然列總和 的多餘列數，從前到後依序均分（每多分 1 直到分完）。 */
function distributeSlack(natural: number[], budget: number): number[] {
    const sum = natural.reduce((a, b) => a + b, 0);
    const slack = budget - sum;
    if (slack <= 0) return natural;
    const n = natural.length;
    const extra = Math.floor(slack / n);
    const remainder = slack - extra * n;
    return natural.map((v, i) => v + extra + (i < remainder ? 1 : 0));
}
