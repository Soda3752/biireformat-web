/**
 * 對帳 2.0 Excel 輸出
 *
 * 1 張 sheet 含 3 個區塊：
 *  1. 客戶對帳表（主表，未收狀態的客戶不輸出）
 *  2. 需人工複核（多重匹配 / 未對應 / 未設 storeCode）
 *  3. 原始交易明細
 *
 * 區塊之間留 2 行空白。狀態欄依 ReconcileStatus 套底色；現金客戶整列灰底淡化。
 */

import type ExcelJS from 'exceljs';

import {buildStyle, createWorkbook, workbookToBlob} from '@/infra/excel-service';
import type {
    CustomerReconcileRow,
    ManualReviewItem,
    ReconcileMatchedRow,
    ReconcileResult,
    ReconcileStatus,
} from '@/domain/bank-reconcile-service';
import type {BankRowMatch} from '@/domain/bank-match-service';

const SHEET_NAME = '對帳匯總';
const FILE_END_FIX = '.xlsx';

const SECTION_BLANK_ROWS = 2;

const ARGB_HEADER_BG = 'FF3E5F8A';
const ARGB_HEADER_FG = 'FFFFFFFF';
const ARGB_TITLE_FG = 'FF1F2937';
const ARGB_BORDER = 'FFD0D7DE';
const ARGB_NA_BG = 'FFF1F3F5';
const ARGB_STATUS_MATCHED = 'FFD8F2DF';
const ARGB_STATUS_UNPAID = 'FFFADBD8';
const ARGB_STATUS_PARTIAL = 'FFFFF4D6';
const ARGB_STATUS_OVERPAID = 'FFE8DAEF';

const HEADER_STYLE = buildStyle({
    font: {bold: true, color: ARGB_HEADER_FG, size: 12},
    fill: ARGB_HEADER_BG,
    align: 'center',
});

const SECTION_TITLE_STYLE = buildStyle({
    font: {bold: true, size: 14, color: ARGB_TITLE_FG},
    align: 'left',
});

const SUMMARY_STYLE = buildStyle({
    font: {size: 11, color: 'FF4B5563'},
    align: 'left',
});

const CUSTOMER_HEADER = [
    '編號', '客戶', '線別', '結帳模式', '應收', '已收', '差額', '狀態', '匯款詳情',
] as const;

const MANUAL_HEADER = [
    '類型', '匯款詳情', '摘要', '存入', '配對候選', '原因',
] as const;

const RAW_HEADER = [
    '日期', '摘要', '存入', '配對來源', '對應客戶',
] as const;

const MANUAL_REASON_LABEL: Record<string, string> = {
    'multi-match': '多重匹配',
    'no-store-code': '未設店家編號',
    'unmatched': '未配對',
};

export interface ReconcileFile {
    filename: string;
    blob: Blob;
}

export async function writeReconcileWorkbook(result: ReconcileResult): Promise<ReconcileFile> {
    const wb = createWorkbook();
    const sheet = wb.addWorksheet(SHEET_NAME);

    writeTitle(sheet, result);
    writeBlankRows(sheet, 1);

    writeCustomerSection(sheet, result.customers);
    writeBlankRows(sheet, SECTION_BLANK_ROWS);

    writeManualReviewSection(sheet, result.manualReviewItems);
    writeBlankRows(sheet, SECTION_BLANK_ROWS);

    writeRawSection(sheet, result.rawRows);

    sheet.views = [{state: 'frozen', ySplit: 2}];
    setColumnWidths(sheet);

    return {
        filename: buildReconcileFilename(result),
        blob: await workbookToBlob(wb),
    };
}

export function buildReconcileFilename(result: ReconcileResult): string {
    const month = result.bill.billDateInfo?.month;
    if (month) return `${month}月_對帳匯總${FILE_END_FIX}`;
    const d = new Date();
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `對帳匯總_${yyyy}${mm}${dd}${FILE_END_FIX}`;
}

function writeTitle(sheet: ExcelJS.Worksheet, result: ReconcileResult): void {
    const month = result.bill.billDateInfo?.month ?? '';
    const titleRow = sheet.addRow([`${month}月_對帳匯總`]);
    titleRow.getCell(1).style = SECTION_TITLE_STYLE;
    titleRow.height = 22;

    const s = result.summary;
    const summary = `應收合計 ${formatMoney(s.totalReceivable)}　／　已收合計 ${formatMoney(s.totalReceived)}　／　差額 ${formatMoney(s.totalDiff)}　／　已配對 ${s.matchedCount} 戶　／　部分 ${s.partialCount} 戶　／　超收 ${s.overpaidCount} 戶　／　待覆核 ${s.manualReviewCount} 筆`;
    const summaryRow = sheet.addRow([summary]);
    summaryRow.getCell(1).style = SUMMARY_STYLE;
    summaryRow.height = 18;
}

function writeCustomerSection(
    sheet: ExcelJS.Worksheet,
    customers: ReadonlyArray<CustomerReconcileRow>,
): void {
    const sectionRow = sheet.addRow(['◤ 客戶對帳表']);
    sectionRow.getCell(1).style = SECTION_TITLE_STYLE;

    const headerRow = sheet.addRow([...CUSTOMER_HEADER]);
    headerRow.eachCell((cell) => {
        cell.style = HEADER_STYLE;
        cell.border = thinBorderAll();
    });
    headerRow.height = 22;

    const visibleCustomers = customers.filter(
        (r): r is VisibleCustomerReconcileRow => r.status !== 'unpaid',
    );
    for (const r of visibleCustomers) {
        const receiptText = formatReceiptCell(r.matchedRows);
        const row = sheet.addRow([
            r.customerCode,
            r.customerName,
            r.customerLine,
            formatPayMode(r),
            r.receivable,
            r.received,
            r.diff,
            statusLabel(r.status),
            receiptText,
        ]);

        const isNa = r.status === 'na' || r.isCashUser;
        row.eachCell({includeEmpty: true}, (cell, colNumber) => {
            cell.style = buildStyle({font: {size: 11}, align: {vertical: 'middle'}});
            cell.border = thinBorderAll();
            if (colNumber === 5 || colNumber === 6 || colNumber === 7) {
                cell.alignment = {...cell.alignment, horizontal: 'right'};
                cell.numFmt = '#,##0';
                if (colNumber === 7 && r.diff !== 0) {
                    cell.font = {...cell.font, bold: true, color: {argb: r.diff < 0 ? 'FFC23E3E' : 'FF8E44AD'}};
                }
            } else if (colNumber === 8) {
                cell.alignment = {...cell.alignment, horizontal: 'center'};
                cell.font = {...cell.font, bold: true, color: {argb: statusColor(r.status)}};
            } else if (colNumber === 9) {
                cell.alignment = {...cell.alignment, horizontal: 'left', vertical: 'top', wrapText: true};
                cell.font = {...cell.font, size: 10, color: {argb: 'FF4B5563'}};
            } else if (colNumber === 1 || colNumber === 3 || colNumber === 4) {
                cell.alignment = {...cell.alignment, horizontal: 'center'};
            }

            if (isNa) {
                cell.fill = solidFill(ARGB_NA_BG);
                cell.font = {...cell.font, color: {argb: 'FF8B95A1'}};
            } else if (colNumber === 8) {
                cell.fill = solidFill(statusBgColor(r.status));
            }
        });
    }
}

function formatReceiptCell(rows: ReadonlyArray<ReconcileMatchedRow>): string {
    if (rows.length === 0) return '';
    return rows.map((r) => formatReceiptLine(r, r.crossMonth)).join('\n');
}

function formatReceiptLine(row: BankRowMatch, crossMonth = false): string {
    const date = (row.date ?? '').trim() || '—';
    const account = (row.account ?? '').trim();
    const base = account ? `${date}  #${account}` : date;
    return crossMonth ? `${base} (跨月)` : base;
}

function writeManualReviewSection(
    sheet: ExcelJS.Worksheet,
    items: ReadonlyArray<ManualReviewItem>,
): void {
    const sectionRow = sheet.addRow([`◤ 需人工複核（${items.length} 筆）`]);
    sectionRow.getCell(1).style = SECTION_TITLE_STYLE;

    const headerRow = sheet.addRow([...MANUAL_HEADER]);
    headerRow.eachCell((cell) => {
        cell.style = HEADER_STYLE;
        cell.border = thinBorderAll();
    });
    headerRow.height = 22;

    if (items.length === 0) {
        const placeholder = sheet.addRow(['（無）', '', '', '', '', '']);
        placeholder.eachCell({includeEmpty: true}, (cell) => {
            cell.style = buildStyle({font: {size: 11, color: 'FF8B95A1'}, align: 'center'});
            cell.border = thinBorderAll();
        });
        return;
    }

    for (const item of items) {
        const candidate = item.candidates.length > 0
            ? item.candidates.map((c) => {
                const line = c.customerLine ? `/${c.customerLine}` : '';
                return `${c.customerName}${line}`;
            }).join('、')
            : '(無)';

        const row = sheet.addRow([
            MANUAL_REASON_LABEL[item.reason] ?? item.reason,
            formatReceiptLine(item.row),
            item.row.summary,
            parseAmount(item.row.deposit),
            candidate,
            buildManualReason(item),
        ]);
        row.eachCell({includeEmpty: true}, (cell, colNumber) => {
            cell.style = buildStyle({font: {size: 11}, align: {vertical: 'middle'}});
            cell.border = thinBorderAll();
            cell.fill = solidFill(ARGB_STATUS_PARTIAL);
            if (colNumber === 4) {
                cell.alignment = {...cell.alignment, horizontal: 'right'};
                cell.numFmt = '#,##0';
            } else if (colNumber === 1) {
                cell.alignment = {...cell.alignment, horizontal: 'center'};
            } else if (colNumber === 2) {
                cell.alignment = {...cell.alignment, horizontal: 'left'};
                cell.font = {...cell.font, size: 10, color: {argb: 'FF4B5563'}};
            }
        });
    }
}

function writeRawSection(
    sheet: ExcelJS.Worksheet,
    rows: ReadonlyArray<BankRowMatch>,
): void {
    const sectionRow = sheet.addRow([`◤ 原始交易明細（${rows.length} 筆）`]);
    sectionRow.getCell(1).style = SECTION_TITLE_STYLE;

    const headerRow = sheet.addRow([...RAW_HEADER]);
    headerRow.eachCell((cell) => {
        cell.style = HEADER_STYLE;
        cell.border = thinBorderAll();
    });
    headerRow.height = 22;

    for (const row of rows) {
        const customers = row.matches
            .map((m) => `${m.customerName}${m.customerLine ? `/${m.customerLine}` : ''}`)
            .join('、');
        const dataRow = sheet.addRow([
            row.date,
            row.summary,
            parseAmount(row.deposit),
            row.account,
            row.matches.length === 0 ? '(未配對)' : customers,
        ]);
        const unmatched = row.matches.length === 0;
        dataRow.eachCell({includeEmpty: true}, (cell, colNumber) => {
            cell.style = buildStyle({font: {size: 11}, align: {vertical: 'middle'}});
            cell.border = thinBorderAll();
            if (colNumber === 3) {
                cell.alignment = {...cell.alignment, horizontal: 'right'};
                cell.numFmt = '#,##0';
            } else if (colNumber === 1) {
                cell.alignment = {...cell.alignment, horizontal: 'center'};
            }
            if (unmatched) cell.fill = solidFill(ARGB_STATUS_UNPAID);
        });
    }
}

function buildManualReason(item: ManualReviewItem): string {
    switch (item.reason) {
        case 'multi-match':
            return '同末五碼匹配多家，請人工指定';
        case 'no-store-code':
            return '匹配到末五碼但該筆未設店家編號';
        case 'unmatched':
            return '末五碼對照表中無此帳號';
    }
}

type VisibleReconcileStatus = Exclude<ReconcileStatus, 'unpaid'>;

type VisibleCustomerReconcileRow = CustomerReconcileRow & { status: VisibleReconcileStatus };

function statusLabel(s: VisibleReconcileStatus): string {
    switch (s) {
        case 'matched':
            return '✓ 已收';
        case 'partial':
            return '⚠ 部分';
        case 'overpaid':
            return '⚠ 超收';
        case 'na':
            return '—';
    }
}

function statusColor(s: VisibleReconcileStatus): string {
    switch (s) {
        case 'matched':
            return 'FF1F8A4C';
        case 'partial':
            return 'FF8B6914';
        case 'overpaid':
            return 'FF8E44AD';
        case 'na':
            return 'FF8B95A1';
    }
}

function statusBgColor(s: VisibleReconcileStatus): string {
    switch (s) {
        case 'matched':
            return ARGB_STATUS_MATCHED;
        case 'partial':
            return ARGB_STATUS_PARTIAL;
        case 'overpaid':
            return ARGB_STATUS_OVERPAID;
        case 'na':
            return ARGB_NA_BG;
    }
}

function formatPayMode(r: CustomerReconcileRow): string {
    let s = '';
    if (r.isCashUser) s += '現';
    if (r.isMonthly) s += '月';
    if (r.isNeedTex) s += '稅';
    return s || '—';
}

function formatMoney(n: number): string {
    return n.toLocaleString('zh-TW');
}

function parseAmount(raw: string): number | string {
    const s = String(raw ?? '').replace(/,/g, '').trim();
    if (s === '') return '';
    const n = Number(s);
    return Number.isFinite(n) ? n : raw;
}

function thinBorderAll(): Partial<ExcelJS.Borders> {
    return {
        top: {style: 'thin', color: {argb: ARGB_BORDER}},
        left: {style: 'thin', color: {argb: ARGB_BORDER}},
        bottom: {style: 'thin', color: {argb: ARGB_BORDER}},
        right: {style: 'thin', color: {argb: ARGB_BORDER}},
    };
}

function solidFill(argb: string): ExcelJS.FillPattern {
    return {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb},
    };
}

function writeBlankRows(sheet: ExcelJS.Worksheet, count: number): void {
    for (let i = 0; i < count; i++) sheet.addRow([]);
}

function setColumnWidths(sheet: ExcelJS.Worksheet): void {
    // 主表寬度為主基準（9 欄）：編號 / 客戶 / 線別 / 模式 / 應收 / 已收 / 差額 / 狀態 / 匯款詳情
    // 同 sheet 含其他區塊（5 欄 / 6 欄），最寬欄沿用 9 欄基準
    const widths = [10, 20, 8, 8, 14, 14, 14, 12, 32];
    widths.forEach((w, i) => {
        sheet.getColumn(i + 1).width = w;
    });
}
