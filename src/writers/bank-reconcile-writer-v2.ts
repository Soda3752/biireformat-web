/**
 * 對帳 2.0 Excel 輸出
 *
 * 1 個檔案含 3 張工作表：
 *  1. 對帳匯總（客戶主表，未收狀態的客戶不輸出）
 *  2. 需人工複核（多重匹配 / 未對應 / 未設 storeCode）
 *  3. 原始交易明細
 *
 * 各 sheet 從表頭開始、獨立凍結首列與欄寬。狀態欄依 ReconcileStatus 套底色；現金客戶整列灰底淡化。
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

const SHEET_CUSTOMER = '對帳匯總';
const SHEET_MANUAL = '需人工複核';
const SHEET_RAW = '原始交易明細';
const FILE_END_FIX = '.xlsx';

const ARGB_HEADER_BG = 'FF3E5F8A';
const ARGB_HEADER_FG = 'FFFFFFFF';
const ARGB_BORDER = 'FFD0D7DE';
const ARGB_NA_BG = 'FFF1F3F5';
const ARGB_STATUS_MATCHED = 'FFD8F2DF';
const ARGB_STATUS_UNPAID = 'FFFADBD8';
const ARGB_STATUS_PARTIAL = 'FFFFF4D6';
const ARGB_STATUS_OVERPAID = 'FFE8DAEF';
const ARGB_HYPERLINK = 'FF1A56DB';

const HEADER_STYLE = buildStyle({
    font: {bold: true, color: ARGB_HEADER_FG, size: 12},
    fill: ARGB_HEADER_BG,
    align: 'center',
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
    const rawRowIndex = buildRawRowIndex(result.rawRows);

    const customerSheet = wb.addWorksheet(SHEET_CUSTOMER);
    writeCustomerSection(customerSheet, result.customers, rawRowIndex);
    customerSheet.views = [{state: 'frozen', ySplit: 1}];
    setCustomerColumnWidths(customerSheet);

    const manualSheet = wb.addWorksheet(SHEET_MANUAL);
    writeManualReviewSection(manualSheet, result.manualReviewItems, rawRowIndex);
    manualSheet.views = [{state: 'frozen', ySplit: 1}];
    setManualColumnWidths(manualSheet);

    const rawSheet = wb.addWorksheet(SHEET_RAW);
    writeRawSection(rawSheet, result.rawRows);
    rawSheet.views = [{state: 'frozen', ySplit: 1}];
    setRawColumnWidths(rawSheet);

    return {
        filename: buildReconcileFilename(result),
        blob: await workbookToBlob(wb),
    };
}

type RawRowIndex = ReadonlyMap<number, number>;

function buildRawRowIndex(rows: ReadonlyArray<BankRowMatch>): RawRowIndex {
    // raw sheet 第 1 列為表頭，第 i 筆資料對應 row number = i + 2
    const m = new Map<number, number>();
    rows.forEach((r, i) => {
        m.set(r.fileLineNumber, i + 2);
    });
    return m;
}

function rawHyperlink(rowNumber: number): string {
    return `#'${SHEET_RAW}'!A${rowNumber}`;
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

function writeCustomerSection(
    sheet: ExcelJS.Worksheet,
    customers: ReadonlyArray<CustomerReconcileRow>,
    rawRowIndex: RawRowIndex,
): void {
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
        writeCustomerRowGroup(sheet, r, rawRowIndex);
    }
}

function writeCustomerRowGroup(
    sheet: ExcelJS.Worksheet,
    r: VisibleCustomerReconcileRow,
    rawRowIndex: RawRowIndex,
): void {
    const isNa = r.status === 'na' || r.isCashUser;
    const matched = r.matchedRows;
    const rowsCount = Math.max(1, matched.length);
    let groupStartRow = 0;

    for (let i = 0; i < rowsCount; i++) {
        const m: ReconcileMatchedRow | undefined = matched[i];
        const isFirst = i === 0;
        const dataRow = sheet.addRow([
            isFirst ? r.customerCode : '',
            isFirst ? r.customerName : '',
            isFirst ? r.customerLine : '',
            isFirst ? formatPayMode(r) : '',
            isFirst ? r.receivable : '',
            isFirst ? r.received : '',
            isFirst ? r.diff : '',
            isFirst ? statusLabel(r.status) : '',
            null,
        ]);
        if (isFirst) groupStartRow = dataRow.number;

        const detailCell = dataRow.getCell(9);
        let hasHyperlink = false;
        if (m) {
            const text = formatReceiptLine(m, m.crossMonth);
            const target = rawRowIndex.get(m.fileLineNumber);
            if (target !== undefined) {
                detailCell.value = {text, hyperlink: rawHyperlink(target)};
                hasHyperlink = true;
            } else {
                detailCell.value = text;
            }
        } else {
            detailCell.value = '';
        }

        dataRow.eachCell({includeEmpty: true}, (cell, colNumber) => {
            cell.style = buildStyle({font: {size: 11}, align: {vertical: 'middle'}});
            cell.border = thinBorderAll();
            if (colNumber === 5 || colNumber === 6 || colNumber === 7) {
                cell.alignment = {...cell.alignment, horizontal: 'right'};
                cell.numFmt = '#,##0';
                if (colNumber === 7 && isFirst && r.diff !== 0) {
                    cell.font = {...cell.font, bold: true, color: {argb: r.diff < 0 ? 'FFC23E3E' : 'FF8E44AD'}};
                }
            } else if (colNumber === 8) {
                cell.alignment = {...cell.alignment, horizontal: 'center'};
                if (isFirst) {
                    cell.font = {...cell.font, bold: true, color: {argb: statusColor(r.status)}};
                }
            } else if (colNumber === 9) {
                cell.alignment = {...cell.alignment, horizontal: 'left', vertical: 'middle'};
                cell.font = {...cell.font, size: 10, color: {argb: 'FF4B5563'}};
            } else if (colNumber === 1 || colNumber === 3 || colNumber === 4) {
                cell.alignment = {...cell.alignment, horizontal: 'center'};
            }

            if (isNa) {
                cell.fill = solidFill(ARGB_NA_BG);
                cell.font = {...cell.font, color: {argb: 'FF8B95A1'}};
            } else if (colNumber === 8 && isFirst) {
                cell.fill = solidFill(statusBgColor(r.status));
            }

            if (colNumber === 9 && hasHyperlink) {
                cell.font = {...cell.font, color: {argb: ARGB_HYPERLINK}, underline: 'single'};
            }
        });
    }

    if (rowsCount > 1 && groupStartRow > 0) {
        for (let c = 1; c <= 8; c++) {
            sheet.mergeCells(groupStartRow, c, groupStartRow + rowsCount - 1, c);
        }
    }
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
    rawRowIndex: RawRowIndex,
): void {
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

        const dataRow = sheet.addRow([
            MANUAL_REASON_LABEL[item.reason] ?? item.reason,
            null,
            item.row.summary,
            parseAmount(item.row.deposit),
            candidate,
            buildManualReason(item),
        ]);

        const receiptText = formatReceiptLine(item.row);
        const target = rawRowIndex.get(item.row.fileLineNumber);
        const hasHyperlink = target !== undefined;
        dataRow.getCell(2).value = hasHyperlink
            ? {text: receiptText, hyperlink: rawHyperlink(target!)}
            : receiptText;

        dataRow.eachCell({includeEmpty: true}, (cell, colNumber) => {
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
                if (hasHyperlink) {
                    cell.font = {...cell.font, color: {argb: ARGB_HYPERLINK}, underline: 'single'};
                }
            }
        });
    }
}

function writeRawSection(
    sheet: ExcelJS.Worksheet,
    rows: ReadonlyArray<BankRowMatch>,
): void {
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

function applyWidths(sheet: ExcelJS.Worksheet, widths: ReadonlyArray<number>): void {
    widths.forEach((w, i) => {
        sheet.getColumn(i + 1).width = w;
    });
}

function setCustomerColumnWidths(sheet: ExcelJS.Worksheet): void {
    // 編號 / 客戶 / 線別 / 模式 / 應收 / 已收 / 差額 / 狀態 / 匯款詳情
    applyWidths(sheet, [10, 20, 8, 8, 14, 14, 14, 12, 32]);
}

function setManualColumnWidths(sheet: ExcelJS.Worksheet): void {
    // 類型 / 匯款詳情 / 摘要 / 存入 / 配對候選 / 原因
    applyWidths(sheet, [12, 22, 28, 12, 28, 28]);
}

function setRawColumnWidths(sheet: ExcelJS.Worksheet): void {
    // 日期 / 摘要 / 存入 / 配對來源 / 對應客戶
    applyWidths(sheet, [12, 30, 14, 14, 32]);
}
