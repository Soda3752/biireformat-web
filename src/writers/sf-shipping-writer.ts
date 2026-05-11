import ExcelJS from 'exceljs';

import type {SfShippingSettings} from '@/infra/sf-shipping-settings';

/**
 * 順豐託運單輸出：載入 public/assets/sf_template.xlsx 模板，保留所有原有格式
 * （欄寬、字型、填色、資料驗證下拉、隱藏的 Datadict 表），從第 3 列起覆寫資料。
 */

export interface SfShippingOrder {
    orderNo: string;       // A
    productName: string;   // AB
    productQty: number;    // AC
    productPrice: number;  // AE
    parcelCount: number;   // AF
    totalWeight: number;   // AG
}

export interface SfShippingReceiver {
    name: string;     // N
    phone: string;    // O
    address: string;  // Q
}

export interface SfShippingPayload {
    sheetName: string;          // 主表 sheet 名（建議用收件人姓名）
    receiver: SfShippingReceiver;
    orders: ReadonlyArray<SfShippingOrder>;
    settings: SfShippingSettings;
}

const TEMPLATE_URL = `${import.meta.env.BASE_URL}assets/sf_template.xlsx`;
const DATA_START_ROW = 3;

/** 欄位 → ExcelJS column number（A=1, B=2, ..., AB=28, BI=61） */
const COL = {
    orderNo: 1,           // A
    shipperName: 2,       // B
    shipperPhone: 3,      // C
    shipperAddress: 5,    // E
    shipperCity: 7,       // G
    shipperState: 8,      // H
    shipperCountry: 9,    // I
    shipperZip: 10,       // J
    shipperType: 12,      // L
    shipperCompany: 13,   // M
    receiverName: 14,     // N
    receiverPhone: 15,    // O
    receiverAddress: 17,  // Q
    receiverCity: 21,     // U
    receiverState: 22,    // V
    receiverCountry: 23,  // W
    receiverZip: 24,      // X
    productName: 28,      // AB
    productQty: 29,       // AC
    productUnit: 30,      // AD
    productPrice: 31,     // AE
    parcelCount: 32,      // AF
    totalWeight: 33,      // AG
    currency: 38,         // AL
    expressType: 39,      // AM
    pickupMethod: 55,     // BC
    paymentMethod: 58,    // BF
    monthlyCardNo: 59,    // BG
} as const;

export async function buildSfShippingWorkbook(
    payload: SfShippingPayload
): Promise<Blob> {
    const resp = await fetch(TEMPLATE_URL);
    if (!resp.ok) {
        throw new Error(`載入模板失敗：${resp.status} ${resp.statusText}`);
    }
    const buffer = await resp.arrayBuffer();

    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const mainSheet = wb.worksheets[0];
    if (!mainSheet) throw new Error('模板缺少主工作表');
    if (payload.sheetName.trim().length > 0) {
        mainSheet.name = sanitizeSheetName(payload.sheetName);
    }

    fillData(mainSheet, payload);

    const out = await wb.xlsx.writeBuffer();
    return new Blob([out], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

function fillData(sheet: ExcelJS.Worksheet, payload: SfShippingPayload): void {
    const {receiver, orders, settings} = payload;
    const templateRow = sheet.getRow(DATA_START_ROW);
    const templateStyles = snapshotRowStyles(templateRow);

    // 清除模板中第 3 列之後的所有資料列
    const lastRow = sheet.actualRowCount;
    if (lastRow > DATA_START_ROW) {
        sheet.spliceRows(DATA_START_ROW + 1, lastRow - DATA_START_ROW);
    }

    orders.forEach((order, idx) => {
        const rowIndex = DATA_START_ROW + idx;
        const row = sheet.getRow(rowIndex);

        if (idx > 0) {
            applyRowStyles(row, templateStyles);
        }

        // 寄件方（固定，從 settings）
        setCell(row, COL.shipperName, settings.shipperName);
        setCell(row, COL.shipperPhone, settings.shipperPhone);
        setCell(row, COL.shipperAddress, settings.shipperAddress);
        setCell(row, COL.shipperCity, settings.shipperCity);
        setCell(row, COL.shipperState, settings.shipperState);
        setCell(row, COL.shipperCountry, settings.shipperCountry);
        setCell(row, COL.shipperZip, settings.shipperZip);
        setCell(row, COL.shipperType, settings.shipperType);
        setCell(row, COL.shipperCompany, settings.shipperCompany);

        // 收件方（共用一份收件人資料，多筆訂單同一支手機）
        setCell(row, COL.receiverName, receiver.name);
        setCell(row, COL.receiverPhone, receiver.phone);
        setCell(row, COL.receiverAddress, receiver.address);
        setCell(row, COL.receiverCity, settings.receiverCity);
        setCell(row, COL.receiverState, settings.receiverState);
        setCell(row, COL.receiverCountry, settings.receiverCountry);
        setCell(row, COL.receiverZip, settings.receiverZip);

        // 訂單號（A 欄為文字格式 @）
        setCell(row, COL.orderNo, order.orderNo);

        // 商品（每筆可獨立）
        setCell(row, COL.productName, order.productName);
        setCell(row, COL.productQty, order.productQty);
        setCell(row, COL.productUnit, settings.productUnit);
        setCell(row, COL.productPrice, order.productPrice);
        setCell(row, COL.parcelCount, order.parcelCount);
        setCell(row, COL.totalWeight, order.totalWeight);

        // 運送（固定）
        setCell(row, COL.currency, settings.currency);
        setCell(row, COL.expressType, settings.expressType);
        setCell(row, COL.pickupMethod, settings.pickupMethod);
        setCell(row, COL.paymentMethod, settings.paymentMethod);
        setCell(row, COL.monthlyCardNo, settings.monthlyCardNo);

        row.commit();
    });
}

/** 取得整列 cell 的 style 與 dataValidation 快照（給後續新列複製用） */
interface CellStyleSnapshot {
    style: Partial<ExcelJS.Style>;
    dataValidation: ExcelJS.DataValidation | undefined;
}

function snapshotRowStyles(row: ExcelJS.Row): Map<number, CellStyleSnapshot> {
    const map = new Map<number, CellStyleSnapshot>();
    // 模板 row 3 預期有完整 61 欄樣式
    for (let col = 1; col <= 61; col++) {
        const cell = row.getCell(col);
        map.set(col, {
            style: deepCloneStyle(cell.style),
            dataValidation: cell.dataValidation,
        });
    }
    return map;
}

function applyRowStyles(
    row: ExcelJS.Row,
    snapshot: Map<number, CellStyleSnapshot>
): void {
    for (const [col, snap] of snapshot) {
        const cell = row.getCell(col);
        cell.style = deepCloneStyle(snap.style);
        if (snap.dataValidation) {
            cell.dataValidation = {...snap.dataValidation};
        }
    }
}

function deepCloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
    if (!style) return {};
    return JSON.parse(JSON.stringify(style)) as Partial<ExcelJS.Style>;
}

function setCell(row: ExcelJS.Row, col: number, value: string | number): void {
    row.getCell(col).value = value;
}

/** Excel sheet 名稱不可含 \ / ? * [ ] : 且長度 <= 31 */
function sanitizeSheetName(raw: string): string {
    const cleaned = raw.replace(/[\\/?*[\]:]/g, '').trim();
    return cleaned.length === 0 ? 'information' : cleaned.slice(0, 31);
}

/** 產生訂單號：YYYYMMDD + 3 位流水（001 起算） */
export function buildOrderNumbers(
    date: Date,
    count: number,
    startSeq = 1
): string[] {
    const y = date.getFullYear().toString();
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const d = date.getDate().toString().padStart(2, '0');
    const prefix = `${y}${m}${d}`;
    const out: string[] = [];
    for (let i = 0; i < count; i++) {
        out.push(`${prefix}${(startSeq + i).toString().padStart(3, '0')}`);
    }
    return out;
}
