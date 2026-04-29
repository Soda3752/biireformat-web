/**
 * 設定頁的本地儲存層（localStorage）。
 *
 * 兩份排序 CSV（cargo_sort、daily_report_list）原本只能由 public/assets 取得，
 * 現在允許使用者在設定頁覆寫；覆寫後將整段 CSV 文字寫進 localStorage。
 *
 * 規則：
 * - 沒寫過 → loader 回退到 public/assets 內建檔
 * - 寫過 → loader 直接拿 localStorage 中的 CSV 文字解析
 * - clear → 移除覆寫，下次回退到內建
 */

const STORAGE_KEYS = {
    cargoSort: 'bii.settings.cargoSort.csv',
    dailyReportList: 'bii.settings.dailyReportList.csv',
    lastFiveDigit: 'bii.settings.lastFiveDigit.csv',
    customerOrderBill: 'bii.settings.customerOrder.bill.csv',
    customerOrderOverview: 'bii.settings.customerOrder.overview.csv',
} as const;

/** 舊版單一 customerOrder key，僅供 customer-order-loader 做一次性遷移使用。 */
export const LEGACY_CUSTOMER_ORDER_KEY = 'bii.settings.customerOrder.csv';

export type SettingsKey = keyof typeof STORAGE_KEYS;

function readKey(key: SettingsKey): string | null {
    try {
        return localStorage.getItem(STORAGE_KEYS[key]);
    } catch {
        return null;
    }
}

function writeKey(key: SettingsKey, value: string): void {
    localStorage.setItem(STORAGE_KEYS[key], value);
}

function removeKey(key: SettingsKey): void {
    try {
        localStorage.removeItem(STORAGE_KEYS[key]);
    } catch {
        // ignore
    }
}

export const localSettings = {
    hasCargoSort(): boolean {
        return readKey('cargoSort') !== null;
    },
    getCargoSort(): string | null {
        return readKey('cargoSort');
    },
    setCargoSort(csv: string): void {
        writeKey('cargoSort', csv);
    },
    clearCargoSort(): void {
        removeKey('cargoSort');
    },

    hasDailyReportList(): boolean {
        return readKey('dailyReportList') !== null;
    },
    getDailyReportList(): string | null {
        return readKey('dailyReportList');
    },
    setDailyReportList(csv: string): void {
        writeKey('dailyReportList', csv);
    },
    clearDailyReportList(): void {
        removeKey('dailyReportList');
    },

    hasLastFiveDigit(): boolean {
        return readKey('lastFiveDigit') !== null;
    },
    getLastFiveDigit(): string | null {
        return readKey('lastFiveDigit');
    },
    setLastFiveDigit(csv: string): void {
        writeKey('lastFiveDigit', csv);
    },
    clearLastFiveDigit(): void {
        removeKey('lastFiveDigit');
    },

    hasCustomerOrderBill(): boolean {
        return readKey('customerOrderBill') !== null;
    },
    getCustomerOrderBill(): string | null {
        return readKey('customerOrderBill');
    },
    setCustomerOrderBill(csv: string): void {
        writeKey('customerOrderBill', csv);
    },
    clearCustomerOrderBill(): void {
        removeKey('customerOrderBill');
    },

    hasCustomerOrderOverview(): boolean {
        return readKey('customerOrderOverview') !== null;
    },
    getCustomerOrderOverview(): string | null {
        return readKey('customerOrderOverview');
    },
    setCustomerOrderOverview(csv: string): void {
        writeKey('customerOrderOverview', csv);
    },
    clearCustomerOrderOverview(): void {
        removeKey('customerOrderOverview');
    },
};

// ============================================================================
// 全部設定的匯出 / 匯入（單一 JSON 檔）
// ============================================================================

export const SETTINGS_EXPORT_VERSION = 1;

export interface SettingsExportPayload {
    version: typeof SETTINGS_EXPORT_VERSION;
    exportedAt: string;
    cargoSort: string | null;
    dailyReportList: string | null;
    lastFiveDigit: string | null;
    customerOrderBill: string | null;
    customerOrderOverview: string | null;
}

const EXPORT_KEYS: ReadonlyArray<Exclude<keyof SettingsExportPayload, 'version' | 'exportedAt'>> = [
    'cargoSort',
    'dailyReportList',
    'lastFiveDigit',
    'customerOrderBill',
    'customerOrderOverview',
];

export function exportAllSettings(): SettingsExportPayload {
    return {
        version: SETTINGS_EXPORT_VERSION,
        exportedAt: new Date().toISOString(),
        cargoSort: readKey('cargoSort'),
        dailyReportList: readKey('dailyReportList'),
        lastFiveDigit: readKey('lastFiveDigit'),
        customerOrderBill: readKey('customerOrderBill'),
        customerOrderOverview: readKey('customerOrderOverview'),
    };
}

export class SettingsImportError extends Error {
    constructor(message: string) {
        super(message);
        this.name = 'SettingsImportError';
    }
}

export function parseSettingsExportPayload(raw: unknown): SettingsExportPayload {
    if (!raw || typeof raw !== 'object') {
        throw new SettingsImportError('檔案內容不是有效的 JSON 物件');
    }
    const obj = raw as Record<string, unknown>;
    if (obj.version !== SETTINGS_EXPORT_VERSION) {
        throw new SettingsImportError(
            `不支援的設定版本（version=${String(obj.version)}，預期 ${SETTINGS_EXPORT_VERSION}）`
        );
    }
    const result: SettingsExportPayload = {
        version: SETTINGS_EXPORT_VERSION,
        exportedAt: typeof obj.exportedAt === 'string' ? obj.exportedAt : '',
        cargoSort: null,
        dailyReportList: null,
        lastFiveDigit: null,
        customerOrderBill: null,
        customerOrderOverview: null,
    };
    for (const key of EXPORT_KEYS) {
        const value = obj[key];
        if (value === null || value === undefined) continue;
        if (typeof value !== 'string') {
            throw new SettingsImportError(`欄位 ${key} 必須是字串或 null`);
        }
        result[key] = value;
    }
    return result;
}

/** 整包覆寫：欄位為 string → 寫入；欄位為 null → 清除（回退到內建預設）。 */
export function importAllSettings(payload: SettingsExportPayload): void {
    for (const key of EXPORT_KEYS) {
        const value = payload[key];
        if (typeof value === 'string') {
            writeKey(key, value);
        } else {
            removeKey(key);
        }
    }
}

export function summarizeImportPayload(payload: SettingsExportPayload): {
    label: string;
    rowCount: number | null;
}[] {
    const labelMap: Record<typeof EXPORT_KEYS[number], string> = {
        cargoSort: '帳單排序',
        dailyReportList: '品項分類',
        lastFiveDigit: '末五碼',
        customerOrderBill: '帳單客戶排序',
        customerOrderOverview: '明細客戶排序',
    };
    return EXPORT_KEYS.map((key) => {
        const value = payload[key];
        return {
            label: labelMap[key],
            rowCount: typeof value === 'string' ? countCsvDataRows(value) : null,
        };
    });
}

/** 粗略計算 CSV 資料列數（扣除 header 與空行）。匯入確認對話框用。 */
function countCsvDataRows(csv: string): number {
    const lines = csv.split(/\r?\n/);
    let count = 0;
    for (let i = 1; i < lines.length; i++) {
        if (lines[i].trim().length > 0) count += 1;
    }
    return count;
}
