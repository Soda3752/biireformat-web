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
    defaultExcludedCustomers: 'bii.settings.defaultExcludedCustomers.csv',
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

    hasDefaultExcludedCustomers(): boolean {
        return readKey('defaultExcludedCustomers') !== null;
    },
    getDefaultExcludedCustomers(): string | null {
        return readKey('defaultExcludedCustomers');
    },
    setDefaultExcludedCustomers(csv: string): void {
        writeKey('defaultExcludedCustomers', csv);
    },
    clearDefaultExcludedCustomers(): void {
        removeKey('defaultExcludedCustomers');
    },
};

// ============================================================================
// 全部設定的匯出 / 匯入（單一 JSON 檔）
// ============================================================================

export const SETTINGS_EXPORT_VERSION = 1;

/**
 * 順豐託運單分頁兩個 JSON-based key：
 * 在 import 時直接以 raw JSON 字串寫回 localStorage，由各自模組驗證 schema。
 */
const SF_SHIPPING_SETTINGS_KEY = 'bii.sfShipping.settings.json';
const SF_FAVORITE_RECEIVERS_KEY = 'bii.sfShipping.favoriteReceivers.json';

export interface SettingsExportPayload {
    version: typeof SETTINGS_EXPORT_VERSION;
    exportedAt: string;
    cargoSort: string | null;
    dailyReportList: string | null;
    lastFiveDigit: string | null;
    customerOrderBill: string | null;
    customerOrderOverview: string | null;
    defaultExcludedCustomers: string | null;
    sfShippingSettings: string | null;
    sfFavoriteReceivers: string | null;
}

const EXPORT_KEYS: ReadonlyArray<Exclude<keyof SettingsExportPayload, 'version' | 'exportedAt'>> = [
    'cargoSort',
    'dailyReportList',
    'lastFiveDigit',
    'customerOrderBill',
    'customerOrderOverview',
    'defaultExcludedCustomers',
    'sfShippingSettings',
    'sfFavoriteReceivers',
];

function readRawKey(key: string): string | null {
    try {
        return localStorage.getItem(key);
    } catch {
        return null;
    }
}

function writeRawKey(key: string, value: string): void {
    try {
        localStorage.setItem(key, value);
    } catch {
        // ignore
    }
}

function removeRawKey(key: string): void {
    try {
        localStorage.removeItem(key);
    } catch {
        // ignore
    }
}

export function exportAllSettings(): SettingsExportPayload {
    return {
        version: SETTINGS_EXPORT_VERSION,
        exportedAt: new Date().toISOString(),
        cargoSort: readKey('cargoSort'),
        dailyReportList: readKey('dailyReportList'),
        lastFiveDigit: readKey('lastFiveDigit'),
        customerOrderBill: readKey('customerOrderBill'),
        customerOrderOverview: readKey('customerOrderOverview'),
        defaultExcludedCustomers: readKey('defaultExcludedCustomers'),
        sfShippingSettings: readRawKey(SF_SHIPPING_SETTINGS_KEY),
        sfFavoriteReceivers: readRawKey(SF_FAVORITE_RECEIVERS_KEY),
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
        defaultExcludedCustomers: null,
        sfShippingSettings: null,
        sfFavoriteReceivers: null,
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
        if (key === 'sfShippingSettings') {
            if (typeof value === 'string') writeRawKey(SF_SHIPPING_SETTINGS_KEY, value);
            else removeRawKey(SF_SHIPPING_SETTINGS_KEY);
        } else if (key === 'sfFavoriteReceivers') {
            if (typeof value === 'string') writeRawKey(SF_FAVORITE_RECEIVERS_KEY, value);
            else removeRawKey(SF_FAVORITE_RECEIVERS_KEY);
        } else if (typeof value === 'string') {
            writeKey(key as SettingsKey, value);
        } else {
            removeKey(key as SettingsKey);
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
        defaultExcludedCustomers: '預設排除店家',
        sfShippingSettings: '順豐託運單設定',
        sfFavoriteReceivers: '收件人最愛',
    };
    return EXPORT_KEYS.map((key) => {
        const value = payload[key];
        return {
            label: labelMap[key],
            rowCount:
                typeof value === 'string'
                    ? key === 'sfShippingSettings'
                        ? value.length > 0
                            ? 1
                            : 0
                        : key === 'sfFavoriteReceivers'
                            ? countJsonArrayItems(value)
                            : countCsvDataRows(value)
                    : null,
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

function countJsonArrayItems(json: string): number {
    try {
        const arr = JSON.parse(json);
        return Array.isArray(arr) ? arr.length : 0;
    } catch {
        return 0;
    }
}
