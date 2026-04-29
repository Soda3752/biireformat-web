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
