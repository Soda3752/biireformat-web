/**
 * 順豐託運單分頁的本地設定（localStorage）。
 *
 * 寄件方資訊、收件方預設、商品預設、付款資訊等固定值統一存在
 * 一個 JSON key `bii.sfShipping.settings.json`。
 */

export const SF_SHIPPING_SETTINGS_KEY = 'bii.sfShipping.settings.json';
const STORAGE_KEY = SF_SHIPPING_SETTINGS_KEY;

export interface SfShippingSettings {
    // 寄件方
    shipperName: string;
    shipperPhone: string;
    shipperAddress: string;
    shipperCity: string;       // G
    shipperState: string;      // H
    shipperCountry: string;    // I
    shipperZip: string;        // J
    shipperType: string;       // L  公司件 / 個人件
    shipperCompany: string;    // M
    // 收件方預設
    receiverCity: string;      // U
    receiverState: string;     // V
    receiverCountry: string;   // W
    receiverZip: string;       // X
    // 商品預設
    productName: string;       // AB
    productQty: number;        // AC
    productUnit: string;       // AD
    productPrice: number;      // AE
    parcelCount: number;       // AF
    totalWeight: number;       // AG
    // 運送
    currency: string;          // AL  NTD/...
    expressType: string;       // AM  順豐特快
    pickupMethod: string;      // BC  快遞員上門取件
    paymentMethod: string;     // BF  寄付/到付/第三方付
    monthlyCardNo: string;     // BG  月結卡號
}

export const DEFAULT_SF_SETTINGS: SfShippingSettings = {
    shipperName: '趙家儀',
    shipperPhone: '0908112588',
    shipperAddress: '彰化縣和美鎮中興路二段29號',
    shipperCity: '.',
    shipperState: '.',
    shipperCountry: '中國臺灣',
    shipperZip: '0000000',
    shipperType: '公司件',
    shipperCompany: '青坊食品有限公司',
    receiverCity: '.',
    receiverState: '.',
    receiverCountry: '中國臺灣',
    receiverZip: '886000',
    productName: '麵包',
    productQty: 1,
    productUnit: '箱',
    productPrice: 500,
    parcelCount: 1,
    totalWeight: 1,
    currency: 'NTD',
    expressType: '順豐特快',
    pickupMethod: '快遞員上門取件',
    paymentMethod: '寄付',
    monthlyCardNo: '8860746803',
};

export function loadSfSettings(): SfShippingSettings {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return {...DEFAULT_SF_SETTINGS};
        const parsed = JSON.parse(raw) as Partial<SfShippingSettings>;
        return {...DEFAULT_SF_SETTINGS, ...parsed};
    } catch {
        return {...DEFAULT_SF_SETTINGS};
    }
}

export function saveSfSettings(settings: SfShippingSettings): void {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function resetSfSettings(): void {
    try {
        localStorage.removeItem(STORAGE_KEY);
    } catch {
        // ignore
    }
}
