/**
 * 「已結清」視窗的快速原因按鈕清單（設定，使用者可自己新增／刪除）。
 */

export const DEFAULT_REASON_CHIPS: readonly string[] = ['已補款', '回饋金自扣'];

export const MAX_REASON_LENGTH = 20;

export const defaultReasonChips = (): string[] => [...DEFAULT_REASON_CHIPS];

/** 壞資料（非陣列、非字串、空白、重複）一律濾掉；整份不是陣列時用預設。存成空陣列是合法的（使用者全刪）。 */
export function normalizeReasonChips(raw: unknown): string[] {
    if (!Array.isArray(raw)) return defaultReasonChips();
    const out: string[] = [];
    for (const v of raw) {
        if (typeof v !== 'string') continue;
        const t = v.trim().slice(0, MAX_REASON_LENGTH);
        if (t && !out.includes(t)) out.push(t);
    }
    return out;
}

/** 空白或重複不加，回傳原清單。 */
export function addReasonChip(chips: readonly string[], text: string): string[] {
    const t = text.trim().slice(0, MAX_REASON_LENGTH);
    return !t || chips.includes(t) ? [...chips] : [...chips, t];
}

export const removeReasonChip = (chips: readonly string[], text: string): string[] => chips.filter((c) => c !== text);
