/**
 * 圖表配色常數。集中於此，避免 hex 散落於各 chart-option。
 */

export const PALETTE = {
    revenue: '#5b8def',
    profit: '#7ac74f',
    cost: '#f5a623',
    accent: '#9b59b6',
} as const;

/** Treemap 毛利率色階（從負到正）。 */
export const MARGIN_COLORS = {
    loss: '#e74c3c',       // < 0%
    breakeven: '#f39c12',  // 0% ~ 10%
    low: '#f4d03f',        // 10% ~ 25%
    mid: '#7ac74f',        // 25% ~ 40%
    high: '#27ae60',       // > 40%
    unset: '#bdc3c7',      // 全部商品未填成本
} as const;

/** 依毛利率（百分比 0~100，可為負）對應 treemap tile 顏色。 */
export function marginColor(pct: number): string {
    if (pct < 0) return MARGIN_COLORS.loss;
    if (pct < 10) return MARGIN_COLORS.breakeven;
    if (pct < 25) return MARGIN_COLORS.low;
    if (pct < 40) return MARGIN_COLORS.mid;
    return MARGIN_COLORS.high;
}

/** 分類 treemap 用的 categorical palette，色相豐富以利視覺辨識。 */
export const CATEGORY_PALETTE = [
    '#5b8def', '#7ac74f', '#f5a623', '#9b59b6',
    '#1abc9c', '#e74c3c', '#3498db', '#e67e22',
    '#16a085', '#8e44ad', '#27ae60', '#d35400',
] as const;

/** 毛利率邊框狀態：虧損紅、損平橙、未填灰、其餘白（健康）。 */
export interface MarginBorder {
    color: string;
    width: number;
}

export function marginBorder(pct: number | null, allCostUnset: boolean): MarginBorder {
    if (allCostUnset || pct === null) return {color: MARGIN_COLORS.unset, width: 4};
    if (pct < 0) return {color: MARGIN_COLORS.loss, width: 4};
    if (pct < 10) return {color: MARGIN_COLORS.breakeven, width: 4};
    return {color: '#ffffff', width: 2};
}

/** 將 hex 顏色朝白色混合，amount=0 原色、amount=1 全白。 */
export function lightenHex(hex: string, amount: number): string {
    const clean = hex.startsWith('#') ? hex.slice(1) : hex;
    const r = parseInt(clean.slice(0, 2), 16);
    const g = parseInt(clean.slice(2, 4), 16);
    const b = parseInt(clean.slice(4, 6), 16);
    const a = Math.max(0, Math.min(1, amount));
    const lr = Math.round(r + (255 - r) * a);
    const lg = Math.round(g + (255 - g) * a);
    const lb = Math.round(b + (255 - b) * a);
    return `rgb(${lr},${lg},${lb})`;
}
