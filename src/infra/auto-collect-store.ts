/**
 * 「自動催款」記住的各月明細，存在這台電腦的瀏覽器 localStorage（每月一份）。
 */

import {defaultReasonChips, normalizeReasonChips} from '@/domain/auto-collect/reason-chips';
import {defaultTemplates, type MessageTemplates, normalizeTemplates} from '@/domain/auto-collect/templates';
import {emptyHistory, emptyMarks, extractLegacyRules, type History, type Marks, migrateLegacyRules, normalizeMarks} from '@/domain/auto-collect/ledger';

const KEY = 'bii.auto-collect.history';
const MARKS_KEY = 'bii.auto-collect.marks';
const TEMPLATES_KEY = 'bii.auto-collect.templates';
const REASON_CHIPS_KEY = 'bii.auto-collect.reason-chips';

export function loadHistory(): History {
    try {
        const raw = localStorage.getItem(KEY);
        if (!raw) return emptyHistory();
        const h = JSON.parse(raw) as History;
        if (!h || typeof h.months !== 'object' || !h.months) return emptyHistory();
        // 舊資料或壞掉的月份（缺 periods）直接丟掉，免得整個分頁畫不出來
        const months = Object.fromEntries(Object.entries(h.months).filter(([k, m]) =>
            /^\d{4}-\d{2}$/.test(k) && !!m && typeof m === 'object' && !!m.periods && typeof m.periods === 'object'
            && (['上', '下'] as const).every((half) => !m.periods[half] || typeof m.periods[half] === 'object')));
        return {version: 1, months};
    } catch (err) {
        console.warn('[auto-collect-store] load failed', err);
        return emptyHistory();
    }
}

/** 寫入失敗（例如瀏覽器空間滿了）會丟出錯誤，由呼叫端提示。 */
export function saveHistory(h: History): void {
    localStorage.setItem(KEY, JSON.stringify(h));
}

export function clearHistory(): void {
    localStorage.removeItem(KEY);
}

/**
 * 「已結清」標記與手動欠款，和月份記錄分開存；舊資料沒有這個鍵或缺欄位時視為空。
 * 舊版「有收到款就算付清」設定會在這裡轉成一般已結清標記（涵蓋目前記住的月份），並把舊欄位從儲存的資料清掉。
 */
export function loadMarks(history: History): Marks {
    try {
        const raw = localStorage.getItem(MARKS_KEY);
        if (!raw) return emptyMarks();
        const parsed: unknown = JSON.parse(raw);
        const marks = migrateLegacyRules(history, normalizeMarks(parsed), extractLegacyRules(parsed));
        if (parsed && typeof parsed === 'object' && 'paid_if_received' in parsed) {
            try {
                saveMarks(marks);
            } catch (err) {
                console.warn('[auto-collect-store] save migrated marks failed', err);
            }
        }
        return marks;
    } catch (err) {
        console.warn('[auto-collect-store] load marks failed', err);
        return emptyMarks();
    }
}

export function saveMarks(m: Marks): void {
    localStorage.setItem(MARKS_KEY, JSON.stringify(m));
}

export function clearMarks(): void {
    localStorage.removeItem(MARKS_KEY);
}

/** 訊息範本（設定，不是資料記錄）：沒有、壞掉或存取被擋時回預設。清除全部記錄時不會動它。 */
export function loadTemplates(): MessageTemplates {
    try {
        const raw = localStorage.getItem(TEMPLATES_KEY);
        return raw ? normalizeTemplates(JSON.parse(raw)) : defaultTemplates();
    } catch {
        return defaultTemplates();
    }
}

/** 寫入失敗會丟出錯誤，由呼叫端提示。 */
export function saveTemplates(t: MessageTemplates): void {
    localStorage.setItem(TEMPLATES_KEY, JSON.stringify(t));
}

export function resetTemplates(): void {
    localStorage.removeItem(TEMPLATES_KEY);
}

/** 「已結清」快速原因按鈕（設定）：沒有、壞掉或存取被擋時回預設。清除全部記錄時不會動它。 */
export function loadReasonChips(): string[] {
    try {
        const raw = localStorage.getItem(REASON_CHIPS_KEY);
        return raw ? normalizeReasonChips(JSON.parse(raw)) : defaultReasonChips();
    } catch {
        return defaultReasonChips();
    }
}

export function saveReasonChips(chips: string[]): void {
    try {
        localStorage.setItem(REASON_CHIPS_KEY, JSON.stringify(chips));
    } catch (err) {
        console.warn('[auto-collect-store] save reason chips failed', err);
    }
}
