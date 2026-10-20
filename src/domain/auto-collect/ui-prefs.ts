/**
 * 自動催款分頁的畫面偏好（各組收合／展開），和欠款資料分開存。
 * 總表與訊息區的開合各自獨立；已結清只出現在總表那一側。
 */

export type SectionKey = 'four' | 'three' | 'two' | 'one_eom' | 'one' | 'settled';
export type Side = 'table' | 'messages';

/** true = 收起 */
export type Collapsed = Record<SectionKey, boolean>;
export interface UiPrefs {
    version: 1;
    table: Collapsed;
    messages: Collapsed;
}

export const SECTION_KEYS: readonly SectionKey[] = ['four', 'three', 'two', 'one_eom', 'one', 'settled'];

/** 預設全部收起；使用者自己展開過的偏好存下後優先。 */
export const defaultCollapsed = (): Collapsed => ({four: true, three: true, two: true, one_eom: true, one: true, settled: true});

export const defaultUiPrefs = (): UiPrefs => ({version: 1, table: defaultCollapsed(), messages: defaultCollapsed()});

const normalizeSide = (raw: unknown): Collapsed => {
    const out = defaultCollapsed();
    if (raw && typeof raw === 'object') {
        const r = raw as Record<string, unknown>;
        for (const k of SECTION_KEYS) {
            const v = r[k];
            if (typeof v === 'boolean') out[k] = v;
        }
        // 舊的「3 期以上」偏好：新的「4 期以上」沿用它（新組沒存過時）
        if (typeof r.four !== 'boolean' && typeof r.three === 'boolean') out.four = r.three;
    }
    return out;
};

/** 壞掉、缺欄位或型別不對的部分一律用預設。 */
export function normalizeUiPrefs(raw: unknown): UiPrefs {
    if (!raw || typeof raw !== 'object') return defaultUiPrefs();
    const r = raw as Record<string, unknown>;
    return {version: 1, table: normalizeSide(r.table), messages: normalizeSide(r.messages)};
}
