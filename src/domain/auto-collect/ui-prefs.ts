/**
 * 自動催款分頁的畫面狀態（各組收合／展開）。只存在記憶體，不寫入 localStorage：
 * 每次載入（含重新整理、重新匯入）一律全部收起。
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

/** 全部收起。 */
export const defaultCollapsed = (): Collapsed => ({four: true, three: true, two: true, one_eom: true, one: true, settled: true});

export const defaultUiPrefs = (): UiPrefs => ({version: 1, table: defaultCollapsed(), messages: defaultCollapsed()});

/** 全部展開的暫時版本（點「需確認」篩選時用），不改原物件。 */
export const expandAll = (p: UiPrefs): UiPrefs => ({
    ...p,
    table: Object.fromEntries(SECTION_KEYS.map((k) => [k, false])) as Collapsed,
    messages: Object.fromEntries(SECTION_KEYS.map((k) => [k, false])) as Collapsed,
});
