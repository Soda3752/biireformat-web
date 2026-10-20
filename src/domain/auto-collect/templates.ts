/**
 * 催款訊息範本：依欠款期數換結尾句（語氣由輕到重），使用者可自己改。
 * 結尾句裡的 {期數} 會換成這位客人的欠款期數。
 */

export interface MessageTemplates {
    version: 1;
    /** 欠 4 期以上 */
    four_plus: string;
    /** 欠 2～3 期 */
    two_three: string;
    /** 欠 1 期（月底前要催） */
    one: string;
}

export type TemplateKey = 'four_plus' | 'two_three' | 'one';

export const PERIODS_PLACEHOLDER = '{期數}';

export const DEFAULT_TEMPLATES: Readonly<MessageTemplates> = {
    version: 1,
    four_plus: '目前已累積 {期數} 期帳款未付，再麻煩儘速匯款，謝謝配合',
    two_three: '帳款已逾期，再麻煩儘早安排匯款，若已經匯款請忽略此訊息，謝謝！',
    one: '麻煩您方便時匯款，若已經匯款請忽略此訊息，謝謝！',
};

export const defaultTemplates = (): MessageTemplates => ({...DEFAULT_TEMPLATES});

/** 壞掉、缺欄位或空白的部分一律用預設。 */
export function normalizeTemplates(raw: unknown): MessageTemplates {
    const out = defaultTemplates();
    if (raw && typeof raw === 'object') {
        for (const k of ['four_plus', 'two_three', 'one'] as const) {
            const v = (raw as Record<string, unknown>)[k];
            if (typeof v === 'string' && v.trim()) out[k] = v.trim();
        }
    }
    return out;
}

export const templateKeyFor = (periods: number): TemplateKey => (periods >= 4 ? 'four_plus' : periods >= 2 ? 'two_three' : 'one');

/** 這位客人的結尾句（已把 {期數} 換成實際期數）。 */
export function closingFor(periods: number, t: MessageTemplates): string {
    return t[templateKeyFor(periods)].split(PERIODS_PLACEHOLDER).join(String(periods));
}
