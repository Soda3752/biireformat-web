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
    /** 月結：本月該付（{期數} 帶月結月數） */
    monthly_due: string;
}

export type TemplateKey = 'four_plus' | 'two_three' | 'monthly_due';

/** 舊版「欠 1 期（月底前要催）」欄位（鍵名 one）的預設句；存的值還是它就換成新預設，她改過的則搬到新欄位。 */
const LEGACY_ONE_DEFAULT = '麻煩您方便時匯款，若已經匯款請忽略此訊息，謝謝！';

export const PERIODS_PLACEHOLDER = '{期數}';

export const DEFAULT_TEMPLATES: Readonly<MessageTemplates> = {
    version: 1,
    four_plus: '目前已累積 {期數} 期帳款未付，再麻煩儘速匯款，謝謝配合',
    two_three: '帳款已逾期，再麻煩儘早安排匯款，若已經匯款請忽略此訊息，謝謝！',
    monthly_due: '您好～您的月結帳款已累積 {期數} 個月未付，再麻煩儘速安排匯款，謝謝！',
};

export const defaultTemplates = (): MessageTemplates => ({...DEFAULT_TEMPLATES});

/** 壞掉、缺欄位或空白的部分一律用預設。 */
export function normalizeTemplates(raw: unknown): MessageTemplates {
    const out = defaultTemplates();
    if (raw && typeof raw === 'object') {
        const r = raw as Record<string, unknown>;
        for (const k of ['four_plus', 'two_three', 'monthly_due'] as const) {
            const v = r[k];
            if (typeof v === 'string' && v.trim()) out[k] = v.trim();
        }
        // 舊欄位 one 移轉：沒有新欄位時才看；還是舊預設句就用新預設
        const legacy = r.one;
        if (!(typeof r.monthly_due === 'string' && r.monthly_due.trim()) && typeof legacy === 'string' && legacy.trim() && legacy.trim() !== LEGACY_ONE_DEFAULT) {
            out.monthly_due = legacy.trim();
        }
    }
    return out;
}

/** 月結本月該付的客人固定用專用句；其餘依期數（1 期只顯示的不會產生訊息，但仍歸到 two_three 以下最輕的一句）。 */
export const templateKeyFor = (periods: number, monthlyDue = false): TemplateKey =>
    (monthlyDue ? 'monthly_due' : periods >= 4 ? 'four_plus' : periods >= 2 ? 'two_three' : 'monthly_due');

/** 這位客人的結尾句（已把 {期數} 換成實際期數）。 */
export function closingFor(periods: number, t: MessageTemplates, monthlyDue = false): string {
    return t[templateKeyFor(periods, monthlyDue)].split(PERIODS_PLACEHOLDER).join(String(periods));
}
