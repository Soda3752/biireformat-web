/**
 * 「自動催款」：記住的各月明細與跨期欠款計算（對應獨立工具 auto_collect/ledger.py）。
 *
 * 規則：
 * - 一期的欠款 = 帳單金額 − 實收金額；實收空白＝整期沒付。
 * - 差額 ≤ 15 元視為銀行手續費，算付清。
 * - 多付的不拿去抵其他期，只在備註提醒。
 * - 同客人（依代號）跨好幾個月、好幾期的欠款加總；店名用最新月份的。
 * - 依欠幾期分組：4 期以上／3 期／2 期／月結：本月該付／1 期（只顯示、不催、不算進總額）。月結客人同月上下期合算 1 期；手動加的欠款不算期數。
 * - 月結客人「這個月還沒結完」（該月下期被略過，或就是今天所在月份且下期尚未記住）的期別＝未到期：不算欠款、不算期數、不進訊息與下載檔，只在畫面標示。
 * - 月結客人的訊息、總表與下載檔期別只寫月份（「8月」），同月上下期合成一行、金額相加；手動加的欠款各自一行。
 * - 月結客人只欠「上個月」（今天所在月份的前一個月）1 期、沒有更早的欠款＝「月結：本月該付」，列入要催的名單；欠更早的月份就照期數分組。
 * - 還沒結帳的期別（整期沒有人付款，或下期和上個月下期幾乎一樣＝複製來的佔位）略過，不算欠款。
 */

import {
    type Half,
    NotStatementError,
    type Periods,
    type StatementRecord,
    needsConfirm,
    parseStatement,
    unmatchedReceipts,
} from './statement-parser';
import {closingFor, DEFAULT_TEMPLATES, type MessageTemplates} from './templates';

export const FEE = 15;
export const KEEP = 'keep';
const HALF_LABEL: Record<Half, string> = {上: '上期', 下: '下期'};

export interface MonthEntry {
    source: string;
    file_mtime: string;
    imported_at: string;
    periods: Periods;
}

export interface History {
    version: 1;
    months: Record<string, MonthEntry>;
}

export interface DebtItem {
    period: string;
    month_key: string;
    half: Half;
    amt: number;
    paid: number;
    owed: number;
    partial: boolean;
    date_text: string;
    memo: string;
    /** 手動加的欠款才有（id 對應 ManualDebt）。 */
    manual?: {id: string; reason: string};
}

/** 手動「已結清」：記下客人＋當時顯示的那幾期（期別鍵 = `${月份鍵}|${上/下}`），之後新的月份不受影響。 */
export interface Settlement {
    id: string;
    /** 客人鍵：代號，沒有代號用店名。 */
    customer: string;
    code: string;
    name: string;
    keys: string[];
    /** 同時標掉的手動欠款 id。 */
    manual_ids: string[];
    reason: string;
    marked_at: string;
}

/** 手動加的一筆欠款（例如少付 64 元）：掛在客人下面，期別不一定要在記住的月份裡。 */
export interface ManualDebt {
    id: string;
    customer: string;
    code: string;
    name: string;
    month_key: string;
    half: Half;
    amount: number;
    reason: string;
    added_at: string;
}

/** 舊版「有收到款就算付清」設定（已拿掉；載入時轉成一般已結清標記）。 */
export interface PaidIfReceived {
    code: string;
    name: string;
    reason: string;
    marked_at: string;
}

export interface Marks {
    version: 1;
    settled: Settlement[];
    manual_debts: ManualDebt[];
    /** 手動設定的月結：客人鍵 → true（設為月結）／false（取消月結，含取消自動認到的）；沒有的照自動判斷。 */
    monthly: Record<string, boolean>;
    /** 手動改過的催款訊息：客人鍵 → 她改的整段文字，加上儲存當時自動產生的版本（之後比對欠款是否變動）。 */
    edits: Record<string, MessageEdit>;
}

export interface MessageEdit {
    text: string;
    /** 儲存當時自動產生的訊息。 */
    base: string;
    edited_at: string;
}

export type LegacyRules = Record<string, PaidIfReceived>;

/** 被排除（不催）的欠款：依已結清標記分組（legacy 只在轉換舊設定時出現）。 */
export interface SettledGroup {
    id: string;
    kind: 'manual' | 'legacy';
    customer: string;
    code: string;
    name: string;
    items: DebtItem[];
    /** 顯示用的期別（見 displayItems）。 */
    lines: DebtItem[];
    total: number;
    reason: string;
    marked_at: string;
}

/** 欠款期數分組：4 期以上／3 期／2 期／月結本月該付／1 期只顯示（含 0 期＝只有手動欠款）。 */
export type DebtGroup = 'four' | 'three' | 'two' | 'monthly_due' | 'one';

export const GROUPS: ReadonlyArray<{key: DebtGroup; label: string}> = [
    {key: 'four', label: '欠 4 期以上'},
    {key: 'three', label: '欠 3 期'},
    {key: 'two', label: '欠 2 期'},
    {key: 'monthly_due', label: '月結：本月該付'},
    {key: 'one', label: '欠 1 期（只顯示，不催款）'},
];

export interface Debtor {
    code: string;
    name: string;
    items: DebtItem[];
    /** 顯示用的期別（訊息、總表、下載檔）：月結客人寫月份、同月上下期合成一行；一般客人同 items。 */
    lines: DebtItem[];
    notes: string[];
    total: number;
    /** 實際使用的訊息：有手動改過就是她改的版本，否則是自動產生的。畫面、複製、下載都用這個。 */
    message: string;
    /** 目前自動產生的訊息（不受手動修改影響）。 */
    auto_message: string;
    /** 手動修改紀錄；stale＝現在自動產生的內容和她編輯當時不同（欠款有變動）。 */
    edit: (MessageEdit & {stale: boolean}) | null;
    /** 欠款期數：月結客人同月上下期合成 1 期，手動加的欠款不算。 */
    periods: number;
    group: DebtGroup;
    /** 目前是否當月結客人（手動設定優先，否則看自動判斷）。 */
    monthly: boolean;
    /** 備註裡有例行標記「月」，自動認成月結。 */
    monthly_auto: boolean;
    /** 月結客人還沒到期的期別（這個月還沒結完）：不算欠款、不算期數。 */
    pending: DebtItem[];
    pending_total: number;
}

/** 略過的期別：整期沒有人付款（no_payment），或下期和上個月下期幾乎一樣（same_as_prev，ref＝比對的月份鍵）。 */
export interface SkippedPeriod {
    half: Half;
    reason: 'no_payment' | 'same_as_prev';
    ref?: string;
}

export interface MonthInfo {
    key: string;
    label: string;
    source: string;
    file_mtime: string;
    imported_at: string;
    count: number;
    /** 還沒結帳而略過、不計入欠款的期別。 */
    skipped: SkippedPeriod[];
}

export interface CollectResult {
    /** 所有有欠款的客人（含欠 1 期的），照分組、需確認、代號排序。 */
    customers: Debtor[];
    months: MonthInfo[];
    /** customers 的欠款合計（含欠 1 期的）。畫面與下載用的是下面的 chase_total。 */
    grand_total: number;
    /** 要催款的客人（欠 2 期以上，或月結本月該付）與合計；其餘欠 1 期的在 light，只顯示、不算進總額。 */
    chase: Debtor[];
    chase_total: number;
    light: Debtor[];
    light_total: number;
    settled: SettledGroup[];
    settled_total: number;
}

export const LEGACY_REASON = '回饋金自扣';

export interface ImportFile {
    name: string;
    /** 檔案修改時間（毫秒，File.lastModified）。 */
    mtime: number;
    data: ArrayBuffer | Uint8Array;
}

export interface ImportReportRow {
    name: string;
    month_key?: string;
    status: 'used' | 'skipped' | 'error';
    reason: string;
}

export interface MonthSummary {
    upper: number;
    lower: number;
    paid: number;
    debtors: number;
    owed_total: number;
}

export interface ConflictOption extends MonthSummary {
    /** 這批檔案的索引，或 KEEP（保留之前記住的）。 */
    value: number | typeof KEEP;
    kind: 'new' | 'old';
    name: string;
    file_mtime: string;
    imported_at?: string;
    same_as_old?: boolean;
}

export interface MonthConflict {
    month_key: string;
    label: string;
    options: ConflictOption[];
}

export type Choices = Record<string, number | typeof KEEP>;

export const emptyHistory = (): History => ({version: 1, months: {}});

export const emptyMarks = (): Marks => ({version: 1, settled: [], manual_debts: [], monthly: {}, edits: {}});

/** 讀進來的標記（可能是舊資料、缺欄位）整理成完整的 Marks；讀不懂的部分視為空。舊的 paid_if_received 不在這裡處理。 */
export function normalizeMarks(raw: unknown): Marks {
    const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>;
    const str = (v: unknown): string => (typeof v === 'string' ? v : '');
    const strs = (v: unknown): string[] => (Array.isArray(v) ? v.filter((k): k is string => typeof k === 'string') : []);
    const settled = (Array.isArray(r.settled) ? r.settled : []).flatMap((x): Settlement[] => {
        const o = (x && typeof x === 'object' ? x : {}) as Record<string, unknown>;
        const keys = strs(o.keys);
        const manual_ids = strs(o.manual_ids);
        if (!str(o.customer) || !(keys.length || manual_ids.length)) return [];
        return [{id: str(o.id) || `${str(o.customer)}|${str(o.marked_at)}|${[...keys, ...manual_ids].join(',')}`, customer: str(o.customer),
            code: str(o.code), name: str(o.name), keys, manual_ids, reason: str(o.reason), marked_at: str(o.marked_at)}];
    });
    const manual_debts = (Array.isArray(r.manual_debts) ? r.manual_debts : []).flatMap((x): ManualDebt[] => {
        const o = (x && typeof x === 'object' ? x : {}) as Record<string, unknown>;
        const amount = typeof o.amount === 'number' ? o.amount : NaN;
        const half = o.half === '上' || o.half === '下' ? o.half : null;
        if (!str(o.customer) || !/^\d{4}-\d{2}$/.test(str(o.month_key)) || !half || !Number.isInteger(amount) || amount <= 0) return [];
        return [{id: str(o.id) || `${str(o.customer)}|${str(o.added_at)}|${str(o.month_key)}|${half}|${amount}`, customer: str(o.customer),
            code: str(o.code), name: str(o.name), month_key: str(o.month_key), half, amount, reason: str(o.reason), added_at: str(o.added_at)}];
    });
    const monthly: Record<string, boolean> = {};
    const mraw = r.monthly && typeof r.monthly === 'object' && !Array.isArray(r.monthly) ? r.monthly as Record<string, unknown> : {};
    for (const [k, v] of Object.entries(mraw)) if (typeof v === 'boolean') monthly[k] = v;
    const edits: Record<string, MessageEdit> = {};
    const draw = r.edits && typeof r.edits === 'object' && !Array.isArray(r.edits) ? r.edits as Record<string, unknown> : {};
    for (const [k, v] of Object.entries(draw)) {
        const o = (v && typeof v === 'object' ? v : {}) as Record<string, unknown>;
        if (typeof o.text === 'string' && o.text.trim()) edits[k] = {text: o.text, base: str(o.base), edited_at: str(o.edited_at)};
    }
    return {version: 1, settled, manual_debts, monthly, edits};
}

/** 從讀進來的舊資料取出已拿掉的「有收到款就算付清」設定（沒有就回空）。 */
export function extractLegacyRules(raw: unknown): LegacyRules {
    const r = (raw && typeof raw === 'object' ? raw : {}) as Record<string, unknown>;
    const str = (v: unknown): string => (typeof v === 'string' ? v : '');
    const out: LegacyRules = {};
    const src = r.paid_if_received && typeof r.paid_if_received === 'object' ? r.paid_if_received as Record<string, unknown> : {};
    for (const [k, v] of Object.entries(src)) {
        const o = (v && typeof v === 'object' ? v : {}) as Record<string, unknown>;
        out[k] = {code: str(o.code), name: str(o.name), reason: str(o.reason), marked_at: str(o.marked_at)};
    }
    return out;
}

/**
 * 把舊設定「目前涵蓋的期別」轉成一般已結清標記（原因「回饋金自扣」），之後不再自動套用到新月份。
 * 目前沒有涵蓋任何期別的設定直接丟掉。
 */
export function migrateLegacyRules(history: History, marks: Marks, legacy: LegacyRules, now: Date = new Date()): Marks {
    if (!Object.keys(legacy).length) return marks;
    const added = compute(history, marks, legacy).settled
        .filter((g) => g.kind === 'legacy' && g.items.length)
        .map((g): Settlement => {
            const keys = g.items.map(periodKey);
            const at = g.marked_at || localIso(now);
            return {id: `${g.customer}|${at}|${keys.join(',')}`, customer: g.customer, code: g.code, name: g.name,
                keys, manual_ids: [], reason: LEGACY_REASON, marked_at: at};
        });
    return added.length ? {...marks, settled: [...marks.settled, ...added]} : marks;
}

export const customerKey = (c: {code: string; name: string}): string => c.code || c.name;
export const periodKey = (i: Pick<DebtItem, 'month_key' | 'half'>): string => `${i.month_key}|${i.half}`;

/** 把這位客人目前顯示的幾期（含手動加的欠款）標成已結清。 */
export function settleDebtor(marks: Marks, c: Pick<Debtor, 'code' | 'name' | 'items'>, reason: string, now: Date = new Date()): Marks {
    const keys = c.items.filter((i) => !i.manual).map(periodKey);
    const manual_ids = c.items.flatMap((i) => (i.manual ? [i.manual.id] : []));
    if (!keys.length && !manual_ids.length) return marks;
    const at = localIso(now);
    const customer = customerKey(c);
    const id = `${customer}|${at}|${[...keys, ...manual_ids].join(',')}`;
    return {...marks, settled: [...marks.settled, {id, customer, code: c.code, name: c.name, keys, manual_ids, reason: reason.trim(), marked_at: at}]};
}

/** 設定客人是否月結。和自動判斷一樣就移除手動設定（之後跟著自動判斷走）。 */
export function setMonthly(marks: Marks, c: Pick<Debtor, 'code' | 'name' | 'monthly_auto'>, monthly: boolean): Marks {
    const next = {...marks.monthly};
    const key = customerKey(c);
    if (monthly === c.monthly_auto) delete next[key];
    else next[key] = monthly;
    return {...marks, monthly: next};
}

const AMOUNT_RE = /(\d{1,3}(?:,\d{3})+|\d+)\s*元/;
const TOTAL_LINE_RE = /^(\s*合計\s*[：:]\s*)(\d[\d,]*)(\s*元.*)$/;

/** 一行「・…」的欠款金額：先拿掉括號內容（如「（帳單 7,871 元，已收 7,811 元）」），再取冒號後第一個「N 元」；沒冒號就取最後一個。 */
export function lineAmount(line: string): number | null {
    if (!/^\s*・/.test(line)) return null;
    const bare = line.replace(/（[^）]*）|\([^)]*\)/g, '');
    const colon = bare.search(/[：:]/);
    let m: RegExpMatchArray | null;
    if (colon >= 0) m = AMOUNT_RE.exec(bare.slice(colon + 1));
    else m = [...bare.matchAll(new RegExp(AMOUNT_RE.source, 'g'))].pop() ?? null;
    return m ? Number(m[1].replace(/,/g, '')) : null;
}

/** 訊息裡每一行金額的加總（只算最後一個「合計」行之前的行）；沒有任何金額行回傳 null。 */
export function sumMessageLines(text: string): number | null {
    const lines = text.split('\n');
    const totalAt = lines.map((l) => TOTAL_LINE_RE.test(l)).lastIndexOf(true);
    let sum: number | null = null;
    for (const l of totalAt >= 0 ? lines.slice(0, totalAt) : lines) {
        const a = lineAmount(l);
        if (a !== null) sum = (sum ?? 0) + a;
    }
    return sum;
}

/** 手動改過訊息後重算「合計：N 元」那行；沒有合計行不新增、沒有金額行不動。千分位沿用原本的寫法。 */
export function recalcMessageTotal(text: string): string {
    const sum = sumMessageLines(text);
    if (sum === null) return text;
    const lines = text.split('\n');
    const at = lines.map((l) => TOTAL_LINE_RE.test(l)).lastIndexOf(true);
    if (at < 0) return text;
    const m = TOTAL_LINE_RE.exec(lines[at])!;
    const plain = m[2].length > 3 && !m[2].includes(',');
    lines[at] = `${m[1]}${plain ? String(sum) : money(sum)}${m[3]}`;
    return lines.join('\n');
}

/** 儲存手動修改的訊息；空白或和自動產生的一樣就等於沒改（清掉修改紀錄）。只動訊息文字，不影響金額與分組。 */
export function setMessageEdit(marks: Marks, c: Pick<Debtor, 'code' | 'name' | 'auto_message'>, text: string, now: Date = new Date()): Marks {
    const clean = recalcMessageTotal(text.replace(/\r\n?/g, '\n').replace(/\s+$/, ''));
    if (!clean.trim() || clean === c.auto_message) return clearMessageEdit(marks, c);
    return {...marks, edits: {...marks.edits, [customerKey(c)]: {text: clean, base: c.auto_message, edited_at: localIso(now)}}};
}

/** 還原成自動產生的訊息。 */
export function clearMessageEdit(marks: Marks, c: Pick<Debtor, 'code' | 'name'>): Marks {
    const k = customerKey(c);
    if (!marks.edits?.[k]) return marks;
    const {[k]: _gone, ...rest} = marks.edits;
    return {...marks, edits: rest};
}

export function unsettle(marks: Marks, id: string): Marks {
    return {...marks, settled: marks.settled.filter((s) => s.id !== id)};
}

export interface ManualDebtInput {
    month_key: string;
    half: Half;
    amount: number;
    reason: string;
}

/** 檢查手動欠款的輸入；回傳錯誤訊息，沒問題回 null。 */
export function validateManualDebt(v: {month_key: string; half: string; amount: number}): string | null {
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(v.month_key)) return '請選擇期別（或輸入年月）';
    if (v.half !== '上' && v.half !== '下') return '請選擇上期或下期';
    if (!Number.isInteger(v.amount) || v.amount <= 0) return '金額要填正整數';
    return null;
}

/** 給客人加一筆手動欠款。 */
export function addManualDebt(marks: Marks, c: Pick<Debtor, 'code' | 'name'>, v: ManualDebtInput, now: Date = new Date()): Marks {
    const at = localIso(now);
    const customer = customerKey(c);
    let id = `${customer}|${at}|${v.month_key}|${v.half}|${v.amount}`;
    for (let n = 2; marks.manual_debts.some((d) => d.id === id); n++) id = `${customer}|${at}|${v.month_key}|${v.half}|${v.amount}|${n}`;
    const debt: ManualDebt = {id, customer, code: c.code, name: c.name, month_key: v.month_key, half: v.half,
        amount: v.amount, reason: v.reason.trim(), added_at: at};
    return {...marks, manual_debts: [...marks.manual_debts, debt]};
}

export function updateManualDebt(marks: Marks, id: string, v: ManualDebtInput): Marks {
    return {...marks, manual_debts: marks.manual_debts.map((d) => (d.id === id
        ? {...d, month_key: v.month_key, half: v.half, amount: v.amount, reason: v.reason.trim()} : d))};
}

export function removeManualDebt(marks: Marks, id: string): Marks {
    return {...marks, manual_debts: marks.manual_debts.filter((d) => d.id !== id)};
}

const pad = (n: number, w = 2): string => String(n).padStart(w, '0');

/** 本地時間 ISO（到秒，不帶時區），同 Python datetime.isoformat(timespec="seconds")。 */
export function localIso(d: Date): string {
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

export const money = (n: number): string => Math.round(n).toLocaleString('en-US');

/** 檔名只有「9月」：預設今年；若月份比本月晚超過 1 個月，當成去年（例：1 月匯入 12 月檔）。 */
export function guessYear(month: number, today: Date = new Date()): number {
    return month > today.getMonth() + 2 ? today.getFullYear() - 1 : today.getFullYear();
}

export const monthKey = (year: number, month: number): string => `${pad(year, 4)}-${pad(month)}`;

export const monthLabel = (key: string): string => `${+key.slice(0, 4)}年${+key.slice(5)}月`;

/** 鍵排序後的 JSON，用來判斷兩份明細內容是否一樣。 */
function stableJson(v: unknown): string {
    if (Array.isArray(v)) return `[${v.map(stableJson).join(',')}]`;
    if (v && typeof v === 'object') {
        return `{${Object.keys(v).sort().map((k) => `${JSON.stringify(k)}:${stableJson((v as Record<string, unknown>)[k])}`).join(',')}}`;
    }
    return JSON.stringify(v);
}

/** 一個月份明細的摘要，讓使用者在同月多份時判斷要用哪份。 */
export function monthSummary(periods: Periods): MonthSummary {
    const recs = (['上', '下'] as const).flatMap((h) => Object.values(periods[h] ?? {}));
    const res = compute({version: 1, months: {'0000-00': {source: '', file_mtime: '', imported_at: '', periods}}});
    return {
        upper: Object.keys(periods.上 ?? {}).length,
        lower: Object.keys(periods.下 ?? {}).length,
        paid: recs.filter((r) => r.paid).length,
        debtors: res.customers.length,
        owed_total: res.grand_total,
    };
}

/**
 * 匯入一批檔案（不改動傳入的 history，回傳新的）。
 *
 * 同一個月有兩份以上不同內容（同一批好幾份，或已記住該月）→ 不自動挑，要由使用者選：
 * choices = {月份鍵: 這批檔案的索引（用那份）或 KEEP（保留原本記住的）}。
 * 還有沒選的月份時整批都不匯入，回傳的 conflicts 列出要選的月份與各份的摘要。
 */
export async function importFiles(
    files: ReadonlyArray<ImportFile>,
    history: History,
    opts: {today?: Date; choices?: Choices; now?: Date} = {},
): Promise<{history: History; report: ImportReportRow[]; conflicts: MonthConflict[]}> {
    const choices = opts.choices ?? {};
    const report: ImportReportRow[] = [];
    const byMonth = new Map<string, Array<{i: number; f: ImportFile; periods: Periods}>>();

    for (const [i, f] of files.entries()) {
        let st;
        try {
            st = await parseStatement(f.data, f.name);
        } catch (e) {
            if (!(e instanceof NotStatementError)) throw e;
            report.push({name: f.name, status: 'error', reason: e.message});
            continue;
        }
        if (!st.month) {
            report.push({name: f.name, status: 'error', reason: '從檔名看不出是幾月（檔名要有「9月」這種字）'});
            continue;
        }
        if (!Object.keys(st.periods.上).length && !Object.keys(st.periods.下).length) {
            report.push({name: f.name, status: 'error', reason: '上期／下期都讀不到任何客人'});
            continue;
        }
        const key = monthKey(guessYear(st.month, opts.today), st.month);
        if (!byMonth.has(key)) byMonth.set(key, []);
        byMonth.get(key)!.push({i, f, periods: st.periods});
    }
    const keys = [...byMonth.keys()].sort();

    const conflicts: MonthConflict[] = [];
    const picks = new Map<string, number | typeof KEEP>();
    for (const key of keys) {
        const items = byMonth.get(key)!;
        items.sort((a, b) => b.f.mtime - a.f.mtime);
        const old = history.months[key];
        const contents = new Set(items.map((x) => stableJson(x.periods)));
        if (old) contents.add(stableJson(old.periods));
        if (contents.size === 1) { // 內容都一樣，不用問
            picks.set(key, old ? KEEP : items[0].i);
            continue;
        }
        const choice = choices[key];
        if ((choice === KEEP && old) || items.some((x) => x.i === choice)) {
            picks.set(key, choice);
            continue;
        }
        const oldJson = old ? stableJson(old.periods) : '';
        const options: ConflictOption[] = items.map(({i, f, periods}) => ({
            value: i, kind: 'new', name: f.name, file_mtime: localIso(new Date(f.mtime)),
            same_as_old: !!old && stableJson(periods) === oldJson,
            ...monthSummary(periods),
        }));
        if (old) {
            options.push({
                value: KEEP, kind: 'old', name: old.source, file_mtime: old.file_mtime,
                imported_at: old.imported_at, ...monthSummary(old.periods),
            });
        }
        conflicts.push({month_key: key, label: monthLabel(key), options});
    }
    if (conflicts.length) return {history, report, conflicts};

    const next: History = {version: 1, months: {...history.months}};
    const now = localIso(opts.now ?? new Date());
    for (const key of keys) {
        const items = byMonth.get(key)!;
        const m = +key.slice(5);
        const old = history.months[key];
        const pick = picks.get(key)!;
        if (pick === KEEP) {
            for (const {f} of items) {
                report.push({
                    name: f.name, month_key: key, status: 'skipped',
                    reason: choices[key] === KEEP
                        ? `${m} 月保留原本記住的「${old.source}」`
                        : `內容和已記住的「${old.source}」一樣，不用更新`,
                });
            }
            continue;
        }
        const {f, periods} = items.find((x) => x.i === pick)!;
        next.months[key] = {source: f.name, file_mtime: localIso(new Date(f.mtime)), imported_at: now, periods};
        const nu = Object.keys(periods.上).length;
        const nl = Object.keys(periods.下).length;
        let reason = `讀到 ${nu + nl} 筆（上期 ${nu}、下期 ${nl}）`;
        if (old) reason += `；取代原本記住的「${old.source}」`;
        report.push({name: f.name, month_key: key, status: 'used', reason});
        for (const {i, f: other} of items) {
            if (i === pick) continue;
            const why = key in choices ? `你選了用「${f.name}」` : `內容和「${f.name}」一樣`;
            report.push({name: other.name, month_key: key, status: 'skipped', reason: `同樣是 ${m} 月，${why}`});
        }
    }
    return {history: next, report, conflicts: []};
}

export function removeMonth(history: History, key: string): History {
    const months = {...history.months};
    delete months[key];
    return {version: 1, months};
}

// ---------- 欠款計算

export function periodLabel(key: string, half: Half, multiYear: boolean): string {
    const y = +key.slice(0, 4);
    const m = +key.slice(5);
    return multiYear ? `${y}年${m}月${HALF_LABEL[half]}` : `${m}月${HALF_LABEL[half]}`;
}

/** 對帳 2.0 例行寫在備註欄的標記，不用另外提醒。 */
const ROUTINE_MEMO = new Set(['月', '扣手續費15', '2筆']);

/**
 * 備註欄裡要使用者先確認的部分（以「；」分段）：去掉例行標記、已由「金額對不上」規則列出的那段，
 * 以及已列「需確認」時的「待確認」；剩下的原文用「；」接回，都沒有就回空字串。
 */
export function memoToConfirm(memo: string, confirmNoted = false): string {
    return memo.split('；').map((x) => x.trim())
        .filter((x) => x && !ROUTINE_MEMO.has(x) && !unmatchedReceipts({date_text: '', memo: x}).length
            && !(confirmNoted && x === '待確認'))
        .join('；');
}

/** 兩個月的下期中，兩邊都有的客人金額相同的比例達此值，就判定新的那份下期是複製上個月的佔位（還沒結帳）。 */
export const UNSETTLED_SAME_RATIO = 0.9;
/** 樣本少於此數（兩邊都有的客人數、或整期的筆數）不自動判斷。 */
export const UNSETTLED_MIN = 3;

const recordsOf = (e: MonthEntry | undefined, half: Half): StatementRecord[] => Object.values(e?.periods?.[half] ?? {});

/**
 * 還沒結帳的期別（對應對帳 2.0 的 isLowerUnsettled，另加「整期沒有人付款」）：
 * - 整期（至少 UNSETTLED_MIN 筆有金額）沒有任何一筆有實收、收款日、實收欄文字 → no_payment
 * - 下期：與上一個記住的月份的下期，兩邊都有的客人（至少 UNSETTLED_MIN 位）金額相同比例 ≥ UNSETTLED_SAME_RATIO → same_as_prev
 * 回傳 `${月份鍵}|${上/下}` → 原因。
 */
export function unsettledPeriods(history: History): Map<string, SkippedPeriod> {
    const keys = Object.keys(history.months).sort();
    const out = new Map<string, SkippedPeriod>();
    keys.forEach((k, i) => {
        const entry = history.months[k];
        for (const half of ['上', '下'] as const) {
            const recs = recordsOf(entry, half).filter((r) => r.amt > 0);
            if (recs.length >= UNSETTLED_MIN && !recs.some((r) => r.paid !== null || r.paid_text || r.date_text)) {
                out.set(`${k}|${half}`, {half, reason: 'no_payment'});
                continue;
            }
            const prevKey = keys[i - 1];
            if (half !== '下' || !prevKey) continue;
            const prev = history.months[prevKey]?.periods?.下 ?? {};
            let common = 0;
            let same = 0;
            for (const [rk, rec] of Object.entries(entry.periods?.下 ?? {})) {
                const b = prev[rk];
                if (!b) continue;
                common++;
                if (b.amt === rec.amt) same++;
            }
            if (common >= UNSETTLED_MIN && same >= common * UNSETTLED_SAME_RATIO) {
                out.set(`${k}|${half}`, {half, reason: 'same_as_prev', ref: prevKey});
            }
        }
    });
    return out;
}

/**
 * 顯示用的期別：月結客人只寫月份（「8月」），同月的上下期合成一行（帳單、已收、欠款相加，有收到款就算部分付款）；
 * 手動加的欠款各自一行、一樣只寫月份。一般客人原樣回傳。
 */
export function displayItems(items: DebtItem[], monthly: boolean): DebtItem[] {
    if (!monthly) return items;
    const out: DebtItem[] = [];
    const byMonth = new Map<string, DebtItem>();
    const join = (a: string, b: string): string => [...new Set([a, b].filter(Boolean))].join('；');
    for (const i of items) {
        const period = i.period.replace(/[上下]期$/, '');
        const same = i.manual ? undefined : byMonth.get(i.month_key);
        if (!same) {
            const line = {...i, period};
            if (!i.manual) byMonth.set(i.month_key, line);
            out.push(line);
            continue;
        }
        same.amt += i.amt;
        same.paid += i.paid;
        same.owed += i.owed;
        same.partial = same.paid > 0;
        same.date_text = join(same.date_text, i.date_text);
        same.memo = join(same.memo, i.memo);
    }
    return out;
}

/** 欠款明細裡各期別的金額（月結寫法：期別去掉上下期、同期相加）＋合計行，用來判斷改寫法前後欠款是否一樣。 */
function monthlyDebtSig(msg: string): string {
    const sums = new Map<string, number>();
    const lines = messageBody(msg).split('\n');
    for (const l of lines) {
        const a = lineAmount(l);
        if (a === null) continue;
        const label = l.replace(/^\s*・/, '').split(/[：:]/)[0].trim().replace(/^((?:\d+年)?\d+月)[上下]期/, '$1');
        sums.set(label, (sums.get(label) ?? 0) + a);
    }
    const total = lines.find((l) => TOTAL_LINE_RE.test(l)) ?? '';
    return JSON.stringify([[...sums].sort(), total.replace(/\s/g, '')]);
}

const GROUP_RANK: Record<DebtGroup, number> = {four: 0, three: 1, two: 2, monthly_due: 3, one: 4};
const groupOf = (periods: number, monthlyDue: boolean): DebtGroup =>
    (periods >= 4 ? 'four' : periods === 3 ? 'three' : periods === 2 ? 'two' : monthlyDue ? 'monthly_due' : 'one');

/**
 * 「這個月還沒結完」的月份：下期被判定未結帳而略過；或就是今天所在的月份且下期尚未記住。
 * 月結客人這些月份的期別算未到期。
 */
function unfinishedMonths(history: History, unsettledMap: Map<string, SkippedPeriod>, today: Date): Set<string> {
    const nowKey = monthKey(today.getFullYear(), today.getMonth() + 1);
    const out = new Set<string>();
    for (const k of Object.keys(history.months)) {
        const lowerCount = Object.keys(history.months[k].periods?.下 ?? {}).length;
        if (unsettledMap.has(`${k}|下`) || (k === nowKey && !lowerCount)) out.add(k);
    }
    return out;
}
/** 備註有一段剛好是「月」＝對帳 2.0 認成月結客人。 */
const isMonthlyMemo = (memo: string | undefined): boolean => (memo ?? '').split('；').some((x) => x.trim() === '月');

export function compute(history: History, marks: Marks = emptyMarks(), legacy: LegacyRules = {}, today: Date = new Date(),
    templates: MessageTemplates = DEFAULT_TEMPLATES): CollectResult {
    const keys = Object.keys(history.months).sort();
    const unsettledMap = unsettledPeriods(history);
    const unfinished = unfinishedMonths(history, unsettledMap, today);
    // 上個月（1 月時是去年 12 月）：月結客人只欠這個月＝本月該付
    const prevKey = today.getMonth() === 0 ? monthKey(today.getFullYear() - 1, 12) : monthKey(today.getFullYear(), today.getMonth());
    const multiYear = new Set([...keys, ...marks.manual_debts.map((d) => d.month_key)].map((k) => k.slice(0, 4))).size > 1;
    const halves = ['上', '下'] as const;

    // 沒代號的列：名稱只對到一個代號就併過去
    const nameToCode = new Map<string, Set<string>>();
    for (const k of keys) {
        for (const half of halves) {
            for (const rec of Object.values(history.months[k].periods[half] ?? {})) {
                if (!rec.code) continue;
                if (!nameToCode.has(rec.name)) nameToCode.set(rec.name, new Set());
                nameToCode.get(rec.name)!.add(rec.code);
            }
        }
    }

    type Cust = Pick<Debtor, 'code' | 'name' | 'items' | 'notes' | 'monthly_auto'> & {noteMonths: string[]};
    const custs = new Map<string, Cust>();
    // 被排除的欠款：客人鍵 → {手動：標記 id → 期別；自動：期別}
    const settledBy = new Map<string, {manual: Map<string, DebtItem[]>; auto: DebtItem[]}>();
    for (const k of keys) {
        for (const half of halves) {
            for (const [rkey, rec] of Object.entries(history.months[k].periods[half] ?? {})) {
                let ckey = rkey;
                if (!rec.code) {
                    const hit = nameToCode.get(rec.name) ?? new Set<string>();
                    ckey = hit.size === 1 ? [...hit][0] : rkey;
                }
                if (!custs.has(ckey)) {
                    custs.set(ckey, {code: rec.code || (/^\d+$/.test(ckey) ? ckey : ''), name: rec.name, items: [], notes: [], noteMonths: [], monthly_auto: false});
                }
                const c = custs.get(ckey)!;
                const note = (text: string): void => {
                    c.notes.push(text);
                    c.noteMonths.push(k);
                };
                c.name = rec.name; // 以最新月份的店名為準
                if (isMonthlyMemo(rec.memo)) c.monthly_auto = true; // 任一已記住月份有「月」就算（包含略過的期別）
                if (unsettledMap.has(`${k}|${half}`)) continue; // 還沒結帳的期別不算欠款
                const {amt, paid} = rec;
                if (amt <= 0) continue;
                const label = periodLabel(k, half, multiYear);
                const short = amt - (paid ?? 0);
                const owed = short > FEE ? short : 0;
                const um = unmatchedReceipts(rec);
                if (owed) {
                    const who = c.code || rec.name;
                    const item: DebtItem = {
                        period: label, month_key: k, half, amt, paid: paid ?? 0, owed,
                        partial: !!paid, date_text: rec.date_text, memo: rec.memo ?? '',
                    };
                    const pk = periodKey(item);
                    const manual = marks.settled.find((x) => x.customer === who && x.keys.includes(pk));
                    // 有收到款 = 實收有數字 >0，或有「金額對不上」的入帳
                    const auto = !manual && legacy[who] && ((paid ?? 0) > 0 || um.length > 0);
                    if (manual || auto) {
                        if (!settledBy.has(who)) settledBy.set(who, {manual: new Map(), auto: []});
                        const sb = settledBy.get(who)!;
                        if (manual) sb.manual.set(manual.id, [...(sb.manual.get(manual.id) ?? []), item]);
                        else sb.auto.push(item);
                        continue; // 不算欠款，也不列這期的注意事項
                    }
                }
                if (um.length) {
                    const lst = um.map(([d, a]) => `${d} ${money(a)}`).join('、');
                    const tot = um.length > 1 ? `，合計 ${money(um.reduce((t, [, a]) => t + a, 0))} 元` : ' 元';
                    note(`${label} 有 ${um.length} 筆入帳金額對不上、沒算進已收（${lst}${tot}），催款前請先確認`);
                }
                if (paid !== null && short < 0) note(`${label} 多收 ${money(-short)} 元（未拿來抵其他期）`);
                const confirmNoted = !!(owed && paid && needsConfirm(rec));
                if (confirmNoted) note(`${label} 的收款對帳 2.0 標「需確認」`);
                if (rec.paid_text) note(`${label} 實收欄寫「${rec.paid_text}」，不是金額，催款前請先確認`);
                const memo = memoToConfirm(rec.memo ?? '', confirmNoted);
                if (memo) note(`${label} 備註：${memo}，催款前請先確認`);
                if (owed) {
                    c.items.push({
                        period: label, month_key: k, half, amt, paid: paid ?? 0, owed,
                        partial: !!paid, date_text: rec.date_text, memo: rec.memo ?? '',
                    });
                }
            }
        }
    }

    // 手動加的欠款：掛在名單上（記住的月份裡出現過）的客人；被已結清標記帶走的歸到已結清區
    for (const d of marks.manual_debts) {
        const c = [...custs.values()].find((x) => (x.code || x.name) === d.customer);
        if (!c) continue;
        const item: DebtItem = {
            period: periodLabel(d.month_key, d.half, multiYear), month_key: d.month_key, half: d.half,
            amt: d.amount, paid: 0, owed: d.amount, partial: false, date_text: '', memo: '',
            manual: {id: d.id, reason: d.reason},
        };
        const st = marks.settled.find((x) => x.customer === d.customer && x.manual_ids.includes(d.id));
        if (st) {
            if (!settledBy.has(d.customer)) settledBy.set(d.customer, {manual: new Map(), auto: []});
            const sb = settledBy.get(d.customer)!;
            sb.manual.set(st.id, [...(sb.manual.get(st.id) ?? []), item]);
        } else {
            c.items.push(item);
        }
    }

    const customers: Debtor[] = [];
    for (const c0 of custs.values()) {
        const {noteMonths, ...c} = c0;
        if (!c.items.length) continue;
        const monthly = marks.monthly?.[customerKey(c)] ?? c.monthly_auto;
        // 月結客人還沒結完的月份：期別算未到期（不算欠款、不算期數、不進訊息），相關注意事項也一併拿掉
        const isPending = (i: DebtItem): boolean => monthly && !i.manual && unfinished.has(i.month_key);
        const pending = c.items.filter(isPending);
        if (pending.length) {
            c.items = c.items.filter((i) => !isPending(i));
            c.notes = c.notes.filter((_, n) => !unfinished.has(noteMonths[n]));
        }
        if (!c.items.length) continue;
        // 依期別時間順序（同一期原本的欠款在前、手動加的在後）
        c.items = c.items.map((it, n) => ({it, n}))
            .sort((a, b) => (a.it.month_key + a.it.half < b.it.month_key + b.it.half ? -1
                : a.it.month_key + a.it.half > b.it.month_key + b.it.half ? 1 : a.n - b.n)).map((x) => x.it);
        const total = c.items.reduce((t, i) => t + i.owed, 0);
        // 期數：每個有欠款的期別 1 期，月結客人同月上下期合成 1 期；手動加的欠款不算
        const periods = new Set(c.items.filter((i) => !i.manual).map((i) => (monthly ? i.month_key : periodKey(i)))).size;
        const monthlyDue = monthly && periods === 1 && c.items.every((i) => i.month_key === prevKey);
        const d: Debtor = {...c, lines: displayItems(c.items, monthly), total, message: '', auto_message: '', edit: null, periods,
            group: groupOf(periods, monthlyDue), monthly,
            pending, pending_total: pending.reduce((t, i) => t + i.owed, 0)};
        d.auto_message = buildMessage(d, templates);
        const edit = marks.edits?.[customerKey(c)];
        d.message = edit ? edit.text : d.auto_message;
        // 只比對欠款明細（到「合計」那行為止）；結尾句因範本改變不算欠款變動。
        // 月結客人只比各期別金額：期別改寫成月份、同月上下期合併（舊版編輯當時是「8月上期」「8月下期」）不算變動
        const stale = !!edit && messageBody(edit.base) !== messageBody(d.auto_message)
            && !(monthly && monthlyDebtSig(edit.base) === monthlyDebtSig(d.auto_message));
        d.edit = edit ? {...edit, stale} : null;
        customers.push(d);
    }
    const isCode = (x: string): boolean => /^\d+$/.test(x);
    // 欠越多期越上面；同組內有注意事項（需要先確認）的排最上面，其餘照代號
    customers.sort((a, b) => {
        if (a.group !== b.group) return GROUP_RANK[a.group] - GROUP_RANK[b.group];
        if (!a.notes.length !== !b.notes.length) return a.notes.length ? -1 : 1;
        if (isCode(a.code) !== isCode(b.code)) return isCode(a.code) ? -1 : 1;
        if (isCode(a.code)) return +a.code - +b.code;
        return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
    });
    const months = keys.map((k) => {
        const m = history.months[k];
        return {
            key: k, label: monthLabel(k), source: m.source, file_mtime: m.file_mtime, imported_at: m.imported_at,
            count: Object.values(m.periods).reduce((t, v) => t + Object.keys(v).length, 0),
            skipped: (['上', '下'] as const).flatMap((h) => unsettledMap.get(`${k}|${h}`) ?? []),
        };
    });

    const settled: SettledGroup[] = [];
    const sumOwed = (items: DebtItem[]): number => items.reduce((t, i) => t + i.owed, 0);
    for (const [who, sb] of settledBy) {
        const base = [...custs.values()].find((c) => (c.code || c.name) === who)!;
        const monthly = marks.monthly?.[customerKey(base)] ?? base.monthly_auto;
        for (const [id, items] of sb.manual) {
            const m = marks.settled.find((x) => x.id === id)!;
            settled.push({id, kind: 'manual', customer: who, code: base.code, name: base.name, items,
                lines: displayItems(items, monthly), total: sumOwed(items),
                reason: m.reason, marked_at: m.marked_at});
        }
    }
    for (const [who, p] of Object.entries(legacy)) {
        const base = [...custs.values()].find((c) => (c.code || c.name) === who);
        const items = settledBy.get(who)?.auto ?? [];
        const monthly = base ? marks.monthly?.[customerKey(base)] ?? base.monthly_auto : false;
        settled.push({id: `legacy:${who}`, kind: 'legacy', customer: who, code: base?.code ?? p.code, name: base?.name ?? p.name, items,
            lines: displayItems(items, monthly),
            total: sumOwed(items), reason: LEGACY_REASON, marked_at: p.marked_at});
    }
    settled.sort((a, b) => {
        if (isCode(a.code) !== isCode(b.code)) return isCode(a.code) ? -1 : 1;
        if (isCode(a.code) && +a.code !== +b.code) return +a.code - +b.code;
        return a.name < b.name ? -1 : a.name > b.name ? 1 : a.kind === b.kind ? 0 : a.kind === 'manual' ? -1 : 1;
    });
    const chase = customers.filter((c) => c.group !== 'one');
    const light = customers.filter((c) => c.group === 'one');
    const sumTotal = (list: Debtor[]): number => list.reduce((t, c) => t + c.total, 0);
    return {
        customers, months, grand_total: sumTotal(customers),
        chase, chase_total: sumTotal(chase), light, light_total: sumTotal(light),
        settled, settled_total: settled.reduce((t, g) => t + g.total, 0),
    };
}

/** 畫面上的篩選：關鍵字比對代號或店名；onlyNotes 時只留有注意事項的。 */
export function filterCustomers(customers: Debtor[], keyword: string, onlyNotes = false): Debtor[] {
    const kw = keyword.trim();
    return customers.filter((c) => (!onlyNotes || c.notes.length > 0)
        && (!kw || c.code.includes(kw) || c.name.includes(kw)));
}

/** 訊息裡欠款明細的部分（到「合計」那行為止），不含結尾句。 */
export const messageBody = (msg: string): string => /^[\s\S]*?\n合計：[^\n]*/.exec(msg)?.[0] ?? msg;

export function buildMessage(c: Pick<Debtor, 'name' | 'items' | 'total' | 'periods' | 'group'> & Partial<Pick<Debtor, 'lines'>>,
    templates: MessageTemplates = DEFAULT_TEMPLATES): string {
    const lines = [`${c.name} 您好～`, '提醒您以下帳款尚未收到：'];
    for (const i of c.lines ?? c.items) {
        let line = `・${i.period}${i.manual?.reason ? `（${i.manual.reason}）` : ''}：${money(i.owed)} 元`;
        if (i.partial) line += `（帳單 ${money(i.amt)} 元，已收 ${money(i.paid)} 元）`;
        lines.push(line);
    }
    lines.push(`合計：${money(c.total)} 元`, '', closingFor(c.periods, templates, c.group === 'monthly_due'));
    return lines.join('\n');
}

const itemLabel = (i: DebtItem): string => `${i.period}${i.manual?.reason ? `（${i.manual.reason}）` : ''}`;

export const periodsText = (c: Debtor): string => c.lines.map(itemLabel).join('、');

export const amountsText = (c: Debtor): string => c.lines.map((i) => `${itemLabel(i)} ${money(i.owed)}`).join('、');
