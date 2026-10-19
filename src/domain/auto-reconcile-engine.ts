/**
 * 對帳 2.0 自動對帳引擎（移植自使用者自寫的「中信入帳自動對帳」工具）
 *
 * 輸入：
 *  - 銀行入帳列（日期 / 摘要 / 存入金額 / 備註 / 註記）
 *  - 月結明細（每月一份 .xlsx，含「上期」「下期」工作表；B~D、H~J 為 代號/客戶名稱/金額，
 *    其後兩欄為 實收金額 / 收款日），可放多個月以對出付舊帳的客人
 *  - 末五碼：設定頁對照表（以「店家編號 = 明細代號」直接對）＋選用的末五碼資料 .xlsx（含「歷史 / 已確認 / 現付」標記）
 *
 * 比對：先依 已確認 → 歷史 → 末五碼 → 備註名稱 找出候選客戶，再用金額去湊期別組合
 * （連續 1~6 期、跳期 2~3 期、扣手續費 15、多店合併、同客戶分筆加總）。
 * 帳號/備註都沒有紀錄的，最後再用金額在尚未收款的期別裡推測客戶（單期或連續 2~4 期，扣 0/15/30）。
 *
 * 輸出：回填各月明細的「實收金額 / 收款日」，收款日右邊插入「備註」欄寫標記與使用者備註，並在最新月份檔加上「銀行對帳結果」工作表；
 * 下次放入回填後的檔案會依該表略過已處理的入帳（接續模式）。
 */

import type ExcelJS from 'exceljs';

import type {BankInfo} from '@/domain/models/bank-info';
import {findMatchesFromCandidates} from '@/domain/bank-match-service';

export const FEE = 15;
export const RESULT_SHEET = '銀行對帳結果';
/** 使用者確認「這是新客戶（清單沒有）」時的標記：照打的名字存，不做模糊比對。 */
export const NEW_CUSTOMER_FLAG = '已確認（新客戶）';
/** 現付店家（叫貨時直接付款，不歸任何期別）的標記。 */
export const CASH_FLAG = '現付（叫貨直接轉帳）';
export const isCashFlag = (flag: string): boolean => flag.includes('現付');
/** 不在明細的入帳：單次付款（記住帳號）、現金存款（自己存的，只算這一筆）。 */
export const ONE_OFF_FLAG = '單次付款';
export const DEPOSIT_FLAG = '現金存款';
export type SpecialKind = typeof ONE_OFF_FLAG | typeof DEPOSIT_FLAG;
export const specialOf = (flag: string): SpecialKind | undefined =>
    flag.includes(ONE_OFF_FLAG) ? ONE_OFF_FLAG : flag.includes(DEPOSIT_FLAG) ? DEPOSIT_FLAG : undefined;

export type AutoReconcileStatus = 'OK' | '需確認' | '金額不符' | '找不到' | '不在明細';

const COLORS: Record<AutoReconcileStatus, string> = {
    OK: 'FFE2EFDA',
    需確認: 'FFFFF2CC',
    金額不符: 'FFFFF2CC',
    找不到: 'FFF8CBAD',
    不在明細: 'FFDDEBF7',
};

export interface BankEntry {
    date: string;
    kind: string;
    bank: string;
    key: string;
    amt: number;
}

export interface PeriodCell {
    sheet: string;
    row: number;
    /** 客戶名稱欄。 */
    col: number;
    amt: number;
    /** 「金額」欄（回填的字級照它）。 */
    amtCol: number;
    /** 回填用的「實收金額」「收款日」欄。 */
    paidCol: number;
    dateCol: number;
    /** 明細上原本已填的實收金額（依金額推測時略過已收款的期別）。 */
    paid?: number;
}

export interface StmtCustomer {
    name: string;
    p: Record<string, PeriodCell>;
}

export type StmtData = Record<string, StmtCustomer>;

export interface AcctRow {
    name: string;
    reg: number | null;
    val: string;
    flag: string;
}

export interface MemoryEntry {
    name: string;
    flag: string;
    at?: string;
}

export interface ReconcileEntry extends BankEntry {
    idx: number;
    /** [客戶代號, 期別陣列]；期別空陣列表示只知道客戶、金額對不上。 */
    alloc: Array<[string, string[]]>;
    fee: number;
    status: AutoReconcileStatus;
    note: string;
    cand: string[];
    src: string;
    confirmed: boolean;
    who: string;
    custNames: string[];
    /** 使用者在本次對帳對「金額不符」填的原因備註（填了即視為已對上）。 */
    userNote?: string;
    /** 金額不符可回填的期別（該客戶在明細有的期別）。 */
    periodOptions?: string[];
    /** 金額不符要回填的期別（由舊到新）：使用者勾的，否則預設原本紅字提醒那一期。 */
    notePeriods?: string[];
    /** 多期回填時各期的實收金額（與 notePeriods 同序）；最後一期＝入帳金額減前面各期。 */
    noteShares?: number[];
    /** 勾的期別無法回填的原因（例：勾太多期，剩餘金額 ≤ 0）。 */
    noteError?: string;
    /**
     * 記為現付的帳號：ok＝當現付算已對上（不回填）；warn＝金額剛好等於某期明細，待使用者確認；
     * period＝使用者選這筆照月結回填。
     */
    cash?: 'ok' | 'warn' | 'period';
    /** 使用者確認為單次付款 / 現金存款：算已對上、不回填明細。 */
    special?: SpecialKind;
    /** 不在明細時帳號對應到的名稱（記單次付款時當客戶名稱）。 */
    nameHint?: string;
    /** 帳號/備註沒有紀錄、依金額推測出的候選（客戶期別）；只有一個時已放在 alloc。 */
    guesses?: GuessPick[];
    /** 候選太多沒列出的數量。 */
    guessMore?: number;
}

/** 依金額推測的一個候選：客戶代號、期別（由舊到新）、扣的手續費。 */
export interface GuessPick {
    code: string;
    run: string[];
    fee: number;
    label?: string;
}

export interface AutoReconcileResult {
    results: ReconcileEntry[];
    cust: Record<string, StmtCustomer & { alias: string[] }>;
    PERIODS: string[];
    skipped: number;
    names: string[];
}

// ---------- 小工具

const s = (v: unknown): string => (v === null || v === undefined ? '' : String(v).trim());

export const cellVal = (c: ExcelJS.Cell | undefined | null): unknown => {
    let v: unknown = c?.value;
    if (v && typeof v === 'object') {
        const o = v as unknown as Record<string, unknown>;
        if ('result' in o) v = o.result;
        else if (Array.isArray(o.richText)) v = (o.richText as Array<{ text: string }>).map((t) => t.text).join('');
        else if ('text' in o) v = o.text;
    }
    return v;
};

const norm = (x: unknown): string => s(x).replace(/[\s　_\-.．、]/g, '').toLowerCase();

function sim(a0: string, b0: string): number {
    const a = norm(a0);
    const b = norm(b0);
    if (!a || !b) return 0;
    if (a === b) return 1;
    if (a.includes(b) || b.includes(a)) return 0.9;
    const [sh, lo] = a.length <= b.length ? [a, b] : [b, a];
    if (sh.length >= 2) {
        let j = 0;
        for (const ch of lo) if (ch === sh[j]) j++;
        if (j === sh.length) return 0.85;
    }
    const A = new Set(a);
    const B = new Set(b);
    let inter = 0;
    A.forEach((ch) => {
        if (B.has(ch)) inter++;
    });
    return (0.8 * inter) / new Set([...A, ...B]).size;
}

const NUM = '[0-9０-９一二三四五六七八九十]';
const mkey = (k: string): string => k.replace(new RegExp(`${NUM}+月?(?=[月上下初底份前後])|${NUM}+月`, 'g'), '#');
const stripMemo = (k: string): string => k
    .replace(new RegExp(`${NUM}+[月∕/．.]?(份|底|初|前半|後半|後)?[上下]?(半)?(貨款|貨錢|下款|上款|款|麵包錢)?`, 'g'), '')
    .replace(/(貨款|貨錢|麵包錢|款項)/g, '')
    .trim();
const acct4 = (k: string): string | null => {
    const m = /\*\*(\d{4})\*/.exec(k);
    return m ? m[1] : null;
};
const CN = '零一二三四五六七八九十';

function mforms(m0: number): string[] {
    const m = ((m0 - 1) % 12 + 12) % 12 + 1;
    const fw = String(m).replace(/\d/g, (d) => '０１２３４５６７８９'[Number(d)]);
    return [String(m), fw, m <= 10 ? CN[m] : '十' + CN[m - 10]];
}

const md = (d: string): string => {
    const p = d.split('/');
    return `${+p[1]}/${+p[2]}`;
};
const dateKey = (d: string): number[] => d.split('/').map(Number);

function cmpArr(a: ReadonlyArray<number>, b: ReadonlyArray<number>): number {
    for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return a[i] - b[i];
    return 0;
}

function cmpDate(a: string, b: string): number {
    return cmpArr(dateKey(a), dateKey(b));
}

function combos<T>(arr: ReadonlyArray<T>, n: number, start = 0, acc: T[] = [], out: T[][] = []): T[][] {
    if (acc.length === n) {
        out.push([...acc]);
        return out;
    }
    for (let i = start; i < arr.length; i++) {
        acc.push(arr[i]);
        combos(arr, n, i + 1, acc, out);
        acc.pop();
    }
    return out;
}

export function sheetByName(wb: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
    return wb.worksheets.find((w) => w.name === name) ?? wb.worksheets.find((w) => w.name.trim() === name.trim());
}

export const isStatementWorkbook = (wb: ExcelJS.Workbook): boolean =>
    !!sheetByName(wb, '上期') && !!sheetByName(wb, '下期');

// ---------- 讀檔

/** 日期統一成 YYYY/M/D（xlsx 對帳單可能是 YYYY-MM-DD）。 */
const normalizeDate = (raw: string): string => {
    const m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/.exec(raw.trim());
    return m ? `${m[1]}/${+m[2]}/${+m[3]}` : raw.trim();
};

/**
 * 由 trans-record-reader 產出的 string[][] 取出入帳列（只取存入金額 > 0）。
 * 標題列需含 日期 / 摘要 / 存入金額 / 備註 / 註記；缺標題時退回標準 9 欄位置。
 */
export function readBankRows(rows: ReadonlyArray<ReadonlyArray<string>>): BankEntry[] {
    const STANDARD = ['日期', '摘要', '幣別', '支出金額', '存入金額', '餘額', '備註', '轉出入帳號', '註記'];
    const headerIdx = rows.findIndex((r) => r.some((x) => s(x) === '日期') && r.some((x) => s(x) === '存入金額'));
    const h = headerIdx >= 0 ? rows[headerIdx].map((x) => s(x)) : STANDARD;
    const ix = (n: string): number => {
        const i = h.indexOf(n);
        return i >= 0 ? i : STANDARD.indexOf(n);
    };
    return rows.slice(headerIdx + 1).map((r) => ({
        date: normalizeDate(s(r[ix('日期')])),
        kind: s(r[ix('摘要')]),
        bank: s(r[ix('備註')]),
        key: s(r[ix('註記')]),
        amt: parseInt(s(r[ix('存入金額')]).replace(/,/g, '') || '0', 10) || 0,
    })).filter((r) => r.amt > 0 && /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(r.date));
}

/** 明細沒有店家編號的客戶，用此前綴加客戶名稱當內部鍵（不會與數字編號衝突）。 */
export const NAME_KEY_PREFIX = '名:';
const isNameKey = (code: string): boolean => code.startsWith(NAME_KEY_PREFIX);

interface StmtBlock {
    /** 店家編號欄（0 表示沒有）。 */
    code: number;
    /** 標題列寫明了店家編號欄。 */
    codeLabeled: boolean;
    name: number;
    amt: number;
    paid: number;
    date: number;
}

/** 舊版面：A/G=店家編號、B/H=客戶名稱、C/I=金額、D/J=實收金額、E/K=收款日。 */
const LEGACY_BLOCKS: StmtBlock[] = [2, 8].map((bc) => ({code: bc - 1, codeLabeled: false, name: bc, amt: bc + 1, paid: bc + 2, date: bc + 3}));
const STMT_JUNK_NAMES = /^(小計|總計|合計|總營收)$/;

/** 依標題列（含「客戶名稱」的那列）找出左右區塊各欄位置；找不到標題列就用舊版面固定欄號。 */
function stmtLayout(ws: ExcelJS.Worksheet): { headerRow: number; blocks: StmtBlock[] } {
    for (let r = 1; r <= 5; r++) {
        const row = ws.getRow(r);
        const h: string[] = [];
        row.eachCell({includeEmpty: false}, (c, col) => {
            h[col] = s(cellVal(c));
        });
        const nameCols = h.map((x, i) => (x === '客戶名稱' ? i : 0)).filter((i) => i > 0);
        if (!nameCols.length) continue;
        const blocks = nameCols.map((nc, bi) => {
            const endCol = nameCols[bi + 1] ?? h.length;
            const find = (label: string, def: number): number => {
                for (let c = nc + 1; c < endCol; c++) if (h[c] === label) return c;
                return def;
            };
            const amt = find('金額', nc + 1);
            const paid = find('實收金額', amt + 1);
            const date = find('收款日', paid + 1);
            return {code: nc > 1 ? nc - 1 : 0, codeLabeled: /編號|代號|代碼/.test(h[nc - 1] ?? ''), name: nc, amt, paid, date};
        });
        return {headerRow: r, blocks};
    }
    return {headerRow: 2, blocks: LEGACY_BLOCKS};
}

/**
 * 讀一個月的明細。有店家編號的區塊只收編號為數字的列（舊行為）；
 * 沒有店家編號的區塊改以 NAME_KEY_PREFIX + 客戶名稱 當鍵，由 reconcile 依名稱併到其他月份的同名客戶。
 */
export function readStmt(wb: ExcelJS.Workbook, month: number): StmtData {
    const out: StmtData = {};
    for (const [half, sname] of [['上', '上期'], ['下', '下期']] as const) {
        const ws = sheetByName(wb, sname);
        if (!ws) throw new Error(`明細表找不到「${sname}」工作表`);
        const {headerRow, blocks} = stmtLayout(ws);
        const rows: Array<[ExcelJS.Row, number]> = [];
        ws.eachRow({includeEmpty: false}, (row, r) => {
            if (r > headerRow) rows.push([row, r]);
        });
        const codeOf = (b: StmtBlock, row: ExcelJS.Row): string => (b.code ? s(cellVal(row.getCell(b.code))) : '');
        const filled = (b: StmtBlock, row: ExcelJS.Row): boolean =>
            !!s(cellVal(row.getCell(b.name))) && typeof cellVal(row.getCell(b.amt)) === 'number';
        const coded = (b: StmtBlock, row: ExcelJS.Row): boolean => /^\d+$/.test(codeOf(b, row)) && filled(b, row);
        // 標題寫了店家編號，或過半的客戶列帶數字編號，就是舊版面（有編號），行為與以前相同；
        // 沒編號的版面在編號位置可能有零星數字（頁碼標記等），不能一有數字就當成有編號
        const count = (f: (b: StmtBlock, row: ExcelJS.Row) => boolean): number =>
            blocks.reduce((t, b) => t + rows.filter(([row]) => f(b, row)).length, 0);
        const hasCode = blocks.some((b) => b.codeLabeled) || count(coded) * 2 >= count(filled);
        for (const b of blocks) {
            for (const [row, r] of rows) {
                const name = s(cellVal(row.getCell(b.name)));
                const amtCell = row.getCell(b.amt);
                const amt = cellVal(amtCell);
                if (!name || typeof amt !== 'number') continue;
                let key: string;
                if (hasCode) {
                    if (!coded(b, row)) continue;
                    key = codeOf(b, row);
                } else {
                    // 沒編號的版面：略過小計/總計列與公式加總列（各區域合計）
                    if (STMT_JUNK_NAMES.test(name) || /^\d+$/.test(name) || amtCell.formula) continue;
                    key = NAME_KEY_PREFIX + name;
                }
                out[key] = out[key] || {name, p: {}};
                const paid = cellVal(row.getCell(b.paid));
                out[key].p[`${month}${half}`] = {
                    sheet: ws.name, row: r, col: b.name, amt: Math.round(amt), amtCol: b.amt, paidCol: b.paid, dateCol: b.date,
                    ...(typeof paid === 'number' && paid > 0 ? {paid} : {}),
                };
            }
        }
    }
    return out;
}

/** 兩份明細的下期中，兩邊都有的客戶金額相同的比例達此值，就判定新的那份下期還沒結帳（是複製上個月的佔位）。 */
export const UNSETTLED_SAME_RATIO = 0.9;
/** 兩邊都有的客戶少於此數時樣本太少，不自動判斷。 */
export const UNSETTLED_MIN_COMMON = 3;

/**
 * 判斷 cur 這份明細的「下期」是否還沒結帳：與上一個月份檔 prev 的下期幾乎一樣
 * （兩邊都有的客戶中，金額相同的比例 ≥ UNSETTLED_SAME_RATIO）。
 */
export function isLowerUnsettled(prev: StmtData, prevMonth: number, cur: StmtData, curMonth: number): boolean {
    let common = 0;
    let same = 0;
    for (const [key, v] of Object.entries(cur)) {
        const a = v.p[`${curMonth}下`];
        const b = prev[key]?.p[`${prevMonth}下`];
        if (!a || !b) continue;
        common++;
        if (a.amt === b.amt) same++;
    }
    return common >= UNSETTLED_MIN_COMMON && same >= common * UNSETTLED_SAME_RATIO;
}

/** 略過某月下期：拿掉該期的明細格（不產生期別、不比對、不回填）。 */
export function dropLowerHalf(data: StmtData, month: number): StmtData {
    const per = `${month}下`;
    const out: StmtData = {};
    for (const [key, v] of Object.entries(data)) {
        const p = {...v.p};
        delete p[per];
        if (Object.keys(p).length) out[key] = {...v, p};
    }
    return out;
}

export function readAcct(wb: ExcelJS.Workbook): AcctRow[] {
    const ws = wb.worksheets[0];
    const rows: AcctRow[] = [];
    if (!ws) return rows;
    ws.eachRow({includeEmpty: false}, (row) => {
        const nm = s(cellVal(row.getCell(1)));
        const reg = cellVal(row.getCell(2));
        const v = s(cellVal(row.getCell(3)));
        const flag = s(cellVal(row.getCell(4)));
        if (!nm || !v || nm === '店名') return;
        rows.push({
            name: nm,
            reg: typeof reg === 'number' ? reg : (/^\d$/.test(s(reg)) ? +s(reg) : null),
            val: v,
            flag,
        });
    });
    return rows;
}

export const doneKey = (b: Pick<BankEntry, 'date' | 'amt' | 'key'>): string => `${b.date}|${b.amt}|${b.key}`;

export function readDone(wb: ExcelJS.Workbook): Set<string> {
    const ws = sheetByName(wb, RESULT_SHEET);
    const done = new Set<string>();
    if (!ws) return done;
    ws.eachRow((row, r) => {
        if (r === 1) return;
        const d = s(cellVal(row.getCell(1)));
        const a = cellVal(row.getCell(3));
        const k = s(cellVal(row.getCell(5)));
        if (d && typeof a === 'number') done.add(`${d}|${a}|${k}`);
    });
    return done;
}

// ---------- 對帳

export interface ReconcileInput {
    bank: ReadonlyArray<BankEntry>;
    /** 由舊到新排序；skipLower＝該月下期還沒結帳，整個略過。 */
    stmts: ReadonlyArray<{ month: number; data: StmtData; skipLower?: boolean }>;
    acct: ReadonlyArray<AcctRow>;
    bankInfos: ReadonlyArray<BankInfo>;
    memory?: Record<string, MemoryEntry>;
    rowMemory?: Record<number, MemoryEntry>;
    /** 本次對帳的金額不符備註（以 idx 為鍵）；有備註的金額不符改算 OK。 */
    rowNotes?: Record<number, string>;
    /** 本次對帳金額不符的回填期別（以 idx 為鍵，可多期）；沒選用預設期別。 */
    rowNotePeriods?: Record<number, string[]>;
    /** 本次對帳記為現付、但金額等於某期明細的入帳：cash＝確認當現付；period＝這筆照月結回填。 */
    rowCashMode?: Record<number, 'cash' | 'period'>;
    /** 本次對帳使用者確認的金額推測（以 idx 為鍵）：照選的客戶期別回填。 */
    rowPick?: Record<number, GuessPick>;
    done?: Set<string>;
}

const add = (o: Record<string, Set<string>>, k: string, v: string) => {
    (o[k] = o[k] || new Set()).add(v);
};

export function reconcile(input: ReconcileInput): AutoReconcileResult {
    const {bank, acct, bankInfos} = input;
    const memory = input.memory ?? {};
    const rowMemory = input.rowMemory ?? {};
    const rowNotes = input.rowNotes ?? {};
    const rowNotePeriods = input.rowNotePeriods ?? {};
    const rowCashMode = input.rowCashMode ?? {};
    const rowPick = input.rowPick ?? {};
    const done = input.done ?? new Set<string>();
    if (input.stmts.length === 0) throw new Error('沒有月結明細');

    const stmts = input.stmts.map((x) => (x.skipLower ? {...x, data: dropLowerHalf(x.data, x.month)} : x));
    const MC = stmts[stmts.length - 1].month;
    const cust: Record<string, StmtCustomer & { alias: string[] }> = {};
    for (const {data} of stmts) {
        for (const [code, v] of Object.entries(data)) {
            if (isNameKey(code)) continue;
            cust[code] = cust[code] || {name: v.name, p: {}, alias: []};
            Object.assign(cust[code].p, v.p);
            cust[code].name = v.name;
            cust[code].alias = [...new Set([...cust[code].alias, v.name])];
        }
    }
    // 沒有店家編號的明細列：名稱只對到一個有編號的客戶就併進去，否則保留名稱鍵自成一個客戶
    const byName: Record<string, Set<string>> = {};
    for (const [c, v] of Object.entries(cust)) for (const a of v.alias) add(byName, norm(a), c);
    for (const {data} of stmts) {
        for (const [key, v] of Object.entries(data)) {
            if (!isNameKey(key)) continue;
            const hit = byName[norm(v.name)];
            const code = hit?.size === 1 ? [...hit][0] : key;
            cust[code] = cust[code] || {name: v.name, p: {}, alias: []};
            Object.assign(cust[code].p, v.p);
            cust[code].alias = [...new Set([...cust[code].alias, v.name])];
        }
    }
    const PERIODS = stmts.flatMap(({month, skipLower}) => (skipLower ? [`${month}上`] : [`${month}上`, `${month}下`]));
    const monthsTxt = stmts.map((x) => x.month).join('、');
    const name2code: Record<string, string> = {};
    for (const [c, v] of Object.entries(cust)) for (const a of v.alias) if (!(a in name2code)) name2code[a] = c;

    // 期別結帳日：上=15日，下=月底；付款日早於結帳日就不可能付那一期
    const perMonth = (per: string): number => parseInt(per, 10);
    const closeOf = (per: string, payDate: string): Date => {
        const [py, pm] = payDate.split('/').map(Number);
        const m = perMonth(per);
        let y = py;
        if (m > pm + 6) y -= 1;
        if (m < pm - 6) y += 1;
        return per.endsWith('上') ? new Date(y, m - 1, 15) : new Date(y, m, 0);
    };
    const toDate = (d: string): Date => {
        const [y, m, dd] = d.split('/').map(Number);
        return new Date(y, m - 1, dd);
    };
    const avail = (per: string, date: string): boolean => !date || closeOf(per, date) <= toDate(date);
    const isCur = (per: string): boolean => perMonth(per) === MC;
    const names = Object.keys(name2code);
    const toCodes = (n: string): string[] => {
        const sc = names.map((x) => [sim(n, x), x] as const).sort((a, b) => b[0] - a[0]);
        if (!sc.length || sc[0][0] < 0.6) return [];
        const top = sc[0][0];
        return [...new Set(sc.filter(([v]) => v >= top - 0.05 && v >= 0.6).map(([, x]) => name2code[x]))];
    };

    // 對照資料（選用的末五碼資料 .xlsx）
    const hist: Record<string, Set<string>> = {};
    const hist2: Record<string, Set<string>> = {};
    const tail: Record<string, Set<string>> = {};
    const namekey: Record<string, Set<string>> = {};
    const override: Record<string, MemoryEntry> = {};
    for (const r of acct) {
        if (/已確認|現付|單次付款/.test(r.flag)) override[r.val] = {name: r.name, flag: r.flag};
        if (r.flag === '歷史') {
            add(hist, r.val, r.name);
            continue;
        }
        let v = r.val;
        if (v.includes('**')) v = v.split('**')[1];
        const d = v.replace(/\D/g, '');
        if (d && d.length === v.length) {
            if (d.length === 4) add(tail, d, JSON.stringify([r.name, r.reg]));
            else if (d.length === 5) add(tail, d.slice(0, 4), JSON.stringify([r.name, r.reg]));
        } else if (!d) add(namekey, v, r.name);
    }
    for (const [k, v] of Object.entries(hist)) {
        if (!/^\d{6,}/.test(k) && mkey(k).includes('#')) v.forEach((n) => add(hist2, mkey(k), n));
    }
    for (const [k, v] of Object.entries(memory)) override[k] = v;

    /** 設定頁末五碼：店家編號存在於明細時直接用代號；否則退回店名比對（以線別篩選）。 */
    const codesFromInfo = (info: BankInfo): string[] => {
        const code = info.storeCode.trim();
        if (code && cust[code]) return [code];
        let t = toCodes(info.customerName);
        const line = /\d/.exec(info.customerLine)?.[0];
        if (line) {
            const f = t.filter((x) => x[0] === line);
            if (f.length) t = f;
        }
        return t;
    };

    const codesOfName = (n: string): string[] => {
        const exact = cust[n] ? n : name2code[n];
        return exact ? [exact] : toCodes(n);
    };

    function candidates(key: string, idx: number): { c: string[]; missing: string[]; src: string } {
        const ov = rowMemory[idx] || override[key];
        if (ov) {
            // 現付帳號走到這裡表示使用者選了這筆照月結回填，照客戶名稱找期別
            const t = ov.flag.includes('新客戶') ? [] : codesOfName(ov.name);
            return {c: t, missing: t.length ? [] : [ov.name], src: '已確認'};
        }
        let c: string[] = [];
        const missing: string[] = [];
        let src = '';
        for (const h of (hist[key] || hist2[mkey(key)] || [])) {
            const t = toCodes(h);
            c.push(...t);
            if (!t.length) missing.push(h);
        }
        if (c.length || missing.length) src = '歷史紀錄';
        if (!c.length && !missing.length) {
            const a = acct4(key);
            const infos = key ? findMatchesFromCandidates([key], bankInfos) : [];
            for (const info of infos) {
                const t = codesFromInfo(info);
                c.push(...t);
                if (!t.length) missing.push(info.customerName);
                src = a ? '末五碼' : '末五碼(戶名)';
            }
            if (a && tail[a]) {
                for (const j of tail[a]) {
                    const [nm, reg] = JSON.parse(j) as [string, number | null];
                    let t = toCodes(nm);
                    if (reg) {
                        const f = t.filter((x) => x[0] === String(reg));
                        if (f.length) t = f;
                    }
                    c.push(...t);
                    if (!t.length) missing.push(nm);
                }
                src = '末五碼';
            }
            for (const nm of (namekey[key] || [])) {
                c.push(...toCodes(nm));
                src = '末五碼(戶名)';
            }
            // 有命中代號的就不再列「不在明細」的同名資料
            if (c.length) missing.length = 0;
        }
        if (!c.length && !missing.length && !acct4(key)) {
            const k2 = stripMemo(key);
            if (k2) {
                c = toCodes(k2);
                src = '備註名稱';
            }
        }
        return {c: [...new Set(c)], missing: [...new Set(missing)], src};
    }

    const RUNS: string[][] = [];
    for (let L = 1; L <= Math.min(6, PERIODS.length); L++) {
        for (let i = 0; i + L <= PERIODS.length; i++) RUNS.push(PERIODS.slice(i, i + L));
    }
    // 跳期付款（例：8上+9上，中間那期之前已付）：最多 3 期的非連續組合
    const contiguous = (r: ReadonlyArray<string>): boolean =>
        r.every((p, i) => i === 0 || PERIODS.indexOf(p) === PERIODS.indexOf(r[i - 1]) + 1);
    for (const n of [2, 3]) for (const cs of combos(PERIODS, n)) if (!contiguous(cs)) RUNS.push(cs);
    const prefKey = (r: ReadonlyArray<string>): number[] => [
        contiguous(r) ? 0 : 1,
        (r.length > 1 && !r.every(isCur)) ? 1 : 0,
        -PERIODS.indexOf(r[r.length - 1]),
        r.length,
    ];
    const PREF = new Map([...RUNS].sort((a, b) => cmpArr(prefKey(a), prefKey(b))).map((r, i) => [r.join(), i]));
    const pref = (run: ReadonlyArray<string>): number => PREF.get(run.join()) ?? 0;

    function classify(amt: number, code: string, date: string): { run: string[]; fee: number } | null {
        const p = cust[code].p;
        let best: { k: number[]; run: string[]; fee: number } | null = null;
        for (const run of RUNS) {
            if (!run.every((x) => p[x])) continue;
            if (!avail(run[run.length - 1], date)) continue;
            const tot = run.reduce((t, x) => t + p[x].amt, 0);
            for (const fee of [0, FEE]) {
                if (tot - fee === amt) {
                    const k = [pref(run), fee];
                    if (!best || cmpArr(k, best.k) < 0) best = {k, run, fee};
                }
            }
        }
        return best ? {run: best.run, fee: best.fee} : null;
    }

    function combo(amt: number, codes: ReadonlyArray<string>, date: string):
        { alloc: Array<[string, string[]]>; fee: number } | null {
        for (const n of [2, 3]) {
            for (const cs of combos(codes, n)) {
                const choices = cs.map((c) => RUNS.filter((r) =>
                    r.length <= 2 && r.every((x) => cust[c].p[x]) && avail(r[r.length - 1], date)));
                const rec = (i: number, pick: string[][]): { alloc: Array<[string, string[]]>; fee: number } | null => {
                    if (i === cs.length) {
                        const tot = cs.reduce((t, c, j) => t + pick[j].reduce((u, x) => u + cust[c].p[x].amt, 0), 0);
                        for (const fee of [0, FEE]) {
                            if (tot - fee === amt) return {alloc: cs.map((c, j) => [c, pick[j]]), fee};
                        }
                        return null;
                    }
                    for (const r of choices[i]) {
                        const v = rec(i + 1, [...pick, r]);
                        if (v) return v;
                    }
                    return null;
                };
                const v = rec(0, []);
                if (v) return v;
            }
        }
        return null;
    }

    // 依金額推測：單期或連續 2~4 期，容許扣手續費 15 / 30
    const GUESS_RUNS = RUNS.filter((r) => r.length <= 4 && contiguous(r));
    const GUESS_FEES = [0, FEE, FEE * 2];
    const GUESS_LIMIT = 8;
    function guess(amt: number, date: string, used: ReadonlySet<string>): GuessPick[] {
        const out: Array<GuessPick & { k: number[] }> = [];
        for (const [code, v] of Object.entries(cust)) {
            for (const run of GUESS_RUNS) {
                if (!run.every((x) => v.p[x] && !v.p[x].paid && !used.has(code + '|' + x))) continue;
                if (!avail(run[run.length - 1], date)) continue;
                const tot = run.reduce((t, x) => t + v.p[x].amt, 0);
                const fee = GUESS_FEES.find((f) => tot - f === amt);
                if (fee !== undefined) out.push({code, run, fee, k: [fee, run.length, pref(run)]});
            }
        }
        return out.sort((a, b) => cmpArr(a.k, b.k)).map(({code, run, fee}) => ({code, run, fee}));
    }

    const OLD = new RegExp('(' + [MC - 1, MC - 2]
        .flatMap((m) => mforms(m))
        .flatMap((f) => ['月', '上', '下'].map((x) => f + x))
        .join('|') + ')');
    const fmtAlloc = (alloc: Array<[string, string[]]>): string => alloc
        .map(([c, run]) => `${cust[c].name} ${run.length < 3 || !contiguous(run)
            ? run.join('+')
            : run[0] + '~' + run[run.length - 1]}`)
        .join(' + ');

    let skipped = 0;
    const results: ReconcileEntry[] = [];
    /** 現付列對不到明細客戶時，顯示記住的名稱 */
    const cashNames: Record<number, string> = {};
    const toGuess: ReconcileEntry[] = [];
    bank.forEach((b, idx) => {
        if (done.has(doneKey(b))) {
            skipped++;
            return;
        }
        const spOv = rowMemory[idx] || override[b.key];
        const special = spOv && specialOf(spOv.flag);
        if (special) {
            const memo = rowNotes[idx]?.trim();
            results.push({
                ...b, idx, alloc: [], fee: 0, status: 'OK', cand: [], src: `已確認（${special}）`,
                confirmed: true, custNames: [], special,
                who: special === ONE_OFF_FLAG ? `${spOv.name}（單次付款）` : DEPOSIT_FLAG,
                note: (memo ? `備註：${memo}；` : '') + special,
                ...(memo ? {userNote: memo} : {}),
            });
            return;
        }
        const cashOv = spOv;
        if (cashOv && isCashFlag(cashOv.flag) && rowCashMode[idx] !== 'period') {
            // 現付：不歸期別，直接算已對上；但金額剛好等於某期（或可組合期別）明細時，可能改月結了，先請使用者確認
            const codes = codesOfName(cashOv.name);
            const res: ReconcileEntry = {
                ...b, idx, alloc: [], fee: 0, status: 'OK', note: '', cand: codes, src: '已確認（現付）',
                confirmed: true, who: '', custNames: [],
            };
            const hit = rowCashMode[idx] === 'cash' ? undefined
                : codes.map((x) => ({x, r: classify(b.amt, x, b.date)}))
                    .find((h): h is { x: string; r: { run: string[]; fee: number } } => h.r !== null);
            if (hit) {
                Object.assign(res, {
                    alloc: [[hit.x, hit.r.run]], fee: hit.r.fee, status: '需確認', cash: 'warn',
                    note: `此帳號記為現付，但金額剛好等於 ${hit.r.run.join('+')} 明細，可能改月結了`,
                });
            } else {
                const memo = rowNotes[idx]?.trim();
                Object.assign(res, {
                    alloc: codes.length ? [[codes[0], []]] : [], cash: 'ok',
                    note: (memo ? `備註：${memo}；` : '') + CASH_FLAG,
                    ...(memo ? {userNote: memo} : {}),
                });
                cashNames[idx] = cashOv.name;
            }
            results.push(res);
            return;
        }
        const pick = rowPick[idx];
        if (pick && cust[pick.code] && pick.run.length && pick.run.every((x) => cust[pick.code].p[x])) {
            results.push({
                ...b, idx, alloc: [[pick.code, pick.run]], fee: pick.fee, status: 'OK', cand: [pick.code],
                src: '已確認', note: '依金額推測，已確認', confirmed: true, who: '', custNames: [],
            });
            return;
        }
        const {c, missing, src} = candidates(b.key, idx);
        const res: ReconcileEntry = {
            ...b, idx, alloc: [], fee: 0, status: '找不到', note: '', cand: c, src,
            confirmed: false, who: '', custNames: [],
        };
        const hits = c.map((x) => ({x, r: classify(b.amt, x, b.date)}))
            .filter((h): h is { x: string; r: { run: string[]; fee: number } } => h.r !== null);
        if (hits.length) {
            hits.sort((a, b2) => cmpArr([pref(a.r.run), a.r.fee], [pref(b2.r.run), b2.r.fee]));
            const h = hits[0];
            res.alloc = [[h.x, h.r.run]];
            res.fee = h.r.fee;
            res.status = ((hits.length === 1 && src !== '末五碼') || c.length === 1) ? 'OK' : '需確認';
            if (hits.length > 1 && c.length > 1) res.note = '其他可能：' + hits.slice(1).map((h2) => cust[h2.x].name).join('、');
            if (src === '末五碼' && c.length > 1 && hits.length === 1) res.status = 'OK';
        } else {
            const cm = c.length > 1 ? combo(b.amt, c, b.date) : null;
            if (cm) {
                Object.assign(res, {alloc: cm.alloc, fee: cm.fee, status: 'OK', note: '多家店合併匯款'});
            } else if (c.length === 1) {
                Object.assign(res, {alloc: [[c[0], []]], status: '金額不符'});
            } else if (c.length) {
                Object.assign(res, {status: '需確認', note: '可能客戶：' + c.map((x) => cust[x].name).join('、') + '，金額皆對不上'});
            } else if (missing.length && src === '已確認') {
                const ov = rowMemory[idx] || override[b.key];
                Object.assign(res, {status: '不在明細', note: `${missing[0]}：${ov?.flag ?? ''}`, nameHint: missing[0]});
            } else if (missing.length) {
                Object.assign(res, {status: '不在明細', nameHint: missing[0], note: `帳號對應到「${missing.join('、')}」，但不在${monthsTxt}月明細表`});
            } else if (b.kind === '現金' && b.bank === '存款機') {
                Object.assign(res, {status: '不在明細', note: '存款機現金存入（可能是自己存現金/收現）'});
            } else {
                // 等其他入帳都比對完，再用金額在尚未收款的期別裡推測（避免搶走已明確對到的期別）
                Object.assign(res, {status: '找不到', note: '帳號/備註沒有紀錄'});
                toGuess.push(res);
            }
        }
        if (src === '已確認' && res.status === '需確認' && res.alloc.length) res.status = 'OK';
        res.confirmed = src === '已確認';
        if (cashOv && isCashFlag(cashOv.flag)) {
            res.cash = 'period';
            res.note = '此帳號記為現付，這筆照月結比對' + (res.note ? '；' + res.note : '');
        }
        results.push(res);
    });

    // 同客戶多筆加總
    const byC: Record<string, ReconcileEntry[]> = {};
    for (const r of results) if (r.status === '金額不符') (byC[r.alloc[0][0]] = byC[r.alloc[0][0]] || []).push(r);
    for (const [code, lst] of Object.entries(byC)) {
        if (lst.length < 2) continue;
        const tot = lst.reduce((t, r) => t + r.amt, 0);
        const r = classify(tot, code, lst[lst.length - 1].date);
        if (r) {
            lst.forEach((x) => Object.assign(x, {
                alloc: [[code, r.run]], fee: r.fee, status: 'OK',
                note: `分${lst.length}筆匯款，合計${tot.toLocaleString()}`,
            }));
        }
    }
    // 找不到的入帳：依金額推測
    const taken = new Set<string>();
    for (const r of results) {
        if (r.status !== 'OK' && r.status !== '需確認') continue;
        r.alloc.forEach(([c, run]) => run.forEach((per) => taken.add(c + '|' + per)));
    }
    for (const r of toGuess) {
        const gs = guess(r.amt, r.date, taken);
        if (!gs.length) continue;
        gs.forEach((g) => {
            g.label = fmtAlloc([[g.code, g.run]]) + (g.fee ? `（扣手續費${g.fee}）` : '');
        });
        r.status = '需確認';
        r.src = '金額推測';
        r.guesses = gs.slice(0, GUESS_LIMIT);
        if (gs.length > GUESS_LIMIT) r.guessMore = gs.length - GUESS_LIMIT;
        if (gs.length === 1) {
            r.alloc = [[gs[0].code, gs[0].run]];
            r.fee = gs[0].fee;
            r.note = '帳號/備註沒有紀錄，依金額推測';
            gs[0].run.forEach((per) => taken.add(gs[0].code + '|' + per));
        } else {
            r.note = `帳號/備註沒有紀錄，依金額推測有 ${gs.length} 個可能：`
                + r.guesses.map((g) => g.label).join('、')
                + (r.guessMore ? `…等（其他請用「改成」自己輸入）` : '');
        }
    }
    for (const r of results) {
        if (r.special) continue;
        if (r.status === '金額不符') {
            const p = cust[r.alloc[0][0]].p;
            r.note = (OLD.test(r.key) ? '備註寫舊月份；' : '') + '明細金額：'
                + PERIODS.filter((k) => p[k]).map((k) => `${k} ${p[k].amt.toLocaleString()}`).join('、');
        }
        if (r.fee) r.note = (`扣手續費${r.fee}；` + r.note).replace(/；$/, '');
        r.who = (r.alloc.length && r.alloc[0][1].length)
            ? fmtAlloc(r.alloc)
            : r.alloc.map(([c]) => cust[c].name).join('、');
        r.custNames = r.alloc.map(([c]) => cust[c].name);
        if (r.cash === 'ok') r.who = (r.who || cashNames[r.idx]) + '（現付）';
    }

    // 金額不符：決定回填期別（預設＝原本紅字提醒那一期），填了備註就照一般 OK 回填
    const used = new Set<string>();
    for (const r of results) {
        if (r.status !== 'OK' && r.status !== '需確認') continue;
        r.alloc.forEach(([c, run]) => run.forEach((per) => used.add(c + '|' + per)));
    }
    for (const r of results) {
        if (r.status !== '金額不符') continue;
        const code = r.alloc[0][0];
        const opts = PERIODS.filter((k) => cust[code].p[k]);
        if (!opts.length) continue;
        const def = [...opts].reverse().find((k) => !used.has(code + '|' + k)) ?? opts[opts.length - 1];
        const want = opts.filter((k) => rowNotePeriods[r.idx]?.includes(k));
        r.periodOptions = opts;
        r.notePeriods = want.length ? want : [def];
        const memo = rowNotes[r.idx]?.trim();
        if (!memo) continue;
        r.userNote = memo;
        const pers = r.notePeriods;
        // 多期：前面各期填明細金額，差額落在最後一期
        const before = pers.slice(0, -1).reduce((t, k) => t + cust[code].p[k].amt, 0);
        const last = r.amt - before;
        if (last <= 0) {
            r.noteError = `勾了 ${pers.join('、')} 共${pers.length}期，前${pers.length - 1}期明細合計 ${before.toLocaleString()} 已超過入帳 ${r.amt.toLocaleString()}，請少勾幾期`;
            r.note = `備註：${memo}；${r.note}`;
            continue;
        }
        r.noteShares = [...pers.slice(0, -1).map((k) => cust[code].p[k].amt), last];
        r.status = 'OK';
        r.alloc = [[code, pers]];
        r.who = fmtAlloc(r.alloc);
        const fillTxt = pers.length > 1
            ? '回填 ' + pers.map((k, i) => `${k} ${r.noteShares![i].toLocaleString()}`).join('、')
            : `回填 ${pers[0]}`;
        r.note = `備註：${memo}；${fillTxt}；${r.note}`;
    }
    return {results, cust, PERIODS, skipped, names: names.sort()};
}

/** 依相似度列出最接近的客戶名稱（「改成」找不到完全一致的名字時給使用者挑）。 */
export function suggestNames(n: string, names: ReadonlyArray<string>, limit = 6): string[] {
    return names.map((x) => [sim(n, x), x] as const)
        .filter(([v]) => v >= 0.25)
        .sort((a, b) => b[0] - a[0])
        .slice(0, limit)
        .map(([, x]) => x);
}

// ---------- 寫回 Excel

const fill = (argb: string): ExcelJS.Fill => ({type: 'pattern', pattern: 'solid', fgColor: {argb}});

/** 收款日「一次付多期或多店合併」改用底色標示（與綠/黃底區分）。 */
const MULTI_DATE_FILL = 'FFDDEBF7';
const RED = 'FFC00000';
const MEMO_HEADER = '備註';
const MEMO_WIDTH = 24;

// ---------- 明細插入「備註」欄（各區塊收款日右邊）

const MAX_COL = 16384;
const colNum = (l: string): number => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const colLetter = (n: number): string => {
    let out = '';
    for (; n > 0; n = Math.floor((n - 1) / 26)) out = String.fromCharCode(65 + ((n - 1) % 26)) + out;
    return out;
};

/** 純參照字串（A1、A1:K9、合併範圍、列印範圍）：欄號 ≥ at 的右移一欄；跨過插入點的範圍因此變寬（同 Excel 插入欄）。 */
const shiftRefs = (ref: string, at: number): string =>
    ref.replace(/(\$?)([A-Z]{1,3})(\$?\d+)/g, (m, d: string, c: string, r: string) =>
        (colNum(c) >= at ? `${d}${colLetter(colNum(c) + 1)}${r}` : m));

const unquote = (n: string): string => (n.startsWith("'") ? n.slice(1, -1).replace(/''/g, "'") : n);

/** 公式內指向 target 工作表的參照右移（略過字串常值與外部活頁簿參照 [1]工作表!A1）。 */
function shiftFormula(f: string, own: string, target: string, at: number): string {
    const re = /("(?:[^"]|"")*")|(?<![\w.$\u0080-￿])((\[\d+\])?('(?:[^']|'')+'|[A-Za-z_\u0080-￿][\w.\u0080-￿]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(])/g;
    return f.replace(re, (m, str: string | undefined, prefix: string | undefined, ext: string | undefined, sheet: string | undefined, ref: string) => {
        if (str || ext) return m;
        const to = prefix ? unquote(sheet!) : own;
        return to === target ? (prefix ?? '') + shiftRefs(ref, at) : m;
    });
}

/** 在 ws 第 at 欄插入空白欄：搬移儲存格、合併、欄寬，並調整整本活頁簿指向此工作表的公式。 */
function insertCol(ws: ExcelJS.Worksheet, at: number): void {
    const name = ws.name;
    // 其他工作表只需改公式
    for (const other of ws.workbook.worksheets) {
        if (other === ws) continue;
        other.eachRow((row) => row.eachCell((c) => {
            // 合併格的從屬格會讀到主格的公式，只改主格；共用公式的從屬格跟著主格走
            if (c.isMerged && c.master.address !== c.address) return;
            const v = c.value as {formula?: string; sharedFormula?: string} | null;
            if (!v || typeof v !== 'object' || typeof v.formula !== 'string' || v.sharedFormula) return;
            const f = shiftFormula(v.formula, other.name, name, at);
            if (f !== v.formula) c.value = {...v, formula: f} as ExcelJS.CellValue;
        }));
    }
    const m = ws.model as ExcelJS.WorksheetModel & {
        rows: Array<{cells: Array<Record<string, unknown> & {address: string}>}>;
        cols?: Array<{min: number; max: number}>;
        merges: string[];
        mergeCells?: string[];
    };
    const merges = [...m.merges];
    for (const r of merges) ws.unMergeCells(r);
    for (const row of m.rows) {
        for (const c of row.cells) {
            c.address = shiftRefs(c.address, at);
            if (typeof c.formula === 'string') c.formula = shiftFormula(c.formula, name, name, at);
            if (typeof c.sharedFormula === 'string') c.sharedFormula = shiftRefs(c.sharedFormula, at);
            if (typeof c.ref === 'string') c.ref = shiftRefs(c.ref, at);
        }
    }
    m.mergeCells = merges.map((r) => shiftRefs(r, at));
    m.cols = (m.cols ?? [])
        .map((c) => ({...c, min: c.min >= at ? c.min + 1 : c.min, max: Math.min(MAX_COL, c.max >= at ? c.max + 1 : c.max)}))
        .filter((c) => c.min <= MAX_COL);
    // 原本跨過插入點的欄設定要拆開，新欄自己一段
    m.cols = m.cols.flatMap((c) => (c.min < at && c.max >= at
        ? [{...c, max: at - 1}, ...(c.max > at ? [{...c, min: at + 1}] : [])]
        : [c]));
    if (m.pageSetup?.printArea) m.pageSetup.printArea = shiftRefs(m.pageSetup.printArea, at);
    ws.model = m;
    ws.getColumn(at).width = MEMO_WIDTH;
}

const noteText = (note: ExcelJS.Cell['note']): string => {
    if (!note) return '';
    if (typeof note === 'string') return note;
    return (note.texts ?? []).map((t) => t.text).join('');
};

/** 舊版回填加的註解（每行都是「2026/10/17 匯入 …」）；使用者自己加的註解不動。 */
const OLD_FILL_NOTE = /^\d{4}\/\d{1,2}\/\d{1,2} 匯入 /;

/** 清掉舊版回填加在實收金額格的紅色註解（改成備註欄後不再需要）。 */
function clearOldFillNotes(ws: ExcelJS.Worksheet): void {
    ws.eachRow((row) => row.eachCell((c) => {
        const lines = noteText(c.note).split('\n').filter((x) => x.trim());
        if (!lines.length || !lines.every((x) => OLD_FILL_NOTE.test(x))) return;
        // ExcelJS 沒有移除註解的 API：清內部欄位，連同 model 快取裡的 comment（寫檔時讀的是它）
        const raw = c as unknown as {_comment?: unknown; _value?: {model?: {comment?: unknown}}};
        raw._comment = undefined;
        if (raw._value?.model) delete raw._value.model.comment;
    }));
}

/**
 * 明細各區塊收款日右邊插入「備註」欄（已有就沿用，接續對帳上傳回填檔時不會重複插）。
 * 回傳本次插入位置（原欄號），供換算對帳時讀到的欄號；找不到標題列的舊版面不插欄（回傳 null）。
 */
function addMemoCols(ws: ExcelJS.Worksheet): number[] | null {
    const {headerRow, blocks} = stmtLayout(ws);
    if (blocks === LEGACY_BLOCKS) return null;
    const at = blocks.map((b) => b.date + 1)
        .filter((c) => s(cellVal(ws.getCell(headerRow, c))) !== MEMO_HEADER)
        .sort((a, b) => b - a);
    for (const c of at) {
        insertCol(ws, c);
        // 列印範圍剛好到收款日的，把新的備註欄也包進去
        const pa = /^(.*:\$?)([A-Z]{1,3})(\$?\d+)$/.exec(ws.pageSetup.printArea ?? '');
        if (pa && colNum(pa[2]) === c - 1) ws.pageSetup.printArea = pa[1] + colLetter(c) + pa[3];
        ws.getCell(headerRow, c).value = MEMO_HEADER;
    }
    // 已有備註欄（含 2.12.8 版下載、框線缺漏的檔）也重新套一次，修好舊檔
    for (const b of stmtLayout(ws).blocks) styleMemoCol(ws, headerRow, b.date + 1);
    return at;
}

/** 回填時加上的底色；備註欄沿用收款日格樣式時不跟著帶過去。 */
const FILL_MARKS = new Set([...Object.values(COLORS), MULTI_DATE_FILL]);

/**
 * 備註欄每一格的框線、字型、底色都照左邊收款日格（含沒有儲存格、靠欄樣式畫框的空白列與合計區），
 * 欄樣式也照抄，表格框線跟插欄前同一個形狀，只是多一欄。
 */
function styleMemoCol(ws: ExcelJS.Worksheet, headerRow: number, c: number): void {
    const memoStyle = (src: Partial<ExcelJS.Style>, isHeader: boolean): Partial<ExcelJS.Style> => {
        const st = JSON.parse(JSON.stringify(src ?? {})) as Partial<ExcelJS.Style>;
        const fg = (st.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb;
        if (fg && FILL_MARKS.has(fg)) delete st.fill;
        st.alignment = {...st.alignment, horizontal: isHeader ? 'center' : 'left', vertical: 'middle', shrinkToFit: false};
        return st;
    };
    const col = ws.getColumn(c);
    col.style = memoStyle(ws.getColumn(c - 1).style, false);
    for (let r = 1; r <= ws.rowCount; r++) {
        const cell = ws.getCell(r, c);
        const src = ws.getCell(r, c - 1);
        // 下方每段表格重複的標題列也補上「備註」
        const isHeader = r === headerRow || s(cellVal(src)) === '收款日';
        const font = cell.font;
        const keepFont = !isHeader && !!s(cellVal(cell));
        cell.style = memoStyle(src.style, isHeader);
        if (isHeader && !s(cellVal(cell))) cell.value = MEMO_HEADER;
        // 已寫好的備註保留自己的字型（convertOldDateTags 會另外把小字改回來）
        if (keepFont && font) cell.font = font;
    }
}

type MemoPart = {text: string; red?: boolean};

/** 備註欄寫入：標記黑字、使用者備註紅字，以「；」接在原有內容後面（front 時放前面；同樣的字不重複寫）。 */
function writeMemo(c: ExcelJS.Cell, parts: ReadonlyArray<MemoPart>, front = false): void {
    const prev = s(cellVal(c));
    const have = new Set(prev.split('；'));
    const add = parts.filter((p) => p.text && !have.has(p.text));
    if (!add.length) return;
    const font = c.font ?? {};
    const old: ExcelJS.RichText[] = [];
    if (prev) {
        const v = c.value as {richText?: ExcelJS.RichText[]} | null;
        old.push(...(v && typeof v === 'object' && v.richText ? v.richText : [{text: prev, font}]));
    }
    const rich: ExcelJS.RichText[] = front ? [] : [...old];
    add.forEach((p, i) => {
        if ((!front && prev) || i) rich.push({text: '；', font});
        rich.push({text: p.text, font: p.red ? {...font, color: {argb: RED}} : font});
    });
    if (front && old.length) rich.push({text: '；', font}, ...old);
    c.value = {richText: rich};
    c.alignment = {...c.alignment, horizontal: 'left', vertical: 'middle', shrinkToFit: false};
}

/** 舊版回填的收款日文字：日期後接 (-15)/(月)/(2筆) 與結尾 ?，例如「10/12(-15)(月)」「10/5(2筆)?」。 */
const OLD_DATE_TAGS = /^(\d{1,2}\/\d{1,2})((?:\((?:-\d+|月|\d+筆)\))*)(\?)?$/;

/** 格子的字級（沒設定就回空物件，不覆蓋原字型）。 */
const sizeOf = (c: ExcelJS.Cell): Partial<ExcelJS.Font> => (c.font?.size ? {size: c.font.size} : {});

/** 舊版小字收款日轉成新格式：收款日只留日期、取消縮字並放大到金額欄字級，(月) 改淺藍底，標記移到備註欄。 */
function convertOldDateTags(ws: ExcelJS.Worksheet): void {
    const {headerRow, blocks} = stmtLayout(ws);
    ws.eachRow((row, r) => {
        if (r <= headerRow) return;
        for (const b of blocks) {
            const ec = row.getCell(b.date);
            const v = cellVal(ec);
            const m = typeof v === 'string' ? OLD_DATE_TAGS.exec(v.trim()) : null;
            if (!m || (!m[2] && !m[3])) continue;
            const tags = m[2].slice(1, -1).split(')(').filter(Boolean);
            const fee = tags.find((x) => x.startsWith('-'));
            const cnt = tags.find((x) => x.endsWith('筆'));
            const multi = tags.includes('月');
            const parts: MemoPart[] = [
                ...(multi ? [{text: '月'}] : []),
                ...(fee ? [{text: `扣手續費${fee.slice(1)}`}] : []),
                ...(cnt ? [{text: cnt}] : []),
                ...(m[3] ? [{text: '待確認'}] : []),
            ];
            ec.value = m[1];
            const memoFont = {...ec.font, ...sizeOf(row.getCell(b.paid))};
            ec.style = {
                ...ec.style,
                font: {...ec.font, ...sizeOf(row.getCell(b.amt))},
                alignment: {...ec.alignment, shrinkToFit: false},
                ...(multi ? {fill: fill(MULTI_DATE_FILL)} : {}),
            };
            // 備註格插欄時沿用了收款日的小字，改回實收金額欄原本的字級
            const mc = row.getCell(b.date + 1);
            mc.style = {...mc.style, font: memoFont};
            writeMemo(mc, parts, true);
        }
    });
}

/** 舊版已回填的實收金額、收款日（帶回填底色、內容是數字或單一日期）放大到同列金額格的字級。 */
function enlargeOldFills(ws: ExcelJS.Worksheet): void {
    const {headerRow, blocks} = stmtLayout(ws);
    ws.eachRow((row, r) => {
        if (r <= headerRow) return;
        for (const b of blocks) {
            const size = sizeOf(row.getCell(b.amt));
            if (!size.size) continue;
            for (const c of [row.getCell(b.paid), row.getCell(b.date)]) {
                const fg = (c.fill as ExcelJS.FillPattern | undefined)?.fgColor?.argb;
                const v = cellVal(c);
                if (!fg || !FILL_MARKS.has(fg) || !(typeof v === 'number' || /^\d{1,2}\/\d{1,2}$/.test(s(v)))) continue;
                c.style = {...c.style, font: {...c.font, ...size}, alignment: {...c.alignment, shrinkToFit: false}};
            }
        }
    });
}

/**
 * 把對帳結果回填到各月明細 workbook（就地修改），並在最新月份 workbook 加上「銀行對帳結果」工作表。
 * @param wbs 月份 → workbook
 * @param latestMonth 最新月份（放結果表）
 */
export function writeBack(wbs: Record<number, ExcelJS.Workbook>, latestMonth: number, rec: AutoReconcileResult): void {
    const {results, cust} = rec;
    const wbC = wbs[latestMonth];
    const wbOf = (per: string): ExcelJS.Workbook | undefined => wbs[parseInt(per, 10)];

    // 先插好備註欄，再把對帳時讀到的欄號換算成插欄後的位置
    const inserted = new Map<ExcelJS.Worksheet, number[] | null>();
    for (const wb of Object.values(wbs)) {
        for (const n of ['上期', '下期']) {
            const ws = sheetByName(wb, n);
            if (!ws) continue;
            clearOldFillNotes(ws);
            const at = addMemoCols(ws);
            inserted.set(ws, at);
            if (at) convertOldDateTags(ws);
            enlargeOldFills(ws);
        }
    }
    const colOf = (ws: ExcelJS.Worksheet, c: number): number => c + (inserted.get(ws) ?? []).filter((p) => p <= c).length;
    /** 備註格：收款日右邊；沒插欄的舊版面只在該格空白時寫。 */
    const memoCell = (ws: ExcelJS.Worksheet, info: PeriodCell): ExcelJS.Cell | undefined => {
        const c = ws.getCell(info.row, colOf(ws, info.dateCol) + 1);
        return inserted.get(ws) || !s(cellVal(c)) ? c : undefined;
    };

    const fills = new Map<string, Array<{ r: ReconcileEntry; share: number; multi: boolean }>>();
    for (const r of results) {
        if (!(r.status === 'OK' || r.status === '需確認') || !r.alloc.length || !r.alloc[0][1].length) continue;
        r.alloc.forEach(([code, run], ai) => run.forEach((per, i) => {
            let share = cust[code].p[per].amt;
            if (r.fee && i === run.length - 1 && ai === r.alloc.length - 1) share -= r.fee;
            if (r.alloc.length === 1 && run.length === 1) share = r.amt;
            if (r.noteShares) share = r.noteShares[i];
            const k = code + '|' + per;
            if (!fills.has(k)) fills.set(k, []);
            fills.get(k)!.push({r, share, multi: run.length > 1 || r.alloc.length > 1});
        }));
    }
    for (const [k, lst0] of fills) {
        const cut = k.lastIndexOf('|');
        const code = k.slice(0, cut);
        const per = k.slice(cut + 1);
        const info = cust[code].p[per];
        const wb = wbOf(per);
        if (!wb) continue;
        const ws = sheetByName(wb, info.sheet);
        if (!ws) continue;
        const dc = ws.getCell(info.row, colOf(ws, info.paidCol));
        const ec = ws.getCell(info.row, colOf(ws, info.dateCol));
        // 實收金額、收款日字級與同列金額格一致
        const size = sizeOf(ws.getCell(info.row, colOf(ws, info.amtCol)));
        const big = (c: ExcelJS.Cell): Partial<ExcelJS.Style> => (size.size ? {font: {...c.font, ...size}} : {});
        const ok = lst0.every((x) => x.r.status === 'OK');
        const prevVal = cellVal(dc);
        const prev = typeof prevVal === 'number' ? prevVal : 0;
        dc.value = prev + lst0.reduce((t, x) => t + x.share, 0);
        const n = lst0.length + (prev ? 1 : 0);
        const multi = lst0.some((x) => x.multi);
        const feeOf = lst0.find((x) => x.r.fee)?.r.fee;
        const parts: MemoPart[] = [
            ...(multi ? [{text: '月'}] : []),
            ...(feeOf ? [{text: `扣手續費${feeOf}`}] : []),
            ...(n > 1 ? [{text: `${n}筆`}] : []),
            ...(ok ? [] : [{text: '待確認'}]),
            ...[...new Set(lst0.map((x) => x.r.userNote).filter((x): x is string => !!x))].map((text) => ({text, red: true})),
        ];
        ec.value = md(lst0[lst0.length - 1].r.date);
        const color = ok ? COLORS.OK : COLORS.需確認;
        dc.style = {...dc.style, ...big(dc), fill: fill(color)};
        ec.style = {...ec.style, ...big(ec), fill: fill(multi ? MULTI_DATE_FILL : color), alignment: {...ec.alignment, shrinkToFit: false}};
        const mc = memoCell(ws, info);
        if (mc && parts.length) writeMemo(mc, parts);
    }

    // 金額不符：收款日標黃底，備註欄寫紅字提醒
    const mm: Record<string, ReconcileEntry[]> = {};
    for (const r of results) if (r.status === '金額不符') (mm[r.alloc[0][0]] = mm[r.alloc[0][0]] || []).push(r);
    for (const [code, lst] of Object.entries(mm)) {
        const p = cust[code].p;
        const per = [...rec.PERIODS].reverse().find((x) => p[x] && !fills.has(code + '|' + x));
        if (!per) continue;
        const wb = wbOf(per);
        if (!wb) continue;
        const info = p[per];
        const ws = sheetByName(wb, info.sheet);
        if (!ws) continue;
        const ec = ws.getCell(info.row, colOf(ws, info.dateCol));
        ec.style = {...ec.style, fill: fill(COLORS.金額不符), alignment: {...ec.alignment, shrinkToFit: false}};
        const mc = memoCell(ws, info);
        if (mc) writeMemo(mc, lst.map((r) => ({text: `${md(r.date)}收${r.amt.toLocaleString()}，金額對不上`, red: true})));
    }

    // 結果工作表
    let ws = sheetByName(wbC, RESULT_SHEET);
    const header = ['日期', '摘要', '存入金額', '銀行', '帳號/備註', '對應客戶與期別', '狀態', '依據', '說明', '確認'];
    if (ws) {
        ws.addRow([]);
        const r = ws.addRow([`—— 新增：${results.length ? results[0].date : ''} 起的入帳 ——`]);
        r.getCell(1).font = {bold: true, color: {argb: 'FF1F4E79'}};
    } else {
        ws = wbC.addWorksheet(RESULT_SHEET);
        (ws as unknown as { orderNo: number }).orderNo = -1;
        ws.addRow(header).font = {bold: true};
        const widths = [11, 8, 10, 8, 20, 30, 9, 10, 46, 7];
        widths.forEach((w, i) => {
            ws!.getColumn(i + 1).width = w;
        });
        ws.views = [{state: 'frozen', ySplit: 1}];
    }
    const order: Record<AutoReconcileStatus, number> = {找不到: 0, 需確認: 1, 金額不符: 2, 不在明細: 3, OK: 4};
    const sorted = [...results].sort((a, b) => (order[a.status] - order[b.status]) || cmpDate(a.date, b.date));
    for (const r of sorted) {
        const row = ws.addRow([r.date, r.kind, r.amt, r.bank, r.key, r.who, r.special ?? r.status, r.src, r.note, r.confirmed ? '✓' : '']);
        row.eachCell({includeEmpty: true}, (c) => {
            c.fill = fill(COLORS[r.status]);
        });
    }
}
