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
 *
 * 輸出：回填各月明細的「實收金額 / 收款日」，並在最新月份檔加上「銀行對帳結果」工作表；
 * 下次放入回填後的檔案會依該表略過已處理的入帳（接續模式）。
 */

import type ExcelJS from 'exceljs';

import type {BankInfo} from '@/domain/models/bank-info';
import {findMatchesFromCandidates} from '@/domain/bank-match-service';

export const FEE = 15;
export const RESULT_SHEET = '銀行對帳結果';
/** 使用者確認「這是新客戶（清單沒有）」時的標記：照打的名字存，不做模糊比對。 */
export const NEW_CUSTOMER_FLAG = '已確認（新客戶）';

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
    col: number;
    amt: number;
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

export function readStmt(wb: ExcelJS.Workbook, month: number): StmtData {
    const out: StmtData = {};
    for (const [half, sname] of [['上', '上期'], ['下', '下期']] as const) {
        const ws = sheetByName(wb, sname);
        if (!ws) throw new Error(`明細表找不到「${sname}」工作表`);
        ws.eachRow({includeEmpty: false}, (row, r) => {
            if (r < 3) return;
            for (const bc of [2, 8]) {
                const code = s(cellVal(row.getCell(bc - 1)));
                const name = s(cellVal(row.getCell(bc)));
                const amt = cellVal(row.getCell(bc + 1));
                if (code && /^\d+$/.test(code) && name && typeof amt === 'number') {
                    out[code] = out[code] || {name, p: {}};
                    out[code].p[`${month}${half}`] = {sheet: ws.name, row: r, col: bc, amt: Math.round(amt)};
                }
            }
        });
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
    /** 由舊到新排序。 */
    stmts: ReadonlyArray<{ month: number; data: StmtData }>;
    acct: ReadonlyArray<AcctRow>;
    bankInfos: ReadonlyArray<BankInfo>;
    memory?: Record<string, MemoryEntry>;
    rowMemory?: Record<number, MemoryEntry>;
    /** 本次對帳的金額不符備註（以 idx 為鍵）；有備註的金額不符改算 OK。 */
    rowNotes?: Record<number, string>;
    /** 本次對帳金額不符的回填期別（以 idx 為鍵，可多期）；沒選用預設期別。 */
    rowNotePeriods?: Record<number, string[]>;
    done?: Set<string>;
}

export function reconcile(input: ReconcileInput): AutoReconcileResult {
    const {bank, stmts, acct, bankInfos} = input;
    const memory = input.memory ?? {};
    const rowMemory = input.rowMemory ?? {};
    const rowNotes = input.rowNotes ?? {};
    const rowNotePeriods = input.rowNotePeriods ?? {};
    const done = input.done ?? new Set<string>();
    if (stmts.length === 0) throw new Error('沒有月結明細');

    const MC = stmts[stmts.length - 1].month;
    const cust: Record<string, StmtCustomer & { alias: string[] }> = {};
    for (const {data} of stmts) {
        for (const [code, v] of Object.entries(data)) {
            cust[code] = cust[code] || {name: v.name, p: {}, alias: []};
            Object.assign(cust[code].p, v.p);
            cust[code].name = v.name;
            cust[code].alias = [...new Set([...cust[code].alias, v.name])];
        }
    }
    const PERIODS = stmts.flatMap(({month}) => [`${month}上`, `${month}下`]);
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
    const add = (o: Record<string, Set<string>>, k: string, v: string) => {
        (o[k] = o[k] || new Set()).add(v);
    };
    for (const r of acct) {
        if (/已確認|現付/.test(r.flag)) override[r.val] = {name: r.name, flag: r.flag};
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

    function candidates(key: string, idx: number): { c: string[]; missing: string[]; src: string } {
        const ov = rowMemory[idx] || override[key];
        if (ov) {
            const exact = cust[ov.name] ? ov.name : name2code[ov.name];
            const t = (ov.flag.includes('現付') || ov.flag.includes('新客戶')) ? []
                : (exact ? [exact] : toCodes(ov.name));
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
    bank.forEach((b, idx) => {
        if (done.has(doneKey(b))) {
            skipped++;
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
                Object.assign(res, {status: '不在明細', note: `${missing[0]}：${ov?.flag ?? ''}`});
            } else if (missing.length) {
                Object.assign(res, {status: '不在明細', note: `帳號對應到「${missing.join('、')}」，但不在${monthsTxt}月明細表`});
            } else if (b.kind === '現金' && b.bank === '存款機') {
                Object.assign(res, {status: '不在明細', note: '存款機現金存入（可能是自己存現金/收現）'});
            } else {
                const g: Array<[string, string[]]> = [];
                for (const x of Object.keys(cust)) {
                    const r = classify(b.amt, x, b.date);
                    if (r && !r.fee && r.run.length === 1 && isCur(r.run[0])) g.push([x, r.run]);
                }
                if (g.length === 1 && !OLD.test(b.key)) {
                    Object.assign(res, {alloc: g, status: '需確認', note: '帳號沒有紀錄，依金額推測'});
                } else {
                    Object.assign(res, {
                        status: '找不到',
                        note: '帳號/備註沒有紀錄' + (g.length ? '；金額相同的有：' + g.slice(0, 8).map(([x]) => cust[x].name).join('、') : ''),
                    });
                }
            }
        }
        if (src === '已確認' && res.status === '需確認' && res.alloc.length) res.status = 'OK';
        res.confirmed = src === '已確認';
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
    for (const r of results) {
        if (r.status === '金額不符') {
            const p = cust[r.alloc[0][0]].p;
            r.note = (OLD.test(r.key) ? '備註寫舊月份；' : '') + '明細金額：'
                + PERIODS.filter((k) => p[k]).map((k) => `${k} ${p[k].amt.toLocaleString()}`).join('、');
        }
        if (r.fee) r.note = ('扣手續費15；' + r.note).replace(/；$/, '');
        r.who = (r.alloc.length && r.alloc[0][1].length)
            ? fmtAlloc(r.alloc)
            : r.alloc.map(([c]) => cust[c].name).join('、');
        r.custNames = r.alloc.map(([c]) => cust[c].name);
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

const noteText = (note: ExcelJS.Cell['note']): string => {
    if (!note) return '';
    if (typeof note === 'string') return note;
    return (note.texts ?? []).map((t) => t.text).join('');
};

/**
 * 把對帳結果回填到各月明細 workbook（就地修改），並在最新月份 workbook 加上「銀行對帳結果」工作表。
 * @param wbs 月份 → workbook
 * @param latestMonth 最新月份（放結果表）
 */
export function writeBack(wbs: Record<number, ExcelJS.Workbook>, latestMonth: number, rec: AutoReconcileResult): void {
    const {results, cust} = rec;
    const wbC = wbs[latestMonth];
    const wbOf = (per: string): ExcelJS.Workbook | undefined => wbs[parseInt(per, 10)];

    const fills = new Map<string, Array<{ r: ReconcileEntry; share: number; multi: boolean }>>();
    for (const r of results) {
        if (!(r.status === 'OK' || r.status === '需確認') || !r.alloc.length || !r.alloc[0][1].length) continue;
        r.alloc.forEach(([code, run], ai) => run.forEach((per, i) => {
            let share = cust[code].p[per].amt;
            if (r.fee && i === run.length - 1 && ai === r.alloc.length - 1) share -= FEE;
            if (r.alloc.length === 1 && run.length === 1) share = r.amt;
            if (r.noteShares) share = r.noteShares[i];
            const k = code + '|' + per;
            if (!fills.has(k)) fills.set(k, []);
            fills.get(k)!.push({r, share, multi: run.length > 1 || r.alloc.length > 1});
        }));
    }
    for (const [k, lst0] of fills) {
        const [code, per] = k.split('|');
        const info = cust[code].p[per];
        const wb = wbOf(per);
        if (!wb) continue;
        const ws = sheetByName(wb, info.sheet);
        if (!ws) continue;
        const dc = ws.getCell(info.row, info.col + 2);
        const ec = ws.getCell(info.row, info.col + 3);
        const ok = lst0.every((x) => x.r.status === 'OK');
        const prevVal = cellVal(dc);
        const prev = typeof prevVal === 'number' ? prevVal : 0;
        dc.value = prev + lst0.reduce((t, x) => t + x.share, 0);
        const n = lst0.length + (prev ? 1 : 0);
        let tag = '';
        if (lst0.some((x) => x.r.fee)) tag += '(-15)';
        if (lst0.some((x) => x.multi)) tag += '(月)';
        if (n > 1) tag += `(${n}筆)`;
        if (!ok) tag += '?';
        ec.value = md(lst0[lst0.length - 1].r.date) + tag;
        const color = ok ? COLORS.OK : COLORS.需確認;
        dc.style = {...dc.style, fill: fill(color)};
        ec.style = {...ec.style, fill: fill(color), ...(tag ? {font: {...(ec.font || {}), size: 8}} : {})};
        const old = dc.note ? noteText(dc.note) + '\n' : '';
        dc.note = old + lst0.map((x) => `${x.r.date} 匯入 ${x.r.amt.toLocaleString()}（${x.r.key}）${x.r.who} ${x.r.note}`).join('\n');
    }

    // 金額不符：在收款日欄寫紅字提醒
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
        const ec = ws.getCell(info.row, info.col + 3);
        const dc = ws.getCell(info.row, info.col + 2);
        const ev = cellVal(ec);
        const prevTxt = typeof ev === 'string' && ev ? ev + '；' : '';
        ec.value = prevTxt + lst.map((r) => `${md(r.date)}收${r.amt}?`).join('；');
        ec.style = {...ec.style, fill: fill(COLORS.金額不符), font: {size: 8, color: {argb: 'FFC00000'}}};
        dc.note = lst.map((r) => `${r.date} 匯入 ${r.amt.toLocaleString()}（${r.key}）金額對不上；${r.note}`).join('\n');
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
        const row = ws.addRow([r.date, r.kind, r.amt, r.bank, r.key, r.who, r.status, r.src, r.note, r.confirmed ? '✓' : '']);
        row.eachCell({includeEmpty: true}, (c) => {
            c.fill = fill(COLORS[r.status]);
        });
    }
}
