/**
 * 「自動催款」收款登記：一次登記某位客人好幾期的收款，寫回各月的「X月_已自動對帳.xlsx」。
 *
 * - 列出客人在已記住月份中還沒付清（差額超過手續費 15 元）的期別；還沒結帳的期別、已標「已結清」的期別不列。
 * - 月結客人同月上下期合成一行（「8月」），金額先補上期、剩下的算下期。
 * - 實收金額預設＝這期還欠的金額（沒收過就是帳單金額）；寫回時加在原本的實收金額上。
 */

import {compute, customerKey, FEE, type History, type Marks, periodKey, unsettledPeriods} from '@/domain/auto-collect/ledger';
import type {Half} from '@/domain/auto-collect/statement-parser';
import type {PaymentEntry} from '@/domain/auto-reconcile-engine';

export interface RegisterPeriod {
    month_key: string;
    half: Half;
    /** 明細裡的客人鍵（代號或「名:店名」），寫回 Excel 時用來找列。 */
    rkey: string;
    amt: number;
    /** 原本的實收金額（沒有為 0）。 */
    paid: number;
    owed: number;
}

/** 畫面上的一行（可勾選）：一般客人一期一行；月結客人同月上下期合成一行。 */
export interface RegisterLine {
    id: string;
    label: string;
    month_key: string;
    periods: RegisterPeriod[];
    amt: number;
    paid: number;
    owed: number;
}

export interface RegisterCustomer {
    /** 客人鍵（同 customerKey：代號，沒代號用店名）。 */
    ckey: string;
    code: string;
    name: string;
    monthly: boolean;
    lines: RegisterLine[];
    total: number;
}

const HALF_LABEL: Record<Half, string> = {上: '上期', 下: '下期'};

/** 每位客人目前還沒付清的期別（沒有欠款的客人也列出，lines 為空）。 */
export function registerCustomers(history: History, marks: Marks, today: Date = new Date()): RegisterCustomer[] {
    const keys = Object.keys(history.months).sort();
    const unsettled = unsettledPeriods(history);
    const multiYear = new Set(keys.map((k) => k.slice(0, 4))).size > 1;
    const halves = ['上', '下'] as const;
    const debtors = new Map(compute(history, marks, {}, today).customers.map((c) => [customerKey(c), c]));

    // 沒代號的列：名稱只對到一個代號就併過去（同 compute）
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

    type Acc = {code: string; name: string; periods: RegisterPeriod[]; monthlyMemo: boolean};
    const custs = new Map<string, Acc>();
    for (const k of keys) {
        for (const half of halves) {
            for (const [rkey, rec] of Object.entries(history.months[k].periods[half] ?? {})) {
                let code = rec.code;
                if (!code) {
                    const hit = nameToCode.get(rec.name);
                    if (hit?.size === 1) code = [...hit][0];
                }
                const ckey = code || rec.name;
                if (!custs.has(ckey)) custs.set(ckey, {code, name: rec.name, periods: [], monthlyMemo: false});
                const c = custs.get(ckey)!;
                c.name = rec.name; // 以最新月份的店名為準
                if ((rec.memo ?? '').split('；').some((x) => x.trim() === '月')) c.monthlyMemo = true;
                if (unsettled.has(`${k}|${half}`) || rec.amt <= 0) continue;
                const paid = rec.paid ?? 0;
                const owed = rec.amt - paid;
                if (owed <= FEE) continue;
                const who = rec.code || rec.name; // 同 compute 記已結清時用的鍵
                if (marks.settled.some((x) => (x.customer === who || x.customer === ckey) && x.keys.includes(periodKey({month_key: k, half})))) continue;
                c.periods.push({month_key: k, half, rkey, amt: rec.amt, paid, owed});
            }
        }
    }

    const monthName = (k: string): string => (multiYear ? `${+k.slice(0, 4)}年${+k.slice(5)}月` : `${+k.slice(5)}月`);
    const out: RegisterCustomer[] = [];
    for (const [ckey, c] of custs) {
        const monthly = marks.monthly?.[ckey] ?? debtors.get(ckey)?.monthly ?? c.monthlyMemo;
        const lines: RegisterLine[] = [];
        for (const p of c.periods) {
            const last = lines[lines.length - 1];
            if (monthly && last && last.month_key === p.month_key) {
                last.periods.push(p);
                last.amt += p.amt;
                last.paid += p.paid;
                last.owed += p.owed;
                last.id += `,${p.half}`;
                continue;
            }
            lines.push({
                id: `${p.month_key}|${p.half}`,
                label: monthName(p.month_key) + (monthly ? '' : HALF_LABEL[p.half]),
                month_key: p.month_key, periods: [p], amt: p.amt, paid: p.paid, owed: p.owed,
            });
        }
        out.push({ckey, code: c.code, name: c.name, monthly, lines, total: lines.reduce((t, l) => t + l.owed, 0)});
    }
    const isCode = (x: string): boolean => /^\d+$/.test(x);
    return out.sort((a, b) => {
        if (isCode(a.code) !== isCode(b.code)) return isCode(a.code) ? -1 : 1;
        if (isCode(a.code) && a.code !== b.code) return +a.code - +b.code;
        return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
    });
}

/** 搜尋代號或店名（忽略空白）；代號完全相同的排最前，其次是還有欠款的。 */
export function searchRegister(list: ReadonlyArray<RegisterCustomer>, keyword: string, limit = 8): RegisterCustomer[] {
    const kw = keyword.replace(/\s+/g, '');
    if (!kw) return [];
    const hit = list.filter((c) => c.code.includes(kw) || c.name.replace(/\s+/g, '').includes(kw));
    const rank = (c: RegisterCustomer): number => (c.code === kw ? 0 : c.lines.length ? 1 : 2);
    return hit.map((c, i) => ({c, i})).sort((a, b) => rank(a.c) - rank(b.c) || a.i - b.i).slice(0, limit).map((x) => x.c);
}

/** 一行收到的金額分給這行的各期：照順序補滿每期欠款，多的算在最後一期。 */
export function allocate(line: Pick<RegisterLine, 'periods'>, amount: number): number[] {
    let left = Math.round(amount);
    return line.periods.map((p, i) => {
        const share = i === line.periods.length - 1 ? left : Math.min(left, p.owed);
        left -= share;
        return share;
    });
}

/** 日期（input type=date 的 YYYY-MM-DD）轉成收款日格式 M/D。 */
export const toMd = (iso: string): string => {
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(iso);
    return m ? `${+m[2]}/${+m[3]}` : '';
};

export const todayIso = (d: Date = new Date()): string =>
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

/** 勾選的行（含各行實收金額）→ 依月份分好的寫回清單；金額 0 的期別不寫。 */
export function buildEntries(picks: ReadonlyArray<{line: RegisterLine; amount: number}>, dateIso: string): Map<string, PaymentEntry[]> {
    const date = toMd(dateIso);
    const byMonth = new Map<string, PaymentEntry[]>();
    for (const {line, amount} of picks) {
        allocate(line, amount).forEach((share, i) => {
            if (share <= 0) return;
            const p = line.periods[i];
            if (!byMonth.has(p.month_key)) byMonth.set(p.month_key, []);
            byMonth.get(p.month_key)!.push({key: p.rkey, half: p.half, amount: share, date, expectAmt: p.amt, expectPaid: p.paid});
        });
    }
    return byMonth;
}

/** 月份鍵在資料夾裡可能的檔名：先找當初匯入的檔名，再找對帳 2.0 預設的「X月_已自動對帳.xlsx」。 */
export function candidateFilenames(history: History, key: string): string[] {
    const src = history.months[key]?.source;
    return [...new Set([src, `${+key.slice(5)}月_已自動對帳.xlsx`].filter((x): x is string => !!x))];
}
