/**
 * 「自動催款」：讀對帳 2.0 回填後的月結明細（X月_已自動對帳.xlsx）。
 *
 * 對應獨立工具 auto-collect/auto_collect/parser.py，檔案格式照 auto-reconcile-engine.ts 的 stmtLayout / writeBack：
 * - 每個月一份，含「上期」「下期」工作表（名稱可能帶空白；「上期_副本」等其他表不讀）。
 * - 標題列（前 5 列內含「客戶名稱」那列）可有左右多個區塊，每塊：代號／客戶名稱／金額／實收金額／收款日（／備註）。
 * - 對帳 2.0 只回填「實收金額」「收款日」（新版另寫「備註」）：
 *     實收金額有數字 → 有收到款（少 15 元＝被扣手續費）
 *     實收金額空白   → 沒對到任何入帳
 *     舊版：收款日「9/30收2065?」＝有一筆入帳金額對不上；收款日結尾「?」＝需確認
 *     新版：收款日只寫日期，標記寫在右邊的「備註」欄：「月」「扣手續費15」「2筆」「待確認」「9/30收2,065，金額對不上」
 */

import ExcelJS from 'exceljs';

export type Half = '上' | '下';

export const HALVES: ReadonlyArray<readonly [Half, string]> = [['上', '上期'], ['下', '下期']];

export interface StatementRecord {
    code: string;
    name: string;
    amt: number;
    /** 實收金額；空白（或非數字）為 null。 */
    paid: number | null;
    paid_text: string;
    /** 收款日欄原文（日期轉成 9/30）。 */
    date_text: string;
    /** 新版「備註」欄原文；舊版沒有這欄為空字串。 */
    memo: string;
}

export type Periods = Record<Half, Record<string, StatementRecord>>;

export interface Statement {
    month: number | null;
    periods: Periods;
}

export class NotStatementError extends Error {}

const JUNK_NAMES = /^(小計|總計|合計|總營收)$/;
const UNMATCHED_RE = /(\d{1,2}\/\d{1,2})收([\d,]+)(?:\?|，金額對不上)/g;
const MEMO_HEADER = '備註';

interface Block {
    code: number;
    labeled: boolean;
    name: number;
    amt: number;
    paid: number;
    date: number;
    memo: number;
}

type Raw = ExcelJS.CellValue | undefined;

const isFormula = (v: Raw): boolean =>
    !!v && typeof v === 'object' && !(v instanceof Date) && ('formula' in v || 'sharedFormula' in v);

/** 公式取快取結果（同 openpyxl data_only）。 */
function plain(v: Raw): unknown {
    if (v === undefined || v === null) return null;
    if (typeof v !== 'object' || v instanceof Date) return v;
    if (isFormula(v)) return plain((v as {result?: ExcelJS.CellValue}).result);
    if ('richText' in v && Array.isArray(v.richText)) return v.richText.map((t) => t.text).join('');
    if ('text' in v && typeof v.text === 'string') return v.text;
    if ('error' in v) return String(v.error);
    return String(v);
}

const s = (v: unknown): string => (v === null || v === undefined ? '' : String(v).trim());
const isNum = (v: unknown): v is number => typeof v === 'number' && Number.isFinite(v);

/** Python round()：.5 取偶數，讓金額跟獨立工具算的一模一樣。 */
export function pyRound(x: number): number {
    const r = Math.round(x);
    return Math.abs(x % 1) === 0.5 && r % 2 !== 0 ? r - 1 : r;
}

function dateText(v: unknown): string {
    if (v instanceof Date) return `${v.getUTCMonth() + 1}/${v.getUTCDate()}`;
    return s(v);
}

function sheetByName(wb: ExcelJS.Workbook, name: string): ExcelJS.Worksheet | undefined {
    return wb.worksheets.find((ws) => ws.name === name) ?? wb.worksheets.find((ws) => ws.name.trim() === name);
}

function layout(ws: ExcelJS.Worksheet): {header: number; blocks: Block[]} {
    for (let r = 1; r <= 5; r++) {
        const h = new Map<number, string>();
        ws.getRow(r).eachCell({includeEmpty: false}, (c, col) => {
            if (c.value !== null && c.value !== undefined) h.set(col, s(plain(c.value)));
        });
        if (!h.size) continue;
        const nameCols = [...h].filter(([, v]) => v === '客戶名稱').map(([c]) => c).sort((a, b) => a - b);
        if (!nameCols.length) continue;
        const end = Math.max(...h.keys()) + 1;
        const blocks = nameCols.map((nc, i) => {
            const stop = i + 1 < nameCols.length ? nameCols[i + 1] : end;
            const find = (label: string, def: number): number => {
                for (let c = nc + 1; c < stop; c++) if (h.get(c) === label) return c;
                return def;
            };
            const amt = find('金額', nc + 1);
            const paid = find('實收金額', amt + 1);
            const date = find('收款日', paid + 1);
            return {
                code: nc > 1 ? nc - 1 : 0,
                labeled: /編號|代號|代碼/.test(h.get(nc - 1) ?? ''),
                name: nc, amt, paid, date,
                memo: h.get(date + 1) === MEMO_HEADER ? date + 1 : 0,
            };
        });
        return {header: r, blocks};
    }
    // 舊版面：A/G=代號、B/H=名稱、C/I=金額、D/J=實收、E/K=收款日；對帳 2.0 新版把備註寫在收款日右邊（原本空白的那欄）
    return {
        header: 2,
        blocks: [2, 8].map((bc) => ({
            code: bc - 1, labeled: false, name: bc, amt: bc + 1, paid: bc + 2, date: bc + 3, memo: bc + 4,
        })),
    };
}

/** 從檔名（或工作表前兩列）找「9月」。 */
export function detectMonth(filename: string, wb?: ExcelJS.Workbook): number | null {
    const pick = (text: string): number | null => {
        const m = /(\d{1,2})\s*月/.exec(text);
        return m && +m[1] >= 1 && +m[1] <= 12 ? +m[1] : null;
    };
    const fromName = pick(filename || '');
    if (fromName) return fromName;
    if (!wb) return null;
    for (const [, sname] of HALVES) {
        const ws = sheetByName(wb, sname);
        if (!ws) continue;
        for (let r = 1; r <= 2; r++) {
            const row = ws.getRow(r);
            for (let c = 1; c <= row.cellCount; c++) {
                const m = pick(s(plain(row.getCell(c).value)));
                if (m) return m;
            }
        }
    }
    return null;
}

/** 解析一份月結明細。 */
export async function parseStatement(data: ArrayBuffer | Uint8Array, filename = ''): Promise<Statement> {
    const wb = new ExcelJS.Workbook();
    try {
        await wb.xlsx.load(data as ArrayBuffer);
    } catch (e) {
        throw new NotStatementError(`打不開這個檔案（不是 Excel .xlsx？）：${e instanceof Error ? e.message : e}`);
    }
    if (!sheetByName(wb, '上期') || !sheetByName(wb, '下期')) {
        throw new NotStatementError('這個檔案沒有「上期」「下期」工作表，不是對帳 2.0 的月結明細');
    }

    const periods = {} as Periods;
    for (const [half, sname] of HALVES) {
        const ws = sheetByName(wb, sname)!;
        const {header, blocks} = layout(ws);
        const rows: Array<[number, ExcelJS.Row]> = [];
        for (let r = header + 1; r <= ws.rowCount; r++) rows.push([r, ws.getRow(r)]);

        const raw = (row: ExcelJS.Row, col: number): Raw => (col > 0 ? row.getCell(col).value : null);
        const cell = (row: ExcelJS.Row, col: number): unknown => plain(raw(row, col));
        const filled = (b: Block, row: ExcelJS.Row): boolean => !!s(cell(row, b.name)) && isNum(cell(row, b.amt));
        const coded = (b: Block, row: ExcelJS.Row): boolean =>
            !!b.code && /^\d+$/.test(s(cell(row, b.code))) && filled(b, row);

        let nCoded = 0;
        let nFilled = 0;
        for (const b of blocks) {
            for (const [, row] of rows) {
                if (coded(b, row)) nCoded++;
                if (filled(b, row)) nFilled++;
            }
        }
        const hasCode = blocks.some((b) => b.labeled) || nCoded * 2 >= nFilled;

        const out: Record<string, StatementRecord> = {};
        for (const b of blocks) {
            for (const [, row] of rows) {
                const name = s(cell(row, b.name));
                const amt = cell(row, b.amt);
                if (!name || !isNum(amt)) continue;
                // 名稱整個是數字＝版面角落的統計格（公式結果），不是客人（10 月下期右下角就有一整塊）
                if (/^\d+$/.test(name)) continue;
                let code = '';
                let key: string;
                if (hasCode) {
                    if (!coded(b, row)) continue;
                    code = s(cell(row, b.code));
                    key = code;
                } else {
                    if (JUNK_NAMES.test(name)) continue;
                    if (isFormula(raw(row, b.amt))) continue;
                    key = '名:' + name;
                }
                const paidRaw = cell(row, b.paid);
                out[key] = {
                    code,
                    name,
                    amt: pyRound(amt),
                    paid: isNum(paidRaw) ? pyRound(paidRaw) : null,
                    paid_text: isNum(paidRaw) ? '' : s(paidRaw),
                    date_text: dateText(cell(row, b.date)),
                    memo: b.memo ? s(cell(row, b.memo)) : '',
                };
            }
        }
        periods[half] = out;
    }
    return {month: detectMonth(filename, wb), periods};
}

/** 金額對不上、沒算進實收的入帳（舊版收款日「9/30收2065?」、新版備註「9/30收2,065，金額對不上」）→ [日期, 金額]。 */
export function unmatchedReceipts(rec: Pick<StatementRecord, 'date_text' | 'memo'>): Array<[string, number]> {
    const text = [rec.date_text, rec.memo].filter(Boolean).join('；');
    return [...text.matchAll(UNMATCHED_RE)].map((m) => [m[1], parseInt(m[2].replace(/,/g, ''), 10)]);
}

/** 對帳 2.0 標「需確認」（舊版收款日結尾「?」、新版備註「待確認」）。 */
export function needsConfirm(rec: Pick<StatementRecord, 'date_text' | 'memo'>): boolean {
    return rec.date_text.trimEnd().endsWith('?') || rec.memo.split('；').includes('待確認');
}
