/**
 * 「支出」收支表解析測試：bun test tests/expense-parser.test.ts
 * 實際範本檔不放進 repo；以環境變數 EXPENSE_XLSX 指定路徑，找不到就略過。
 */
import {describe, expect, test} from 'bun:test';
import {existsSync, readFileSync} from 'node:fs';

import {type GridCell, parseExpenseWorkbook, parseMonthSheet, type SheetGrid} from '../src/expense/expense-parser';
import {readExpenseGrids} from '../src/expense/expense-workbook';
import {EXPENSE_CATEGORIES, sumTotals} from '../src/expense/expense-model';

function grid(name: string, cells: Record<string, string | number | { f: string; v?: number }>): SheetGrid {
    const map = new Map<string, GridCell>();
    let rowCount = 0;
    let colCount = 0;
    for (const [addr, raw] of Object.entries(cells)) {
        const m = /^([A-Z]+)(\d+)$/.exec(addr)!;
        const col = [...m[1]].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
        const row = Number(m[2]);
        rowCount = Math.max(rowCount, row);
        colCount = Math.max(colCount, col);
        map.set(`${row}:${col}`, typeof raw === 'object' ? {value: raw.v ?? null, formula: raw.f} : {value: raw});
    }
    return {name, rowCount, colCount, get: (r, c) => map.get(`${r}:${c}`) ?? null};
}

describe('parseMonthSheet', () => {
    test('以標題＋SUM 小計辨識區塊，子標題（銀行／人事）另開小節', () => {
        const g = grid('1月', {
            A1: '114/1月',
            A3: '彰化', B3: 100, D3: '原物料', H3: '營運', L3: '車輛', P3: '人事', S3: '其他',
            A4: '和美', B4: 50, D4: '麵粉', E4: 1000, F4: '權信', G4: '10月',
            H4: '房租', I4: 500,
            L4: '車油', M4: 300, L5: '修車', M5: 200, M6: {f: 'SUM(M4:M5)', v: 500},
            L7: '人事', L8: '紅包', M8: 20, N8: '阿光', M9: {f: 'SUM(M8:M8)'},
            P4: '薪水', Q4: 42000, R4: '阿光', Q5: 1000, R5: '彰化愛買', Q6: {f: 'SUM(Q4:Q5)', v: 43000},
            S4: '妹學費', T4: 10000, T5: {f: 'SUM(T4:T4)', v: 10000},
            S7: '銀行', S8: '台新', T8: 5000, U8: '房車貸', T9: {f: 'SUM(T8:T8)', v: 5000},
            A6: '合計', B6: {f: 'SUM(B3:B4)', v: 150},
            E7: {f: 'SUM(E4:E6)', v: 1000}, I7: {f: 'SUM(I4:I6)', v: 500},
        });
        const m = parseMonthSheet(g, {fallbackYear: 0, sourceFile: 't.xlsx'});
        expect(m.sortKey).toBe('114-01');
        expect(m.totals).toEqual({原物料: 1000, 營運: 500, 車輛: 500, 人事: 43020, 其他: 10000, 銀行: 5000});
        expect(m.warnings).toEqual([]);
        expect(m.incomeTotal).toBe(150);
        expect(m.items.find((it) => it.cell === 'Q5')?.name).toBe('彰化愛買');
        expect(m.items.find((it) => it.cell === 'E4')?.note).toBe('權信・10月');
        expect(m.items.find((it) => it.cell === 'M8')?.category).toBe('人事');
    });
});

const SAMPLE = process.env.EXPENSE_XLSX ?? '';

describe.skipIf(!SAMPLE || !existsSync(SAMPLE))('實際收支表範本（EXPENSE_XLSX）', () => {
    test('1~9 月各類金額與年度總表一致', async () => {
        const buf = readFileSync(SAMPLE);
        const grids = await readExpenseGrids(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
        const res = parseExpenseWorkbook(grids, 'sample.xlsx', 114);
        expect(res.errors).toEqual([]);
        expect(res.hasSummary).toBe(true);
        expect(res.months.map((m) => m.month)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9]);
        expect(res.summaryMismatches).toEqual([]);
        for (const m of res.months) {
            for (const c of EXPENSE_CATEGORIES) expect(Number.isFinite(m.totals[c])).toBe(true);
        }
        const jan = res.months[0];
        expect(jan.totals).toEqual({原物料: 1075325, 營運: 314746, 車輛: 168690, 人事: 796753, 其他: 147508, 銀行: 105000});
        expect(sumTotals(jan.totals)).toBe(2608022);
        expect(res.months.reduce((s, m) => s + sumTotals(m.totals), 0)).toBe(23594092);
    });
});
