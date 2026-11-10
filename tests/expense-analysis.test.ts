/**
 * 「支出」結餘三層＋分析區塊計算測試：bun test tests/expense-analysis.test.ts
 * 實際範本以環境變數 EXPENSE_XLSX 指定（搭配模擬營收），找不到就略過該段。
 */
import {describe, expect, test} from 'bun:test';
import {existsSync, readFileSync} from 'node:fs';

import {EXPENSE_CATEGORIES, type ExpenseCategory, type ExpenseItem, type MonthExpense, sumTotals} from '../src/expense/expense-model';
import {type MonthSummary, summarizeMonth, summarizeYear} from '../src/expense/expense-summary';
import {
    breakeven,
    detectAnomalies,
    householdShares,
    revenueShares,
    rollingMargins,
} from '../src/expense/expense-analysis';
import {parseExpenseWorkbook} from '../src/expense/expense-parser';
import {readExpenseGrids} from '../src/expense/expense-workbook';

const NO_SHOP = new Set<string>();

interface Spec {
    m: number;
    rev?: number;
    原物料?: number;
    營運?: number;
    車輛?: number;
    人事?: number;
    銀行?: number;
    /** 「其他」項目：家用（預設全勾） */
    家用?: number;
    /** 「其他」項目：取消勾選＝店務 */
    其他店務?: number;
}

function row(s: Spec): MonthSummary {
    const totals = Object.fromEntries(EXPENSE_CATEGORIES.map((c) => [c, 0])) as Record<ExpenseCategory, number>;
    const items: ExpenseItem[] = [];
    for (const c of ['原物料', '營運', '車輛', '人事', '銀行'] as const) {
        totals[c] = s[c] ?? 0;
        if (totals[c]) items.push({category: c, name: c, amount: totals[c], note: '', cell: 'A1'});
    }
    if (s.家用) items.push({category: '其他', name: '學費', amount: s.家用, note: '', cell: 'T4'});
    if (s.其他店務) items.push({category: '其他', name: '店用雜支', amount: s.其他店務, note: '', cell: 'T5'});
    totals['其他'] = (s.家用 ?? 0) + (s.其他店務 ?? 0);
    const exp = {
        year: 115, month: s.m, sortKey: `115-${String(s.m).padStart(2, '0')}`, sheetName: `${s.m}月`,
        totals, items, income: [], incomeTotal: 0, warnings: [], sourceFile: 't', uploadedAt: 0,
    } as MonthExpense;
    return summarizeMonth(exp, s.rev === undefined ? undefined : {amount: s.rev, fileNames: []}, new Set(['店用雜支']));
}

describe('結餘三層', () => {
    test('營業費用不含還款、家用；三層結餘依序扣', () => {
        const r = row({m: 1, rev: 1_000_000, 原物料: 400_000, 營運: 100_000, 車輛: 20_000, 人事: 200_000, 銀行: 120_000, 家用: 50_000, 其他店務: 10_000});
        expect(r.repayment).toBe(120_000);
        expect(r.otherShop).toBe(10_000);
        expect(r.operatingExpense).toBe(730_000);
        expect(r.operatingBalance).toBe(270_000);
        expect(r.afterRepayBalance).toBe(150_000);
        expect(r.netBalance).toBe(100_000);
        expect(r.netBalance).toBe(r.revenue! - r.totalExpense);
        // 營業結餘 = 舊店務結餘（營收 − 全部支出 + 家用）+ 銀行
        expect(r.operatingBalance).toBe(r.revenue! - r.shopExpense + r.repayment);
    });

    test('沒營收：三層都是 null；年度只累計有營收月份', () => {
        const a = row({m: 1, rev: 100, 營運: 10, 銀行: 5, 家用: 1});
        const b = row({m: 2, 營運: 10, 銀行: 5});
        expect(b.operatingBalance).toBeNull();
        expect(b.netBalance).toBeNull();
        const y = summarizeYear([a, b]);
        expect(y.repayment).toBe(10);
        expect(y.operatingExpense).toBe(20);
        expect(y.operatingBalance).toBe(90);
        expect(y.afterRepayBalance).toBe(85);
        expect(y.netBalance).toBe(84);
    });
});

describe('分析', () => {
    const rows = [
        row({m: 1, rev: 1000, 原物料: 400, 人事: 200, 營運: 100, 車輛: 50, 其他店務: 10, 銀行: 100, 家用: 40}),
        row({m: 2, rev: 800, 原物料: 480, 人事: 200, 營運: 100, 車輛: 50, 其他店務: 10, 銀行: 100, 家用: 40}),
        row({m: 3, rev: 1200, 原物料: 360, 人事: 200, 營運: 150, 車輛: 50, 其他店務: 10, 銀行: 100, 家用: 40}),
        row({m: 4, 原物料: 999, 營運: 99_999}), // 沒營收：分析全部略過
    ];

    test('各項占營收 %：每月與年度，含原物料＋人事', () => {
        const {months, year} = revenueShares(rows);
        expect(months.map((r) => r.month)).toEqual([1, 2, 3]);
        expect(months[0].pct['原物料']).toBeCloseTo(0.4);
        expect(months[0].pct['還款']).toBeCloseTo(0.1);
        expect(months[0].pct['營業結餘']).toBeCloseTo(0.24);
        expect(months[0].materialLabor).toBeCloseTo(0.6);
        expect(months[1].materialLabor).toBeCloseTo(0.85);
        expect(year!.revenue).toBe(3000);
        expect(year!.pct['原物料']).toBeCloseTo(1240 / 3000);
        // 各項 + 實際剩下 = 100%
        const sumExceptBalance = (['原物料', '人事', '營運', '車輛', '其他店務'] as const).reduce((s, k) => s + months[0].pct[k]!, 0);
        expect(sumExceptBalance + months[0].pct['營業結餘']!).toBeCloseTo(1);
    });

    test('近 3 個月平均毛利率＝合計毛利 ÷ 合計營收，年初不足 3 個月照用', () => {
        const m = rollingMargins(rows);
        expect(m).toHaveLength(3);
        expect(m[0].average).toBeCloseTo(0.6);
        expect(m[0].monthsUsed).toBe(1);
        expect(m[1].margin).toBeCloseTo(0.4);
        expect(m[1].average).toBeCloseTo((600 + 320) / 1800);
        expect(m[2].average).toBeCloseTo((600 + 320 + 840) / 3000);
        expect(m[2].monthsUsed).toBe(3);
    });

    test('滾動視窗依日曆月份：中間缺帳單的月份不往前多抓', () => {
        const gap = [row({m: 1, rev: 100, 原物料: 50}), row({m: 2}), row({m: 3, rev: 100, 原物料: 10}), row({m: 4, rev: 100, 原物料: 10})];
        const m = rollingMargins(gap);
        expect(m.map((x) => x.monthsUsed)).toEqual([1, 2, 2]);
        expect(m[2].average).toBeCloseTo(0.9); // 4 月只看 2~4 月（2 月沒營收）
    });

    test('支出異常：比該類其他月份平均高 20% 以上才列', () => {
        const a = detectAnomalies(rows);
        // 3 月營運 150，其他月份平均 100 → 高 50%
        expect(a).toEqual([{sortKey: '115-03', month: 3, key: '營運', amount: 150, average: 100, over: 0.5}]);
        // 剛好 20% 也算
        const edge = detectAnomalies([row({m: 1, rev: 1, 車輛: 120}), row({m: 2, rev: 1, 車輛: 100})]);
        expect(edge.map((x) => x.key)).toEqual(['車輛']);
        expect(detectAnomalies([row({m: 1, rev: 1, 車輛: 119}), row({m: 2, rev: 1, 車輛: 100})])).toEqual([]);
        expect(detectAnomalies([rows[0]])).toEqual([]);
    });

    test('損益兩平：固定支出 ÷ 年度毛利率，再加還款、再加家用', () => {
        const b = breakeven(rows)!;
        expect(b.months).toBe(3);
        expect(b.grossMargin).toBeCloseTo(1760 / 3000);
        expect(b.avgRevenue).toBeCloseTo(1000);
        expect(b.avgFixed).toBeCloseTo((360 + 360 + 410) / 3);
        expect(b.operating).toBeCloseTo(b.avgFixed / b.grossMargin!);
        expect(b.withRepayment).toBeCloseTo((b.avgFixed + 100) / b.grossMargin!);
        expect(b.withHousehold).toBeCloseTo((b.avgFixed + 140) / b.grossMargin!);
        // 毛利率 ≤ 0：算不出
        const loss = breakeven([row({m: 1, rev: 100, 原物料: 150, 營運: 10})])!;
        expect(loss.operating).toBeNull();
        expect(breakeven([row({m: 1})])).toBeNull();
    });

    test('家用占還款後結餘：負值月份為 null，年度用合計', () => {
        const list = [
            row({m: 1, rev: 1000, 營運: 500, 銀行: 300, 家用: 100}), // 還款後 200 → 50%
            row({m: 2, rev: 500, 營運: 500, 銀行: 300, 家用: 100}), // 還款後 −300 → null
        ];
        const {months, year} = householdShares(list);
        expect(months[0].ratio).toBeCloseTo(0.5);
        expect(months[1].afterRepayBalance).toBe(-300);
        expect(months[1].ratio).toBeNull();
        expect(year!.afterRepayBalance).toBe(-100);
        expect(year!.ratio).toBeNull();
    });
});

const SAMPLE = process.env.EXPENSE_XLSX ?? '';

describe.skipIf(!SAMPLE || !existsSync(SAMPLE))('實際收支表範本（EXPENSE_XLSX）＋模擬營收', () => {
    test('營業結餘 = 舊店務結餘 + 銀行；實際剩下 = 舊含家用結餘；分析各項可算', async () => {
        const buf = readFileSync(SAMPLE);
        const grids = await readExpenseGrids(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
        const {months} = parseExpenseWorkbook(grids);
        expect(months.length).toBeGreaterThan(1);
        const rows = months.map((m, i) => summarizeMonth(m, {amount: 3_000_000 + i * 100_000, fileNames: []}, NO_SHOP));
        for (const r of rows) {
            const oldShopBalance = r.revenue! - (sumTotals(r.totals) - r.household);
            const oldAllBalance = r.revenue! - sumTotals(r.totals);
            expect(r.operatingBalance).toBe(oldShopBalance + r.totals['銀行']);
            expect(r.netBalance).toBe(oldAllBalance);
            expect(r.afterRepayBalance! - r.household).toBe(r.netBalance!);
        }
        const y = summarizeYear(rows);
        expect(y.operatingBalance - y.repayment - y.household).toBe(y.netBalance);
        const b = breakeven(rows)!;
        expect(b.months).toBe(rows.length);
        console.log(
            `[範本] ${rows.length} 個月；年度營業結餘 ${Math.round(y.operatingBalance)}、還款 ${y.repayment}、還款後 ${Math.round(y.afterRepayBalance)}、實際剩下 ${Math.round(y.netBalance)}；` +
            `毛利率 ${(b.grossMargin! * 100).toFixed(1)}%、營業不虧 ${Math.round(b.operating ?? 0)}、要還款 ${Math.round(b.withRepayment ?? 0)}、連家用 ${Math.round(b.withHousehold ?? 0)}；` +
            `異常 ${detectAnomalies(rows).map((a) => `${a.month}月${a.key}+${Math.round(a.over * 100)}%`).join('、') || '無'}`,
        );
    });
});
