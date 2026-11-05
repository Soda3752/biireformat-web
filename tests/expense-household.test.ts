/**
 * 「支出」家用標記測試：bun test tests/expense-household.test.ts
 * 規則：「其他」項目預設全部算家用；被取消勾選（=店務）的名稱存在 `bii.expense.shop.v1`，同名項目跨月套用。
 * 實際範本以環境變數 EXPENSE_XLSX 指定，找不到就略過該段。
 */
import {beforeEach, describe, expect, test} from 'bun:test';
import {existsSync, readFileSync} from 'node:fs';

import {EXPENSE_CATEGORIES, type ExpenseCategory, type ExpenseItem, type MonthExpense, sumTotals} from '../src/expense/expense-model';
import {isHouseholdName, loadShopNames, setHousehold} from '../src/expense/expense-store';
import {summarizeMonth, summarizeYear} from '../src/expense/expense-summary';
import {parseExpenseWorkbook} from '../src/expense/expense-parser';
import {readExpenseGrids} from '../src/expense/expense-workbook';

const store = new Map<string, string>();
(globalThis as unknown as {localStorage: Storage}).localStorage = {
    getItem: (k: string) => store.get(k) ?? null,
    setItem: (k: string, v: string) => void store.set(k, v),
    removeItem: (k: string) => void store.delete(k),
    clear: () => store.clear(),
    key: (i: number) => [...store.keys()][i] ?? null,
    get length() {
        return store.size;
    },
} as Storage;

function month(m: number, other: Array<[string, number]>, shopAmount = 1000): MonthExpense {
    const totals = Object.fromEntries(EXPENSE_CATEGORIES.map((c) => [c, 0])) as Record<ExpenseCategory, number>;
    totals['營運'] = shopAmount;
    const items: ExpenseItem[] = [{category: '營運', name: '房租', amount: shopAmount, note: '', cell: 'I4'}];
    for (const [name, amount] of other) {
        items.push({category: '其他', name, amount, note: '', cell: 'T4'});
        totals['其他'] += amount;
    }
    return {year: 114, month: m, sortKey: `114-${String(m).padStart(2, '0')}`, sheetName: `${m}月`, totals, items, income: [], incomeTotal: 0};
}

beforeEach(() => store.clear());

describe('家用預設全勾', () => {
    test('沒取消任何項：「其他」全算家用', () => {
        const r = summarizeMonth(month(1, [['妹學費', 10000], ['保險', 3000]]), {sortKey: '114-01', amount: 50000} as never, loadShopNames());
        expect(r.household).toBe(13000);
        expect(r.shopExpense).toBe(1000);
        expect(r.totalExpense).toBe(14000);
        expect(r.shopBalance).toBe(49000);
        expect(r.allBalance).toBe(36000);
    });

    test('取消某項後該項變店務，他月同名（含前後空白）也套用；重新勾回又變家用', () => {
        setHousehold('保險', false);
        expect(JSON.parse(store.get('bii.expense.shop.v1')!)).toEqual(['保險']);
        const shop = loadShopNames();
        expect(isHouseholdName(' 保險 ', shop)).toBe(false);
        expect(isHouseholdName('新項目', shop)).toBe(true);
        const jan = summarizeMonth(month(1, [['妹學費', 10000], ['保險', 3000]]), undefined, shop);
        const feb = summarizeMonth(month(2, [['保險 ', 2000], ['新項目', 500]]), undefined, shop);
        expect(jan.household).toBe(10000);
        expect(jan.shopExpense).toBe(4000);
        expect(feb.household).toBe(500);
        expect(feb.shopExpense).toBe(3000);
        const y = summarizeYear([jan, feb]);
        expect(y.household).toBe(10500);
        expect(y.shopExpense).toBe(7000);

        setHousehold('保險', true);
        expect(loadShopNames().size).toBe(0);
    });

    test('舊 key（household.v1）不再讀取，不影響預設全勾', () => {
        store.set('bii.expense.household.v1', JSON.stringify(['妹學費']));
        const r = summarizeMonth(month(1, [['妹學費', 10000], ['保險', 3000]]), undefined, loadShopNames());
        expect(r.household).toBe(13000);
    });
});

const SAMPLE = process.env.EXPENSE_XLSX ?? '';

describe.skipIf(!SAMPLE || !existsSync(SAMPLE))('實際收支表範本（EXPENSE_XLSX）家用規則', () => {
    test('未取消時「其他」全算家用；取消一項後各月同名都變店務', async () => {
        const buf = readFileSync(SAMPLE);
        const grids = await readExpenseGrids(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
        const {months} = parseExpenseWorkbook(grids);
        const otherSum = (m: MonthExpense) => m.items.filter((it) => it.category === '其他').reduce((s, it) => s + it.amount, 0);

        for (const m of months) {
            const r = summarizeMonth(m, undefined, loadShopNames());
            expect(r.household).toBe(otherSum(m));
            expect(r.shopExpense).toBe(sumTotals(m.totals) - otherSum(m));
        }

        // 挑一個在最多月份出現的「其他」項目名稱來取消
        const count = new Map<string, number>();
        for (const m of months) for (const n of new Set(m.items.filter((it) => it.category === '其他').map((it) => it.name.trim()))) count.set(n, (count.get(n) ?? 0) + 1);
        const [target, appear] = [...count.entries()].sort((a, b) => b[1] - a[1])[0];
        expect(appear).toBeGreaterThan(1);
        setHousehold(target, false);
        const shop = loadShopNames();
        for (const m of months) {
            const targetAmt = m.items.filter((it) => it.category === '其他' && it.name.trim() === target).reduce((s, it) => s + it.amount, 0);
            const r = summarizeMonth(m, undefined, shop);
            expect(r.household).toBe(otherSum(m) - targetAmt);
        }
        console.log(`[範本] ${months.length} 個月，取消「${target}」（出現 ${appear} 個月）後各月家用皆正確扣除`);
    });
});
