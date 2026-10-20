/**
 * 「自動催款」計算邏輯測試：bun test tests/
 * 對應獨立工具 auto-collect/tests/test_core.py，另加對帳 2.0 新格式（收款日右邊插「備註」欄）。
 */
import {describe, expect, test} from 'bun:test';
import {existsSync, readFileSync} from 'node:fs';

import ExcelJS from 'exceljs';

import {addReasonChip, defaultReasonChips, normalizeReasonChips, removeReasonChip} from '../src/domain/auto-collect/reason-chips';
import {expandAll, normalizeUiPrefs} from '../src/domain/auto-collect/ui-prefs';
import {closingFor, DEFAULT_TEMPLATES, defaultTemplates, normalizeTemplates} from '../src/domain/auto-collect/templates';

import {
    addManualDebt,
    buildMessage,
    clearMessageEdit,
    compute,
    emptyHistory,
    emptyMarks,
    filterCustomers,
    messageBody,
    guessYear,
    type History,
    type ImportFile,
    importFiles,
    KEEP,
    extractLegacyRules,
    migrateLegacyRules,
    normalizeMarks,
    removeManualDebt,
    setMessageEdit,
    setMonthEnd,
    setMonthly,
    settleDebtor,
    unsettle,
    updateManualDebt,
    validateManualDebt,
    type Choices,
} from '../src/domain/auto-collect/ledger';
import {parseStatement, type StatementRecord} from '../src/domain/auto-collect/statement-parser';
import {buildMessagesText, buildSummaryXlsx} from '../src/writers/auto-collect-writer';

const TODAY = new Date(2026, 9, 18);

type Row = [string | null, string, number, number | null, string | null, string?];

/**
 * upper/lower = [代號, 店名, 金額, 實收, 收款日, 備註?]，照對帳 2.0 版面（左右兩塊）。
 * memo: true → 新格式，各區塊收款日右邊多一欄「備註」（紅字用 richText）。
 */
async function makeStmt(upper: Row[], lower: Row[], memo = false): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    for (const [title, rows] of [['上期', upper], ['下期 ', lower]] as const) {
        const ws = wb.addWorksheet(title);
        ws.addRow([null]);
        const head = ['代號', '客戶名稱', '金額', '實收金額', '收款日', ...(memo ? ['備註'] : []), null];
        ws.addRow([...head, ...head]);
        for (const r of rows) {
            const vals: unknown[] = [r[0], r[1], r[2], r[3], r[4]];
            if (memo) vals.push(r[5] ? {richText: [{text: r[5], font: {color: {argb: 'FFC00000'}}}]} : null);
            ws.addRow(vals);
        }
        const total = ws.addRow([null, '總計']);
        total.getCell(3).value = {formula: 'SUM(C3:C9)', result: 0};
    }
    wb.addWorksheet('上期_副本').addRow(['代號', '客戶名稱', '金額']);
    return new Uint8Array(await wb.xlsx.writeBuffer());
}

const file = (name: string, mtime: number, data: Uint8Array): ImportFile => ({name, mtime, data});

const run = (files: ImportFile[], history: History = emptyHistory(), choices?: Choices) =>
    importFiles(files, history, {today: TODAY, choices});

describe('欠款計算', () => {
    test('未付、付清、手續費、部分付款、多收、對不上的入帳、跨月加總', async () => {
        const aug = await makeStmt(
            [['1001', '甲店', 1000, null, null], // 未付
                ['1002', '乙店', 2000, 2000, '8/20'], // 付清
                ['1003', '丙店', 3000, 2985, '8/20(-15)']], // 扣手續費 = 付清
            [['1001', '甲店', 1100, 600, '9/1'], // 部分付款，差 500
                ['1002', '乙店', 2100, null, '9/2收999?'], // 未付＋對不上的入帳
                ['1003', '丙店', 3100, 3500, '9/3']], // 多付
        );
        const sep = await makeStmt(
            [['1001', '甲店新名', 1200, null, null], ['1002', '乙店', 2200, 2200, '9/20?']],
            [['1001', '甲店新名', 1300, 1300, '10/1'], ['1003', '丙店', 0, null, null]],
        );
        const {history, report} = await run([file('8月_已自動對帳.xlsx', 1, aug), file('9月_已自動對帳.xlsx', 2, sep)]);
        expect(report.map((r) => r.status)).toEqual(['used', 'used']);
        const res = compute(history);
        const by = Object.fromEntries(res.customers.map((c) => [c.code, c]));
        expect(Object.keys(by).sort()).toEqual(['1001', '1002']);
        const a = by['1001'];
        expect(a.name).toBe('甲店新名');
        expect(a.items.map((i) => [i.period, i.owed, i.partial])).toEqual(
            [['8月上期', 1000, false], ['8月下期', 500, true], ['9月上期', 1200, false]]);
        expect(a.total).toBe(2700);
        expect(a.message).toBe([
            '甲店新名 您好～',
            '提醒您以下帳款尚未收到：',
            '・8月上期：1,000 元',
            '・8月下期：500 元（帳單 1,100 元，已收 600 元）',
            '・9月上期：1,200 元',
            '合計：2,700 元',
            '',
            '帳款已逾期，再麻煩儘早安排匯款，若已經匯款請忽略此訊息，謝謝！',
        ].join('\n'));
        const b = by['1002'];
        expect(b.total).toBe(2100);
        expect(b.notes.some((n) => n.includes('9/2 999'))).toBe(true);
        expect(res.grand_total).toBe(4800);
        // 丙店多付的不會拿去抵，但也不欠 → 不在名單
        expect(by['1003']).toBeUndefined();
    });

    test('差 16 元算欠款、差 15 元算手續費', async () => {
        const st = await makeStmt([['1001', '甲店', 1000, 984, '9/1'], ['1002', '乙店', 1000, 985, '9/1']], []);
        const {history} = await run([file('9月.xlsx', 1, st)]);
        const res = compute(history);
        expect(res.customers.map((c) => [c.code, c.total, c.items[0].partial])).toEqual([['1001', 16, true]]);
    });

    test('沒有代號的列（零售、小計）不算', async () => {
        const st = await makeStmt([['1001', '甲店', 1000, null, null], [null, '零售', 500, null, null],
            [null, '小計', 1500, null, null]], []);
        const {history} = await run([file('9月.xlsx', 1, st)]);
        expect(compute(history).customers.map((c) => c.name)).toEqual(['甲店']);
    });

    test('跨年時期別帶年份', async () => {
        const dec = await makeStmt([['1001', '甲店', 100, null, null]], []);
        const jan = await makeStmt([['1001', '甲店', 200, null, null]], []);
        const {history} = await importFiles([file('12月.xlsx', 1, dec), file('1月.xlsx', 2, jan)], emptyHistory(),
            {today: new Date(2027, 0, 5)});
        expect(compute(history).customers[0].items.map((i) => i.period)).toEqual(['2026年12月上期', '2027年1月上期']);
    });
});

describe('需確認的客人', () => {
    // 1002 有對不上的入帳、1004 多收 → 有注意事項；1001、1003、1005 沒有
    const build = async () => {
        const st = await makeStmt(
            [['1001', '甲店', 1000, null, null], ['1002', '乙店', 2000, null, '9/2收999?'],
                ['1003', '丙店', 3000, null, null], ['1004', '丁店', 4000, 4500, '9/3'], ['1005', '戊店', 500, null, null]],
            [['1004', '丁店', 100, null, null]],
        );
        const {history} = await run([file('9月.xlsx', 1, st)]);
        return compute(history);
    };

    test('有注意事項的排最上面，其餘照代號', async () => {
        const res = await build();
        expect(res.customers.map((c) => [c.code, c.notes.length > 0])).toEqual(
            [['1002', true], ['1004', true], ['1001', false], ['1003', false], ['1005', false]]);
        expect(res.grand_total).toBe(1000 + 2000 + 3000 + 500 + 100);
    });

    test('只看需確認、可和搜尋一起用', async () => {
        const {customers} = await build();
        const codes = (cs: typeof customers) => cs.map((c) => c.code);
        expect(codes(filterCustomers(customers, '', true))).toEqual(['1002', '1004']);
        expect(codes(filterCustomers(customers, '', false))).toEqual(['1002', '1004', '1001', '1003', '1005']);
        expect(codes(filterCustomers(customers, '丁', true))).toEqual(['1004']);
        expect(codes(filterCustomers(customers, ' 甲 ', true))).toEqual([]);
        expect(codes(filterCustomers(customers, ' 甲 ', false))).toEqual(['1001']);
        expect(codes(filterCustomers(customers, '100', true))).toEqual(['1002', '1004']);
    });
});

describe('對帳 2.0 新格式（收款日右邊的備註欄）', () => {
    test('讀得到金額與備註，規則和舊格式一樣', async () => {
        const st = await makeStmt(
            [['1001', '甲店', 1000, 985, '9/1', '扣手續費15'],
                ['1002', '乙店', 2000, null, null, '9/30收2,065，金額對不上'],
                ['1003', '丙店', 3000, 1000, '9/5', '月；2筆；待確認'],
                ['1004', '丁店', 500, 500, '9/6', '月']],
            [], true);
        const parsed = await parseStatement(st, '9月_已自動對帳.xlsx');
        expect(parsed.periods.上['1002']).toMatchObject({amt: 2000, paid: null, date_text: '', memo: '9/30收2,065，金額對不上'});
        expect(parsed.periods.上['1003']).toMatchObject({paid: 1000, date_text: '9/5', memo: '月；2筆；待確認'});
        const {history} = await run([file('9月_已自動對帳.xlsx', 1, st)]);
        const res = compute(history);
        const by = Object.fromEntries(res.customers.map((c) => [c.code, c]));
        expect(Object.keys(by)).toEqual(['1002', '1003']);
        expect(by['1002'].notes).toEqual(['9月上期 有 1 筆入帳金額對不上、沒算進已收（9/30 2,065 元），催款前請先確認']);
        expect(by['1003'].notes).toEqual(['9月上期 的收款對帳 2.0 標「需確認」']);
        expect(by['1003'].items[0]).toMatchObject({owed: 2000, partial: true});
        expect(res.grand_total).toBe(4000);
    });

    test('同樣的資料新舊格式算出來一樣', async () => {
        const oldF = await makeStmt([['1001', '甲店', 1000, 600, '9/1(-15)?'], ['1002', '乙店', 800, null, '9/2收700?']],
            [['1001', '甲店', 1200, null, null]]);
        const newF = await makeStmt([['1001', '甲店', 1000, 600, '9/1', '扣手續費15；待確認'],
            ['1002', '乙店', 800, null, null, '9/2收700，金額對不上']], [['1001', '甲店', 1200, null, null]], true);
        const a = compute((await run([file('9月.xlsx', 1, oldF)])).history);
        const b = compute((await run([file('9月.xlsx', 1, newF)])).history);
        const strip = (r: typeof a) => r.customers.map((c) => [c.code, c.total, c.notes, c.message]);
        expect(strip(b)).toEqual(strip(a));
        expect(a.grand_total).toBe(2400);
    });
});

describe('備註欄與實收欄文字列為需確認', () => {
    test('例行標記不列；非例行備註、例行＋其他混合、實收欄非數字會列；欠款金額不變', async () => {
        const rows: Row[] = [
            ['1001', '甲店', 1000, 985, '9/1', '扣手續費15；月；2筆'], // 全例行、付清 → 不在名單
            ['1002', '乙店', 2000, null, null, '月；2筆'], // 全例行、未付 → 欠款但無注意
            ['1003', '丙店', 3000, null, null, '客人說月底一起付'], // 非例行
            ['1004', '丁店', 4000, null, null, '月；跟1005合併付'], // 例行＋其他混合
            ['1005', '戊店', 500, null, '9/5', undefined],
            ['1006', '己店', 600, 600, '9/6', '老闆娘付現'], // 付清但備註非例行
            ['1007', '庚店', 700, null, null, '9/30收2,065，金額對不上；月'], // 已被金額對不上抓到，不重複
        ];
        const st = await makeStmt(rows, [['1006', '己店', 100, null, null]], true);
        const {history} = await run([file('9月_已自動對帳.xlsx', 1, st)]);
        const res = compute(history);
        const by = Object.fromEntries(res.customers.map((c) => [c.code, c]));
        expect(by['1001']).toBeUndefined();
        expect(by['1002'].notes).toEqual([]);
        expect(by['1003'].notes).toEqual(['9月上期 備註：客人說月底一起付，催款前請先確認']);
        expect(by['1004'].notes).toEqual(['9月上期 備註：跟1005合併付，催款前請先確認']);
        expect(by['1006'].notes).toEqual(['9月上期 備註：老闆娘付現，催款前請先確認']);
        expect(by['1006'].items.map((i) => [i.period, i.owed])).toEqual([['9月下期', 100]]);
        expect(by['1007'].notes).toEqual(['9月上期 有 1 筆入帳金額對不上、沒算進已收（9/30 2,065 元），催款前請先確認']);
        expect(res.grand_total).toBe(2000 + 3000 + 4000 + 500 + 100 + 700);
        expect(res.customers.map((c) => [c.code, c.total])).toEqual(
            [['1003', 3000], ['1004', 4000], ['1006', 100], ['1007', 700], ['1002', 2000], ['1005', 500]]);
    });

    test('實收欄寫字不是金額：列需確認，仍整筆算欠', async () => {
        const st = await makeStmt([['1001', '甲店', 1000, null, '9/1'], ['1002', '乙店', 800, null, null]], [], true);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(st.buffer as ArrayBuffer);
        wb.getWorksheet('上期')!.getCell('D3').value = '已付現';
        const edited = new Uint8Array(await wb.xlsx.writeBuffer());
        const res = compute((await run([file('9月.xlsx', 1, edited)])).history);
        expect(res.customers.map((c) => [c.code, c.total, c.notes])).toEqual([
            ['1001', 1000, ['9月上期 實收欄寫「已付現」，不是金額，催款前請先確認']],
            ['1002', 800, []],
        ]);
    });

    test('待確認已列「需確認」時不重複；沒列時當成要確認的備註', async () => {
        const st = await makeStmt([['1001', '甲店', 1000, 500, '9/1', '月；待確認'],
            ['1002', '乙店', 1000, null, null, '待確認']], [], true);
        const res = compute((await run([file('9月.xlsx', 1, st)])).history);
        const by = Object.fromEntries(res.customers.map((c) => [c.code, c.notes]));
        expect(by['1001']).toEqual(['9月上期 的收款對帳 2.0 標「需確認」']);
        expect(by['1002']).toEqual(['9月上期 備註：待確認，催款前請先確認']);
    });

    test('舊格式（沒有備註欄）不變', async () => {
        const st = await makeStmt([['1001', '甲店', 1000, 600, '9/1(-15)?'], ['1002', '乙店', 800, null, '9/2收700?'],
            ['1003', '丙店', 900, null, null], ['1004', '丁店', 400, 400, '9/4(月)']], []);
        const res = compute((await run([file('9月.xlsx', 1, st)])).history);
        expect(res.customers.map((c) => [c.code, c.total, c.notes])).toEqual([
            ['1001', 400, ['9月上期 的收款對帳 2.0 標「需確認」']],
            ['1002', 800, ['9月上期 有 1 筆入帳金額對不上、沒算進已收（9/2 700 元），催款前請先確認']],
            ['1003', 900, []],
        ]);
    });
});

describe('同月多份與已記住的月份', () => {
    test('同一批同月多份：不自動挑，要使用者選', async () => {
        const a = await makeStmt([['1001', '甲店', 1000, null, null]], []);
        const b = await makeStmt([['1001', '甲店', 1000, 1000, '9/5']], [['1001', '甲店', 500, null, null]]);
        const files = [file('9月_已自動對帳.xlsx', 5, b), file('9月_已自動對帳 (1).xlsx', 3, a), file('8月_已自動對帳.xlsx', 1, a)];
        let r = await run(files);
        // 沒選之前整批都不匯入，也不會自動挑修改時間新的
        expect(r.history.months).toEqual({});
        expect(r.conflicts.length).toBe(1);
        const c = r.conflicts[0];
        expect(c.month_key).toBe('2026-09');
        expect(c.options.map((o) => [o.value, o.kind, o.name]))
            .toEqual([[0, 'new', '9月_已自動對帳.xlsx'], [1, 'new', '9月_已自動對帳 (1).xlsx']]);
        const {upper, lower, paid, debtors, owed_total} = c.options[0];
        expect({upper, lower, paid, debtors, owed_total}).toEqual({upper: 1, lower: 1, paid: 1, debtors: 1, owed_total: 500});
        expect(c.options[1].owed_total).toBe(1000);
        // 選舊的那份（修改時間較早）也照選擇
        r = await run(files, emptyHistory(), {'2026-09': 1});
        expect(r.conflicts).toEqual([]);
        expect(Object.fromEntries(r.report.map((x) => [x.name, x.status]))).toEqual({
            '9月_已自動對帳.xlsx': 'skipped', '9月_已自動對帳 (1).xlsx': 'used', '8月_已自動對帳.xlsx': 'used',
        });
        expect(r.history.months['2026-09'].source).toBe('9月_已自動對帳 (1).xlsx');
        expect(compute(r.history).grand_total).toBe(2000);
        // 不屬於這個月的選擇無效 → 還是要問
        r = await run(files, emptyHistory(), {'2026-09': 2});
        expect(r.conflicts.length).toBe(1);
    });

    test('同月兩份內容一模一樣：不用問', async () => {
        const a = await makeStmt([['1001', '甲店', 1000, null, null]], []);
        const r = await run([file('9月_已自動對帳.xlsx', 1, a), file('9月_已自動對帳 (1).xlsx', 2, a)]);
        expect(r.conflicts).toEqual([]);
        expect(r.report.map((x) => x.status)).toEqual(['used', 'skipped']);
    });

    test('已記住的月份再丟新的一份：問取代或保留', async () => {
        const first = await makeStmt([['1001', '甲店', 1000, null, null]], []);
        let r = await run([file('9月_已自動對帳.xlsx', 1, first)]);
        const h0 = r.history;
        const nu = [file('9月_已自動對帳 (1).xlsx', 9, await makeStmt([['1001', '甲店', 1000, 1000, '9/5']], []))];
        r = await run(nu, h0);
        expect(r.conflicts[0].options.map((o) => [o.value, o.kind])).toEqual([[0, 'new'], [KEEP, 'old']]);
        expect(r.conflicts[0].options[0].same_as_old).toBe(false);
        expect(r.history.months['2026-09'].source).toBe('9月_已自動對帳.xlsx'); // 沒選之前不會蓋掉
        r = await run(nu, h0, {'2026-09': KEEP});
        expect(r.report[0].status).toBe('skipped');
        expect(compute(r.history).grand_total).toBe(1000);
        r = await run(nu, h0, {'2026-09': 0});
        expect(r.report[0].reason).toContain('取代');
        expect(compute(r.history).customers).toEqual([]);
        // 傳入的紀錄不會被改到
        expect(h0.months['2026-09'].source).toBe('9月_已自動對帳.xlsx');
        // 再丟一模一樣的檔：不用問
        const again = await run(nu, r.history);
        expect([again.conflicts, again.report[0].status]).toEqual([[], 'skipped']);
    });

    test('讀不了的檔案列為錯誤', async () => {
        const blank = new ExcelJS.Workbook();
        blank.addWorksheet('Sheet1');
        const r = await run([
            file('x.xlsx', 1, new Uint8Array(await blank.xlsx.writeBuffer())),
            file('y.xlsx', 1, new TextEncoder().encode('not excel')),
            file('無月份.xlsx', 1, await makeStmt([['1', 'a', 1, null, null]], [])),
        ]);
        expect(r.report.map((x) => x.status)).toEqual(['error', 'error', 'error']);
    });

    test('年份推測', () => {
        expect(guessYear(9, TODAY)).toBe(2026);
        expect(guessYear(12, new Date(2027, 0, 5))).toBe(2026);
    });
});

describe('已結清標記', () => {
    // 甲店：8 月上期未付、8 月下期付一部分(600/1100)；乙店：8 月上期對不上入帳、下期完全沒收到
    const build = async () => {
        const aug = await makeStmt(
            [['1001', '甲店', 1000, null, null], ['1002', '乙店', 2000, null, '8/5收1999?']],
            [['1001', '甲店', 1100, 600, '9/1'], ['1002', '乙店', 2100, null, null]],
        );
        return (await run([file('8月.xlsx', 1, aug)])).history;
    };
    const NOW = new Date(2026, 9, 19, 10, 0, 0);

    test('標已結清：只排除目前顯示的期別，總數、合計、需確認人數更新', async () => {
        const h = await build();
        const before = compute(h);
        expect(before.grand_total).toBe(1000 + 500 + 2000 + 2100);
        const a = before.customers.find((c) => c.code === '1001')!;
        const marks = settleDebtor(emptyMarks(), a, '已補款', NOW);
        const res = compute(h, marks);
        expect(res.customers.map((c) => c.code)).toEqual(['1002']);
        expect(res.grand_total).toBe(4100);
        expect(res.settled).toHaveLength(1);
        expect(res.settled[0]).toMatchObject({kind: 'manual', code: '1001', total: 1500, reason: '已補款', marked_at: '2026-10-19T10:00:00'});
        expect(res.settled[0].items.map((i) => i.period)).toEqual(['8月上期', '8月下期']);
        expect(res.settled_total).toBe(1500);
    });

    test('標完之後新的月份沒付仍會出現', async () => {
        const h = await build();
        const marks = settleDebtor(emptyMarks(), compute(h).customers.find((c) => c.code === '1001')!, '');
        const sep = await makeStmt([['1001', '甲店', 700, null, null]], []);
        const {history} = await run([file('9月.xlsx', 2, sep)], h);
        const res = compute(history, marks);
        const a = res.customers.find((c) => c.code === '1001')!;
        expect(a.items.map((i) => i.period)).toEqual(['9月上期']);
        expect(a.total).toBe(700);
        expect(res.settled[0].items.map((i) => i.period)).toEqual(['8月上期', '8月下期']);
    });

    test('取消後恢復', async () => {
        const h = await build();
        const marks = settleDebtor(emptyMarks(), compute(h).customers.find((c) => c.code === '1001')!, '已補款');
        const gone = compute(h, marks);
        expect(gone.customers.some((c) => c.code === '1001')).toBe(false);
        const back = compute(h, unsettle(marks, gone.settled[0].id));
        expect(back.customers.map((c) => c.code)).toEqual(compute(h).customers.map((c) => c.code));
        expect(back.grand_total).toBe(compute(h).grand_total);
        expect(back.settled).toEqual([]);
    });

    // 舊版「有收到款就算付清」設定（localStorage 裡可能還留著）
    const legacyRaw = (rules: Record<string, {code: string; name: string; reason: string}>) => ({
        version: 1, settled: [],
        paid_if_received: Object.fromEntries(Object.entries(rules).map(([k, v]) => [k, {...v, marked_at: '2026-10-01T09:00:00'}])),
    });

    test('舊的回饋金設定：載入時轉成一般已結清（原因回饋金自扣），不會有客人冒出來、之後不套用新月份', async () => {
        const h = await build();
        const raw = legacyRaw({'1001': {code: '1001', name: '甲店', reason: '回饋金'}, '1002': {code: '1002', name: '乙店', reason: ''}});
        const legacy = extractLegacyRules(raw);
        // 轉換前（舊行為）：甲只剩 8 月上期、乙只剩 8 月下期
        const before = compute(h, normalizeMarks(raw), legacy);
        const migrated = migrateLegacyRules(h, normalizeMarks(raw), legacy, NOW);
        const after = compute(h, migrated);
        expect(after.customers.map((c) => [c.code, c.items.map((i) => [i.period, i.owed])]))
            .toEqual(before.customers.map((c) => [c.code, c.items.map((i) => [i.period, i.owed])]));
        expect(after.customers.map((c) => [c.code, c.total])).toEqual([['1001', 1000], ['1002', 2100]]);
        expect(after.grand_total).toBe(3100);
        expect(after.settled.map((g) => [g.code, g.kind, g.total, g.reason])).toEqual([
            ['1001', 'manual', 500, '回饋金自扣'], ['1002', 'manual', 2000, '回饋金自扣']]);
        // 轉換後沒有舊欄位，且可以像一般標記一樣取消
        expect('paid_if_received' in migrated).toBe(false);
        const back = compute(h, unsettle(migrated, after.settled[0].id));
        expect(back.customers.find((c) => c.code === '1001')!.total).toBe(1500);
        // 之後新的月份：甲店有付一部分也照樣催（規則不再自動套用）
        const sep = await makeStmt([['1001', '甲店', 700, 300, '9/3']], []);
        const {history} = await run([file('9月.xlsx', 2, sep)], h);
        expect(compute(history, migrated).customers.find((c) => c.code === '1001')!.items.map((i) => [i.period, i.owed]))
            .toEqual([['8月上期', 1000], ['9月上期', 400]]);
    });

    test('舊設定目前沒涵蓋任何期別 → 丟掉；沒有舊設定 → 標記原封不動', async () => {
        const h = await build();
        const raw = legacyRaw({'9999': {code: '9999', name: '丁店', reason: ''}});
        const m = migrateLegacyRules(h, normalizeMarks(raw), extractLegacyRules(raw), NOW);
        expect(m).toEqual(emptyMarks());
        const base = emptyMarks();
        expect(migrateLegacyRules(h, base, {}, NOW)).toBe(base);
    });

    test('只有備註「回扣」、沒有金額也沒有對不上的入帳 → 舊設定也不會把它轉成已結清', async () => {
        const st = await makeStmt([['2035', '大口吐司', 5000, null, null, '回扣']], [], true);
        const {history} = await run([file('9月.xlsx', 1, st)]);
        const raw = legacyRaw({'2035': {code: '2035', name: '大口吐司', reason: ''}});
        const m = migrateLegacyRules(history, normalizeMarks(raw), extractLegacyRules(raw), NOW);
        expect(compute(history, m).customers.map((c) => c.code)).toEqual(['2035']);
    });

    test('復原：標已結清後用同一個 id 取消 → 欠款、總額、已結清區都回到原狀（別人的標記不受影響）', async () => {
        const h = await build();
        const base = compute(h);
        const m1 = settleDebtor(emptyMarks(), base.customers.find((c) => c.code === '1001')!, '已補款', NOW);
        const m2 = settleDebtor(m1, base.customers.find((c) => c.code === '1002')!, '', new Date(2026, 9, 19, 10, 0, 5));
        const id2 = m2.settled[m2.settled.length - 1].id;
        const undone = compute(h, unsettle(m2, id2));
        expect(undone.customers.map((c) => c.code)).toEqual(['1002']);
        expect(undone.settled.map((g) => g.code)).toEqual(['1001']);
        const all = compute(h, unsettle(unsettle(m2, id2), m1.settled[0].id));
        expect(all.customers.map((c) => [c.code, c.total])).toEqual(base.customers.map((c) => [c.code, c.total]));
        expect(all.grand_total).toBe(base.grand_total);
        expect(all.settled).toEqual([]);
        // 同一秒內連按兩位也各有各的 id
        const same = settleDebtor(settleDebtor(emptyMarks(), base.customers[0], '', NOW), base.customers[1], '', NOW);
        expect(new Set(same.settled.map((s) => s.id)).size).toBe(2);
    });

    test('舊資料相容：沒有標記欄位或格式亂掉都視為空', async () => {
        const h = await build();
        for (const raw of [undefined, null, {}, {version: 1}, 'x', {settled: 'x', paid_if_received: 5, manual_debts: 'x'},
            {manual_debts: [{customer: '1001', month_key: '2026-08', half: '下', amount: -3}, {customer: '1001', month_key: '8月', half: '下', amount: 3}]},
            {settled: [{customer: '1001'}, null]}]) {
            const m = normalizeMarks(raw);
            expect(m).toEqual(emptyMarks());
            expect(compute(h, m).grand_total).toBe(compute(h).grand_total);
        }
        // 沒帶標記參數時和原本一樣
        expect(compute(h).settled).toEqual([]);
        // 缺 reason／id 的舊標記補預設值
        const m = normalizeMarks({settled: [{customer: '1001', keys: ['2026-08|上'], marked_at: 'x'}]});
        expect(m.settled[0]).toMatchObject({customer: '1001', reason: '', keys: ['2026-08|上']});
        expect(compute(h, m).customers.find((c) => c.code === '1001')!.items.map((i) => i.period)).toEqual(['8月下期']);
    });
});

describe('手動加一筆欠款', () => {
    // 3060 早安好呷：9 月上期 2,324、下期 2,096 沒付；8 月都付清
    const build = async () => {
        const aug = await makeStmt([['3060', '早安好呷', 1500, 1500, '8/5']], [['3060', '早安好呷', 1800, 1800, '9/5']]);
        const sep = await makeStmt([['3060', '早安好呷', 2324, null, null], ['1001', '甲店', 1000, null, null]], [['3060', '早安好呷', 2096, null, null]]);
        return (await run([file('8月.xlsx', 1, aug), file('9月.xlsx', 2, sep)])).history;
    };
    const NOW = new Date(2026, 9, 19, 10, 0, 0);
    const input = {month_key: '2026-08', half: '下' as const, amount: 64, reason: '少付'};

    test('範例：9月上下期 + 8月下期少付64 → 合計 4,484，訊息依期別時間順序多一行', async () => {
        const h = await build();
        const c0 = compute(h).customers.find((c) => c.code === '3060')!;
        expect(c0.total).toBe(4420);
        const marks = addManualDebt(emptyMarks(), c0, input, NOW);
        const res = compute(h, marks);
        const c = res.customers.find((x) => x.code === '3060')!;
        expect(c.total).toBe(4484);
        expect(c.items.map((i) => [i.period, i.owed])).toEqual([['8月下期', 64], ['9月上期', 2324], ['9月下期', 2096]]);
        expect(res.grand_total).toBe(4484 + 1000);
        expect(c.message).toContain('・8月下期（少付）：64 元\n・9月上期：2,324 元\n・9月下期：2,096 元\n合計：4,484 元');
    });

    test('訊息：原因空白不加括號；夾在其他期之間按時間排', async () => {
        const h = await build();
        const c0 = compute(h).customers.find((x) => x.code === '3060')!;
        const res = compute(h, addManualDebt(emptyMarks(), c0, {...input, month_key: '2026-09', half: '上', reason: ''}, NOW));
        const c = res.customers.find((x) => x.code === '3060')!;
        expect(c.items.map((i) => i.period)).toEqual(['9月上期', '9月上期', '9月下期']);
        expect(c.message).toContain('・9月上期：2,324 元\n・9月上期：64 元\n・9月下期：2,096 元');
        expect(buildMessage(c)).toBe(c.message);
    });

    test('刪除、修改', async () => {
        const h = await build();
        const c0 = compute(h).customers.find((x) => x.code === '3060')!;
        const marks = addManualDebt(emptyMarks(), c0, input, NOW);
        const id = marks.manual_debts[0].id;
        const edited = compute(h, updateManualDebt(marks, id, {month_key: '2026-07', half: '上', amount: 100, reason: ''}));
        expect(edited.customers.find((x) => x.code === '3060')!.total).toBe(4520);
        expect(edited.customers.find((x) => x.code === '3060')!.items[0]).toMatchObject({period: '7月上期', owed: 100});
        const removed = compute(h, removeManualDebt(marks, id));
        expect(removed.customers.find((x) => x.code === '3060')!.total).toBe(4420);
        expect(removed.grand_total).toBe(compute(h).grand_total);
    });

    test('同一期同金額連加兩筆 id 不重複；跨年期別帶年份', async () => {
        const h = await build();
        const c0 = compute(h).customers.find((x) => x.code === '3060')!;
        const m = addManualDebt(addManualDebt(emptyMarks(), c0, input, NOW), c0, input, NOW);
        expect(new Set(m.manual_debts.map((d) => d.id)).size).toBe(2);
        const c = compute(h, addManualDebt(emptyMarks(), c0, {...input, month_key: '2025-12'}, NOW)).customers.find((x) => x.code === '3060')!;
        expect(c.items.map((i) => i.period)).toEqual(['2025年12月下期', '2026年9月上期', '2026年9月下期']);
    });

    test('客人所有期都已結清、但有手動欠款：照樣顯示手動欠款；已結清會連手動欠款一起帶走，復原回來', async () => {
        const h = await build();
        const c0 = compute(h).customers.find((x) => x.code === '3060')!;
        // 先結清兩期，再加手動欠款 → 只剩手動欠款
        const m1 = settleDebtor(emptyMarks(), c0, '已補款', NOW);
        const m2 = addManualDebt(m1, c0, input, NOW);
        const res = compute(h, m2);
        const c = res.customers.find((x) => x.code === '3060')!;
        expect(c.items.map((i) => [i.period, i.owed])).toEqual([['8月下期', 64]]);
        expect(c.total).toBe(64);
        // 再按已結清：手動欠款一起標掉
        const m3 = settleDebtor(m2, c, '', new Date(2026, 9, 19, 10, 0, 9));
        const gone = compute(h, m3);
        expect(gone.customers.some((x) => x.code === '3060')).toBe(false);
        expect(gone.settled.map((g) => [g.items.map((i) => i.period), g.total])).toEqual([[['9月上期', '9月下期'], 4420], [['8月下期'], 64]]);
        const back = compute(h, unsettle(m3, m3.settled[1].id));
        expect(back.customers.find((x) => x.code === '3060')!.total).toBe(64);
    });

    test('只做名單上的客人：不在名單的客人，手動欠款不顯示', async () => {
        const h = await build();
        const res = compute(h, addManualDebt(emptyMarks(), {code: '7777', name: '路人'}, input, NOW));
        expect(res.customers.some((c) => c.code === '7777')).toBe(false);
        expect(res.grand_total).toBe(compute(h).grand_total);
    });

    test('輸入檢查：期別、金額要正整數', () => {
        expect(validateManualDebt({month_key: '2026-08', half: '下', amount: 64})).toBeNull();
        expect(validateManualDebt({month_key: '', half: '下', amount: 64})).not.toBeNull();
        expect(validateManualDebt({month_key: '2026-13', half: '下', amount: 64})).not.toBeNull();
        expect(validateManualDebt({month_key: '2026-08', half: '', amount: 64})).not.toBeNull();
        for (const amount of [0, -5, 1.5, NaN]) {
            expect(validateManualDebt({month_key: '2026-08', half: '下', amount})).not.toBeNull();
        }
    });

    test('標記存取：手動欠款經 JSON 存回來一樣；舊資料沒有這個欄位視為空', async () => {
        const h = await build();
        const c0 = compute(h).customers.find((x) => x.code === '3060')!;
        const marks = addManualDebt(emptyMarks(), c0, input, NOW);
        expect(normalizeMarks(JSON.parse(JSON.stringify(marks)))).toEqual(marks);
        expect(normalizeMarks({version: 1, settled: []}).manual_debts).toEqual([]);
    });
});

// ---------- 欠款期數分組、月結、未結帳期別

type Spec = [code: string, name: string, amt: number, paid: number | null, memo?: string];

const rec = ([code, name, amt, paid, memo]: Spec): StatementRecord =>
    ({code, name, amt, paid, paid_text: '', date_text: paid === null ? '' : '9/1', memo: memo ?? ''});

/** 手工組記錄：months = {月份鍵: {上: Spec[], 下: Spec[]}}。每期自帶 3 位付清的客人（9001–9003），免得整期被當成沒人付款。 */
function hist(months: Record<string, {上?: Spec[]; 下?: Spec[]}>): History {
    const out: History = {version: 1, months: {}};
    for (const [k, v] of Object.entries(months)) {
        // 金額隨月份不同，免得下期被當成複製上個月的佔位（same_as_prev）
        const amt = 100 + +k.slice(5);
        const filler: Spec[] = [['9001', '付清一', amt, amt], ['9002', '付清二', amt, amt], ['9003', '付清三', amt, amt]];
        const periods = {上: {}, 下: {}} as History['months'][string]['periods'];
        for (const h of ['上', '下'] as const) {
            for (const s of [...filler, ...(v[h] ?? [])]) periods[h][s[0]] = rec(s);
        }
        out.months[k] = {source: `${k}.xlsx`, file_mtime: '', imported_at: '', periods};
    }
    return out;
}

const owe = (code: string, name: string, amt: number, memo?: string): Spec => [code, name, amt, null, memo];

describe('欠款期數分組', () => {
    const h = hist({
        '2026-08': {上: [owe('1002', '乙', 200), owe('1003', '丙', 300)], 下: [owe('1003', '丙', 310)]},
        '2026-09': {上: [owe('1001', '甲', 100), owe('1002', '乙', 210), owe('1003', '丙', 320)]},
    });
    const res = compute(h);

    test('欠 1 期／2 期／3 期以上分開，欠越多越上面', () => {
        expect(res.customers.map((c) => [c.code, c.periods, c.group])).toEqual(
            [['1003', 3, 'three'], ['1002', 2, 'two'], ['1001', 1, 'one']]);
        expect(res.chase.map((c) => c.code)).toEqual(['1003', '1002']);
        expect(res.light.map((c) => c.code)).toEqual(['1001']);
    });

    test('欠 1 期的不算進催款總額，分開加總', () => {
        expect(res.chase_total).toBe(930 + 410);
        expect(res.light_total).toBe(100);
        expect(res.grand_total).toBe(1440); // 原算法（含欠 1 期）不變
    });

    test('同組內需確認的排最上面', () => {
        const h2 = hist({
            '2026-08': {上: [owe('1002', '乙', 200), owe('1004', '丁', 400)]},
            '2026-09': {上: [owe('1002', '乙', 210), owe('1004', '丁', 410, '待確認')]},
        });
        const r2 = compute(h2);
        expect(r2.chase.map((c) => c.code)).toEqual(['1004', '1002']);
        expect(r2.chase[0].notes.length).toBeGreaterThan(0);
    });

    test('下載的 Excel 與文字檔只有要催的，欠 1 期的不在裡面', async () => {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await (await buildSummaryXlsx(res)).arrayBuffer());
        const sheet = wb.getWorksheet('催款總表')!;
        const rows: unknown[][] = [];
        sheet.eachRow((r, n) => n > 1 && rows.push((r.values as unknown[]).slice(1)));
        expect(rows.map((r) => r[0])).toEqual(['1003', '1002', '']);
        expect(rows[2][5]).toBe(1340);
        expect(JSON.stringify(wb.getWorksheet('各期明細')!.getSheetValues())).not.toContain('甲');
        const txt = buildMessagesText(res, '2026-10-19 12:00');
        expect(txt).toContain('共 2 位，合計 1,340 元');
        expect(txt).toContain('欠 3 期（1 位）');
        expect(txt).toContain('欠 2 期');
        expect(txt).not.toContain('甲');
    });
});

describe('月結客人', () => {
    const base = (memo?: string) => hist({'2026-09': {上: [owe('1001', '甲', 100, memo)], 下: [owe('1001', '甲', 200)]}});
    const one = (h: History, marks = emptyMarks()) => compute(h, marks).customers.find((c) => c.code === '1001')!;

    test('不是月結：9 上＋9 下算 2 期', () => {
        const c = one(base());
        expect([c.periods, c.group, c.monthly, c.monthly_auto]).toEqual([2, 'two', false, false]);
    });

    test('自動認：備註有一段剛好是「月」→ 同月上下期合算 1 期', () => {
        for (const memo of ['月', '月；2筆', '2筆；月']) {
            const c = one(base(memo));
            expect([c.periods, c.group, c.monthly, c.monthly_auto, c.total]).toEqual([1, 'one', true, true, 300]);
        }
        expect(one(base('月餅禮盒')).monthly).toBe(false); // 不是剛好「月」
    });

    test('月結標記在任一個已記住的月份就算（不限欠款那一期）', () => {
        const h = hist({'2026-07': {上: [['1001', '甲', 50, 50, '月']]}, '2026-09': {上: [owe('1001', '甲', 100)], 下: [owe('1001', '甲', 200)]}});
        expect(one(h).monthly).toBe(true);
    });

    test('跨月不合併：月結客人欠 8 月和 9 月是 2 期', () => {
        const h = hist({'2026-08': {上: [owe('1001', '甲', 90, '月')]}, '2026-09': {上: [owe('1001', '甲', 100)], 下: [owe('1001', '甲', 200)]}});
        const c = one(h);
        expect([c.periods, c.group]).toEqual([2, 'two']);
    });

    test('手動設為月結、取消月結（含取消自動認到的）', () => {
        const plain = base();
        const on = setMonthly(emptyMarks(), one(plain), true);
        expect(on.monthly).toEqual({'1001': true});
        expect([one(plain, on).monthly, one(plain, on).periods]).toEqual([true, 1]);

        const auto = base('月');
        const off = setMonthly(emptyMarks(), one(auto), false);
        expect(off.monthly).toEqual({'1001': false});
        expect([one(auto, off).monthly, one(auto, off).monthly_auto, one(auto, off).periods]).toEqual([false, true, 2]);

        // 設成和自動判斷一樣 → 不留手動設定
        expect(setMonthly(off, one(auto, off), true).monthly).toEqual({});
        expect(setMonthly(on, one(plain, on), false).monthly).toEqual({});
    });

    test('月結設定存回來一樣；舊資料沒有這個欄位視為空', () => {
        const marks = setMonthly(emptyMarks(), {code: '1001', name: '甲', monthly_auto: false}, true);
        expect(normalizeMarks(JSON.parse(JSON.stringify(marks)))).toEqual(marks);
        expect(normalizeMarks({version: 1, settled: [], manual_debts: []}).monthly).toEqual({});
        expect(normalizeMarks({monthly: {a: 'yes', b: false}}).monthly).toEqual({b: false});
    });
});

describe('手動欠款不算期數', () => {
    const NOW = new Date(2026, 9, 19, 10, 0, 0);
    const input = {month_key: '2026-09', half: '下' as const, amount: 64, reason: '少付'};

    test('有 1 期欠款再加手動欠款：期數不變，金額多算', () => {
        const h = hist({'2026-09': {上: [owe('1001', '甲', 100)]}});
        const c0 = compute(h).customers[0];
        const c = compute(h, addManualDebt(emptyMarks(), c0, input, NOW)).customers[0];
        expect([c.periods, c.group, c.total, c.items.length]).toEqual([1, 'one', 164, 2]);
    });

    test('只有手動欠款：期數 0，歸在欠 1 期那組只顯示', () => {
        const h = hist({'2026-09': {上: [['1001', '甲', 100, 100]]}});
        const marks = addManualDebt(emptyMarks(), {code: '1001', name: '甲'}, input, NOW);
        const res = compute(h, marks);
        expect(res.customers.map((c) => [c.code, c.periods, c.group, c.total])).toEqual([['1001', 0, 'one', 64]]);
        expect(res.chase).toEqual([]);
        expect(res.chase_total).toBe(0);
        expect(res.light_total).toBe(64);
    });

    test('手動欠款加到 2 期欠款的客人：還是欠 2 期', () => {
        const h = hist({'2026-08': {上: [owe('1001', '甲', 100)]}, '2026-09': {上: [owe('1001', '甲', 110)]}});
        const c0 = compute(h).customers[0];
        const c = compute(h, addManualDebt(emptyMarks(), c0, input, NOW)).customers[0];
        expect([c.periods, c.group, c.total]).toEqual([2, 'two', 274]);
    });
});

describe('還沒結帳的期別略過', () => {
    const nobodyPaid: Spec[] = [owe('1001', '甲', 100), owe('1002', '乙', 200), owe('1003', '丙', 300)];
    const noFiller = (months: Record<string, {上: Spec[]; 下: Spec[]}>): History => {
        const out: History = {version: 1, months: {}};
        for (const [k, v] of Object.entries(months)) {
            const periods = {上: {}, 下: {}} as History['months'][string]['periods'];
            for (const half of ['上', '下'] as const) for (const s of v[half]) periods[half][s[0]] = rec(s);
            out.months[k] = {source: `${k}.xlsx`, file_mtime: '', imported_at: '', periods};
        }
        return out;
    };
    const paidRows: Spec[] = [['1001', '甲', 100, 100], ['1002', '乙', 200, 200], ['1003', '丙', 300, 300]];

    test('整期沒有任何人付款 → 不算欠款，月份列表標示', () => {
        const h = noFiller({'2026-09': {上: nobodyPaid.map((s, i) => (i ? paidRows[i] : s)), 下: nobodyPaid}});
        const res = compute(h);
        expect(res.months[0].skipped).toEqual([{half: '下', reason: 'no_payment'}]);
        expect(res.customers.map((c) => [c.code, c.items.map((i) => i.period)])).toEqual([['1001', ['9月上期']]]);
    });

    test('只要有一個人付款，這期就照常算', () => {
        const rows: Spec[] = [owe('1001', '甲', 100), owe('1002', '乙', 200), ['1003', '丙', 300, 300]];
        const res = compute(noFiller({'2026-09': {上: rows, 下: rows}}));
        expect(res.months[0].skipped).toEqual([]);
        expect(res.customers.map((c) => c.code).sort()).toEqual(['1001', '1002']);
    });

    test('樣本太少（不到 3 筆）不自動略過', () => {
        const res = compute(noFiller({'2026-09': {上: [owe('1001', '甲', 100)], 下: [owe('1001', '甲', 100)]}}));
        expect(res.months[0].skipped).toEqual([]);
        expect(res.customers[0].periods).toBe(2);
    });

    test('下期和上個月下期幾乎一樣（複製來的佔位）→ 略過；上期照算', () => {
        const sep: Spec[] = [owe('1001', '甲', 100), owe('1002', '乙', 200), ['1003', '丙', 300, 300], ['1004', '丁', 400, 400]];
        const octLower: Spec[] = sep.map((s) => [s[0], s[1], s[2], s[0] === '1003' ? 300 : null]); // 金額全相同、其中有人「付款」
        const octUpper: Spec[] = [owe('1001', '甲', 150), ['1002', '乙', 250, 250], ['1003', '丙', 350, 350]];
        const h = noFiller({'2026-09': {上: sep, 下: sep}, '2026-10': {上: octUpper, 下: octLower}});
        const res = compute(h);
        expect(res.months[0].skipped).toEqual([]);
        expect(res.months[1].skipped).toEqual([{half: '下', reason: 'same_as_prev', ref: '2026-09'}]);
        // 10 月下期的 1001 100、1002 200 沒算；9 月上下期＋10 月上期的 1001 算
        expect(res.customers.find((c) => c.code === '1001')!.items.map((i) => i.period)).toEqual(['9月上期', '9月下期', '10月上期']);
        expect(res.customers.find((c) => c.code === '1002')!.items.map((i) => i.period)).toEqual(['9月上期', '9月下期']);
    });

    test('下期金額有變（真的結過帳）→ 不略過', () => {
        const sep: Spec[] = [owe('1001', '甲', 100), owe('1002', '乙', 200), ['1003', '丙', 300, 300]];
        const oct: Spec[] = [owe('1001', '甲', 111), owe('1002', '乙', 222), ['1003', '丙', 333, 333]];
        const res = compute(noFiller({'2026-09': {上: sep, 下: sep}, '2026-10': {上: oct, 下: oct}}));
        expect(res.months[1].skipped).toEqual([]);
    });

    test('版面角落的統計格（名稱是整個數字）不當成客人讀進來', async () => {
        const st = await makeStmt([['1001', '甲店', 1000, null, null], ['11025', '11025', 11025, 11025, '11025']], []);
        const {history} = await run([file('9月.xlsx', 1, st)]);
        expect(Object.keys(history.months['2026-09'].periods.上)).toEqual(['1001']);
    });
});

describe('月結客人未到期與月底前要催', () => {
    const TODAY_OCT = new Date(2026, 9, 19);
    /** 不自帶付清客人的月份（要測下期整期沒人付款／下期尚未記住）。 */
    const bare = (months: Record<string, {上?: Spec[]; 下?: Spec[]}>): History => {
        const out: History = {version: 1, months: {}};
        for (const [k, v] of Object.entries(months)) {
            const periods = {上: {}, 下: {}} as History['months'][string]['periods'];
            for (const half of ['上', '下'] as const) for (const sp of v[half] ?? []) periods[half][sp[0]] = rec(sp);
            out.months[k] = {source: `${k}.xlsx`, file_mtime: '', imported_at: '', periods};
        }
        return out;
    };
    const paid3: Spec[] = [['9001', '付一', 100, 100], ['9002', '付二', 100, 100], ['9003', '付三', 100, 100]];
    const unpaid3 = (amt: number): Spec[] => [owe('1001', '寶米', amt, '月'), owe('2001', '乙', amt), owe('2002', '丙', amt)];
    const find = (h: History, code: string, marks = emptyMarks(), today = TODAY_OCT) =>
        compute(h, marks, {}, today).customers.find((c) => c.code === code)!;

    // 月結 1001：8 月下期只付一部分、9 月上下期沒付、10 月上期沒付；10 月下期還沒記住
    const oct = bare({
        '2026-08': {上: paid3, 下: [...paid3, ['1001', '寶米', 100, 40, '月']]},
        '2026-09': {上: [...paid3, owe('1001', '寶米', 8062, '月')], 下: [...paid3, owe('1001', '寶米', 6823, '月')]},
        '2026-10': {上: [...paid3, owe('1001', '寶米', 7599, '月'), owe('1002', '一般店', 500)]},
    });

    test('月結客人本月（下期尚未記住）的上期＝未到期：不算欠款、不算期數，只標示', () => {
        const c = find(oct, '1001');
        expect(c.items.map((i) => [i.period, i.owed])).toEqual([['8月下期', 60], ['9月上期', 8062], ['9月下期', 6823]]);
        expect([c.total, c.periods, c.group]).toEqual([14945, 2, 'two']);
        expect(c.pending.map((i) => [i.period, i.owed])).toEqual([['10月上期', 7599]]);
        expect(c.pending_total).toBe(7599);
        expect(c.message).not.toContain('10月');
    });

    test('一般（非月結）客人的本月上期照舊算欠', () => {
        const c = find(oct, '1002');
        expect([c.items.map((i) => i.period), c.total, c.pending]).toEqual([['10月上期'], 500, []]);
        // 取消寶米的月結 → 10 月上期也要算
        const off = setMonthly(emptyMarks(), c0(oct, '1001'), false);
        const p = find(oct, '1001', off);
        expect([p.items.map((i) => i.period), p.periods, p.pending]).toEqual([['8月下期', '9月上期', '9月下期', '10月上期'], 4, []]);
    });
    const c0 = (h: History, code: string) => find(h, code);

    test('下期被判定未結帳而略過的月份（即使已不是本月）也是未到期', () => {
        const h = bare({
            '2026-09': {上: [...paid3, owe('1001', '寶米', 100, '月')], 下: [...paid3, owe('1001', '寶米', 200, '月')]},
            '2026-10': {上: [...paid3, owe('1001', '寶米', 300, '月')], 下: unpaid3(400)},
        });
        const c = find(h, '1001', emptyMarks(), new Date(2026, 10, 3));
        expect([c.items.map((i) => i.period), c.total, c.periods]).toEqual([['9月上期', '9月下期'], 300, 1]);
        expect(c.pending.map((i) => i.period)).toEqual(['10月上期']);
    });

    test('只有未到期、沒有欠款的月結客人不在名單', () => {
        const h = bare({'2026-10': {上: [...paid3, owe('1001', '寶米', 300, '月')]}});
        expect(compute(h, emptyMarks(), {}, TODAY_OCT).customers).toEqual([]);
    });

    test('月底前要催：月結客人只欠上個月 1 期也列入要催（有訊息、算進總額與下載檔）', async () => {
        const h = bare({
            '2026-09': {上: [...paid3, owe('1001', '寶米', 8000, '月')], 下: [...paid3, ['1001', '寶米', 100, 100, '月']]},
            '2026-10': {上: [...paid3, owe('1001', '寶米', 7000, '月')]},
        });
        const plain = compute(h, emptyMarks(), {}, TODAY_OCT);
        expect(plain.customers.map((c) => [c.code, c.group])).toEqual([['1001', 'one']]);
        expect(plain.chase_total).toBe(0);

        const marks = setMonthEnd(emptyMarks(), {code: '1001', name: '寶米'}, true);
        const res = compute(h, marks, {}, TODAY_OCT);
        const c = res.customers[0];
        expect([c.periods, c.group, c.month_end, c.total, c.pending_total]).toEqual([1, 'one_eom', true, 8000, 7000]);
        expect(res.chase.map((x) => x.code)).toEqual(['1001']);
        expect([res.chase_total, res.light.length]).toEqual([8000, 0]);
        expect(c.message).toContain('8,000');
        const txt = buildMessagesText(res, '2026-10-19 12:00');
        expect(txt).toContain('欠 1 期（月底前要催）');
        expect(txt).toContain('寶米');
        expect(txt).not.toContain('7,000');
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await (await buildSummaryXlsx(res)).arrayBuffer());
        expect(JSON.stringify(wb.getWorksheet('催款總表')!.getSheetValues())).toContain('月底前要催');
    });

    test('月底前要催：分組排在欠 2 期之後、只顯示之前；欠 2 期以上照原本分組；一般客人或只欠本月不適用', () => {
        const h = bare({
            '2026-08': {上: [...paid3, owe('1002', '兩期', 110, '月')]},
            '2026-09': {
                上: [...paid3, owe('1001', '一期催', 20, '月'), owe('1002', '兩期', 30, '月'), owe('1003', '只顯示', 40, '月'), owe('1004', '一般', 50)],
                下: [...paid3, ['1001', '一期催', 5, 5, '月']],
            },
        });
        const marks = ['1001', '1002', '1004'].reduce((m, code) => setMonthEnd(m, {code, name: ''}, true), emptyMarks());
        const res = compute(h, marks, {}, TODAY_OCT);
        expect(res.customers.map((c) => [c.code, c.group])).toEqual([['1002', 'two'], ['1001', 'one_eom'], ['1003', 'one'], ['1004', 'one']]);
        expect(res.customers.find((c) => c.code === '1004')!.month_end).toBe(false); // 非月結，設定不生效
        expect(res.chase.map((c) => c.code)).toEqual(['1002', '1001']);
        expect(res.light.map((c) => c.code)).toEqual(['1003', '1004']);
        // 只欠本月（上個月沒欠）→ 未到期，不列入
        const h2 = bare({'2026-10': {上: [...paid3, owe('1001', '寶米', 300, '月')]}});
        expect(compute(h2, marks, {}, TODAY_OCT).chase).toEqual([]);
    });

    test('月底前要催設定：可關閉、存回來一樣、舊資料沒有欄位視為空', () => {
        const on = setMonthEnd(emptyMarks(), {code: '1001', name: '寶米'}, true);
        expect(on.month_end).toEqual({'1001': true});
        expect(normalizeMarks(JSON.parse(JSON.stringify(on)))).toEqual(on);
        expect(setMonthEnd(on, {code: '1001', name: '寶米'}, false).month_end).toEqual({});
        expect(normalizeMarks({version: 1, settled: [], manual_debts: [], monthly: {}}).month_end).toEqual({});
        expect(normalizeMarks({month_end: {a: true, b: false, c: 'x'}}).month_end).toEqual({a: true});
    });

    test('存在 localStorage；清除全部記錄（只清月份）時設定保留', async () => {
        const mem = new Map<string, string>();
        (globalThis as unknown as {localStorage: Storage}).localStorage = {
            getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v),
            removeItem: (k: string) => void mem.delete(k), clear: () => mem.clear(), key: () => null, length: 0,
        } as Storage;
        const store = await import('../src/infra/auto-collect-store');
        const on = setMonthEnd(emptyMarks(), {code: '1001', name: '寶米'}, true);
        store.saveHistory(oct);
        store.saveMarks(on);
        store.clearHistory(); // 清除全部記錄預設只清月份
        expect(store.loadHistory().months).toEqual({});
        expect(store.loadMarks(store.loadHistory()).month_end).toEqual({'1001': true});
        store.clearMarks(); // 勾選「同時清除標記」才清
        expect(store.loadMarks(store.loadHistory()).month_end).toEqual({});
    });
});

describe('手動修改催款訊息', () => {
    const NOW = new Date(2026, 9, 19, 10, 0, 0);
    const build = () => hist({
        '2026-08': {上: [owe('7016', '寶米早餐', 7871)], 下: [owe('7016', '寶米早餐', 100)]},
        '2026-09': {上: [owe('7016', '寶米早餐', 500)], 下: [owe('1001', '甲店', 300)]},
    });
    const pick = (h: History, m = emptyMarks()) => compute(h, m, {}, TODAY).customers.find((c) => c.code === '7016')!;
    const EDITED = '寶米早餐 您好～\n・8月下期：少付 60 元\n麻煩您，謝謝！';

    test('儲存／讀回：message 是她改的版本，auto_message 不變，金額與分組不受影響', () => {
        const h = build();
        const c0 = pick(h);
        const m = setMessageEdit(emptyMarks(), c0, EDITED, NOW);
        const c = pick(h, m);
        expect(c.message).toBe(EDITED);
        expect(c.auto_message).toBe(c0.auto_message);
        expect(c.edit).toMatchObject({text: EDITED, base: c0.auto_message, stale: false});
        const a = compute(h, emptyMarks(), {}, TODAY);
        const b = compute(h, m, {}, TODAY);
        expect(b.chase_total).toBe(a.chase_total);
        expect(b.grand_total).toBe(a.grand_total);
        expect(b.customers.map((x) => [x.code, x.group, x.total])).toEqual(a.customers.map((x) => [x.code, x.group, x.total]));
        // 經 JSON（localStorage）往返後一樣
        const back = normalizeMarks(JSON.parse(JSON.stringify(m)));
        expect(pick(h, back).message).toBe(EDITED);
        expect(back.edits).toEqual(m.edits);
    });

    test('沒代號的客人用店名當鍵；空白或與自動版一樣＝沒改', () => {
        const c0 = {code: '', name: '無號店', auto_message: 'AUTO'};
        const m = setMessageEdit(emptyMarks(), c0, 'X', NOW);
        expect(Object.keys(m.edits)).toEqual(['無號店']);
        expect(setMessageEdit(m, c0, '   \n', NOW).edits).toEqual({});
        expect(setMessageEdit(m, c0, 'AUTO', NOW).edits).toEqual({});
        expect(setMessageEdit(emptyMarks(), c0, 'a\r\nb  \n', NOW).edits['無號店'].text).toBe('a\nb');
    });

    test('還原：回到自動產生；沒改過時回傳同一物件', () => {
        const h = build();
        const c0 = pick(h);
        const m = setMessageEdit(emptyMarks(), c0, EDITED, NOW);
        const r = clearMessageEdit(m, c0);
        expect(pick(h, r).message).toBe(c0.auto_message);
        expect(pick(h, r).edit).toBeNull();
        expect(clearMessageEdit(r, c0)).toBe(r);
    });

    test('資料變動（加一筆欠款、已結清、月結）→ 保留她的版本並標 stale；變回一樣就不提醒', () => {
        const h = build();
        const c0 = pick(h);
        const m = setMessageEdit(emptyMarks(), c0, EDITED, NOW);
        const added = pick(h, addManualDebt(m, c0, {month_key: '2026-08', half: '下', amount: 60, reason: '少付'}, NOW));
        expect(added.message).toBe(EDITED);
        expect(added.edit?.stale).toBe(true);
        expect(added.auto_message).toContain('少付');
        expect(added.auto_message).not.toBe(c0.auto_message);
        const mo = pick(h, setMonthly(m, c0, true));
        expect(mo.message).toBe(EDITED);
        expect(mo.edit?.stale).toBe(mo.auto_message !== c0.auto_message);
        // 新月份進來
        const h2 = hist({...build().months as never, '2026-10': {上: [owe('7016', '寶米早餐', 222)]}} as never);
        expect(pick(h2, m).edit?.stale).toBe(true);
        // 移除手動欠款後自動版回到編輯當時 → 不再提醒
        const withDebt = addManualDebt(m, c0, {month_key: '2026-08', half: '下', amount: 60, reason: ''}, NOW);
        expect(pick(h, removeManualDebt(withDebt, withDebt.manual_debts[0].id)).edit?.stale).toBe(false);
    });

    test('下載文字檔與總表 Excel 的訊息欄用改過的版本', async () => {
        const h = build();
        const c0 = pick(h);
        const res = compute(h, setMessageEdit(emptyMarks(), c0, EDITED, NOW), {}, TODAY);
        const txt = buildMessagesText(res, '2026-10-19 10:00');
        expect(txt).toContain(EDITED);
        expect(txt).not.toContain(c0.auto_message);
        expect(txt).toContain('（合計 ' + c0.total.toLocaleString('en-US') + ' 元）');
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(Buffer.from(await (await buildSummaryXlsx(res)).arrayBuffer()));
        const ws = wb.getWorksheet('催款訊息')!;
        const row = ws.getRows(2, ws.rowCount - 1)!.find((r) => r.getCell(1).value === '7016')!;
        expect(row.getCell(4).value).toBe(EDITED);
        expect(row.getCell(3).value).toBe(c0.total);
    });

    test('舊資料相容：沒有 edits 欄位照常讀；壞掉的項目丟掉', () => {
        const old = {version: 1, settled: [], manual_debts: [], monthly: {'7016': true}, month_end: {}};
        const m = normalizeMarks(old);
        expect(m.edits).toEqual({});
        expect(m.monthly).toEqual({'7016': true});
        const bad = normalizeMarks({...old, edits: {a: {text: 'ok', base: 'b', edited_at: 't'}, b: {text: '  '}, c: 5, d: {base: 'x'}}});
        expect(Object.keys(bad.edits)).toEqual(['a']);
        expect(normalizeMarks({...old, edits: [1]}).edits).toEqual({});
        // 沒有 edits 的 Marks 物件直接丟進 compute 也不會壞
        const legacyMarks = {...emptyMarks()} as Partial<typeof old> & ReturnType<typeof emptyMarks>;
        delete (legacyMarks as {edits?: unknown}).edits;
        expect(pick(build(), legacyMarks as never).edit).toBeNull();
    });

    test('存在 localStorage：重新讀取還在；清除月份記錄時保留、勾選清標記才清', async () => {
        const mem = new Map<string, string>();
        (globalThis as unknown as {localStorage: Storage}).localStorage = {
            getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v),
            removeItem: (k: string) => void mem.delete(k), clear: () => mem.clear(), key: () => null, length: 0,
        } as Storage;
        const store = await import('../src/infra/auto-collect-store');
        const h = build();
        store.saveHistory(h);
        store.saveMarks(setMessageEdit(emptyMarks(), pick(h), EDITED, NOW));
        expect(pick(store.loadHistory(), store.loadMarks(store.loadHistory())).message).toBe(EDITED);
        store.clearHistory();
        expect(store.loadMarks(store.loadHistory()).edits['7016'].text).toBe(EDITED);
        store.clearMarks();
        expect(store.loadMarks(store.loadHistory()).edits).toEqual({});
    });
});

const DL = (m: number): string => `${process.env.HOME}/Downloads/${m}月_已自動對帳.xlsx`;
describe.skipIf(![7, 8, 9, 10].every((m) => existsSync(DL(m))))('實際對帳檔（~/Downloads 7–10 月）', () => {
    test('10 月下期（雜訊統計格濾掉後整期沒人付款）判定未結帳略過', async () => {
        const files = [7, 8, 9, 10].map((m) => file(`${m}月_已自動對帳.xlsx`, m, new Uint8Array(readFileSync(DL(m)))));
        const {history} = await run(files);
        const res = compute(history);
        expect(res.months.map((m) => [m.key, m.skipped.map((s) => `${s.half}:${s.reason}`)])).toEqual([
            ['2026-07', []], ['2026-08', []], ['2026-09', []], ['2026-10', ['下:no_payment']]]);
        expect(res.months[3].count).toBe(263 + 311);
        const sum = (list: typeof res.chase) => list.reduce((t, c) => t + c.total, 0);
        expect(res.chase_total).toBe(sum(res.chase));
        expect(res.customers.length).toBe(res.chase.length + res.light.length);
        expect(res.chase.every((c) => c.periods >= 2)).toBe(true);
        expect(res.light.every((c) => c.periods <= 1)).toBe(true);
    });
});

describe('畫面偏好（各組收合）', () => {
    test('沒有或壞掉的資料用預設：全部收起', () => {
        for (const raw of [undefined, null, 'x', 5, [], {}]) {
            const p = normalizeUiPrefs(raw);
            expect(p.table).toEqual({four: true, three: true, two: true, one_eom: true, one: true, settled: true});
            expect(p.messages).toEqual(p.table);
        }
    });

    test('只採用布林值，其他欄位退回預設；總表與訊息區各自獨立', () => {
        const p = normalizeUiPrefs({table: {three: true, one: 'no', settled: false}, messages: {two: true}, extra: 1});
        expect(p.table).toEqual({four: true, three: true, two: true, one_eom: true, one: true, settled: false});
        expect(p.messages).toEqual({four: true, three: true, two: true, one_eom: true, one: true, settled: true});
        // 使用者自己展開過的（false）要記住
        expect(normalizeUiPrefs({table: {two: false}}).table).toMatchObject({two: false, three: true});
        // 新存過 four 就以它為準，不被舊的 three 蓋掉
        expect(normalizeUiPrefs({table: {three: true, four: false}}).table).toMatchObject({three: true, four: false});
    });

    test('存進再讀回來一致', () => {
        const p = normalizeUiPrefs({table: {three: true}, messages: {one: false}});
        expect(normalizeUiPrefs(JSON.parse(JSON.stringify(p)))).toEqual(p);
    });
});

describe('催款語氣與訊息範本', () => {
    const TODAY2 = new Date(2026, 9, 19, 10, 0, 0);
    // 1004 欠 4 期（6～9 月各 1 期上期）、1003 欠 3 期（7～9 月）、1002 欠 2 期、1001 欠 1 期
    const h = hist({
        '2026-06': {上: [owe('1004', '丁', 100)]},
        '2026-07': {上: [owe('1004', '丁', 110), owe('1003', '丙', 290)]},
        '2026-08': {上: [owe('1004', '丁', 120), owe('1003', '丙', 300), owe('1002', '乙', 200)]},
        '2026-09': {上: [owe('1004', '丁', 130), owe('1003', '丙', 310), owe('1002', '乙', 210), owe('1001', '甲', 50)]},
    });
    const get = (code: string, t = DEFAULT_TEMPLATES, m = emptyMarks(), hh = h) =>
        compute(hh, m, {}, TODAY2, t).customers.find((c) => c.code === code)!;

    test('新分組：4 期以上｜3 期｜2 期｜1 期（只顯示），欠越多越上面', () => {
        const res = compute(h, emptyMarks(), {}, TODAY2);
        expect(res.customers.map((c) => [c.code, c.periods, c.group])).toEqual(
            [['1004', 4, 'four'], ['1003', 3, 'three'], ['1002', 2, 'two'], ['1001', 1, 'one']]);
        expect(res.chase.map((c) => c.code)).toEqual(['1004', '1003', '1002']);
    });

    test('各級結尾句：4 期以上（含期數）／2～3 期／1 期', () => {
        expect(get('1004').message.split('\n').pop()).toBe('目前已累積 4 期帳款未付，再麻煩儘速匯款，謝謝配合');
        const two = '帳款已逾期，再麻煩儘早安排匯款，若已經匯款請忽略此訊息，謝謝！';
        expect(get('1003').message.split('\n').pop()).toBe(two);
        expect(get('1002').message.split('\n').pop()).toBe(two);
        expect(get('1001').message.split('\n').pop()).toBe('麻煩您方便時匯款，若已經匯款請忽略此訊息，謝謝！');
        expect(get('1004').message).toBe([
            '丁 您好～', '提醒您以下帳款尚未收到：', '・6月上期：100 元', '・7月上期：110 元', '・8月上期：120 元', '・9月上期：130 元',
            '合計：460 元', '', '目前已累積 4 期帳款未付，再麻煩儘速匯款，謝謝配合'].join('\n'));
    });

    test('{期數} 換成客人的期數（月結客人同月上下期合算 1 期）', () => {
        const t = {...defaultTemplates(), four_plus: '共 {期數} 期，{期數} 期！'};
        expect(get('1004', t).message.endsWith('共 4 期，4 期！')).toBe(true);
        expect(closingFor(5, t)).toBe('共 5 期，5 期！');
        const m = hist({'2026-06': {上: [owe('1005', '戊', 90, '月')], 下: [owe('1005', '戊', 90)]}, '2026-07': {上: [owe('1005', '戊', 90)], 下: [owe('1005', '戊', 90)]},
            '2026-08': {上: [owe('1005', '戊', 90)], 下: [owe('1005', '戊', 90)]}, '2026-09': {上: [owe('1005', '戊', 90)], 下: [owe('1005', '戊', 90)]}});
        const c = get('1005', DEFAULT_TEMPLATES, emptyMarks(), m);
        expect([c.periods, c.monthly, c.group]).toEqual([4, true, 'four']);
        expect(c.message).toContain('目前已累積 4 期帳款未付');
    });

    test('範本存取、壞資料用預設、恢復預設', async () => {
        const mem = new Map<string, string>();
        (globalThis as unknown as {localStorage: Storage}).localStorage = {
            getItem: (k: string) => mem.get(k) ?? null, setItem: (k: string, v: string) => void mem.set(k, v),
            removeItem: (k: string) => void mem.delete(k), clear: () => mem.clear(), key: () => null, length: 0,
        } as Storage;
        const store = await import('../src/infra/auto-collect-store');
        expect(store.loadTemplates()).toEqual(defaultTemplates());
        const mine = {...defaultTemplates(), one: '請儘快', four_plus: '已 {期數} 期'};
        store.saveTemplates(mine);
        expect(store.loadTemplates()).toEqual(mine);
        store.clearHistory();
        store.clearMarks(); // 清除記錄不動範本
        expect(store.loadTemplates()).toEqual(mine);
        store.resetTemplates();
        expect(store.loadTemplates()).toEqual(defaultTemplates());
        mem.set('bii.auto-collect.templates', '{壞掉');
        expect(store.loadTemplates()).toEqual(defaultTemplates());
        expect(normalizeTemplates({one: '  ', two_three: 5, four_plus: ' 好 '})).toEqual({...defaultTemplates(), four_plus: '好'});
    });

    test('改範本：自動訊息即時變；手動編輯過的維持她的版本，也不出現「欠款已變動」', () => {
        const EDIT = '丁 您好～\n・9月上期：130 元\n合計：460 元\n\n拜託了';
        const c0 = get('1004');
        const marks = setMessageEdit(emptyMarks(), c0, EDIT, TODAY2);
        const t = {...defaultTemplates(), four_plus: '新語氣 {期數}'};
        const c1 = get('1004', t, marks);
        expect(c1.message).toBe(EDIT);
        expect(c1.auto_message.endsWith('新語氣 4')).toBe(true);
        expect(c1.auto_message).not.toBe(c0.auto_message);
        expect(c1.edit?.stale).toBe(false);
        expect(get('1003', t).message.endsWith('帳款已逾期，再麻煩儘早安排匯款，若已經匯款請忽略此訊息，謝謝！')).toBe(true);
        // 欠款真的變了仍會提醒
        const added = compute(h, addManualDebt(marks, c0, {month_key: '2026-09', half: '下', amount: 64, reason: '少付'}, TODAY2), {}, TODAY2, t)
            .customers.find((c) => c.code === '1004')!;
        expect(added.edit?.stale).toBe(true);
    });

    test('messageBody 只留到「合計」那行', () => {
        expect(messageBody('A\n合計：1 元\n\n結尾\n多行')).toBe('A\n合計：1 元');
        expect(messageBody('沒有合計')).toBe('沒有合計');
    });

    test('下載的 Excel／文字檔用新語氣與新分組名稱', async () => {
        const res = compute(h, emptyMarks(), {}, TODAY2);
        const txt = buildMessagesText(res, '2026-10-19 12:00');
        expect(txt).toContain('欠 4 期以上（1 位）');
        expect(txt).toContain('目前已累積 4 期帳款未付，再麻煩儘速匯款，謝謝配合');
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await (await buildSummaryXlsx(res)).arrayBuffer());
        expect(JSON.stringify(wb.getWorksheet('催款訊息')!.getSheetValues())).toContain('再麻煩儘速匯款');
        expect(String(wb.getWorksheet('催款總表')!.getRow(2).getCell(3).value)).toContain('4 期以上');
    });
});

describe('篩選時全展開', () => {
    test('expandAll 全部展開且不改原偏好', () => {
        const p = normalizeUiPrefs({table: {two: true}});
        const e = expandAll(p);
        expect(Object.values(e.table).every((v) => v === false)).toBe(true);
        expect(Object.values(e.messages).every((v) => v === false)).toBe(true);
        expect(p.table.two).toBe(true);
    });
});

describe('已結清快速原因按鈕清單', () => {
    test('預設兩顆，壞資料回預設', () => {
        expect(defaultReasonChips()).toEqual(['已補款', '回饋金自扣']);
        expect(normalizeReasonChips(null)).toEqual(['已補款', '回饋金自扣']);
        expect(normalizeReasonChips('abc')).toEqual(['已補款', '回饋金自扣']);
    });
    test('濾掉非字串、空白、重複；空陣列合法', () => {
        expect(normalizeReasonChips([' 已補款 ', 3, '', '已補款', null, '現金'])).toEqual(['已補款', '現金']);
        expect(normalizeReasonChips([])).toEqual([]);
    });
    test('新增：空白與重複不加；刪除預設也可以', () => {
        const base = defaultReasonChips();
        expect(addReasonChip(base, '  ')).toEqual(base);
        expect(addReasonChip(base, '已補款')).toEqual(base);
        expect(addReasonChip(base, ' 匯款 ')).toEqual(['已補款', '回饋金自扣', '匯款']);
        expect(removeReasonChip(base, '已補款')).toEqual(['回饋金自扣']);
    });
});
