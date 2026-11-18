/**
 * 「自動催款」收款登記測試：bun test tests/payment-register.test.ts
 * 列出未付清期別、月結合併、金額分配，以及寫回 Excel 只動實收金額／收款日（＋備註）、其他格不變。
 */
import {describe, expect, test} from 'bun:test';

import ExcelJS from 'exceljs';

import {emptyHistory, emptyMarks, importFiles, settleDebtor, setMonthly} from '../src/domain/auto-collect/ledger';
import {allocate, buildEntries, candidateFilenames, registerCustomers, searchRegister, toMd} from '../src/domain/auto-collect/payment-register';
import {parseStatement} from '../src/domain/auto-collect/statement-parser';
import {cellVal, registerPayments} from '../src/domain/auto-reconcile-engine';

const TODAY = new Date(2026, 10, 18);

type Row = [string | null, string, number, number | null, string | null, string?];

/** 對帳 2.0 新版面：左右兩塊，各塊 代號／客戶名稱／金額／實收金額／收款日／備註；金額欄 14 號字置中。 */
async function makeStmt(upper: Row[], lower: Row[]): Promise<Uint8Array> {
    const wb = new ExcelJS.Workbook();
    for (const [title, rows] of [['上期', upper], ['下期', lower]] as const) {
        const ws = wb.addWorksheet(title);
        ws.addRow(['青坊食品行 月結明細']);
        const head = ['代號', '客戶名稱', '金額', '實收金額', '收款日', '備註'];
        ws.addRow([...head, ...head]);
        rows.forEach((r, i) => {
            // 前半放左塊、後半放右塊
            const half = Math.ceil(rows.length / 2);
            const rowNo = 3 + (i < half ? i : i - half);
            const base = i < half ? 1 : 7;
            const vals: unknown[] = [r[0], r[1], r[2], r[3], r[4], r[5] ? {richText: [{text: r[5], font: {color: {argb: 'FFC00000'}}}]} : null];
            vals.forEach((v, j) => {
                const c = ws.getCell(rowNo, base + j);
                c.value = v as ExcelJS.CellValue;
                c.border = {top: {style: 'thin'}, bottom: {style: 'thin'}};
            });
            const amt = ws.getCell(rowNo, base + 2);
            amt.font = {size: 14};
            amt.alignment = {horizontal: 'center', vertical: 'middle'};
        });
        const total = ws.getCell(20, 3);
        total.value = {formula: 'SUM(C3:C9)', result: 0};
    }
    wb.addWorksheet('銀行對帳結果').addRow(['日期', '摘要']);
    return new Uint8Array(await wb.xlsx.writeBuffer());
}

const load = async (data: Uint8Array): Promise<ExcelJS.Workbook> => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data.slice().buffer as ArrayBuffer);
    return wb;
};

/** 整本活頁簿每格的值與樣式快照（位址 → JSON）。 */
function snapshot(wb: ExcelJS.Workbook): Map<string, string> {
    const out = new Map<string, string>();
    for (const ws of wb.worksheets) {
        ws.eachRow({includeEmpty: true}, (row) => row.eachCell({includeEmpty: true}, (c) => {
            out.set(`${ws.name}!${c.address}`, JSON.stringify({v: c.value, st: c.style}));
        }));
    }
    return out;
}

async function setup() {
    const aug = await makeStmt(
        [['1001', '賴先生', 12000, null, null], ['1002', '王小姐', 3000, 3000, '8/10']],
        [['1001', '賴先生', 11000, 5000, '9/2', '漏付，8上已付'], ['1002', '王小姐', 3100, null, null]],
    );
    const sep = await makeStmt(
        [['1001', '賴先生', 9000, null, null], ['1002', '王小姐', 2000, 1985, '9/12']],
        [['1001', '賴先生', 9500, null, null], ['1002', '王小姐', 2100, 2100, '10/1']],
    );
    const files = [{name: '8月_已自動對帳.xlsx', mtime: 1, data: aug}, {name: '9月_已自動對帳.xlsx', mtime: 2, data: sep}];
    const {history} = await importFiles(files, emptyHistory(), {today: TODAY});
    return {history, aug, sep};
}

describe('列出未付清的期別', () => {
    test('搜尋「賴」找到賴先生，列出未付清各期（部分付款的列差額）', async () => {
        const {history} = await setup();
        const list = registerCustomers(history, emptyMarks(), TODAY);
        const [lai] = searchRegister(list, '賴');
        expect(lai.name).toBe('賴先生');
        expect(lai.lines.map((l) => [l.label, l.amt, l.paid, l.owed])).toEqual([
            ['8月上期', 12000, 0, 12000], ['8月下期', 11000, 5000, 6000], ['9月上期', 9000, 0, 9000], ['9月下期', 9500, 0, 9500]]);
        expect(lai.total).toBe(36500);
        expect(searchRegister(list, '1001')[0].name).toBe('賴先生');
        // 王小姐 9 月上期扣手續費 15 不算欠
        const wang = list.find((c) => c.code === '1002')!;
        expect(wang.lines.map((l) => l.label)).toEqual(['8月下期']);
    });

    test('已結清的期別不列；月結客人同月上下期合成一行', async () => {
        const {history} = await setup();
        let marks = settleDebtor(emptyMarks(), {code: '1001', name: '賴先生', items: [
            {period: '8月上期', month_key: '2026-08', half: '上', amt: 12000, paid: 0, owed: 12000, partial: false, date_text: '', memo: ''}]}, '');
        let lai = registerCustomers(history, marks, TODAY).find((c) => c.code === '1001')!;
        expect(lai.lines.map((l) => l.label)).toEqual(['8月下期', '9月上期', '9月下期']);
        marks = setMonthly(emptyMarks(), {code: '1001', name: '賴先生', monthly_auto: false}, true);
        lai = registerCustomers(history, marks, TODAY).find((c) => c.code === '1001')!;
        expect(lai.monthly).toBe(true);
        expect(lai.lines.map((l) => [l.label, l.owed, l.periods.length])).toEqual([['8月', 18000, 2], ['9月', 18500, 2]]);
        // 月結一行的金額先補上期，剩下算下期
        expect(allocate(lai.lines[0], 18000)).toEqual([12000, 6000]);
        expect(allocate(lai.lines[0], 10000)).toEqual([10000, 0]);
        expect(allocate(lai.lines[0], 20000)).toEqual([12000, 8000]);
    });

    test('檔名候選：當初匯入的檔名，其次是預設檔名', async () => {
        const {history} = await setup();
        expect(candidateFilenames(history, '2026-08')).toEqual(['8月_已自動對帳.xlsx']);
        expect(toMd('2026-11-05')).toBe('11/5');
    });
});

describe('寫回 Excel', () => {
    test('實收金額與收款日寫到正確格子、其他格不變，再讀回欠款消失', async () => {
        const {history, aug} = await setup();
        const lai = registerCustomers(history, emptyMarks(), TODAY).find((c) => c.code === '1001')!;
        const picks = lai.lines.filter((l) => l.month_key === '2026-08').map((line) => ({line, amount: line.owed}));
        const byMonth = buildEntries(picks, '2026-11-18');
        expect([...byMonth.keys()]).toEqual(['2026-08']);

        // 比對基準也經過一次讀寫（ExcelJS 存檔會把樣式整理成固定寫法）
        const before = snapshot(await load(new Uint8Array(await (await load(aug)).xlsx.writeBuffer())));
        const wb = await load(aug);
        expect(registerPayments(wb, 8, byMonth.get('2026-08')!)).toEqual([]);
        const out = new Uint8Array(await wb.xlsx.writeBuffer());
        const after = snapshot(await load(out));

        const up = (await load(out)).getWorksheet('上期')!;
        const lo = (await load(out)).getWorksheet('下期')!;
        // 上期：賴先生在左塊第 3 列 → D3 實收、E3 收款日
        expect(cellVal(up.getCell('D3'))).toBe(12000);
        expect(cellVal(up.getCell('E3'))).toBe('11/18');
        expect(up.getCell('D3').font?.size).toBe(14);
        expect((up.getCell('E3').fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FFE2EFDA');
        // 下期原本收過 5000 → 加上這次 6000；漏付紅字改成已補付，加 2筆
        expect(cellVal(lo.getCell('D3'))).toBe(11000);
        expect(cellVal(lo.getCell('E3'))).toBe('11/18');
        expect(cellVal(lo.getCell('F3'))).toBe('已補付 11/18；2筆');

        // 只有這 6 格變了（D3/E3 兩張表、下期 F3）；其他格的值與樣式都一樣
        const changed = [...new Set([...before.keys(), ...after.keys()])].filter((k) => before.get(k) !== after.get(k)).sort();
        expect(changed).toEqual(['上期!D3', '上期!E3', '下期!D3', '下期!E3', '下期!F3']);

        // 讀回：8 月賴先生付清，9 月照舊
        const st = await parseStatement(out, '8月_已自動對帳.xlsx');
        expect(st.periods.上['1001']).toMatchObject({paid: 12000, date_text: '11/18'});
        expect(st.periods.下['1001']).toMatchObject({paid: 11000, date_text: '11/18'});
        const next = {...history, months: {...history.months, '2026-08': {...history.months['2026-08'], periods: st.periods}}};
        const again = registerCustomers(next, emptyMarks(), TODAY).find((c) => c.code === '1001')!;
        expect(again.lines.map((l) => l.label)).toEqual(['9月上期', '9月下期']);
    });

    test('右邊區塊的客人也寫到正確欄', async () => {
        const {history, aug} = await setup();
        const wang = registerCustomers(history, emptyMarks(), TODAY).find((c) => c.code === '1002')!;
        const byMonth = buildEntries([{line: wang.lines[0], amount: 3085}], '2026-11-03');
        const wb = await load(aug);
        expect(registerPayments(wb, 8, byMonth.get('2026-08')!)).toEqual([]);
        const lo = wb.getWorksheet('下期')!;
        // 下期兩人 → 王小姐在右塊（G 欄起）第 3 列：J3 實收、K3 收款日
        expect(cellVal(lo.getCell('J3'))).toBe(3085);
        expect(cellVal(lo.getCell('K3'))).toBe('11/3');
        expect(cellVal(lo.getCell('D3'))).toBe(5000); // 左塊賴先生不動
    });

    test('Excel 跟網站記住的不同（已登記過、金額改過）就一筆都不寫', async () => {
        const {history, aug} = await setup();
        const lai = registerCustomers(history, emptyMarks(), TODAY).find((c) => c.code === '1001')!;
        const entries = buildEntries(lai.lines.filter((l) => l.month_key === '2026-08').map((line) => ({line, amount: line.owed})), '2026-11-18').get('2026-08')!;
        const wb = await load(aug);
        expect(registerPayments(wb, 8, entries)).toEqual([]);
        // 同一份檔再登記一次：原本實收對不上 → 擋下
        const before = snapshot(wb);
        const probs = registerPayments(wb, 8, entries);
        expect(probs.map((p) => p.half)).toEqual(['上', '下']);
        expect(probs[0].reason).toContain('已經填了實收');
        expect(snapshot(wb)).toEqual(before);
        // 找不到客人
        const wb2 = await load(aug);
        expect(registerPayments(wb2, 8, [{...entries[0], key: '9999'}])[0].reason).toContain('找不到');
        // 帳單金額不同
        expect(registerPayments(wb2, 8, [{...entries[0], expectAmt: 1}])[0].reason).toContain('帳單金額');
    });
});
