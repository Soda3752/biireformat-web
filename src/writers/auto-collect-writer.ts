/**
 * 「自動催款」產出：催款總表 xlsx 與全部訊息文字檔（對應獨立工具 auto_collect/report.py）。
 */

import ExcelJS from 'exceljs';

import {amountsText, type CollectResult, type Debtor, GROUPS, money, periodsText} from '@/domain/auto-collect/ledger';

const fill = (argb: string): ExcelJS.Fill => ({type: 'pattern', pattern: 'solid', fgColor: {argb}});
const HEADER_FILL = fill('FF3E5F8A');
const PARTIAL_FILL = fill('FFFFF4D6');
const NOTE_FILL = fill('FFFADBD8');
const THIN: Partial<ExcelJS.Border> = {style: 'thin', color: {argb: 'FFD0D7DE'}};
const BORDER: Partial<ExcelJS.Borders> = {top: THIN, left: THIN, bottom: THIN, right: THIN};

function addSheet(wb: ExcelJS.Workbook, title: string, header: string[], widths: number[]): ExcelJS.Worksheet {
    const ws = wb.addWorksheet(title, {views: [{state: 'frozen', ySplit: 1}]});
    ws.addRow(header).eachCell((c) => {
        c.fill = HEADER_FILL;
        c.font = {bold: true, color: {argb: 'FFFFFFFF'}, size: 12};
        c.border = BORDER;
        c.alignment = {horizontal: 'center', vertical: 'middle'};
    });
    widths.forEach((w, i) => (ws.getColumn(i + 1).width = w));
    return ws;
}

function addStyledRow(ws: ExcelJS.Worksheet, values: unknown[], moneyCols: number[] = [], rowFill?: ExcelJS.Fill): ExcelJS.Row {
    const row = ws.addRow(values);
    for (let c = 1; c <= values.length; c++) {
        const cell = row.getCell(c);
        cell.border = BORDER;
        cell.alignment = {wrapText: true, vertical: 'top'};
        if (moneyCols.includes(c)) cell.numFmt = '#,##0';
        if (rowFill) cell.fill = rowFill;
    }
    return row;
}

const groupText = (c: Debtor): string => `${c.group === 'four' ? '4 期以上' : c.group === 'three' ? '3 期' : c.group === 'two' ? '2 期' : '月結本月該付'}（${c.periods} 期${c.monthly ? '，月結' : ''}）`;

export async function buildSummaryXlsx(result: CollectResult): Promise<Blob> {
    const wb = new ExcelJS.Workbook();
    // 只含要催款的（欠 2 期以上，或月結本月該付）；只顯示不催的與未到期的只在畫面上顯示，不進檔案
    const customers = result.chase;

    let ws = addSheet(wb, '催款總表', ['代號', '客人', '欠幾期', '欠哪幾期', '各期欠款', '合計', '備註'], [8, 18, 14, 30, 44, 12, 60]);
    for (const c of customers) {
        const row = addStyledRow(ws, [c.code, c.name, groupText(c), periodsText(c), amountsText(c), c.total, c.notes.join('\n')], [6]);
        if (c.notes.length) row.getCell(7).fill = NOTE_FILL;
    }
    addStyledRow(ws, ['', '合計', '', `${customers.length} 位客人`, '', result.chase_total, ''], [6])
        .eachCell((c) => (c.font = {bold: true}));

    ws = addSheet(wb, '各期明細', ['代號', '客人', '期別', '帳單金額', '已收', '欠款', '收款日／備註（原檔）', '情形'],
        [8, 18, 14, 12, 12, 12, 30, 10]);
    for (const c of customers) {
        for (const i of c.lines) {
            addStyledRow(ws, [c.code, c.name, i.period, i.manual ? null : i.amt, i.paid || null, i.owed,
                i.manual ? i.manual.reason : [i.date_text, i.memo].filter(Boolean).join('；'),
                i.manual ? '手動加入' : i.partial ? '部分付款' : '未付'],
            [4, 5, 6], i.partial ? PARTIAL_FILL : undefined);
        }
    }

    ws = addSheet(wb, '催款訊息', ['代號', '客人', '合計', '訊息（可直接複製貼到 LINE）'], [8, 18, 12, 70]);
    for (const c of customers) addStyledRow(ws, [c.code, c.name, c.total, c.message], [3]);

    ws = addSheet(wb, '資料來源', ['月份', '採用的檔案', '檔案修改時間', '匯入時間', '讀到筆數'], [12, 34, 22, 22, 10]);
    for (const m of result.months) addStyledRow(ws, [m.label, m.source, m.file_mtime, m.imported_at, m.count]);

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}

export function buildMessagesText(result: CollectResult, stamp: string): string {
    const parts = [`催款訊息（${stamp} 產出，共 ${result.chase.length} 位，合計 ${money(result.chase_total)} 元）`, ''];
    for (const g of GROUPS.filter((x) => x.key !== 'one')) {
        const list = result.chase.filter((c) => c.group === g.key);
        if (!list.length) continue;
        parts.push(`########## ${g.label}（${list.length} 位）##########`, '');
        for (const c of list) {
            parts.push(`===== ${c.code} ${c.name}（合計 ${money(c.total)} 元）=====`);
            parts.push(...c.notes.map((n) => `【注意】${n}`));
            parts.push(c.message, '');
        }
    }
    return parts.join('\n');
}
