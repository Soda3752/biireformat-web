/** 驗證：注入 manifest 到副本後，readHandfillBook 是否走 manifest 路徑讀回 38 間。 */
import {readFileSync} from 'node:fs';
import ExcelJS from 'exceljs';
import {HANDFILL_MANIFEST_SHEET} from '@/domain/models/handfill-book';
import {readHandfillBook} from '@/readers/handfill-reader';

const xlsx = process.argv[2];
const manifestPath = process.argv[3];

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(xlsx);
    const manifestRaw = readFileSync(manifestPath, 'utf-8');
    const ws = wb.addWorksheet(HANDFILL_MANIFEST_SHEET);
    ws.state = 'veryHidden';
    ws.getCell(1, 1).value = manifestRaw;
    const buf = await wb.xlsx.writeBuffer();

    // 包成 File 餵 reader
    const file = new File([buf], '6.海線.xlsx', {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const book = await readHandfillBook(file as unknown as File);
    console.log(`讀回客戶數：${book.customers.length}`);
    console.log(`線別：(${book.lineNo})${book.lineName}  ${book.year}年${book.month}月`);
    console.log('前 3 間：', book.customers.slice(0, 3).map(c => `${c.customerId || '—'}/${c.customerName}`).join('  '));
    console.log('梧棲那兩間：', book.customers.filter(c => c.customerName.includes('梧棲')).map(c => `${c.customerId || '—'}/${c.customerName}`).join('  '));
    console.log('末間：', (() => {
        const c = book.customers[book.customers.length - 1];
        return `${c.customerId || '—'}/${c.customerName}`;
    })());
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
