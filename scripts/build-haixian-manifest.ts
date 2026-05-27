/**
 * 以「數量欄(c19) 的 $$」為每間店尾界，切段建 HandfillBook，輸出 manifest.json。
 * 跑法：npx tsx scripts/build-haixian-manifest.ts <xlsx> [--json out.json]
 */
import {writeFileSync} from 'node:fs';
import ExcelJS from 'exceljs';
import {HANDFILL_MANIFEST_VERSION, layoutHashOfWorkbook, textOfExcelJs} from '@/infra/handfill-manifest';
import {createEmptyBook, genId, type HandfillBook, type HandfillCustomer} from '@/domain/models/handfill-book';

const path = process.argv[2];
if (!path) throw new Error('需要檔案路徑');
const jsonIdx = process.argv.indexOf('--json');
const jsonOut = jsonIdx >= 0 ? process.argv[jsonIdx + 1] : '';

const DOLLAR_COL = 19;
/** 店尾界標記：相容 $$ 與 @@（使用者兩種都用過）。 */
const isMarker = (s: string) => s.includes('$') || s.includes('@');
const isCode = (s: string) => /^\d{3,5}$/.test(s.trim());
const isHeaderOrTitle = (a: string, b: string) =>
    a === '客戶名稱' || a === '品名' || b === '品名' || /^[(（]\s*[一二三四五六七八九十\d]+\s*[)）]/.test(a);
const looksLikePhone = (s: string) => {
    const d = s.trim().replace(/-/g, '');
    return /^\d{7,15}$/.test(d);
};

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const sheet = wb.getWorksheet('海線上') ?? wb.worksheets[0];
    // 注意：ExcelJS 的 actualRowCount 在此檔誤報 348，但真正資料到 r504（以 dimensions 為準）。
    // 用 actualRowCount 當上界會截掉尾段約 18 間店，故取 dimensions.bottom 為真實底列。
    const totalRows = Math.max(sheet.dimensions?.bottom ?? 0, sheet.actualRowCount || 0) || sheet.rowCount;

    const A: string[] = [], B: string[] = [], C: string[] = [], D: string[] = [];
    for (let r = 1; r <= totalRows; r++) {
        A[r] = textOfExcelJs(sheet.getRow(r).getCell(1).value).trim();
        B[r] = textOfExcelJs(sheet.getRow(r).getCell(2).value).trim();
        C[r] = textOfExcelJs(sheet.getRow(r).getCell(3).value).trim();
        D[r] = textOfExcelJs(sheet.getRow(r).getCell(DOLLAR_COL).value).trim();
    }

    // 找標記列（店尾界，$$ 或 @@）
    const dollarRows: number[] = [];
    for (let r = 1; r <= totalRows; r++) if (isMarker(D[r])) dollarRows.push(r);

    const book: HandfillBook = createEmptyBook();
    book.id = genId('book-');
    book.lineNo = 6;
    book.lineName = '海線';
    // 從標題列 r1 抓年月（c15~c20）
    for (let c = 14; c <= 20; c++) {
        const n = parseInt(textOfExcelJs(sheet.getRow(1).getCell(c).value).trim(), 10);
        if (Number.isFinite(n) && n > 100 && n < 200) book.year = n;
    }

    const customers: HandfillCustomer[] = [];
    let segStart = 1;
    for (const dRow of dollarRows) {
        // 段 = segStart .. dRow（含），但跳過頁首/標題列
        const cust: HandfillCustomer = {
            id: genId('cust-'), customerId: '', customerName: '',
            products: [], restNotes: [], phones: [], manualSort: false,
        };
        const nameCandidates: string[] = [];
        for (let r = segStart; r <= dRow; r++) {
            const a = A[r], b = B[r], c = C[r];
            if (isHeaderOrTitle(a, b)) continue;
            // col A
            if (a) {
                if (isCode(a) && !cust.customerId) {
                    cust.customerId = a;
                } else if (looksLikePhone(a)) {
                    cust.phones.push(a);
                } else if (!cust.customerName && !isCode(a)) {
                    cust.customerName = a;  // 第一個非代號非電話 → 店名
                } else {
                    cust.restNotes.push(a);
                }
            }
            // col B/C 品名+單價
            const price = c ? parseFloat(c) : NaN;
            if (b || Number.isFinite(price)) {
                cust.products.push({name: b, unitPrice: Number.isFinite(price) ? price : undefined});
            }
        }
        if (cust.products.length === 0) cust.products.push({name: '', unitPrice: undefined});
        while (cust.phones.length < 2) cust.phones.push('');
        // 整段全空（純頁首）→ 跳過
        if (cust.customerId || cust.customerName || cust.products.some(p => p.name)) {
            customers.push(cust);
        }
        segStart = dRow + 1;
    }

    book.customers = customers;

    // 預覽
    customers.forEach((c, i) => {
        const prods = c.products.map(p => p.name + (p.unitPrice != null ? `(${p.unitPrice})` : '')).filter(Boolean).join('、');
        const ph = c.phones.filter(Boolean).join(',');
        const notes = c.restNotes.join('|');
        console.log(`[${String(i + 1).padStart(2)}] 代號=${(c.customerId || '—').padEnd(6)} 店名=${(c.customerName || '(空)').padEnd(10)} 品(${c.products.filter(p => p.name).length}): ${prods}${notes ? `  休/備:${notes}` : ''}${ph ? `  ☎${ph}` : ''}`);
    });
    console.log(`\n=== 共 ${customers.length} 間（$$ 標記 ${dollarRows.length} 個）===`);

    if (jsonOut) {
        const manifest = {version: HANDFILL_MANIFEST_VERSION, book, layoutHash: layoutHashOfWorkbook(wb)};
        writeFileSync(jsonOut, JSON.stringify(manifest, null, 2), 'utf-8');
        console.log(`已輸出 manifest：${jsonOut}（layoutHash=${manifest.layoutHash}）`);
    }
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
