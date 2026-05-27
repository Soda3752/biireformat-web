/**
 * 海線檔草稿解析器：以「數字代號 OR 中文店名」當客戶邊界，拆出比 36 更多的店。
 * 用於產生草稿清單給使用者核對。跑法：npx tsx scripts/parse-haixian-draft.ts <path> [--json out.json]
 *
 * 注意：店名 vs 休息日備註用黑名單啟發式判斷，必然有誤差，需使用者核對。
 */
import {writeFileSync} from 'node:fs';

import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];
if (!path) throw new Error('需要檔案路徑參數');
const jsonIdx = process.argv.indexOf('--json');
const jsonOut = jsonIdx >= 0 ? process.argv[jsonIdx + 1] : '';

const TITLE_RE = /海線|^\(/;

function isId(a: string): boolean {
    return /^\d{3,5}$/.test(a);
}

function isPhone(a: string): boolean {
    const d = a.replace(/[^\d]/g, '');
    return /^\d{7,15}$/.test(d) && !isId(a);
}

/** 休息日 / 付款 / 備註判斷（命中即「非店名」）。盡量精準避免誤殺店名。 */
function isNote(a: string): boolean {
    if (/休|送|連|現|發票|老闆|店長|固/.test(a)) return true;   // 日休/日送/二連/收現/發票/老闆/店長/固定
    if (/[次天]/.test(a) && /\d/.test(a)) return true;            // 2天1次 / 10天一次
    if (/^皮\d/.test(a)) return true;                             // 皮1包 / 皮2包
    if (/削/.test(a)) return true;                                // 3削 日4削（A欄出現削=備註，品名在B欄）
    return false;
}

function isStoreName(a: string): boolean {
    if (!a) return false;
    if (isId(a) || isPhone(a)) return false;
    if (a === '客戶名稱' || a === '品名' || a === '單' || a === '價') return false;
    if (TITLE_RE.test(a)) return false;
    if (isNote(a)) return false;
    return /[一-鿿a-zA-Z]/.test(a);
}

interface Seg {
    id: string;
    name: string;
    products: string[];
    notes: string[];
    phones: string[];
    rows: number[];
    hasCode: boolean;
}

async function main(): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const sheet = wb.getWorksheet('海線上') ?? wb.worksheets[0];
    const totalRows = sheet.actualRowCount || sheet.rowCount;

    const colA: string[] = [], colB: string[] = [], colC: string[] = [];
    for (let r = 1; r <= totalRows; r++) {
        colA[r] = textOfExcelJs(sheet.getRow(r).getCell(1).value).trim();
        colB[r] = textOfExcelJs(sheet.getRow(r).getCell(2).value).trim();
        colC[r] = textOfExcelJs(sheet.getRow(r).getCell(3).value).trim();
    }

    const segs: Seg[] = [];
    let cur: Seg | null = null;
    const newSeg = (id: string, name: string, r: number, hasCode: boolean): Seg => {
        const s: Seg = {id, name, products: [], notes: [], phones: [], rows: [r], hasCode};
        segs.push(s);
        return s;
    };

    for (let r = 1; r <= totalRows; r++) {
        const a = colA[r], b = colB[r];
        // 跳過標題與欄頭列
        if (a === '客戶名稱' || a === '品名' || TITLE_RE.test(a)) continue;

        if (isId(a)) {
            cur = newSeg(a, '', r, true);
            if (b) cur.products.push(b);
            continue;
        }
        if (isStoreName(a)) {
            if (cur && !cur.name && cur.hasCode && cur.rows[cur.rows.length - 1] >= r - 2) {
                // 緊跟代號的第一個店名 → 填入當前（同一間店）
                cur.name = a;
                cur.rows.push(r);
            } else {
                cur = newSeg('', a, r, false);  // 子店（無獨立代號）
            }
            if (b) cur.products.push(b);
            continue;
        }
        // note / phone / 純品名延續
        if (cur) {
            cur.rows.push(r);
            if (isPhone(a)) cur.phones.push(a);
            else if (a) cur.notes.push(a);
            if (b) cur.products.push(b);
        }
    }

    // 輸出清單
    segs.forEach((s, i) => {
        const flag = s.hasCode ? '' : '  ⚠️無代號(疑子店)';
        const rs = `r${s.rows[0]}~${s.rows[s.rows.length - 1]}`;
        console.log(`[${String(i + 1).padStart(2)}] 代號=${(s.id || '—').padEnd(6)} 店名=${(s.name || '(空)').padEnd(12)}${flag}`);
        console.log(`     品名(${s.products.length}): ${s.products.join('、')}`);
        if (s.notes.length) console.log(`     休/備: ${s.notes.join(' | ')}`);
        if (s.phones.length) console.log(`     電話: ${s.phones.join(' | ')}`);
        console.log(`     來源: ${rs}`);
    });
    console.log(`\n=== 草稿共拆出 ${segs.length} 間（其中無代號疑似子店 ${segs.filter((s) => !s.hasCode).length} 間）===`);

    if (jsonOut) {
        writeFileSync(jsonOut, JSON.stringify(segs, null, 2), 'utf-8');
        console.log(`已輸出草稿 JSON：${jsonOut}`);
    }
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
