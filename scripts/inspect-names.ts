/**
 * 統計「海線上」A 欄中「疑似店名」的數量，判斷是否有無代號店家被漏算。
 * 跑法：npx tsx scripts/inspect-names.ts <path>
 */
import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];
if (!path) throw new Error('需要檔案路徑參數');

// 休息日 / 付款 / 備註關鍵字（出現即視為非店名）
const NOTE_KW = ['休', '現', '收', '送', '連', '包', '老闆', '店長', '次', '固定', '發票', '天', '削', '個', '元'];

function looksLikePhone(s: string): boolean {
    const d = s.replace(/-/g, '');
    return /^\d{7,15}$/.test(d);
}

function looksLikeName(s: string): boolean {
    if (!s) return false;
    if (/^\d/.test(s)) return false;               // 數字開頭（代號/電話）
    if (looksLikePhone(s)) return false;
    if (s.length > 10) return false;
    if (NOTE_KW.some((k) => s.includes(k))) return false;
    return /[一-鿿]/.test(s);               // 含中文
}

async function main(): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const sheet = wb.getWorksheet('海線上') ?? wb.worksheets[0];
    const totalRows = sheet.actualRowCount || sheet.rowCount;

    const ids: Array<{ r: number; v: string }> = [];
    const names: Array<{ r: number; v: string }> = [];
    for (let r = 1; r <= totalRows; r++) {
        const a = textOfExcelJs(sheet.getRow(r).getCell(1).value).trim();
        if (/^\d{3,5}$/.test(a)) ids.push({r, v: a});
        else if (looksLikeName(a)) names.push({r, v: a});
    }

    console.log(`代號(3~5位數字) 數量：${ids.length}`);
    console.log(`疑似店名 數量：${names.length}`);
    console.log('\n--- 疑似店名清單（含列號）---');
    names.forEach((n) => console.log(`  r${String(n.r).padStart(3)}  ${n.v}`));
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
