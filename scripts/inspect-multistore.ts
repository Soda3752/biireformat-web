/**
 * 多店候選 flagger：以數字代號切 block，統計每個 block 的電話數。
 * 2+ 電話 = 可能含多間店。跑法：npx tsx scripts/inspect-multistore.ts <path>
 */
import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];
if (!path) throw new Error('需要檔案路徑參數');

function looksLikePhone(s: string): boolean {
    const d = s.replace(/-/g, '').replace(/[^\d]/g, '');
    return /^\d{7,15}$/.test(d) && !/^\d{3,5}$/.test(s.trim());
}

async function main(): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const sheet = wb.getWorksheet('海線上') ?? wb.worksheets[0];
    const totalRows = sheet.actualRowCount || sheet.rowCount;

    const colA: string[] = [];
    const colB: string[] = [];
    for (let r = 1; r <= totalRows; r++) {
        colA[r] = textOfExcelJs(sheet.getRow(r).getCell(1).value).trim();
        colB[r] = textOfExcelJs(sheet.getRow(r).getCell(2).value).trim();
    }

    // 代號列
    const idRows: number[] = [];
    for (let r = 1; r <= totalRows; r++) {
        if (/^\d{3,5}$/.test(colA[r])) idRows.push(r);
    }

    let multi = 0;
    let extraStores = 0;
    for (let i = 0; i < idRows.length; i++) {
        const start = idRows[i];
        const end = (i + 1 < idRows.length ? idRows[i + 1] : totalRows + 1);
        // 排除中間的欄頭列（客戶名稱）
        let phones = 0;
        const aTexts: string[] = [];
        for (let r = start; r < end; r++) {
            if (colA[r] === '客戶名稱') continue;
            if (looksLikePhone(colA[r])) phones++;
            else if (colA[r]) aTexts.push(`${colA[r]}`);
        }
        if (phones >= 2) {
            multi++;
            extraStores += phones - 1;
            console.log(`代號 ${colA[start]} (r${start})  電話數=${phones}  A欄內容: ${aTexts.join(' | ')}`);
        }
    }

    console.log(`\n代號 block 總數：${idRows.length}`);
    console.log(`多電話(2+) block 數：${multi}`);
    console.log(`若每多 1 電話 = 多 1 店，推估總店數：${idRows.length + extraStores}`);
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
