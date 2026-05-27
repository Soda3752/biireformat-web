/**
 * 一次性檔案結構探查腳本。
 * 跑法：npx tsx scripts/inspect-xlsx.ts <path>
 */
import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];
if (!path) throw new Error('需要檔案路徑參數');

async function main(): Promise<void> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);

    console.log('=== 分頁清單 ===');
    wb.worksheets.forEach((s) => {
        console.log(`  "${s.name}"  state=${s.state}  rows=${s.actualRowCount}/${s.rowCount}  cols=${s.actualColumnCount}/${s.columnCount}`);
    });

    for (const s of wb.worksheets) {
        console.log(`\n=== 分頁 "${s.name}" 的 A/B/C 欄逐列（僅印 A 或 B 非空者）===`);
        const totalRows = s.actualRowCount || s.rowCount;
        let idLike = 0;
        for (let r = 1; r <= totalRows; r++) {
            const a = textOfExcelJs(s.getRow(r).getCell(1).value).trim();
            const b = textOfExcelJs(s.getRow(r).getCell(2).value).trim();
            const c = textOfExcelJs(s.getRow(r).getCell(3).value).trim();
            if (a || b) {
                const isId = /^\d{3,5}$/.test(a);
                if (isId) idLike++;
                console.log(`  r${String(r).padStart(3)}  A=${JSON.stringify(a).padEnd(20)} B=${JSON.stringify(b).padEnd(18)} C=${JSON.stringify(c)}${isId ? '   <-- ID樣式' : ''}`);
            }
        }
        console.log(`  >>> 符合「3~5位純數字」(目前 reader 認的客戶起始) 的列數：${idLike}`);
    }
}

main().catch((e) => {
    console.error(e);
    process.exit(1);
});
