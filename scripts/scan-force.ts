import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const s = wb.getWorksheet('海線上')!;
    console.log(`actualRowCount=${s.actualRowCount} rowCount=${s.rowCount} dimensions=${s.dimensions?.toString?.() ?? JSON.stringify(s.dimensions)}`);
    let at = 0, lastContentRow = 0;
    const atRows: number[] = [];
    for (let r = 1; r <= 1000; r++) {
        for (let c = 1; c <= 30; c++) {
            const v = textOfExcelJs(s.getRow(r).getCell(c).value).trim();
            if (v) {
                lastContentRow = r;
            }
            if (v.includes('@')) {
                at++;
                atRows.push(r);
            }
        }
    }
    console.log(`強制掃 1~1000 列 × 1~30 欄：@@=${at}  最後有內容的列=${lastContentRow}`);
    console.log(`@@ 所在列：${atRows.join(',')}`);
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
