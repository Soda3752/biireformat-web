import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    for (const s of wb.worksheets) {
        const totalRows = s.actualRowCount || s.rowCount;
        const totalCols = s.actualColumnCount || s.columnCount;
        const hits: Array<{ r: number; c: number; v: string }> = [];
        for (let r = 1; r <= totalRows; r++) {
            for (let c = 1; c <= totalCols; c++) {
                const v = textOfExcelJs(s.getRow(r).getCell(c).value).trim();
                if (v.includes('$')) hits.push({r, c, v});
            }
        }
        console.log(`=== ${s.name} (rows ${totalRows}, cols ${totalCols}) : $ 命中 ${hits.length} 處 ===`);
        hits.forEach(h => console.log(`  r${String(h.r).padStart(3)} c${h.c} = ${JSON.stringify(h.v)}`));
    }
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
