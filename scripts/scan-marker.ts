import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    for (const s of wb.worksheets) {
        const totalRows = s.actualRowCount || s.rowCount;
        const totalCols = s.actualColumnCount || s.columnCount;
        let dollar = 0, at = 0;
        const cols = new Set<number>();
        for (let r = 1; r <= totalRows; r++) for (let c = 1; c <= totalCols; c++) {
            const v = textOfExcelJs(s.getRow(r).getCell(c).value).trim();
            if (v.includes('$')) {
                dollar++;
                cols.add(c);
            }
            if (v.includes('@')) {
                at++;
                cols.add(c);
            }
        }
        console.log(`${s.name}: $$=${dollar}  @@=${at}  (出現欄:${[...cols].join(',') || '無'})`);
    }
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
