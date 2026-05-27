import ExcelJS from 'exceljs';
import {textOfExcelJs} from '@/infra/handfill-manifest';

const path = process.argv[2];

async function main() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(path);
    const s = wb.getWorksheet('海線上') ?? wb.worksheets[0];
    const totalRows = s.actualRowCount || s.rowCount;
    let codes = 0, dollars = 0;
    for (let r = 1; r <= totalRows; r++) {
        const a = textOfExcelJs(s.getRow(r).getCell(1).value).trim();
        const b = textOfExcelJs(s.getRow(r).getCell(2).value).trim();
        const d = textOfExcelJs(s.getRow(r).getCell(19).value).trim();
        const isCode = /^\d{3,5}$/.test(a);
        if (isCode) codes++;
        if (d.includes('$')) dollars++;
        if (a || b || d) {
            const tag = isCode ? ' <ID>' : '';
            const dm = d.includes('$') ? '  <<<$$' : (d ? `  c19=${JSON.stringify(d)}` : '');
            console.log(`r${String(r).padStart(3)} A=${JSON.stringify(a).padEnd(16)} B=${JSON.stringify(b).padEnd(14)}${tag}${dm}`);
        }
    }
    console.log(`\n代號列=${codes}  $$=${dollars}  總列=${totalRows}`);
}

main().catch(e => {
    console.error(e);
    process.exit(1);
});
