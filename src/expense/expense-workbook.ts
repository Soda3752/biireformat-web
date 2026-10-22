/**
 * ExcelJS Workbook → SheetGrid 轉接層，讓 expense-parser 保持純函式。
 */

import ExcelJS from 'exceljs';
import type {GridCell, SheetGrid} from './expense-parser';

type FormulaValue = { formula?: string; sharedFormula?: string; result?: unknown };

function scalar(v: unknown): string | number | null {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number' || typeof v === 'string') return v;
    if (typeof v === 'boolean') return String(v);
    if (v instanceof Date) return v.toISOString();
    if (typeof v === 'object') {
        const o = v as { richText?: { text: string }[]; text?: string; error?: string };
        if (Array.isArray(o.richText)) return o.richText.map((t) => t.text).join('');
        if (typeof o.text === 'string') return o.text;
        if (o.error) return null;
    }
    return null;
}

function toGrid(ws: ExcelJS.Worksheet): SheetGrid {
    const resolveFormula = (fv: FormulaValue): string | undefined => {
        if (typeof fv.formula === 'string') return fv.formula;
        // 共用公式：記錄的是主儲存格位址，取主儲存格的公式判斷型態（SUM 與否）即可
        if (typeof fv.sharedFormula === 'string') {
            const master = ws.getCell(fv.sharedFormula).value as FormulaValue | null;
            if (master && typeof master === 'object' && typeof master.formula === 'string') return master.formula;
        }
        return undefined;
    };

    return {
        name: ws.name,
        rowCount: ws.actualRowCount > 0 ? ws.rowCount : 0,
        colCount: ws.columnCount,
        get(row: number, col: number): GridCell | null {
            const v = ws.getRow(row).getCell(col).value;
            if (v === null || v === undefined) return null;
            if (typeof v === 'object' && !(v instanceof Date) && ('formula' in v || 'sharedFormula' in v)) {
                const fv = v as FormulaValue;
                return {value: scalar(fv.result), formula: resolveFormula(fv)};
            }
            return {value: scalar(v)};
        },
    };
}

export async function readExpenseGrids(data: ArrayBuffer): Promise<SheetGrid[]> {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);
    return wb.worksheets.filter((ws) => ws.state !== 'veryHidden').map(toGrid);
}
