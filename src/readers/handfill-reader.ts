/**
 * 「生成手填本」.xlsx / .xls 匯入解析器。
 *
 * 解析流程：
 *   1. 依副檔名選擇後端：
 *        - .xlsx → ExcelJS
 *        - .xls  → SheetJS (xlsx) 社群版（dynamic import，僅在需要時載入）
 *   2. 將「上」分頁（或第一個分頁）讀為 2D 字串矩陣
 *   3. 套用共用解析邏輯抽出 HandfillBook
 *
 * 預期內容對應範本 `1.彰化.xls`：
 *   - 標題列：(線別)名稱 / 民國年 / 月份
 *   - 欄位標頭列含「客戶名稱」「品名」字串
 *   - 客戶區塊：col 0 = ID / 名稱 / 休息日 / 電話；col 1 = 品名；col 2 = 單價
 */

import ExcelJS from 'exceljs';

import {createEmptyBook, genId, type HandfillBook, type HandfillCustomer,} from '@/domain/models/handfill-book';

type Cell = string;
type Matrix = Cell[][];   // matrix[rowIdx][colIdx]，0-indexed

/* ====================== 對外 API ======================= */

export async function readHandfillBook(file: File): Promise<HandfillBook> {
    const fname = file.name.toLowerCase();
    let matrix: Matrix;
    if (fname.endsWith('.xls')) {
        matrix = await loadMatrixFromXls(file);
    } else {
        // 預設走 ExcelJS（.xlsx / 未知副檔名）
        matrix = await loadMatrixFromXlsx(file);
    }
    return parseMatrix(matrix);
}

/* ====================== ExcelJS (.xlsx) 適配 ======================= */

async function loadMatrixFromXlsx(file: File): Promise<Matrix> {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const sheet = wb.getWorksheet('上') ?? wb.worksheets[0];
    if (!sheet) throw new Error('找不到資料分頁');

    const matrix: Matrix = [];
    const totalRows = sheet.actualRowCount || sheet.rowCount;
    const totalCols = sheet.actualColumnCount || sheet.columnCount;

    for (let r = 1; r <= totalRows; r++) {
        const row: Cell[] = [];
        for (let c = 1; c <= totalCols; c++) {
            row.push(textOfExcelJs(sheet.getRow(r).getCell(c).value));
        }
        matrix.push(row);
    }
    return matrix;
}

/* ====================== SheetJS (.xls) 適配 ======================= */

async function loadMatrixFromXls(file: File): Promise<Matrix> {
    // dynamic import：xlsx 套件約 750KB，只在需要時才載入
    const XLSX = await import('xlsx');
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, {type: 'array'});

    const sheetName = wb.SheetNames.includes('上') ? '上' : wb.SheetNames[0];
    if (!sheetName) throw new Error('找不到資料分頁');
    const sheet = wb.Sheets[sheetName];

    // sheet_to_json header:1 模式輸出 2D 陣列；defval:'' 確保空格也有 placeholder
    const raw = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
        header: 1,
        defval: '',
        raw: false, // 取格式化後的字串
    });

    const matrix: Matrix = raw.map((row) => row.map((cell) => textOfPrimitive(cell)));
    return matrix;
}

/* ====================== 共用解析邏輯 ======================= */

function parseMatrix(matrix: Matrix): HandfillBook {
    const book = createEmptyBook();
    book.id = genId('book-');

    if (matrix.length === 0) return book;

    // 1. 標題列（matrix[0]）
    parseTitleRow(matrix[0] ?? [], book);

    // 2. 找出所有「客戶名稱 / 品名」欄頭列
    const headerRows = new Set<number>();
    for (let r = 0; r < matrix.length; r++) {
        const c0 = (matrix[r]?.[0] ?? '').trim();
        const c1 = (matrix[r]?.[1] ?? '').trim();
        if (c0 === '客戶名稱' && c1 === '品名') {
            headerRows.add(r);
        }
    }

    // 3. 找出資料起始列：第一個欄頭的下兩列
    let dataStartRow = 4; // 範本預設
    if (headerRows.size > 0) {
        const firstHeader = Math.min(...headerRows);
        dataStartRow = firstHeader + 2;
    }

    // 4. 切塊解析
    book.customers = parseCustomers(matrix, headerRows, dataStartRow);

    return book;
}

function parseTitleRow(row: Cell[], book: HandfillBook): void {
    const titleText = (row[0] ?? '').trim();
    // 期望格式：(一)彰化 或 (1)彰化
    const m = titleText.match(/^[(（]\s*([一二三四五六七八九十\d]+)\s*[)）]\s*(.+)$/);
    if (m) {
        const linePart = m[1];
        const cnMap: Record<string, number> = {
            一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10,
        };
        if (cnMap[linePart]) {
            book.lineNo = cnMap[linePart];
        } else {
            const num = parseInt(linePart, 10);
            if (Number.isFinite(num)) book.lineNo = num;
        }
        book.lineName = (m[2] || '').trim();
    } else if (titleText) {
        book.lineName = titleText;
    }

    // 年：col 14~20 找民國年
    for (let c = 14; c <= 20 && c < row.length; c++) {
        const num = parseInt((row[c] ?? '').trim(), 10);
        if (Number.isFinite(num) && num > 100 && num < 200) {
            book.year = num;
            break;
        }
    }

    // 月：col 14~21 找 1~12 數字
    for (let c = 14; c <= 21 && c < row.length; c++) {
        const num = parseInt((row[c] ?? '').trim(), 10);
        if (Number.isFinite(num) && num >= 1 && num <= 12) {
            book.month = num;
        }
    }
}

function parseCustomers(matrix: Matrix, headerRows: Set<number>, dataStartRow: number): HandfillCustomer[] {
    const customers: HandfillCustomer[] = [];
    const lastDataRow = matrix.length - 1;

    // 客戶區塊起始：col 0 是 3~5 位數字
    const starts: number[] = [];
    for (let r = dataStartRow; r <= lastDataRow; r++) {
        if (headerRows.has(r) || headerRows.has(r - 1)) continue;
        const c0 = (matrix[r]?.[0] ?? '').trim();
        if (isCustomerIdLike(c0)) {
            starts.push(r);
        }
    }

    for (let i = 0; i < starts.length; i++) {
        const startRow = starts[i];
        const nextStart = i + 1 < starts.length ? starts[i + 1] : lastDataRow + 1;
        // block 結束於下一客戶起始前；但若中間有欄頭列則於該列截斷
        let endRowExclusive = nextStart;
        for (let r = startRow + 1; r < nextStart; r++) {
            if (headerRows.has(r)) {
                endRowExclusive = r;
                break;
            }
        }
        customers.push(parseCustomerBlock(matrix, startRow, endRowExclusive));
    }

    return customers;
}

function isCustomerIdLike(s: string): boolean {
    if (!s) return false;
    return /^\d{3,5}$/.test(s.trim());
}

function parseCustomerBlock(matrix: Matrix, startRow: number, endRowExclusive: number): HandfillCustomer {
    const cust: HandfillCustomer = {
        id: genId('cust-'),
        customerId: '',
        customerName: '',
        products: [],
        restNotes: [],
        phones: [],
    };

    const rows = endRowExclusive - startRow;
    if (rows <= 0) return cust;

    // 第一列：代號
    cust.customerId = (matrix[startRow]?.[0] ?? '').trim();
    // 第二列：客戶名稱
    if (rows >= 2) {
        cust.customerName = (matrix[startRow + 1]?.[0] ?? '').trim();
    }

    // 收集所有品名 + 單價
    for (let r = startRow; r < endRowExclusive; r++) {
        const name = (matrix[r]?.[1] ?? '').trim();
        const priceStr = (matrix[r]?.[2] ?? '').trim();
        const price = priceStr ? parseFloat(priceStr) : NaN;
        if (name || Number.isFinite(price)) {
            cust.products.push({
                name,
                unitPrice: Number.isFinite(price) ? price : undefined,
            });
        }
    }
    if (cust.products.length === 0) {
        cust.products.push({name: '', unitPrice: undefined});
    }

    // col 0 的其他列 → 休息日 + 電話
    // 標準佈局：col0[0]=ID, [1]=名稱, [2]=空, [3~5]=休息日, [6]=電話1, [7]=電話2
    const col0Lines: Array<{ idx: number; text: string }> = [];
    for (let r = startRow; r < endRowExclusive; r++) {
        col0Lines.push({
            idx: r - startRow,
            text: (matrix[r]?.[0] ?? '').trim(),
        });
    }

    if (col0Lines.length >= 8) {
        const lastIdx = col0Lines.length - 1;
        const secondLastIdx = col0Lines.length - 2;
        for (let i = 2; i < col0Lines.length; i++) {
            const {text} = col0Lines[i];
            if (!text) continue;
            if (i === lastIdx || i === secondLastIdx) {
                cust.phones.push(text);
            } else {
                cust.restNotes.push(text);
            }
        }
    } else {
        for (let i = 2; i < col0Lines.length; i++) {
            const {text} = col0Lines[i];
            if (text) cust.restNotes.push(text);
        }
    }

    // 補齊電話到 2 個（保留空槽以符合 UI 預設）
    while (cust.phones.length < 2) cust.phones.push('');

    return cust;
}

/* ====================== Cell value 規格化 ======================= */

function textOfExcelJs(value: ExcelJS.CellValue): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    if (value instanceof Date) return value.toISOString();
    const obj = value as { richText?: Array<{ text: string }>; text?: string; result?: unknown };
    if (obj.richText) return obj.richText.map((t) => t.text).join('');
    if (obj.text) return obj.text;
    if (obj.result !== undefined) return textOfExcelJs(obj.result as ExcelJS.CellValue);
    return '';
}

function textOfPrimitive(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    if (value instanceof Date) return value.toISOString();
    return String(value);
}
