/**
 * 「生成手填本」.xlsx / .xls 匯入解析器。
 *
 * 解析流程：
 *   1. .xlsx 路徑：先嘗試讀 `_handfill_meta` 隱藏分頁（web app 寫出的檔案會帶），
 *      該分頁的 A1 是整本 HandfillBook 的 JSON，可直接還原。
 *   2. 找不到 manifest 時 fallback 到版面分析：
 *        - .xlsx → ExcelJS
 *        - .xls  → SheetJS (xlsx) 社群版（dynamic import，僅在需要時載入）
 *   3. 將「上」分頁（或第一個分頁）讀為 2D 字串矩陣，套用共用邏輯抽出 HandfillBook
 *
 * 版面分析預期對應範本 `1.彰化.xls`：
 *   - 標題列：(線別)名稱 / 民國年 / 月份
 *   - 欄位標頭列含「客戶名稱」「品名」字串
 *   - 客戶區塊：col 0 = ID / 名稱 / 休息日 / 電話；col 1 = 品名；col 2 = 單價
 */

import ExcelJS from 'exceljs';

import {
    createEmptyBook,
    genId,
    type HandfillBook,
    type HandfillCustomer,
    type HandfillProduct,
} from '@/domain/models/handfill-book';
import {buildMatrixFromWorkbook, type Cell, hashMatrix, type Matrix, readManifestRaw,} from '@/infra/handfill-manifest';

/* ====================== 對外 API ======================= */

export async function readHandfillBook(file: File): Promise<HandfillBook> {
    const fname = file.name.toLowerCase();

    if (fname.endsWith('.xls')) {
        // 舊範本檔（.xls）一律走版面分析
        return parseMatrix(await loadMatrixFromXls(file));
    }

    // .xlsx：讀 manifest 並以 layoutHash 判斷使用者是否在 Excel 改過版面
    const wb = await loadXlsxWorkbook(file);
    const manifest = readManifestRaw(wb);
    const matrix = buildMatrixFromWorkbook(wb);

    if (manifest) {
        // 舊檔（af24e5d 寫出、無 layoutHash）：維持「manifest 優先」舊行為，不破壞既有檔案
        if (!manifest.layoutHash) {
            return normalizeBook(manifest.book);
        }
        // 新檔：hash 相符代表版面未被手改 → 直接用 JSON 還原
        if (hashMatrix(matrix) === manifest.layoutHash) {
            return normalizeBook(manifest.book);
        }
        // hash 不符代表使用者在 Excel 改過版面 → 以版面分析為準，再補回版面表達不了的內部欄位
        const layoutBook = parseMatrix(matrix);
        mergeInternalFields(layoutBook, manifest.book);
        return layoutBook;
    }

    // 無 manifest（外部範本 / 非 web app 產出）→ 純版面分析
    return parseMatrix(matrix);
}

async function loadXlsxWorkbook(file: File): Promise<ExcelJS.Workbook> {
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    return wb;
}

/* ====================== Manifest 內部欄位補回 ======================= */

/**
 * 使用者改過版面、改走版面分析後，從 manifest book 補回「版面表達不了」的內部欄位。
 * 目前僅 manualSort（手動排序旗標）。
 * 對應 key：優先用 customerId（非空且在 source 中唯一），否則 fallback 用索引位置。
 * id / createdAt / updatedAt 屬內部識別，允許重新生成，不補回。
 */
function mergeInternalFields(target: HandfillBook, source: HandfillBook): void {
    const byId = new Map<string, HandfillCustomer>();
    const dupIds = new Set<string>();
    for (const c of source.customers) {
        const key = c.customerId.trim();
        if (!key) continue;
        if (byId.has(key)) dupIds.add(key);
        else byId.set(key, c);
    }

    target.customers.forEach((tc, i) => {
        const key = tc.customerId.trim();
        const src = (key && byId.has(key) && !dupIds.has(key))
            ? byId.get(key)
            : source.customers[i];
        if (src) tc.manualSort = src.manualSort ?? false;
    });
}

/** 補齊欄位、保證型別正確，避免 manifest 缺欄位時下游 UI 出錯。 */
function normalizeBook(raw: Partial<HandfillBook>): HandfillBook {
    const base = createEmptyBook();
    const customers: HandfillCustomer[] = (raw.customers ?? []).map((c) => {
        const phones = Array.isArray(c.phones) ? [...c.phones] : [];
        while (phones.length < 2) phones.push('');
        const products: HandfillProduct[] = Array.isArray(c.products) && c.products.length > 0
            ? c.products.map((p) => ({name: p.name ?? '', unitPrice: p.unitPrice}))
            : [{name: '', unitPrice: undefined}];
        return {
            id: c.id ?? genId('cust-'),
            customerId: c.customerId ?? '',
            customerName: c.customerName ?? '',
            products,
            restNotes: Array.isArray(c.restNotes) ? [...c.restNotes] : [],
            phones,
            manualSort: c.manualSort ?? false,
        };
    });
    return {
        ...base,
        ...raw,
        id: raw.id ?? base.id,
        customers,
    };
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
    // 注意：writer 把欄頭兩列做了直向 merge，ExcelJS 讀回時兩列 col0 都是「客戶名稱」，
    // 因此 headerRows 會同時包含 r 與 r+1。dataStartRow 已指向第一個客戶 ID 列，
    // 不能再用 headerRows.has(r - 1) 當條件，否則每頁第一個客戶會被誤跳。
    const starts: number[] = [];
    for (let r = dataStartRow; r <= lastDataRow; r++) {
        if (headerRows.has(r)) continue;
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
        // 修剪尾端整列空白：writer 在頁尾的 page-break 前可能留下不屬於本客戶的空白列，
        // 若不修剪會讓 layoutCol0 邏輯把電話位置算錯（誤判成休息日）。
        while (endRowExclusive > startRow + 1 && isRowBlank(matrix[endRowExclusive - 1])) {
            endRowExclusive--;
        }
        customers.push(parseCustomerBlock(matrix, startRow, endRowExclusive));
    }

    return customers;
}

function isCustomerIdLike(s: string): boolean {
    if (!s) return false;
    return /^\d{3,5}$/.test(s.trim());
}

function isRowBlank(row: Cell[] | undefined): boolean {
    if (!row) return true;
    return !(row[0] ?? '').trim() && !(row[1] ?? '').trim() && !(row[2] ?? '').trim();
}

function looksLikePhone(s: string): boolean {
    const trimmed = s.trim();
    if (!trimmed) return false;
    const digits = trimmed.replace(/-/g, '');
    return /^\d{7,15}$/.test(digits);
}

function parseCustomerBlock(matrix: Matrix, startRow: number, endRowExclusive: number): HandfillCustomer {
    const cust: HandfillCustomer = {
        id: genId('cust-'),
        customerId: '',
        customerName: '',
        products: [],
        restNotes: [],
        phones: [],
        manualSort: false,
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

    // col 0 的其他列：以內容樣式判斷而非位置。
    // 原本「最後兩列＝電話」的位置邏輯只在 blockRows==8 時穩定；當客戶區塊被 trim 後
    // 不足 8 列、或電話位置因 writer 演算法變動而非最末列時，會把電話誤讀為休息日。
    // 改用 looksLikePhone：純數字（可含 dash）且 7~15 位數 → 電話，其餘 → 休息日。
    for (let r = startRow + 2; r < endRowExclusive; r++) {
        const text = (matrix[r]?.[0] ?? '').trim();
        if (!text) continue;
        if (looksLikePhone(text)) {
            cust.phones.push(text);
        } else {
            cust.restNotes.push(text);
        }
    }

    // 補齊電話到 2 個（保留空槽以符合 UI 預設）
    while (cust.phones.length < 2) cust.phones.push('');

    return cust;
}

/* ====================== Cell value 規格化 ======================= */

function textOfPrimitive(value: unknown): string {
    if (value === null || value === undefined) return '';
    if (typeof value === 'string') return value;
    if (typeof value === 'number') return String(value);
    if (typeof value === 'boolean') return value ? 'true' : 'false';
    if (value instanceof Date) return value.toISOString();
    return String(value);
}
