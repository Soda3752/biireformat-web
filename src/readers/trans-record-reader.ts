import ExcelJS from 'exceljs';
import Papa from 'papaparse';

/**
 * 對應桌面版 TransRecordReader.kt
 * 讀取銀行對帳單，支援兩種格式：
 *   - .csv（桌面版預設 Big5 編碼，亦相容 UTF-8 / UTF-8 BOM）
 *   - .xlsx（網路銀行常見匯出格式，內含 formula cell）
 *
 * 兩種格式都輸出統一的 string[][]（與桌面版 List<List<String>> 等價），
 * 第一列為標題（日期, 摘要, 幣別, 支出金額, 存入金額, 餘額, 備註, 轉出入帳號, 註記）。
 * 後續 bank-match-service / bank-name-writer 對 CSV 與 XLSX 來源一視同仁。
 */
export const parseTransRecord = async (file: File): Promise<string[][]> => {
  const lower = file.name.toLowerCase();
    const rows = (lower.endsWith('.xlsx') || lower.endsWith('.xls'))
        ? await parseTransRecordXlsx(file)
        : await parseTransRecordCsv(file);

    if (isPostOfficeFormat(rows)) {
        return normalizePostOffice(rows);
    }
    return rows;
};

/**
 * 郵局 CSV 兼容層
 *
 * 郵局存摺明細 CSV 的格式特徵：
 *   - 前 6 列為 metadata（含「6個月內交易明細」與存摺帳號）
 *   - 標題列：交易日期 / 沖銷記號 / 摘要 / 提款金額 / 存款金額 / 備註（共 6 欄 + 3 空欄）
 *   - 每個欄位值前綴 \t、日期為民國年含時間（115/05/12 05:02:45）、金額帶 .00
 *
 * 與標準 9 欄銀行 CSV（日期, 摘要, 幣別, 支出金額, 存入金額, 餘額, 備註, 轉出入帳號, 註記）
 * 欄位數與順序皆不同，下游 bank-match-service 取「最後一個非空欄」當配對目標、
 * 取 raw[4] 當存入金額，因此這裡把郵局 6 欄重新映射到標準 9 欄：
 *
 *   郵局「備註」（含對方帳號末五碼或對方銀行/姓名）→ 標準「註記」（最後一欄、配對目標）
 *   郵局「提款金額」→ 標準「支出金額」(index 3)
 *   郵局「存款金額」→ 標準「存入金額」(index 4)
 *   日期轉西元只留日期、金額去 .00、欄前 \t 一律剝除
 *
 * 「續上一筆」列保留為獨立列，因為它常含帳號末五碼或對方銀行名，可命中配對。
 */

const POST_OFFICE_TITLE = '6個月內交易明細';
const POST_OFFICE_HEADER = ['交易日期', '沖銷記號', '摘要', '提款金額', '存款金額', '備註'];
const STANDARD_HEADER = ['日期', '摘要', '幣別', '支出金額', '存入金額', '餘額', '備註', '轉出入帳號', '註記'];

const stripLeadingTab = (s: string): string => s.replace(/^\t+/, '');

const cleanCell = (s: string | undefined): string => stripLeadingTab(s ?? '').trim();

const findPostOfficeHeaderIndex = (rows: ReadonlyArray<ReadonlyArray<string>>): number => {
    const scanLimit = Math.min(rows.length, 20);
    for (let i = 0; i < scanLimit; i++) {
        const row = rows[i];
        if (POST_OFFICE_HEADER.every((token, j) => cleanCell(row[j]) === token)) {
            return i;
        }
    }
    return -1;
};

const isPostOfficeFormat = (rows: ReadonlyArray<ReadonlyArray<string>>): boolean => {
    if (rows.length === 0) return false;
    if (findPostOfficeHeaderIndex(rows) >= 0) return true;
    // 後備偵測：第一列含「6個月內交易明細」
    const firstRowTokens = rows[0].map((c) => cleanCell(c));
    return firstRowTokens.some((c) => c.includes(POST_OFFICE_TITLE));
};

/** 民國年 `115/05/12 05:02:45` → 西元 `2026/5/12`；非預期格式則回原字串。 */
const convertRocDate = (raw: string): string => {
    const m = raw.match(/^(\d{2,3})\/(\d{1,2})\/(\d{1,2})/);
    if (!m) return raw;
    const year = parseInt(m[1], 10) + 1911;
    const month = parseInt(m[2], 10);
    const day = parseInt(m[3], 10);
    return `${year}/${month}/${day}`;
};

/** `1419.00` → `1419`；空白或非數字維持原樣（空字串）。 */
const cleanAmount = (raw: string): string => {
    if (raw === '') return '';
    const n = Number(raw);
    if (!Number.isFinite(n)) return raw;
    return Number.isInteger(n) ? String(n) : String(n);
};

/**
 * 對帳只關心入帳款項，所以郵局明細的處理規則：
 *   1. 主行：存入金額為「非空、非 0」的列，輸出為標準 9 欄
 *   2. 補充行：「續上一筆 / 續上一行」（金額通常 0.00），其備註合進前一筆主行的「備註」欄
 *      並 append 到主行尾端作為配對候選
 *   3. 其他列（提款、手續費、所有 deposit 空白或 0 的非補充列）一律忽略
 */
const SUPPLEMENT_SUMMARIES = new Set(['續上一筆', '續上一行']);

const isMeaningfulDeposit = (amount: string): boolean => {
    if (amount === '') return false;
    const n = Number(amount);
    return Number.isFinite(n) && n !== 0;
};

const normalizePostOffice = (rows: ReadonlyArray<ReadonlyArray<string>>): string[][] => {
    const headerIdx = findPostOfficeHeaderIndex(rows);
    if (headerIdx < 0) return rows.map((r) => [...r]);

    const out: string[][] = [STANDARD_HEADER.slice()];
    let current: string[] | null = null;
    let supplements: string[] = [];

    const flush = () => {
        if (current) {
            if (supplements.length > 0) {
                current[6] = supplements.join('、');
                for (const s of supplements) current.push(s);
            }
            out.push(current);
        }
        current = null;
        supplements = [];
    };

    for (let i = headerIdx + 1; i < rows.length; i++) {
        const cells = rows[i].map((c) => cleanCell(c));
        if (cells.every((c) => c === '')) continue;

        const summary = cells[2] ?? '';
        const note = cells[5] ?? '';

        if (SUPPLEMENT_SUMMARIES.has(summary)) {
            if (current !== null && note !== '') supplements.push(note);
            continue;
        }

        const deposit = cleanAmount(cells[4] ?? '');
        if (!isMeaningfulDeposit(deposit)) {
            // 非入帳列（提款、手續費、存入為 0 但非補充行的雜列）：丟棄
            flush();
            continue;
        }

        const date = convertRocDate(cells[0] ?? '');
        const withdraw = cleanAmount(cells[3] ?? '');

        flush();
        current = [date, summary, 'TWD', withdraw, deposit, '', '', '-', note];
    }
    flush();
    return out;
};

const parseTransRecordCsv = async (file: File): Promise<string[][]> => {
  const buffer = await file.arrayBuffer();
  const text = decodeCsv(buffer);

  const result = Papa.parse<string[]>(text, {
    header: false,
    skipEmptyLines: 'greedy',
      transform: (value) => {
          const s = typeof value === 'string' ? value : String(value ?? '');
          return stripExcelTextWrap(s);
      },
  });

  if (result.errors.length > 0) {
    const firstFatal = result.errors.find((e) => e.type !== 'FieldMismatch');
    if (firstFatal) {
      throw new Error(`CSV 解析失敗：${firstFatal.message}`);
    }
  }

  return (result.data as string[][]).filter((row) => Array.isArray(row));
};

/**
 * 剝除 Excel 強制文字格式包裹：`="000109450**0954*"` → `000109450**0954*`
 *
 * 部分網路銀行（例：PCMS 開頭的對帳單）為了讓 Excel 不要把帳號開頭 0 吃掉，
 * 會把欄位輸出成 `="..."` 字面形式。PapaParse 不會自動拆這層，
 * 導致註記欄塞滿 `=`、`"` 等非數字字元，後續末五碼比對直接失敗。
 *
 * 只處理「完整以 `="` 開頭並以 `"` 結尾」的字串；中間含換行也保留（s 旗標）。
 */
const stripExcelTextWrap = (s: string): string => {
    if (s.length < 3 || !s.startsWith('="') || !s.endsWith('"')) return s;
    return s.slice(2, -1);
};

const parseTransRecordXlsx = async (file: File): Promise<string[][]> => {
  const buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const sheet = workbook.worksheets[0];
  if (!sheet) {
    throw new Error('xlsx 未含任何工作表');
  }

  const colCount = Math.max(sheet.columnCount, sheet.actualColumnCount ?? 0);
  if (colCount === 0) return [];

  const out: string[][] = [];
  sheet.eachRow({includeEmpty: false}, (row) => {
    const cells: string[] = [];
    for (let c = 1; c <= colCount; c++) {
      cells.push(cellValueToString(row.getCell(c).value));
    }
    if (cells.every((s) => s.trim() === '')) return;
    out.push(cells);
  });
  return out;
};

/**
 * 偵測編碼：UTF-8 BOM 直接走 UTF-8；
 * 否則先嘗試 Big5，若 replacement char 比例過高就退回 UTF-8。
 */
const decodeCsv = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer);
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }

  try {
    const big5 = new TextDecoder('big5', { fatal: false }).decode(bytes);
    if (replacementRatio(big5) < 0.02) return big5;
  } catch {
    // ignore，走 utf-8
  }

  return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
};

const replacementRatio = (s: string): number => {
  if (s.length === 0) return 0;
  let count = 0;
  for (const ch of s) if (ch === '�') count++;
  return count / s.length;
};

/**
 * 將 ExcelJS cell value 轉字串。
 * - null / undefined → ''
 * - Date → yyyy/M/d（與 CSV `2026/4/20` 寫法對齊，不補零）
 * - Formula cell → 取 result 後遞迴
 * - RichText → 串接 text
 */
const cellValueToString = (value: ExcelJS.CellValue): string => {
  if (value === null || value === undefined) return '';
  if (typeof value === 'string') return value;
  if (typeof value === 'number') return String(value);
  if (typeof value === 'boolean') return String(value);
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = value.getMonth() + 1;
    const d = value.getDate();
    return `${y}/${m}/${d}`;
  }
  if (typeof value === 'object') {
    const obj = value as unknown as Record<string, unknown>;
    if (Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>).map((rt) => rt.text ?? '').join('');
    }
    if ('result' in obj) {
      return cellValueToString(obj.result as ExcelJS.CellValue);
    }
    if ('text' in obj) {
      return String(obj.text ?? '');
    }
    if ('error' in obj) {
      return '';
    }
  }
  return String(value);
};
