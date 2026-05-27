/**
 * 「生成手填本」資料模型。
 *
 * 對應範本檔（`1.彰化.xls`）的結構：
 *  - 一份手填本 = 一個線別 + 一個年月 + 多個客戶
 *  - 每個客戶有：代號、名稱、品名清單（含單價）、休息日備註、電話
 *  - 客戶數無上限、品名數無上限（範本固定為 60 / 8，UI 與輸出皆允許超出）
 *
 * 編輯狀態以 `HandfillBook` 物件序列化存入 localStorage。
 */

export interface HandfillProduct {
    name: string;
    unitPrice?: number;
}

export interface HandfillCustomer {
    id: string;
    customerId: string;
    customerName: string;
    products: HandfillProduct[];
    restNotes: string[];
    phones: string[];
    /**
     * 歷史欄位：早期版本用來標記「品名清單已手動排序、停用自動重排」。
     * 自動排序已全面移除（排序改為使用者按「排序」按鈕主動觸發），此旗標
     * 不再驅動任何 UI 行為，僅為相容舊資料與 round-trip 保存而保留。
     * 既有 localStorage 資料可能不存在此欄位，讀取時視為 false。
     */
    manualSort?: boolean;
}

export interface HandfillBook {
    id: string;
    lineNo: number;
    lineName: string;
    year: number;
    month: number;
    customers: HandfillCustomer[];
    createdAt: number;
    updatedAt: number;
}

/**
 * 隱藏 metadata 分頁名稱。
 * Writer 寫出 .xlsx 時會額外加一個 veryHidden 分頁，A1 存整本 HandfillBook 的 JSON，
 * Reader 讀回時若該分頁存在則直接還原，繞過版面分析的不確定性。
 */
export const HANDFILL_MANIFEST_SHEET = '_handfill_meta';

const LINE_CN = ['', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十'];

/** 將線別編號轉為中文括號標示，例 1 → "(一)"。超過 10 顯示原數字。 */
export function lineLabel(no: number): string {
    if (!Number.isFinite(no) || no <= 0) return '';
    if (no >= 1 && no <= 10) return `(${LINE_CN[no]})`;
    return `(${no})`;
}

/** 線別編號 + 名稱組合，例：(一)彰化 */
export function lineFullName(book: Pick<HandfillBook, 'lineNo' | 'lineName'>): string {
    return `${lineLabel(book.lineNo)}${book.lineName ?? ''}`;
}

let _uuidCounter = 0;

/** 簡易 UUID（時間戳 + 累加計數 + random），不需強密碼學等級。 */
export function genId(prefix = ''): string {
    _uuidCounter = (_uuidCounter + 1) % 1_000_000;
    const ts = Date.now().toString(36);
    const rand = Math.floor(Math.random() * 1e8).toString(36);
    const cnt = _uuidCounter.toString(36);
    return `${prefix}${ts}-${cnt}-${rand}`;
}

/** 建立預設客戶（電話預設 2 個空格） */
export function createEmptyCustomer(): HandfillCustomer {
    return {
        id: genId('cust-'),
        customerId: '',
        customerName: '',
        products: [{name: '', unitPrice: undefined}],
        restNotes: [],
        phones: ['', ''],
        manualSort: false,
    };
}

/** 建立預設空白手填本 */
export function createEmptyBook(init?: Partial<Pick<HandfillBook, 'lineNo' | 'lineName' | 'year' | 'month'>>): HandfillBook {
    const now = Date.now();
    const today = new Date(now);
    const rocYear = today.getFullYear() - 1911;
    return {
        id: genId('book-'),
        lineNo: init?.lineNo ?? 1,
        lineName: init?.lineName ?? '',
        year: init?.year ?? rocYear,
        month: init?.month ?? (today.getMonth() + 1),
        customers: [],
        createdAt: now,
        updatedAt: now,
    };
}

/**
 * 就地排序客戶的品名清單（由「排序」按鈕觸發，非自動）。
 * 順序規則：
 *   1. 空白品名 → 置頂
 *   2. 在「帳單排序品項（cargo_sort）」表內的品名 → 依表索引排序
 *   3. 不在表內的非空白品名 → 殿後
 * 同一群組內保留原相對順序（穩定排序）。
 */
export function sortProductsBlanksFirst(
    products: HandfillProduct[],
    cargoNames: ReadonlyArray<string>,
): void {
    const indexMap = new Map<string, number>();
    cargoNames.forEach((name, i) => indexMap.set(name, i));
    const keyed = products.map((p, i) => {
        let key: number;
        if (p.name.trim().length === 0) {
            key = Number.MIN_SAFE_INTEGER; // 空白置頂
        } else {
            key = indexMap.get(p.name) ?? Number.MAX_SAFE_INTEGER; // 表內依序，未知殿後
        }
        return {p, i, key};
    });
    keyed.sort((a, b) => (a.key - b.key) || (a.i - b.i));
    for (let i = 0; i < keyed.length; i++) {
        products[i] = keyed[i].p;
    }
}

/** 客戶是否「實質為空」（用於儲存/匯出時剔除空殼） */
export function isCustomerEmpty(c: HandfillCustomer): boolean {
    if (c.customerId.trim() || c.customerName.trim()) return false;
    if (c.products.some((p) => p.name.trim() || p.unitPrice !== undefined)) return false;
    if (c.restNotes.some((r) => r.trim())) return false;
    if (c.phones.some((p) => p.trim())) return false;
    return true;
}
