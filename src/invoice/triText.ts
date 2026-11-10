// 移植自 invoice-helper/src/triText.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 網頁版三聯「大框框」連續輸入的解析：把一段文字拆成一張一張的金額。
// 不引用 exceljs，畫面（src/tabs/invoice.ts）與測試都可直接用。
//
// 規則（畫面提示與 CLAUDE.md 同步）：
// - 分隔：空白、換行、Tab、全形空白、頓號「、」、分號「;；」，以及「不是千分位」的逗號。
// - 千分位：數字內的逗號若是標準千分位（開頭 1～3 位數字，之後每個逗號後面剛好三位數字，
//   例如 1,290、12,345,678）就當同一個數字；其他逗號（例如 1290,300、12,34）當分隔。
//   要把「500,300」當成 500 與 300 兩張，請用空白分開。
// - 全形數字、全形逗號、全形句點先轉半形；數字前可有 $ / NT$、後面可有「元」（Excel 貨幣格式複製過來）。
// - 單獨的「-」（Excel 會計格式的 0）與 0 不算一張。可有負號（例如折讓 -1,290）與小數。
// - 其他看不懂的內容（例如打錯成字母 12a4）列為錯誤，不納入計算。

export type TriToken = {
  /** 原文（轉半形前）中的這一段 */
  text: string;
  /** 在原文中的起訖位置 [start, end) */
  start: number;
  end: number;
  /** 解析出的金額；null＝無法辨識 */
  value: number | null;
};

export type TriParse = {
  /** 拆出的金額（不含 0 與單獨的「-」） */
  amounts: number[];
  /** 合計 */
  sum: number;
  /** 無法辨識的段落（依出現順序） */
  errors: TriToken[];
  /** 所有段落（含 0、錯誤），給畫面標示用 */
  tokens: TriToken[];
};

/** 全形 → 半形（數字、逗號、句點、負號、錢字號）；長度不變，位置可對回原文 */
export function toHalfWidth(s: string): string {
  return s
    .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .replace(/，/g, ",")
    .replace(/．/g, ".")
    .replace(/[－−]/g, "-")
    .replace(/＄/g, "$");
}

const SEP = /[\s　、;；]/; // \s 已含 Tab、換行、不換行空白
const NUM_GROUPED = /^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?(?!\d)/;
const NUM_PLAIN = /^-?\d+(?:\.\d+)?/;

export function parseTriText(raw: string): TriParse {
  const s = toHalfWidth(raw);
  const tokens: TriToken[] = [];
  let i = 0;
  while (i < s.length) {
    if (SEP.test(s[i]!) || s[i] === ",") {
      i++;
      continue;
    }
    // 一個「塊」：直到下一個空白類分隔
    let j = i;
    while (j < s.length && !SEP.test(s[j]!)) j++;
    splitChunk(s, raw, i, j, tokens);
    i = j;
  }
  const amounts: number[] = [];
  for (const t of tokens) if (t.value !== null && t.value !== 0) amounts.push(t.value);
  const sum = Math.round(amounts.reduce((a, b) => a + b, 0) * 100) / 100;
  return { amounts, sum, errors: tokens.filter((t) => t.value === null), tokens };
}

/** 把 [a, b) 這一塊（沒有空白）依千分位規則拆成數字；看不懂的部分整段當錯誤 */
function splitChunk(s: string, raw: string, a: number, b: number, out: TriToken[]) {
  const chunk = s.slice(a, b);
  // 單獨的「-」＝會計格式的 0
  if (/^-+$/.test(chunk) || /^\$-$/.test(chunk)) {
    out.push({ text: raw.slice(a, b), start: a, end: b, value: 0 });
    return;
  }
  const found: TriToken[] = [];
  let p = 0;
  while (p < chunk.length) {
    if (chunk[p] === ",") {
      p++;
      continue;
    }
    const rest = chunk.slice(p);
    const prefix = /^(?:NT\$|NT|\$)/i.exec(rest)?.[0] ?? "";
    const body = rest.slice(prefix.length);
    const m = NUM_GROUPED.exec(body) ?? NUM_PLAIN.exec(body);
    if (!m) break;
    let len = prefix.length + m[0].length;
    if (chunk[p + len] === "元") len++;
    const next = chunk[p + len];
    if (next !== undefined && next !== ",") break; // 數字後面接了怪字 → 整塊錯
    found.push({ text: raw.slice(a + p, a + p + len), start: a + p, end: a + p + len, value: Number(m[0].replace(/,/g, "")) });
    p += len;
  }
  if (p < chunk.length) {
    // 這一塊有看不懂的字：整塊（到逗號為止的那一小段）標成錯誤，其餘照常
    // 為了讓使用者好找，錯誤段＝從出錯位置往前後延伸到逗號或塊的邊界
    let st = p;
    while (st > 0 && chunk[st - 1] !== ",") st--;
    let en = p;
    while (en < chunk.length && chunk[en] !== ",") en++;
    const keep = found.filter((t) => t.end <= a + st);
    out.push(...keep);
    out.push({ text: raw.slice(a + st, a + en), start: a + st, end: a + en, value: null });
    if (en < chunk.length) splitChunk(s, raw, a + en + 1, b, out);
    return;
  }
  out.push(...found);
}
