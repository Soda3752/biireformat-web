// 移植自 invoice-helper/src/cell.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 儲存格值的共用處理：把 exceljs 各種值型別（數字、字串、richText、公式、日期）統一成好用的形式。
import type ExcelJS from "exceljs";

export type CellInfo = {
  /** 文字內容（字串、richText 合併後），已 trim；非文字為 "" */
  text: string;
  /** 數值（字面數字、公式結果、可解析的數字字串）；沒有則 null */
  num: number | null;
  /** 公式文字（不含 =）；沒有則 null */
  formula: string | null;
  /** 日期值（字面日期或公式結果為日期） */
  date: Date | null;
  /** 是否完全空白 */
  empty: boolean;
};

const EMPTY: CellInfo = { text: "", num: null, formula: null, date: null, empty: true };

/** 去掉前後空白（含全形空白、不換行空白） */
export function cleanName(s: string): string {
  return s.replace(/^[\s　 ]+|[\s　 ]+$/g, "");
}

/** 表頭比對用：去掉所有空白 */
export function normHeader(s: string): string {
  return s.replace(/[\s　 ]+/g, "");
}

function parseNumericString(s: string): number | null {
  const t = s.replace(/[,\s　]/g, "");
  if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
  return Number(t);
}

export function readCell(cell: ExcelJS.Cell | undefined): CellInfo {
  if (!cell) return EMPTY;
  return readValue(cell.value as unknown, cell);
}

export function readValue(v: unknown, cell?: ExcelJS.Cell): CellInfo {
  if (v === null || v === undefined || v === "") return EMPTY;
  if (typeof v === "number") return { text: "", num: v, formula: null, date: null, empty: false };
  if (typeof v === "string") {
    const text = cleanName(v);
    if (text === "") return EMPTY;
    return { text, num: parseNumericString(text), formula: null, date: null, empty: false };
  }
  if (typeof v === "boolean") return { text: String(v), num: null, formula: null, date: null, empty: false };
  if (v instanceof Date) return { text: "", num: null, formula: null, date: v, empty: false };
  if (typeof v === "object") {
    const o = v as Record<string, unknown>;
    if (Array.isArray(o.richText)) {
      const text = cleanName((o.richText as { text?: string }[]).map((r) => r.text ?? "").join(""));
      if (text === "") return EMPTY;
      return { text, num: parseNumericString(text), formula: null, date: null, empty: false };
    }
    if ("formula" in o || "sharedFormula" in o) {
      let formula = (o.formula as string | undefined) ?? null;
      if (!formula && cell && "sharedFormula" in o) {
        // exceljs 會把共用公式轉譯到這一格
        try {
          formula = (cell as unknown as { formula?: string }).formula ?? null;
        } catch {
          formula = null;
        }
      }
      if (!formula) formula = (o.sharedFormula as string | undefined) ?? "";
      const r = readValue(o.result);
      return { ...r, formula, empty: false };
    }
    if ("text" in o) return readValue(o.text);
    if ("error" in o) return { text: String(o.error), num: null, formula: null, date: null, empty: false };
  }
  return EMPTY;
}

/** 公式是不是 SUM(...)（小計列） */
export function isSumFormula(f: string | null): boolean {
  return !!f && /^\s*SUM\s*\(/i.test(f);
}

/** 公式是否引用其他儲存格（例如 J35、C236-C237） */
export function hasCellRef(f: string | null): boolean {
  return !!f && /\$?[A-Z]{1,3}\$?\d+/i.test(f);
}

/** 欄號（1 起算）轉欄名 */
export function colName(n: number): string {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colNumber(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}
