// 移植自 invoice-helper/src/parseCustomers.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 讀結帳軟體匯出的「應收帳款對帳單--明細表」原始檔，找出「統一編號」有填的客戶（有統編＝開三聯，整期不開二聯）。
// 客戶開頭列範例：[客戶名稱, 辭修莫尼, 客戶編號, 1003, 統一編號, 88445345, 電話號碼, ]
// 統編沒填時可能是空字串，或（結帳網站 reader 會略過空格）「統一編號」後面直接接「電話號碼」。
// 判定邏輯同 biireformat-web 的 CustomerModel.newInstanceWithRowData：值不等於下一個標籤、且非空。
import ExcelJS from "exceljs";
import { readCell, cleanName } from "./cell.ts";

const L_NAME = "客戶名稱";
const L_CODE = "客戶編號";
const L_UNIFORM = "統一編號";
const L_PHONE = "電話號碼";
/** 客戶表頭的所有標籤：某標籤後面的「值」若剛好是另一個標籤，代表該值是空的（被 reader 略過） */
const LABELS = new Set([L_NAME, L_CODE, L_UNIFORM, L_PHONE, "客戶地址", "聯絡人員", "傳真電話"]);

export type UniformCustomer = {
  code: string; // 客戶編號
  name: string; // 原始檔上的客戶名稱
  uniform: string; // 統一編號
  sources: string[]; // 出現在哪些檔案（檔名）
};

export type CustomerHeader = { name: string; code: string; uniform: string };

/** 標籤後面的值；空白或剛好是另一個標籤 → "" */
function valueAfter(row: readonly string[], label: string): string {
  const i = row.indexOf(label);
  if (i < 0) return "";
  const v = cleanName(row[i + 1] ?? "");
  return LABELS.has(v) ? "" : v;
}

/** 這一列是不是客戶開頭列；是的話回傳 名稱／編號／統編（統編沒填為 ""） */
export function parseCustomerHeaderRow(row: readonly string[]): CustomerHeader | null {
  if (!row.includes(L_NAME) || !row.includes(L_CODE)) return null;
  return { name: valueAfter(row, L_NAME), code: valueAfter(row, L_CODE), uniform: valueAfter(row, L_UNIFORM) };
}

/** 以客戶編號去重（沒有編號才用名稱） */
const keyOf = (c: { code: string; name: string }) => (c.code ? `#${c.code}` : `@${c.name.replace(/[\s　 ]+/g, "")}`);

export type CustomerSourceParse = {
  file: string;
  customerRows: number; // 讀到的客戶開頭列（含重複）
  customerCount: number; // 去重後客戶數
  uniformCustomers: UniformCustomer[];
  warnings: string[];
};

/** rows：每列已轉成文字的儲存格陣列（空白格為 ""） */
export function parseCustomerRows(rows: readonly (readonly string[])[], label: string): CustomerSourceParse {
  const res: CustomerSourceParse = { file: label, customerRows: 0, customerCount: 0, uniformCustomers: [], warnings: [] };
  const all = new Map<string, CustomerHeader>();
  const uni = new Map<string, UniformCustomer>();
  for (const row of rows) {
    const h = parseCustomerHeaderRow(row);
    if (!h) continue;
    res.customerRows++;
    if (!h.code && !h.name) continue;
    const k = keyOf(h);
    if (!all.has(k)) all.set(k, h);
    if (!h.uniform) continue;
    const prev = uni.get(k);
    if (!prev) uni.set(k, { code: h.code, name: h.name, uniform: h.uniform, sources: [label] });
    else if (prev.uniform !== h.uniform)
      res.warnings.push(`${label}：客戶編號 ${h.code || "(無)"}「${h.name}」出現多次且統編不同（${prev.uniform}／${h.uniform}），以第一筆為準`);
  }
  res.customerCount = all.size;
  res.uniformCustomers = [...uni.values()];
  if (res.customerRows === 0) res.warnings.push(`${label}：找不到任何「客戶名稱／客戶編號」開頭列，可能不是應收帳款對帳單明細表`);
  return res;
}

export function worksheetRows(ws: ExcelJS.Worksheet): string[][] {
  const rows: string[][] = [];
  ws.eachRow({ includeEmpty: false }, (row) => {
    const out: string[] = [];
    for (let c = 1; c <= row.cellCount; c++) {
      const info = readCell(row.getCell(c));
      out.push(info.text || (info.num !== null ? String(info.num) : ""));
    }
    rows.push(out);
  });
  return rows;
}

/** 已載入的結帳原始檔工作簿（網頁版用 wb.xlsx.load 載入） */
export function parseCustomerSourceWorkbook(wb: ExcelJS.Workbook, label: string): CustomerSourceParse {
  return parseCustomerRows(wb.worksheets.flatMap((ws) => worksheetRows(ws)), label);
}

/** 多個原始檔取聯集（以客戶編號去重） */
export function mergeUniformCustomers(parses: CustomerSourceParse[]): { customers: UniformCustomer[]; warnings: string[] } {
  const m = new Map<string, UniformCustomer>();
  const warnings: string[] = [];
  for (const p of parses)
    for (const c of p.uniformCustomers) {
      const k = keyOf(c);
      const prev = m.get(k);
      if (!prev) m.set(k, { ...c, sources: [...c.sources] });
      else {
        for (const s of c.sources) if (!prev.sources.includes(s)) prev.sources.push(s);
        if (prev.uniform !== c.uniform) warnings.push(`客戶編號 ${c.code}「${c.name}」在不同原始檔的統編不同（${prev.uniform}／${c.uniform}），以 ${prev.sources[0]} 為準`);
        if (prev.name !== c.name) warnings.push(`客戶編號 ${c.code} 在不同原始檔名稱不同（「${prev.name}」／「${c.name}」）`);
      }
    }
  return { customers: [...m.values()].sort((a, b) => a.code.localeCompare(b.code, "en", { numeric: true })), warnings };
}
