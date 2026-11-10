// 移植自 invoice-helper/tests/helpers.ts，另加讀真實樣本的小工具。
import ExcelJS from "exceljs";

export type CellSpec = string | number | null | Date | { formula: string; result?: number | string };

/** 用二維陣列建立工作表（第 1 列對應 rows[0]） */
export function addSheet(wb: ExcelJS.Workbook, name: string, rows: CellSpec[][]): ExcelJS.Worksheet {
  const ws = wb.addWorksheet(name);
  rows.forEach((r, i) => {
    r.forEach((v, j) => {
      if (v !== null && v !== undefined) ws.getCell(i + 1, j + 1).value = v as ExcelJS.CellValue;
    });
  });
  return ws;
}

export const sum = (f: string, result: number) => ({ formula: f, result });

// ---- 以下為移植時新增：invoice-helper 原本用 wb.xlsx.readFile 讀檔，這裡改成讀成 ArrayBuffer 再 load（與網頁版相同）。
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { parseWorkbook, type BillFileParse } from "../src/invoice/parseBills.ts";
import { monthFromFilename } from "../src/invoice/dates.ts";
import { parseCustomerSourceWorkbook, type CustomerSourceParse } from "../src/invoice/parseCustomers.ts";

/** invoice-helper 的真實樣本資料夾（唯讀，不在本專案；不存在時相關測試會略過） */
export const IH_SAMPLES = join(import.meta.dir, "..", "..", "invoice-helper", "samples");
export const sample = (f: string) => join(IH_SAMPLES, f);
export const hasSamples = (...fs: string[]) => fs.every((f) => existsSync(sample(f)));
export const IH_EXCLUDE = join(import.meta.dir, "..", "..", "invoice-helper", "config", "exclude.txt");

export function fileBuffer(path: string): ArrayBuffer {
  const b = readFileSync(path);
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength) as ArrayBuffer;
}
export async function loadWorkbook(path: string): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fileBuffer(path));
  return wb;
}
/** 等同 invoice-helper 的 parseBillFile */
export async function parseBillFile(path: string, month?: number): Promise<BillFileParse> {
  const m = month ?? monthFromFilename(path);
  if (!m) throw new Error(`無法從檔名判斷月份：${path}`);
  return parseWorkbook(await loadWorkbook(path), path, m);
}
/** 等同 invoice-helper 的 parseCustomerSourceFile */
export async function parseCustomerSourceFile(path: string): Promise<CustomerSourceParse> {
  return parseCustomerSourceWorkbook(await loadWorkbook(path), path.split(/[\\/]/).pop()!);
}
