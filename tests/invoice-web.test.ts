// 移植自 invoice-helper/tests/web.test.ts：二聯發票分頁的計算（src/invoice/webCalc.ts）與命令列版結果一致。
// 網頁版計算（src/webCalc.ts）與命令列版結果一致
import { test, expect, describe } from "bun:test";
import ExcelJS from "exceljs";
import { readFileSync } from "node:fs";
import { computeFromBuffers, type WebInput } from "../src/invoice/webCalc.ts";
import { mergeUniformCustomers } from "../src/invoice/parseCustomers.ts";
import { parseExcludeList } from "../src/invoice/select.ts";
import { runSelection } from "../src/invoice/core.ts";
import { FIXED_EXCLUDE_TEXT } from "../src/invoice/exclude.ts";
import { fileBuffer, hasSamples, IH_EXCLUDE, loadWorkbook, parseBillFile, parseCustomerSourceFile, sample } from "./invoice-helpers.ts";

const S = sample;
const buf = (f: string) => fileBuffer(S(f));
// 分頁內建的固定排除名單（與 invoice-helper/config/exclude.txt 相同）
const excludeText = FIXED_EXCLUDE_TEXT;
const SAMPLES = ["9月_已自動對帳.xlsx", "10月_已自動對帳_修正版.xlsx", "9月_結帳原始檔.xlsx", "10月_結帳原始檔.xlsx", "二聯發票_使用者工作簿.xlsx"];

/** 使用者工作簿 115.9.10 的三聯逐格金額（G3:I34、Q3:S34） */
async function workbookTriplicate(): Promise<{ month: number; amounts: number[] }[]> {
  const wb = await loadWorkbook(S("二聯發票_使用者工作簿.xlsx"));
  const ws = wb.getWorksheet("115.9.10")!;
  return ([[9, ["G", "H", "I"]], [10, ["Q", "R", "S"]]] as const).map(([month, cols]) => {
    const amounts: number[] = [];
    for (let r = 3; r <= 34; r++)
      for (const c of cols) {
        const v = ws.getCell(`${c}${r}`).value as unknown;
        const x = typeof v === "number" ? v : (v as { result?: unknown } | null)?.result;
        if (typeof x === "number") amounts.push(x);
      }
    return { month, amounts };
  });
}

test("內建固定排除名單與 invoice-helper/config/exclude.txt 相同", () => {
  try {
    expect(FIXED_EXCLUDE_TEXT).toBe(readFileSync(IH_EXCLUDE, "utf8"));
  } catch (e) {
    if ((e as NodeJS.ErrnoException).code !== "ENOENT") throw e;
  }
  expect(parseExcludeList(FIXED_EXCLUDE_TEXT)).toEqual(["石先生", "愛買", "好味到"]);
});

describe.if(hasSamples(...SAMPLES))("網頁版 computeFromBuffers（115 年 9～10 月真實樣本）", async () => {
  const triplicate = hasSamples(...SAMPLES) ? await workbookTriplicate() : [];
  const input: WebInput = {
    bills: [
      { name: "10月_已自動對帳_修正版.xlsx", month: 10, data: buf("10月_已自動對帳_修正版.xlsx") },
      { name: "9月_已自動對帳.xlsx", month: 9, data: buf("9月_已自動對帳.xlsx") },
    ],
    sources: [
      { name: "9月_結帳原始檔.xlsx", data: buf("9月_結帳原始檔.xlsx") },
      { name: "10月_結帳原始檔.xlsx", data: buf("10月_結帳原始檔.xlsx") },
    ],
    triplicate,
    mode: "period",
    total: 1550000,
    threshold: 7000,
    overMin: 300,
    overMax: 5000,
    prefer: "many",
    maxPerHalf: 31,
    year: 115,
    excludeText,
  };
  const out = await computeFromBuffers(input);

  test("三聯逐格加總＝工作簿 G36、Q36", () => {
    expect(triplicate.map((t) => t.amounts.reduce((a, b) => a + b, 0))).toEqual([595667, 474504]);
  });

  test("結果等於命令列版正式參數：60 張、483,511、二+三聯 1,553,682、四期各 15 張", () => {
    const c = out.result.calcs[0]!;
    expect(out.result.selected.length).toBe(60);
    expect(c.triplicateSum).toBe(1070171);
    expect(c.gap).toBe(479829);
    expect(c.selectedSum).toBe(483511);
    expect(c.over).toBe(3682);
    expect(c.grandTotal).toBe(1553682);
    expect(c.halfStats.map((h) => [h.count, h.sum])).toEqual([[15, 125244], [15, 123551], [15, 122349], [15, 112367]]);
    expect(out.periodLabel).toBe("115年09-10月");
    expect(out.uniformLine).toContain("有統編客戶 48 位");
  });

  test("挑出的客戶與命令列版流程（讀檔版解析＋runSelection）逐筆相同", async () => {
    const bills = [await parseBillFile(S("9月_已自動對帳.xlsx")), await parseBillFile(S("10月_已自動對帳_修正版.xlsx"))];
    const sources = [await parseCustomerSourceFile(S("9月_結帳原始檔.xlsx")), await parseCustomerSourceFile(S("10月_結帳原始檔.xlsx"))];
    const r = runSelection(bills, [], {
      mode: "period", total: 1550000, threshold: 7000, deductMode: "exclude", includeRetail: false,
      fixedExclude: parseExcludeList(excludeText), uniformCustomers: mergeUniformCustomers(sources).customers,
      triplicateTotal: 1070171, overMin: 300, overMax: 5000, prefer: "many", maxPerHalf: 31,
    }, [], []);
    const key = (e: { month: number; half: string; code: string; name: string; amount: number }) => `${e.month}${e.half}${e.code}${e.name}${e.amount}`;
    expect(out.result.selected.map(key)).toEqual(r.selected.map(key));
    expect(out.result.excluded.length).toBe(r.excluded.length);
  });

  test("下載的客戶清單 Excel 打得開，二聯清單 60 列＋合計", async () => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out.report);
    expect(wb.worksheets.map((w) => w.name)).toEqual(["二聯清單", "摘要", "未選上候選", "排除名單", "對帳檔核對"]);
    const ws = wb.getWorksheet("二聯清單")!;
    expect(ws.getRow(1).values).toEqual([, "月份", "期別", "客戶代號", "客戶名稱", "二聯金額"]);
    expect((ws.getCell(63, 5).value as { result: number }).result).toBe(483511);
  });

  test("檢查：對帳檔數量、月份重複、總額", async () => {
    await expect(computeFromBuffers({ ...input, bills: input.bills.slice(0, 1) })).rejects.toThrow("兩個月");
    await expect(computeFromBuffers({ ...input, bills: input.bills.map((b) => ({ ...b, month: 9 })) })).rejects.toThrow("月份相同");
    await expect(computeFromBuffers({ ...input, total: undefined })).rejects.toThrow("總額");
  });
});
