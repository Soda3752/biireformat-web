// 移植自 invoice-helper/tests/parseBills.test.ts（二聯發票分頁的計算核心 src/invoice/）。執行：bun test tests/
import { test, expect, describe } from "bun:test";
import ExcelJS from "exceljs";
import { existsSync } from "node:fs";
import { parseSheet, parseWorkbook, monthFromFilename } from "../src/invoice/parseBills.ts";
import { addSheet, sum, parseBillFile, sample } from "./invoice-helpers.ts";

const H7 = ["客戶名稱", "金額", "實收金額", "收款日", "備註", null, "客戶名稱", "金額", "實收金額", "收款日", "備註"];
const H8 = ["代號", "客戶名稱", "金額", "實收金額", "收款日", "備註", null, "代號", "客戶名稱", "金額", "實收金額", "收款日", "備註"];

describe("表頭找欄位", () => {
  test("7月版（無代號）：左右兩塊都讀到", () => {
    const wb = new ExcelJS.Workbook();
    const ws = addSheet(wb, "上期", [
      [],
      H7,
      ["早船長", 3755, 3755, "8/2", null, null, "堡嗝", 1023],
      ["辭修弘爺", 10623, null, null, null, null, "力行弘爺", 5932],
    ]);
    const r = parseSheet(ws, "上期");
    expect(r.customers.map((c) => [c.name, c.amount, c.code])).toEqual([
      ["早船長", 3755, ""],
      ["辭修弘爺", 10623, ""],
      ["堡嗝", 1023, ""],
      ["力行弘爺", 5932, ""],
    ]);
  });

  test("8月版（多一欄代號）：代號跟著客戶，金額不會讀成實收金額", () => {
    const wb = new ExcelJS.Workbook();
    const ws = addSheet(wb, "上期", [
      [],
      H8,
      [1080, "早安雞", 1395, 1395, "8/17", null, null, 1009, "三民呷", 3231, 9999],
    ]);
    const r = parseSheet(ws, "上期");
    expect(r.customers).toEqual([
      { half: "上期", name: "早安雞", code: "1080", amount: 1395, address: "C3" },
      { half: "上期", name: "三民呷", code: "1009", amount: 3231, address: "J3" },
    ]);
  });

  test("欄位順序不同也能找到（金額在實收金額右邊）", () => {
    const wb = new ExcelJS.Workbook();
    const ws = addSheet(wb, "下期", [["客戶名稱", "實收金額", "金額"], ["A客", 100, 200]]);
    expect(parseSheet(ws, "下期").customers[0]!.amount).toBe(200);
  });

  test("工作表名稱尾端空白要 trim；其他工作表略過", () => {
    const wb = new ExcelJS.Workbook();
    addSheet(wb, "銀行對帳結果", [["日期", "摘要", "存入金額"], ["2026/9/26", "轉帳", 2775]]);
    addSheet(wb, "上期", [H7, ["甲", 8000]]);
    addSheet(wb, "下期 ", [H7, ["乙", 9000]]);
    const r = parseWorkbook(wb, "x.xlsx", 10);
    expect(r.halves.map((h) => [h.half, h.customers[0]!.name])).toEqual([
      ["上期", "甲"],
      ["下期", "乙"],
    ]);
    expect(r.skippedSheets).toEqual(["銀行對帳結果"]);
  });

  test("從檔名抓月份", () => {
    expect(monthFromFilename("/a/b/10月_已自動對帳_最新.xlsx")).toBe(10);
    expect(monthFromFilename("9月_已自動對帳 (1).xlsx")).toBe(9);
    expect(monthFromFilename("對帳.xlsx")).toBeNull();
  });
});

describe("略過合計列、表尾小計與附註", () => {
  // 模擬真實檔：多區塊、SUM 小計、總計、零售、表尾石先生／代送、區域總表、附註列
  function build() {
    const wb = new ExcelJS.Workbook();
    const ws = addSheet(wb, "下期", [
      H8, // 1
      [1, "甲", 8000, null, null, null, null, 11, "丙", 3000], // 2
      [2, "乙", 2000, null, null, null, null, 12, "丁", 7001], // 3
      [null, null, null, null, null, null, null, null, "零售", 500], // 4
      [null, null, sum("SUM(C2:C4)", 10000), null, null, null, null, null, null, sum("SUM(J2:J4)", 10501)], // 5
      [null, null, null, null, null, null, null, null, "總計", sum("J5+C5", 20501)], // 6
      H8, // 7
      [3, "戊", 9000, null, null, null, null, 13, "己", 100], // 8
      [null, "總計", sum("SUM(C8:C8)", 9000), null, null, null, null, null, "總計", sum("SUM(J8:J8)", 100)], // 9
      [null, "石先生", 86420], // 10
      [null, "代送", 4523], // 11
      [7001, null, sum("C10-C11", 81897), null, null, null, null, null, "各區", sum("J6+C9+J9", 29601)], // 12
      [null, null, null, null, null, null, null, null, "石先生", sum("C12", 81897)], // 13
      [null, null, null, null, null, null, null, null, "總計", sum("SUM(J12:J13)", 111498)], // 14
      [null, "115年/10月", "115年/10月"], // 15
      [null, "彰化", 176223], // 16 區域總表（字面數字，不能當客戶）
      [null, "愛買+雲雀", sum("J2+J3", 10001)], // 17 附註
    ]);
    return parseSheet(ws, "下期");
  }

  test("只讀到真正的客戶，小計／總計／零售／區域列都不是客戶", () => {
    const r = build();
    expect(r.customers.map((c) => c.name)).toEqual(["甲", "乙", "丙", "丁", "戊", "己"]);
    expect(r.retail.map((x) => x.amount)).toEqual([500]);
  });

  test("表尾 石先生 − 代送 組合辨識為扣代送，不當一般客戶", () => {
    const r = build();
    expect(r.deducts).toHaveLength(1);
    expect(r.deducts[0]).toMatchObject({ name: "石先生", gross: 86420, deduction: 4523, net: 81897 });
  });

  test("附註列（愛買+雲雀）列入 notes 並有警告", () => {
    const r = build();
    expect(r.notes.map((n) => n.name)).toEqual(["愛買+雲雀"]);
    expect(r.warnings.some((w) => w.includes("愛買+雲雀"))).toBe(true);
  });

  test("小計公式用解析結果重算一致；整期總計 = 客戶＋零售＋淨額", () => {
    const r = build();
    expect(r.checks.filter((c) => c.diff !== 0)).toEqual([]);
    expect(r.parsedTotal).toBe(8000 + 2000 + 3000 + 7001 + 500 + 9000 + 100 + 81897);
    expect(r.fileGrandTotal?.value).toBe(111498);
  });

  test("總計公式引用空白儲存格（D35+J34）→ 報出差額與原因", () => {
    const wb = new ExcelJS.Workbook();
    const ws = addSheet(wb, "下期", [
      H8,
      [1, "甲", 8000, null, null, null, null, 11, "丙", 3000],
      [null, null, sum("SUM(C2:C2)", 8000), null, null, null, null, null, null, sum("SUM(J2:J2)", 3000)],
      [null, null, null, null, null, null, null, null, "總計", sum("D3+J3", 3000)],
    ]);
    const r = parseSheet(ws, "下期");
    const bad = r.checks.filter((c) => c.diff !== 0);
    expect(bad).toHaveLength(1);
    expect(bad[0]).toMatchObject({ kind: "區塊總計", fileValue: 3000, parsedValue: 11000, diff: 8000 });
    expect(bad[0]!.note).toContain("D3");
    expect(r.fileBlockTotalSum).toBe(3000);
  });
});

describe("真實樣本回歸測試（invoice-helper/samples/）", () => {
  const f9 = sample("9月_已自動對帳.xlsx");
  const f10 = sample("10月_已自動對帳_最新.xlsx");
  test.skipIf(!existsSync(f9))("9月：上期解析合計等於表尾總計，所有小計一致", async () => {
    const r = await parseBillFile(f9);
    const up = r.halves.find((h) => h.half === "上期")!;
    expect(up.parsedTotal).toBe(up.fileGrandTotal!.value);
    for (const h of r.halves) expect(h.checks.filter((c) => c.diff !== 0)).toEqual([]);
  });
  test.skipIf(!existsSync(f10))("10月最新：略過銀行對帳工作表；下期各區總計少算左半被偵測", async () => {
    const r = await parseBillFile(f10);
    expect(r.skippedSheets).toEqual(["銀行對帳結果"]);
    const down = r.halves.find((h) => h.half === "下期")!;
    expect(down.customers.find((c) => c.name === "早安雞")?.amount).toBe(1645);
    expect(down.customers.some((c) => /總計|小計|零售|代送/.test(c.name))).toBe(false);
    const bad = down.checks.filter((c) => c.diff !== 0);
    expect(bad.length).toBe(6);
    expect(bad.every((c) => c.kind === "區塊總計")).toBe(true);
    expect(down.deducts[0]).toMatchObject({ name: "石先生", code: "7001", gross: 86420, deduction: 4523, net: 81897 });
  });
});
