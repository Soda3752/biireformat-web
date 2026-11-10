// 移植自 invoice-helper/tests/triText.test.ts（二聯發票分頁的計算核心 src/invoice/）。執行：bun test tests/
// 網頁版三聯大框框輸入的解析（src/triText.ts）
import { test, expect, describe } from "bun:test";
import { parseTriText, toHalfWidth } from "../src/invoice/triText.ts";

const amounts = (s: string) => parseTriText(s).amounts;
const errs = (s: string) => parseTriText(s).errors.map((e) => e.text);

describe("parseTriText 分隔", () => {
  test("空白、多個空白、前後空白", () => {
    expect(amounts("  1290 3000   450 ")).toEqual([1290, 3000, 450]);
  });
  test("換行（含 \\r\\n）與空行", () => {
    expect(amounts("1290\n3000\r\n\r\n450\n")).toEqual([1290, 3000, 450]);
  });
  test("Tab（Excel 多欄貼上，含空白格）", () => {
    expect(amounts("1290\t3000\t\n\t\t450\n12,600\t\t\n")).toEqual([1290, 3000, 450, 12600]);
  });
  test("全形空白、頓號、分號", () => {
    expect(amounts("1290　3000、450；600;700")).toEqual([1290, 3000, 450, 600, 700]);
  });
  test("空字串", () => {
    expect(parseTriText("")).toEqual({ amounts: [], sum: 0, errors: [], tokens: [] });
    expect(parseTriText(" \n\t ").amounts).toEqual([]);
  });
});

describe("parseTriText 全形與千分位", () => {
  test("全形數字轉半形", () => {
    expect(toHalfWidth("１２９０")).toBe("1290");
    expect(amounts("１２９０ ３，０００ ４５０")).toEqual([1290, 3000, 450]);
  });
  test("標準千分位當同一個數字", () => {
    expect(amounts("1,290 12,345,678 999")).toEqual([1290, 12345678, 999]);
  });
  test("千分位之間不用空白也能用逗號分隔", () => {
    expect(amounts("1,290,3000")).toEqual([1290, 3000]);
  });
  test("不是千分位的逗號當分隔", () => {
    expect(amounts("1290,300")).toEqual([1290, 300]);
    expect(amounts("12,34,5")).toEqual([12, 34, 5]);
    expect(amounts("1290, 300,")).toEqual([1290, 300]);
    expect(amounts("1,2903")).toEqual([1, 2903]);
  });
  test("逗號後面剛好三位數字＝千分位（500,300 是一個數）", () => {
    expect(amounts("500,300")).toEqual([500300]);
    expect(amounts("500, 300")).toEqual([500, 300]);
  });
  test("貨幣符號、元、小數、負數", () => {
    expect(amounts("$1,290 NT$3,000 450元 1,290.00 -600")).toEqual([1290, 3000, 450, 1290, -600]);
  });
  test("0 與單獨的「-」不算一張", () => {
    const r = parseTriText("1290 0 - 3000 －");
    expect(r.amounts).toEqual([1290, 3000]);
    expect(r.errors).toEqual([]);
  });
  test("合計", () => {
    expect(parseTriText("1,290 3000\n450").sum).toBe(4740);
  });
});

describe("parseTriText 錯字", () => {
  test("數字中夾字母：整段標錯、不納入", () => {
    const r = parseTriText("1290 12a4 3000");
    expect(r.amounts).toEqual([1290, 3000]);
    expect(r.errors.map((e) => e.text)).toEqual(["12a4"]);
    expect(r.errors[0]!.start).toBe(5);
    expect(r.errors[0]!.end).toBe(9);
  });
  test("純文字、錢字號、負號夾中間、多個小數點", () => {
    expect(errs("abc $ 1290-300 1.2.3")).toEqual(["abc", "$", "1290-300", "1.2.3"]);
  });
  test("用逗號連在一起的只標錯的那一小段", () => {
    const r = parseTriText("1290,12o4,3000");
    expect(r.amounts).toEqual([1290, 3000]);
    expect(r.errors.map((e) => e.text)).toEqual(["12o4"]);
  });
  test("錯誤位置對回原文（全形）", () => {
    const raw = "１２９０　１２ｘ４";
    const r = parseTriText(raw);
    expect(r.errors.length).toBe(1);
    expect(raw.slice(r.errors[0]!.start, r.errors[0]!.end)).toBe("１２ｘ４");
  });
});
