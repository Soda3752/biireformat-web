// 移植自 invoice-helper/tests/customers.test.ts（二聯發票分頁的計算核心 src/invoice/）。執行：bun test tests/
import { test, expect, describe } from "bun:test";
import { parseCustomerHeaderRow, parseCustomerRows, mergeUniformCustomers, type UniformCustomer } from "../src/invoice/parseCustomers.ts";
import { parseCustomerSourceFile, sample, hasSamples } from "./invoice-helpers.ts";
import { select, type Options } from "../src/invoice/select.ts";
import type { BillFileParse, Half, HalfParse } from "../src/invoice/parseBills.ts";


function half(h: Half, customers: [string, string, number][]): HalfParse {
  return {
    half: h,
    sheetName: h,
    customers: customers.map(([code, name, amount], i) => ({ half: h, name, code, amount, address: `C${i + 2}` })),
    retail: [],
    deducts: [],
    unresolved: [],
    notes: [],
    checks: [],
    parsedTotal: 0,
    fileGrandTotal: null,
    fileBlockTotalSum: 0,
    warnings: [],
  };
}
function bill(month: number, up: [string, string, number][], down: [string, string, number][] = []): BillFileParse {
  return { file: `${month}月.xlsx`, month, halves: [half("上期", up), half("下期", down)], skippedSheets: [], warnings: [] };
}
const uc = (code: string, name: string, uniform = "12345678"): UniformCustomer => ({ code, name, uniform, sources: ["t.xlsx"] });
const base: Options = { mode: "period", total: 10 ** 7, threshold: 7000, deductMode: "exclude", includeRetail: false, fixedExclude: [] };

describe("結帳原始檔：統編判定", () => {
  test("統編有填", () => {
    expect(parseCustomerHeaderRow(["客戶名稱", "辭修莫尼", "客戶編號", "1003", "統一編號", "88445345", "電話號碼", ""])).toEqual({ name: "辭修莫尼", code: "1003", uniform: "88445345" });
  });
  test("統編是空字串 → 沒有統編", () => {
    expect(parseCustomerHeaderRow(["客戶名稱", "零售", "客戶編號", "1000", "統一編號", "", "電話號碼", ""])!.uniform).toBe("");
    expect(parseCustomerHeaderRow(["客戶名稱", "零售", "客戶編號", "1000", "統一編號", "   ", "電話號碼", ""])!.uniform).toBe("");
  });
  test("空格被 reader 略過、「統一編號」後面直接接「電話號碼」→ 沒有統編", () => {
    expect(parseCustomerHeaderRow(["客戶名稱", "辭修弘爺", "客戶編號", "1001", "統一編號", "電話號碼", "0912"])!.uniform).toBe("");
    expect(parseCustomerHeaderRow(["客戶名稱", "某店", "客戶編號", "1002", "統一編號"])!.uniform).toBe("");
  });
  test("不是客戶開頭列 → null", () => {
    expect(parseCustomerHeaderRow(["客戶地址", "", "聯絡人員", "", "傳真電話", "現"])).toBeNull();
    expect(parseCustomerHeaderRow(["日期", "單號", "發票號碼"])).toBeNull();
  });
  test("同一客戶出現多次要去重；多檔取聯集", () => {
    const rows = [
      ["客戶名稱", "甲", "客戶編號", "1003", "統一編號", "11111111", "電話號碼", ""],
      ["日期", "單號"],
      ["客戶名稱", "甲", "客戶編號", "1003", "統一編號", "11111111", "電話號碼", ""],
      ["客戶名稱", "乙", "客戶編號", "1004", "統一編號", "", "電話號碼", ""],
      ["客戶名稱", "丙", "客戶編號", "1005", "統一編號", "電話號碼"],
    ];
    const a = parseCustomerRows(rows, "9月.xlsx");
    expect(a.customerRows).toBe(4);
    expect(a.customerCount).toBe(3);
    expect(a.uniformCustomers.map((c) => c.code)).toEqual(["1003"]);
    const b = parseCustomerRows([["客戶名稱", "甲", "客戶編號", "1003", "統一編號", "11111111"], ["客戶名稱", "丁", "客戶編號", "2001", "統一編號", "22222222"]], "10月.xlsx");
    const m = mergeUniformCustomers([a, b]);
    expect(m.customers.map((c) => [c.code, c.sources])).toEqual([
      ["1003", ["9月.xlsx", "10月.xlsx"]],
      ["2001", ["10月.xlsx"]],
    ]);
  });
});

describe("有統編客戶排除", () => {
  test("代號優先：名稱不同（愛買／愛買復興店）也排除；同名但代號不同不排除並列需人工確認", () => {
    const b = bill(9, [
      ["8028", "愛買", 80000],
      ["9999", "好婆婆", 9000], // 原始檔好婆婆是 1206 → 代號不同，不排除
      ["1001", "辭修弘爺", 11000],
    ]);
    const r = select([b], [], { ...base, uniformCustomers: [uc("8028", "愛買復興店"), uc("1206", "好婆婆")] });
    expect(r.excluded.map((x) => [x.entry.name, x.reason, x.uniform?.code])).toEqual([["愛買", "有統編", "8028"]]);
    expect(r.candidates.map((c) => c.name)).toEqual(["好婆婆", "辭修弘爺"]);
    const s = r.uniformStatus.find((u) => u.customer.code === "1206")!;
    expect(s.matchedBy).toBeNull();
    expect(r.pending.some((p) => p.includes("好婆婆") && p.includes("9999"))).toBe(true);
    expect(r.uniformStatus.find((u) => u.customer.code === "8028")).toMatchObject({ matchedBy: "代號", billNames: ["愛買"] });
  });

  test("對帳檔沒有代號（7 月格式）→ 用名稱比對，容忍空白；整期兩月上下期都排除", () => {
    const b7 = bill(7, [["", " H咖啡 ", 9000], ["", "辭修莫尼", 8000]], [["", "辭修莫尼", 9000]]);
    const b8 = bill(8, [["", "辭修莫尼", 9500]]);
    const r = select([b7, b8], [], { ...base, uniformCustomers: [uc("4089", "H 咖啡"), uc("1003", "辭修莫尼")] });
    expect(r.candidates).toEqual([]);
    expect(r.excluded.filter((x) => x.reason === "有統編")).toHaveLength(4);
    expect(r.uniformStatus.map((u) => u.matchedBy)).toEqual(["名稱", "名稱"]);
  });

  test("比對不到的統編客戶列入警告", () => {
    const r = select([bill(9, [["1001", "甲", 9000]])], [], { ...base, uniformCustomers: [uc("1082", "悅恩幼兒園")] });
    expect(r.warnings.some((w) => w.includes("比對不到") && w.includes("悅恩幼兒園"))).toBe(true);
  });

  test("固定排除名單照舊生效", () => {
    const r = select([bill(9, [["8048", "好味到", 46536], ["1001", "甲", 9000]])], [], { ...base, fixedExclude: ["好味到"], uniformCustomers: [] });
    expect(r.excluded.map((x) => [x.entry.name, x.reason])).toEqual([["好味到", "固定排除名單"]]);
  });
});

describe("三聯只給總額", () => {
  const b9 = bill(9, [["1", "A", 10000], ["2", "B", 8000]], [["3", "C", 9000]]);
  const b10 = bill(10, [["4", "D", 12000]], [["5", "E", 7500], ["6", "F", 20000]]);

  test("整期三聯總額：差額＝總額−三聯總額", () => {
    const r = select([b9, b10], [], { ...base, total: 100000, triplicateTotal: 60000 });
    expect(r.calcs.map((c) => [c.triplicateSum, c.gap, c.selectedSum, c.remaining])).toEqual([[60000, 40000, 40000, 0]]);
  });
  test("每月三聯總額（monthly 模式）", () => {
    const r = select([b9, b10], [], {
      ...base,
      mode: "monthly",
      monthTotals: new Map([[9, 30000], [10, 40000]]),
      monthTriplicate: new Map([[9, 11500], [10, 10000]]),
    });
    expect(r.calcs.map((c) => [c.label, c.triplicateSum, c.gap, c.selectedSum])).toEqual([
      ["9月", 11500, 18500, 18000],
      ["10月", 10000, 30000, 27500],
    ]);
  });
  test("每月模式用整期三聯總額 → 報錯；缺某月三聯總額 → 報錯", () => {
    const monthTotals = new Map([[9, 1], [10, 1]]);
    expect(() => select([b9, b10], [], { ...base, mode: "monthly", monthTotals, triplicateTotal: 5 })).toThrow("--month-triplicate");
    expect(() => select([b9, b10], [], { ...base, mode: "monthly", monthTotals, monthTriplicate: new Map([[9, 1]]) })).toThrow("10 月");
  });

});

const S9 = sample("9月_結帳原始檔.xlsx");
const S10 = sample("10月_結帳原始檔.xlsx");
describe.if(hasSamples("9月_結帳原始檔.xlsx", "10月_結帳原始檔.xlsx"))("真實結帳原始檔回歸", () => {
  test("9 月 43 位有統編、10 月 46 位、聯集 48 位", async () => {
    const a = await parseCustomerSourceFile(S9);
    const b = await parseCustomerSourceFile(S10);
    expect(a.uniformCustomers).toHaveLength(43);
    expect(b.uniformCustomers).toHaveLength(46);
    const m = mergeUniformCustomers([a, b]);
    expect(m.customers).toHaveLength(48);
    const get = (code: string) => m.customers.find((c) => c.code === code);
    expect(get("1003")?.name).toBe("辭修莫尼");
    expect(get("4030")?.name).toBe("肉蛋吐司");
    expect(get("8028")?.name).toBe("愛買復興店");
    expect(get("8048")?.uniform).toBe("83830262");
    expect(get("1000")).toBeUndefined(); // 零售沒統編
  }, 30000);
});
