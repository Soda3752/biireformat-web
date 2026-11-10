// 移植自 invoice-helper/tests/select.test.ts（二聯發票分頁的計算核心 src/invoice/）。執行：bun test tests/
import { test, expect, describe } from "bun:test";
import ExcelJS from "exceljs";
import { subsetSumMax, select, parseExcludeList, similarNames, type Options } from "../src/invoice/select.ts";
import type { BillFileParse, Half, HalfParse } from "../src/invoice/parseBills.ts";
import type { Triplicate } from "../src/invoice/parseTriplicate.ts";
import { parseTriplicateWorkbook } from "../src/invoice/parseTriplicate.ts";
import { addSheet } from "./invoice-helpers.ts";

function half(h: Half, customers: [string, number][], extra: Partial<HalfParse> = {}): HalfParse {
  return {
    half: h,
    sheetName: h,
    customers: customers.map(([name, amount], i) => ({ half: h, name, code: "", amount, address: `C${i + 2}` })),
    retail: [],
    deducts: [],
    unresolved: [],
    notes: [],
    checks: [],
    parsedTotal: 0,
    fileGrandTotal: null,
    fileBlockTotalSum: 0,
    warnings: [],
    ...extra,
  };
}
function bill(month: number, up: [string, number][], down: [string, number][], extra: Partial<HalfParse> = {}): BillFileParse {
  return { file: `${month}月.xlsx`, month, halves: [half("上期", up, extra), half("下期", down)], skippedSheets: [], warnings: [] };
}
function tri(name: string, amount: number, month: number): Triplicate {
  const sales = Math.round(amount / 1.05);
  return { date: { y: 2026, m: month, d: 15 }, month, name, invoiceNo: "", sales, tax: amount - sales, amount, source: "t" };
}
const base: Options = { mode: "period", total: 0, threshold: 7000, deductMode: "exclude", includeRetail: false, fixedExclude: [] };

describe("子集合加總：不超過差額且最大", () => {
  function brute(a: number[], cap: number): number {
    let best = 0;
    for (let mask = 0; mask < 1 << a.length; mask++) {
      let s = 0;
      for (let i = 0; i < a.length; i++) if (mask & (1 << i)) s += a[i]!;
      if (s <= cap && s > best) best = s;
    }
    return best;
  }
  test("隨機 300 組與暴力法比對", () => {
    let seed = 42;
    const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31);
    for (let t = 0; t < 300; t++) {
      const n = 1 + Math.floor(rnd() * 12);
      const a = Array.from({ length: n }, () => 7001 + Math.floor(rnd() * 20000));
      const cap = Math.floor(rnd() * a.reduce((x, y) => x + y, 0) * 1.1);
      const r = subsetSumMax(a, cap);
      const picked = r.picked.reduce((s, i) => s + a[i]!, 0);
      expect(new Set(r.picked).size).toBe(r.picked.length); // 不重複使用
      expect(picked).toBe(r.sum);
      expect(r.sum).toBeLessThanOrEqual(Math.max(cap, 0));
      expect(r.sum).toBe(brute(a, cap));
    }
  });
  test("貪婪法會失敗的例子：DP 要找到最佳", () => {
    // 貪婪（大的先拿）會拿 10000 → 剩 5000 放不下 6000、6000；最佳是 6000+6000+3000=15000
    const r = subsetSumMax([10000, 6000, 6000, 3000], 15000);
    expect(r.sum).toBe(15000);
  });
  test("差額 ≤ 0 → 不選；全部加總 ≤ 差額 → 全選", () => {
    expect(subsetSumMax([8000], 0)).toEqual({ sum: 0, picked: [] });
    expect(subsetSumMax([8000, 9000], -5)).toEqual({ sum: 0, picked: [] });
    expect(subsetSumMax([8000, 9000], 999999)).toEqual({ sum: 17000, picked: [0, 1] });
  });
});

describe("候選規則", () => {
  test("7000 門檻：要「大於」7000，單期計算（上下期不合併）", () => {
    const b = bill(9, [["剛好", 7000], ["過門檻", 7001], ["分兩期", 5000]], [["分兩期", 5000]]);
    const r = select([b], [], { ...base, total: 10 ** 7 });
    expect(r.candidates.map((c) => c.name)).toEqual(["過門檻"]);
  });

  test("三聯客戶整個排除（兩個月、上下期都排除），名稱前後空白要容忍", () => {
    const b9 = bill(9, [["老船長", 9734]], [["老船長", 9810], ["東方美", 11170]]);
    const b10 = bill(10, [["老船長", 9520]], [["東方美", 12251]]);
    const r = select([b9, b10], [tri("  老船長　", 50000, 9)], { ...base, total: 10 ** 7 });
    expect(r.candidates.map((c) => c.name)).toEqual(["東方美", "東方美"]);
    expect(r.excluded.filter((x) => x.entry.name === "老船長")).toHaveLength(3);
    expect(r.triplicateCustomers).toEqual([{ name: "老船長", matched: true, hints: [] }]);
  });

  test("固定排除名單；找不到時給近似名稱", () => {
    const b = bill(10, [["好味到", 46536], ["愛買", 44755]], []);
    const fixed = parseExcludeList("# 註解\n愛買\n好味道  # 原話\n");
    expect(fixed).toEqual(["愛買", "好味道"]);
    const r = select([b], [], { ...base, total: 10 ** 7, fixedExclude: fixed });
    expect(r.candidates.map((c) => c.name)).toEqual(["好味到"]);
    expect(r.fixedExcludeStatus.find((f) => f.name === "好味道")).toEqual({ name: "好味道", matched: false, hints: ["好味到"] });
    expect(similarNames("好味道", ["好味到", "愛買", "石先生"])).toEqual(["好味到"]);
  });

  test("扣代送客戶預設不列候選；--deduct net 用淨額", () => {
    const deducts = [{ half: "上期" as Half, name: "某客", code: "7001", gross: 86420, deduction: 4523, net: 81897, addresses: [] }];
    const b = bill(10, [["甲", 9000]], [], { deducts });
    const r1 = select([b], [], { ...base, total: 10 ** 7 });
    expect(r1.candidates.map((c) => c.name)).toEqual(["甲"]);
    const r2 = select([b], [], { ...base, total: 10 ** 7, deductMode: "net" });
    expect(r2.candidates.find((c) => c.name === "某客")?.amount).toBe(81897);
  });

  test("同一期同名客戶出現兩次 → 合併並警告", () => {
    const b = bill(9, [["重複", 4000], ["重複", 4000]], []);
    const r = select([b], [], { ...base, total: 10 ** 7 });
    expect(r.candidates.map((c) => [c.name, c.amount])).toEqual([["重複", 8000]]);
    expect(r.warnings.some((w) => w.includes("重複"))).toBe(true);
  });
});

describe("兩種總額模式", () => {
  const b9 = bill(9, [["A", 10000], ["B", 8000]], [["C", 9000]]);
  const b10 = bill(10, [["D", 12000]], [["E", 7500], ["F", 20000]]);
  const tris = [tri("X", 21000, 9), tri("Y", 10500, 10)];

  test("整期模式：兩個月四個半月期一起湊，扣兩個月的三聯", () => {
    const r = select([b9, b10], tris, { ...base, total: 21000 + 10500 + 40000 });
    expect(r.calcs).toHaveLength(1);
    const c = r.calcs[0]!;
    expect(c.gap).toBe(40000);
    // 候選 10000,8000,9000,12000,7500,20000；≤40000 最大 = 40000（例如 20000+12000+8000）
    expect(c.selectedSum).toBe(40000);
    expect(c.remaining).toBe(0);
    expect(r.selected.reduce((a, e) => a + e.amount, 0)).toBe(40000);
  });

  test("每月模式：各月分別扣該月三聯、只用該月候選", () => {
    const r = select([b9, b10], tris, {
      ...base,
      mode: "monthly",
      monthTotals: new Map([
        [9, 21000 + 18500], // 9 月差額 18500：候選 10000/8000/9000 → 最大 18000
        [10, 10500 + 30000], // 10 月差額 30000：候選 12000/7500/20000 → 最大 27500
      ]),
    });
    expect(r.calcs.map((c) => [c.label, c.gap, c.selectedSum, c.remaining])).toEqual([
      ["9月", 18500, 18000, 500],
      ["10月", 30000, 27500, 2500],
    ]);
    expect(r.selected.filter((e) => e.month === 9).every((e) => ["A", "B", "C"].includes(e.name))).toBe(true);
  });

  test("每月模式缺某月總額 → 報錯", () => {
    expect(() => select([b9, b10], [], { ...base, mode: "monthly", monthTotals: new Map([[9, 1]]) })).toThrow("10 月");
  });

  test("三聯超過總額 → 差額為負、不開二聯、有警告", () => {
    const r = select([b9], [tri("X", 50000, 9)], { ...base, total: 40000 });
    expect(r.selected).toEqual([]);
    expect(r.calcs[0]!.gap).toBe(-10000);
    expect(r.warnings.some((w) => w.includes("超過"))).toBe(true);
  });
});

describe("三聯清單：依表頭找欄位", () => {
  test("只有客戶名稱＋銷售額也能讀；金額 = 銷售額 + 5% 稅（四捨五入）；日期判斷月份", () => {
    const wb = new ExcelJS.Workbook();
    addSheet(wb, "工作表1", [
      ["備註文字"],
      ["客戶名稱", "日期", "銷售額"],
      ["甲", "115/09/15", 10001],
      ["乙", new Date(Date.UTC(2026, 9, 31)), 2000],
      ["合計", null, { formula: "SUM(C3:C4)", result: 12001 } as never],
    ]);
    const r = parseTriplicateWorkbook(wb, { defaultYear: 2026, months: [9, 10] });
    expect(r.items.map((t) => [t.name, t.month, t.sales, t.tax, t.amount])).toEqual([
      ["甲", 9, 10001, 500, 10501],
      ["乙", 10, 2000, 100, 2100],
    ]);
  });

  test("範本格式（左右兩月並排、有二聯欄）只讀三聯；有金額時以金額為準", () => {
    const wb = new ExcelJS.Workbook();
    addSheet(wb, "115.9.10", [
      ["115年09月"],
      ["日期", "(二)發票號碼", "金額", null, "日期", "客戶名稱", "(三)發票號碼", "銷售額", "營業稅", "金額", null, "日期", "(二)發票號碼", "金額", null, "日期", "客戶名稱", "(三)發票號碼", "銷售額", "營業稅", "金額"],
      [null, "JE1", 4370, null, null, "盛鑫商行", "JC1", 67728, 3386, 71114, null, null, "JE2", 100, null, null, "九犇", "JC2", 160, 8, 168],
    ]);
    const r = parseTriplicateWorkbook(wb, { defaultYear: 2026, months: [9, 10] });
    expect(r.items.map((t) => [t.name, t.month, t.invoiceNo, t.amount])).toEqual([
      ["盛鑫商行", 9, "JC1", 71114],
      ["九犇", 10, "JC2", 168],
    ]);
  });
});
