// 移植自 invoice-helper/tests/overMax.test.ts（二聯發票分頁的計算核心 src/invoice/）。執行：bun test tests/
import { test, expect, describe } from "bun:test";
import { subsetSumOver, selectBalancedOver, select, type Options } from "../src/invoice/select.ts";
import type { BillFileParse, Half, HalfParse } from "../src/invoice/parseBills.ts";


function half(h: Half, customers: [string, number][]): HalfParse {
  return {
    half: h, sheetName: h,
    customers: customers.map(([name, amount], i) => ({ half: h, name, code: "", amount, address: `C${i + 2}` })),
    retail: [], deducts: [], unresolved: [], notes: [], checks: [], parsedTotal: 0, fileGrandTotal: null, fileBlockTotalSum: 0, warnings: [],
  };
}
const bill = (month: number, up: [string, number][], down: [string, number][] = []): BillFileParse =>
  ({ file: `${month}月.xlsx`, month, halves: [half("上期", up), half("下期", down)], skippedSheets: [], warnings: [] });
const base: Options = { mode: "period", total: 0, threshold: 7000, deductMode: "exclude", includeRetail: false, fixedExclude: [] };

describe("超出規則（不分散）subsetSumOver：[差額+min, 差額+max]、超出最少、尾數不為 00", () => {
  function brute(a: number[], lo: number, hi: number, tri: number): number | null {
    let best: number | null = null;
    for (let mask = 0; mask < 1 << a.length; mask++) {
      let s = 0;
      for (let i = 0; i < a.length; i++) if (mask & (1 << i)) s += a[i]!;
      if (s >= lo && s <= hi && (s + tri) % 100 !== 0 && (best === null || s < best)) best = s;
    }
    return best;
  }
  test("隨機 300 組與暴力法比對", () => {
    let seed = 7;
    const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31);
    for (let t = 0; t < 300; t++) {
      const n = 1 + Math.floor(rnd() * 12);
      const a = Array.from({ length: n }, () => 7001 + Math.floor(rnd() * 20000));
      const gap = Math.floor(rnd() * a.reduce((x, y) => x + y, 0));
      const min = 1 + Math.floor(rnd() * 500);
      const max = min + Math.floor(rnd() * 6000);
      const tri = Math.floor(rnd() * 3) * 1000 + Math.floor(rnd() * 2) * 50;
      const r = subsetSumOver(a, gap, max, tri, min);
      const want = brute(a, gap + min, gap + max, tri);
      if (want === null) expect(r.ok).toBe(false);
      else {
        expect(r.ok).toBe(true);
        if (!r.ok) continue;
        expect(r.sum).toBe(want);
        expect(r.picked.reduce((s, i) => s + a[i]!, 0)).toBe(r.sum);
      }
    }
  });
  test("尾數 00 也不行（例如 15500+三聯 500 = 16000、15050+50 = 15100）", () => {
    const r = subsetSumOver([7500, 8000, 7550, 8100], 15000, 5000, 50, 1);
    // 可湊 15050（7500+7550）+50=15100 尾數 00 → 跳過；15500（7500+8000）+50=15550 ✓
    expect(r.ok && r.sum).toBe(15500);
  });
  test("超出下限：差額 15000、min 300 → 15120（超出 120）不行，取 15530", () => {
    const r = subsetSumOver([8000, 7120, 7530], 15000, 5000, 0, 300);
    expect(r.ok && r.sum).toBe(15530);
  });
  test("範圍內無解 → ok=false，列出低於與高於範圍最接近的方案", () => {
    const r = subsetSumOver([10000, 20530], 12000, 5000, 0, 300);
    expect(r.ok).toBe(false);
    if (r.ok) return;
    expect(r.below?.sum).toBe(10000);
    expect(r.above?.sum).toBe(20530);
  });
});

describe("張數平均的精確挑選 selectBalancedOver", () => {
  // 暴力法：所有子集合中，合計在範圍、尾數非 00 → 求最小「最多−最少張數」
  function bruteSpread(halves: number[][], lo: number, hi: number, tri: number): number | null {
    const flat = halves.flatMap((h, hi_) => h.map((a) => ({ a, h: hi_ })));
    let best: number | null = null;
    for (let mask = 0; mask < 1 << flat.length; mask++) {
      let s = 0;
      const cnt = halves.map(() => 0);
      for (let i = 0; i < flat.length; i++) if (mask & (1 << i)) (s += flat[i]!.a), cnt[flat[i]!.h]!++;
      if (s < lo || s > hi || (s + tri) % 100 === 0) continue;
      const act = cnt.filter((_, i) => halves[i]!.length > 0);
      const sp = Math.max(...act) - Math.min(...act);
      if (best === null || sp < best) best = sp;
    }
    return best;
  }
  test("隨機 150 組與暴力法比對：可行性一致、張數差最小、合計與尾數條件成立", () => {
    let seed = 99;
    const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31);
    for (let t = 0; t < 150; t++) {
      const halves = Array.from({ length: 4 }, () => Array.from({ length: Math.floor(rnd() * 4) }, () => 7001 + Math.floor(rnd() * 15000)));
      const tot = halves.flat().reduce((a, b) => a + b, 0);
      const gap = Math.floor(rnd() * tot);
      const lo = gap + 300, hi = gap + 300 + Math.floor(rnd() * 8000);
      const tri = Math.floor(rnd() * 5) * 50;
      const r = selectBalancedOver(halves, lo, hi, tri);
      const want = bruteSpread(halves, lo, hi, tri);
      if (want === null) {
        expect(r.ok).toBe(false);
        continue;
      }
      expect(r.ok).toBe(true);
      if (!r.ok) continue;
      const sums = r.picks.map((p, h) => p.reduce((a, i) => a + halves[h]![i]!, 0));
      expect(sums.reduce((a, b) => a + b, 0)).toBe(r.sum);
      expect(r.picks.map((p) => p.length)).toEqual(r.counts);
      r.picks.forEach((p) => expect(new Set(p).size).toBe(p.length));
      expect(r.sum).toBeGreaterThanOrEqual(lo);
      expect(r.sum).toBeLessThanOrEqual(hi);
      expect((r.sum + tri) % 100).not.toBe(0);
      const act = r.counts.filter((_, i) => halves[i]!.length > 0);
      expect(Math.max(...act) - Math.min(...act)).toBe(want);
    }
  });
  // 暴力法：所有可行子集合的張數組合中，依（張數差, 平方差, ±總張數）取字典序最小的鍵
  function bruteKey(halves: number[][], lo: number, hi: number, tri: number, prefer: "many" | "few", cap = Infinity): number[] | null {
    const flat = halves.flatMap((h, hi_) => h.map((a) => ({ a, h: hi_ })));
    let best: number[] | null = null;
    for (let mask = 0; mask < 1 << flat.length; mask++) {
      let s = 0;
      const cnt = halves.map(() => 0);
      for (let i = 0; i < flat.length; i++) if (mask & (1 << i)) (s += flat[i]!.a), cnt[flat[i]!.h]!++;
      if (s < lo || s > hi || (s + tri) % 100 === 0 || cnt.some((c) => c > cap)) continue;
      const k = key(halves, cnt, prefer);
      if (best === null || cmpKey(k, best) < 0) best = k;
    }
    return best;
  }
  function key(halves: number[][], cnt: number[], prefer: "many" | "few"): number[] {
    const act = cnt.filter((_, i) => halves[i]!.length > 0);
    const mean = act.reduce((a, b) => a + b, 0) / Math.max(act.length, 1);
    const tot = cnt.reduce((a, b) => a + b, 0);
    return [act.length ? Math.max(...act) - Math.min(...act) : 0, Math.round(act.reduce((a, b) => a + (b - mean) ** 2, 0) * 1e6), prefer === "many" ? -tot : tot];
  }
  const cmpKey = (a: number[], b: number[]) => a[0]! - b[0]! || a[1]! - b[1]! || a[2]! - b[2]!;
  for (const prefer of ["many", "few"] as const)
    test(`prefer ${prefer}：隨機 200 組與暴力法比對（平均優先，再總張數${prefer === "many" ? "最多" : "最少"}）`, () => {
      let seed = prefer === "many" ? 7 : 13;
      const rnd = () => ((seed = (seed * 1103515245 + 12345) % 2 ** 31) / 2 ** 31);
      for (let t = 0; t < 200; t++) {
        const halves = Array.from({ length: 4 }, () => Array.from({ length: Math.floor(rnd() * 5) }, () => 7001 + Math.floor(rnd() * 15000)));
        const tot = halves.flat().reduce((a, b) => a + b, 0);
        const gap = Math.floor(rnd() * tot);
        const lo = gap + 300, hi = gap + 300 + Math.floor(rnd() * 8000);
        const tri = Math.floor(rnd() * 5) * 50;
        const cap = rnd() < 0.3 ? 2 : Infinity;
        const r = selectBalancedOver(halves, lo, hi, tri, cap, prefer);
        const want = bruteKey(halves, lo, hi, tri, prefer, cap);
        if (want === null) {
          expect(r.ok).toBe(false);
          continue;
        }
        expect(r.ok).toBe(true);
        if (!r.ok) continue;
        const sums = r.picks.map((p, h) => p.reduce((a, i) => a + halves[h]![i]!, 0));
        expect(sums.reduce((a, b) => a + b, 0)).toBe(r.sum);
        expect(r.picks.map((p) => p.length)).toEqual(r.counts);
        expect(r.sum).toBeGreaterThanOrEqual(lo);
        expect(r.sum).toBeLessThanOrEqual(hi);
        expect((r.sum + tri) % 100).not.toBe(0);
        expect(key(halves, r.counts, prefer)).toEqual(want);
        expect(r.counts.reduce((a, b) => a + b, 0)).toBeLessThanOrEqual(r.bound.total);
      }
    });
  test("prefer many 選多張小額、few 選少張大額", () => {
    // 每期：一張 30000 或三張 10000+；範圍可容納兩種
    const halves = [[30003, 10001, 10002, 10004], [30005, 10006, 10007, 10008], [30011, 10012, 10013, 10014], [30017, 10018, 10019, 10021]];
    const many = selectBalancedOver(halves, 120000, 125000, 0, Infinity, "many");
    const few = selectBalancedOver(halves, 120000, 125000, 0, Infinity, "few");
    expect(many.ok && many.counts).toEqual([3, 3, 3, 3]);
    expect(few.ok && few.counts).toEqual([1, 1, 1, 1]);
    expect(many.bound).toEqual({ total: 12, equalPerHalf: 3 });
  });
  test("平均優先：可以只從一期湊，但會選四期各 1 張", () => {
    const halves = [[10010, 10020, 10030, 10040], [10001], [10002], [10003]];
    const r = selectBalancedOver(halves, 40000, 41000, 0);
    expect(r.ok && r.counts).toEqual([1, 1, 1, 1]);
  });
  test("合計偏向範圍中間、個位數不是 0", () => {
    const halves = [[8001, 8501, 9001, 9501, 9203, 9707]];
    // 範圍 16000～19000，中間 17500；可湊 17502（8001+9501）、17208（8001+9207?）…
    const r = selectBalancedOver(halves, 16000, 19000, 0);
    expect(r.ok).toBe(true);
    if (!r.ok) return;
    expect(r.sum % 10).not.toBe(0);
    expect(Math.abs(r.sum - 17500)).toBeLessThanOrEqual(10);
  });
  test("每期上限（工作簿每欄格數）", () => {
    const r = selectBalancedOver([[8000, 8001, 8002], [9000]], 24003, 24003, 0, 2);
    expect(r.ok).toBe(false);
  });
});

describe("select 套用 over-min／over-max", () => {
  // 金額刻意不是整百（整百的組合會被「尾數 00」擋掉）
  const b9 = bill(9, [["A", 10010], ["B", 8020], ["G", 8830]], [["C", 9040], ["H", 7750]]);
  const b10 = bill(10, [["D", 12060], ["I", 7370]], [["E", 7580], ["F", 20090]]);

  test("整期：超出落在 300～5000、四期張數平均、摘要數字自洽", () => {
    const r = select([b9, b10], [], { ...base, total: 400000 + 40000, triplicateTotal: 400000, overMin: 300, overMax: 5000 });
    const c = r.calcs[0]!;
    expect(c.gap).toBe(40000);
    expect(c.over).toBeGreaterThanOrEqual(300);
    expect(c.over).toBeLessThanOrEqual(5000);
    expect(c.selectedSum).toBe(r.selected.reduce((a, e) => a + e.amount, 0));
    expect(c.grandTotal).toBe(c.selectedSum + 400000);
    expect(c.grandTotal % 100).not.toBe(0);
    expect(c.halfStats.map((h) => h.count)).toEqual([1, 1, 1, 1]);
    expect(c.halfStats.reduce((a, h) => a + h.sum, 0)).toBe(c.selectedSum);
  });
  test("無解：不開二聯、需人工確認列出兩個最接近方案", () => {
    const r = select([bill(9, [["A", 10000], ["B", 20530]])], [], { ...base, total: 12000, overMin: 300, overMax: 5000 });
    expect(r.selected).toEqual([]);
    expect(r.calcs[0]!.noSolution).toEqual({ below: { sum: 10000, count: 1 }, above: { sum: 20530, count: 1 } });
    expect(r.pending.some((p) => p.includes("找不到超出差額") && p.includes("低於範圍最接近") && p.includes("高於範圍最接近"))).toBe(true);
  });
  test("monthly 模式：每月各自套用（只用該月候選、該月上下期平均）", () => {
    const r = select([b9, b10], [], {
      ...base,
      mode: "monthly",
      monthTotals: new Map([[9, 1000 + 17000], [10, 2000 + 19000]]),
      monthTriplicate: new Map([[9, 1000], [10, 2000]]),
      overMin: 300,
      overMax: 5000,
    });
    for (const c of r.calcs) {
      expect(c.noSolution).toBeUndefined();
      expect(c.over).toBeGreaterThanOrEqual(300);
      expect(c.over).toBeLessThanOrEqual(5000);
      expect(c.grandTotal % 100).not.toBe(0);
      expect(c.halfStats.map((h) => h.count)).toEqual([1, 1]);
      expect(c.halfStats.every((h) => c.months.includes(h.month))).toBe(true);
    }
    expect(r.selected.filter((e) => e.month === 9).reduce((a, e) => a + e.amount, 0)).toBe(r.calcs[0]!.selectedSum);
  });
  test("沒給 over-max → 舊規則（最接近但不超過）", () => {
    const r = select([b9, b10], [], { ...base, total: 40000 });
    expect(r.calcs[0]!.selectedSum).toBeLessThanOrEqual(40000);
    expect(r.calcs[0]!.selectedSum).toBe(39940); // ≤ 40000 的最大可湊合計
    expect(r.selected.reduce((a, e) => a + e.amount, 0)).toBe(r.calcs[0]!.selectedSum);
    expect(r.calcs[0]!.overMax).toBeUndefined();
  });
});

