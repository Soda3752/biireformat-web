// 移植自 invoice-helper/src/select.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 候選產生與挑選：門檻、排除（三聯客戶＋固定排除名單）、子集合加總最大化（精確 DP）。
import { cleanName } from "./cell.ts";
import type { BillFileParse, Half } from "./parseBills.ts";
import type { Triplicate } from "./parseTriplicate.ts";
import type { UniformCustomer } from "./parseCustomers.ts";

/**
 * 子集合加總最大化：從 amounts 選一組，總和最大但不超過 cap。精確動態規劃（金額為整數元）。
 * 回傳選中的索引（依原順序）。
 */
export function subsetSumMax(amounts: number[], cap: number): { sum: number; picked: number[] } {
  if (cap <= 0 || amounts.length === 0) return { sum: 0, picked: [] };
  for (const a of amounts) if (!Number.isInteger(a) || a < 0) throw new Error(`金額必須是非負整數：${a}`);
  const total = amounts.reduce((x, y) => x + y, 0);
  if (total <= cap) return { sum: total, picked: amounts.map((_, i) => i) };
  const C = Math.floor(cap);
  // from[s] = 第一次湊到總和 s 時用到的項目索引（-1 = 尚未湊到）；s=0 用 -2 代表起點
  const from = new Int32Array(C + 1).fill(-1);
  from[0] = -2;
  let best = 0;
  for (let i = 0; i < amounts.length; i++) {
    const a = amounts[i]!;
    if (a === 0 || a > C) continue;
    for (let s = C; s >= a; s--) {
      if (from[s] === -1 && from[s - a] !== -1) {
        from[s] = i;
        if (s > best) best = s;
      }
    }
    if (best === C) break;
  }
  const picked: number[] = [];
  let s = best;
  while (s > 0) {
    const i = from[s]!;
    picked.push(i);
    s -= amounts[i]!;
  }
  picked.sort((x, y) => x - y);
  return { sum: best, picked };
}

/** 可達總和表：from[s] = 第一次湊到 s 時用到的項目索引（-1 未湊到；s=0 為 -2）。精確 0/1 DP。 */
function reachable(amounts: number[], limit: number): Int32Array {
  for (const a of amounts) if (!Number.isInteger(a) || a < 0) throw new Error(`金額必須是非負整數：${a}`);
  const L = Math.max(0, Math.floor(limit));
  const from = new Int32Array(L + 1).fill(-1);
  from[0] = -2;
  for (let i = 0; i < amounts.length; i++) {
    const a = amounts[i]!;
    if (a === 0 || a > L) continue;
    for (let s = L; s >= a; s--) if (from[s] === -1 && from[s - a] !== -1) from[s] = i;
  }
  return from;
}
function rebuild(from: Int32Array, amounts: number[], sum: number): number[] {
  const picked: number[] = [];
  let s = sum;
  while (s > 0) {
    const i = from[s]!;
    picked.push(i);
    s -= amounts[i]!;
  }
  return picked.sort((x, y) => x - y);
}

/** 「二聯＋三聯合計」尾數是 00（含整千）→ 不接受 */
export const isRoundHundred = (n: number) => n % 100 === 0;

/**
 * 不考慮分散時的最接近方案（無解時列給使用者參考，或作為工具函式）：
 * 二聯合計 s 落在 [gap+overMin, gap+overMax] 取超出最少、s+triplicate 尾數不是 00。
 * 無解時回傳低於範圍最接近、高於範圍最接近（尾數不是 00）的方案。
 */
export type OverResult =
  | { ok: true; sum: number; picked: number[]; over: number }
  | { ok: false; below: { sum: number; picked: number[] } | null; above: { sum: number; picked: number[] } | null };

export function subsetSumOver(amounts: number[], gap: number, overMax: number, triplicate = 0, overMin = 1): OverResult {
  if (!Number.isInteger(overMax) || !Number.isInteger(overMin) || overMin < 1 || overMax < overMin) throw new Error(`超出範圍不合理：${overMin}～${overMax}`);
  const total = amounts.reduce((x, y) => x + y, 0);
  const lo = Math.max(gap + overMin, 0);
  const hi = gap + overMax;
  if (hi >= 0 && lo <= total) {
    const from = reachable(amounts, Math.min(hi, total));
    for (let s = lo; s <= Math.min(hi, total); s++)
      if (from[s] !== -1 && !isRoundHundred(s + triplicate)) return { ok: true, sum: s, picked: rebuild(from, amounts, s), over: s - gap };
  }
  const from = reachable(amounts, total);
  let below: { sum: number; picked: number[] } | null = null;
  for (let s = Math.min(lo - 1, total); s >= 0; s--)
    if (from[s] !== -1) {
      below = { sum: s, picked: rebuild(from, amounts, s) };
      break;
    }
  let above: { sum: number; picked: number[] } | null = null;
  for (let s = Math.max(hi + 1, 0); s <= total; s++)
    if (from[s] !== -1 && !isRoundHundred(s + triplicate)) {
      above = { sum: s, picked: rebuild(from, amounts, s) };
      break;
    }
  return { ok: false, below, above };
}

// ---- 四期（或兩期）張數平均的精確挑選 ----

/** dst |= src << a（位元集合，只保留 0..limit） */
function shiftOr(dst: Uint32Array, src: Uint32Array, a: number) {
  const ws = a >>> 5, bs = a & 31, W = dst.length;
  if (bs === 0) for (let j = W - 1; j >= ws; j--) dst[j]! |= src[j - ws]!;
  else
    for (let j = W - 1; j >= ws; j--) {
      const lo = j - ws - 1 >= 0 ? src[j - ws - 1]! >>> (32 - bs) : 0;
      dst[j]! |= (src[j - ws]! << bs) | lo;
    }
}
const bitGet = (b: Uint32Array, i: number) => (b[i >>> 5]! >>> (i & 31)) & 1;

/** 從 init 集合出發，再從 amounts 恰好選 c 個，可達總和集合（≤ limit） */
function stageReach(init: Uint32Array, amounts: number[], c: number, limit: number): Uint32Array {
  const W = init.length;
  const L = Array.from({ length: c + 1 }, (_, k) => (k === 0 ? init.slice() : new Uint32Array(W)));
  amounts.forEach((a, i) => {
    if (a > limit) return;
    for (let k = Math.min(c, i + 1); k >= 1; k--) shiftOr(L[k]!, L[k - 1]!, a);
  });
  // 清掉超過 limit 的位元
  const out = L[c]!;
  for (let i = limit + 1; i < W * 32; i++) if (bitGet(out, i)) out[i >>> 5]! &= ~(1 << (i & 31));
  return out;
}

/** 單一期別：從 amounts 恰好選 c 個、總和 t 且 ok(t) 成立的一組（回傳索引） */
function pickExact(amounts: number[], c: number, limit: number, ok: (t: number) => boolean): { t: number; picked: number[] } | null {
  const N = limit + 1;
  const from = new Int16Array((c + 1) * N).fill(-1);
  from[0] = -2;
  amounts.forEach((a, i) => {
    if (a > limit) return;
    for (let k = Math.min(c, i + 1); k >= 1; k--) {
      const cur = k * N, prev = (k - 1) * N;
      for (let t = limit; t >= a; t--) if (from[cur + t] === -1 && from[prev + t - a] !== -1) from[cur + t] = i;
    }
  });
  for (let t = limit; t >= 0; t--) {
    if (from[c * N + t] === -1 || !ok(t)) continue;
    const picked: number[] = [];
    let k = c, s = t;
    while (k > 0) {
      const i = from[k * N + s]!;
      picked.push(i);
      s -= amounts[i]!;
      k--;
    }
    return { t, picked: picked.sort((x, y) => x - y) };
  }
  return null;
}

export type Prefer = "many" | "few";

/** 張數上限估計（只看金額最小的幾張能否塞進範圍上限，是上界、不保證湊得到） */
export type CountBound = {
  /** 不管分佈，四期合計最多幾張：金額最小的 N 張合計 ≤ hi 的最大 N（受每期上限限制） */
  total: number;
  /** 四期（有候選的期別）各取相同張數 c 時，各期最小 c 張合計 ≤ hi 的最大 c */
  equalPerHalf: number;
};

export type BalancedResult =
  | { ok: true; sum: number; counts: number[]; picks: number[][]; vectorsTried: number; bound: CountBound }
  | { ok: false; vectorsTried: number; bound: CountBound };

/**
 * 二聯合計 s ∈ [lo, hi]、s+triplicate 尾數不是 00；
 * 先求各期張數最平均（最多與最少張數差最小，再來各期張數與平均的平方差最小），
 * 再依 prefer：many＝總張數越多越好（偏小額客戶，預設）；few＝總張數越少越好（偏大客戶），
 * 再在該張數組合可湊到的合計中，選最接近範圍中間、且二聯＋三聯合計個位數不是 0 的（都不行才接受個位 0）。
 * 精確：每個張數組合都用位元集合 DP 檢查可行，不是貪婪。
 */
export function selectBalancedOver(halves: number[][], lo: number, hi: number, triplicate = 0, maxPerHalf = Infinity, prefer: Prefer = "many"): BalancedResult {
  lo = Math.max(lo, 0);
  for (const h of halves) for (const a of h) if (!Number.isInteger(a) || a <= 0) throw new Error(`金額必須是正整數：${a}`);
  const H = halves.length;
  const caps = halves.map((h) => Math.min(h.length, maxPerHalf));
  const asc = halves.map((h) => [0, ...[...h].sort((x, y) => x - y)].map((_, i, arr) => arr.slice(0, i + 1).reduce((p, q) => p + q, 0)));
  const desc = halves.map((h) => [0, ...[...h].sort((x, y) => y - x)].map((_, i, arr) => arr.slice(0, i + 1).reduce((p, q) => p + q, 0)));
  const active = caps.map((c, i) => (c > 0 ? i : -1)).filter((i) => i >= 0);
  const bound: CountBound = { total: 0, equalPerHalf: 0 };
  {
    const smallest = halves.flatMap((h, i) => [...h].sort((x, y) => x - y).slice(0, caps[i]!)).sort((x, y) => x - y);
    let s = 0;
    for (const a of smallest) {
      if (s + a > hi) break;
      s += a;
      bound.total++;
    }
    if (active.length)
      for (let c = 1; c <= Math.min(...active.map((i) => caps[i]!)); c++) {
        if (active.reduce((a, i) => a + asc[i]![c]!, 0) > hi) break;
        bound.equalPerHalf = c;
      }
  }
  if (hi < lo) return { ok: false, vectorsTried: 0, bound };
  // 列出所有張數組合，用最小／最大可能總和先過濾
  const vecs: { v: number[]; spread: number; dev: number }[] = [];
  const v = new Array(H).fill(0);
  const rec = (h: number, minS: number, maxS: number) => {
    if (h === H) {
      if (minS > hi || maxS < lo) return;
      const act = active.map((i) => v[i]!);
      const spread = act.length ? Math.max(...act) - Math.min(...act) : 0;
      const mean = act.reduce((a, b) => a + b, 0) / Math.max(act.length, 1);
      vecs.push({ v: [...v], spread, dev: act.reduce((a, b) => a + (b - mean) ** 2, 0) });
      return;
    }
    for (let c = 0; c <= caps[h]!; c++) {
      if (minS + asc[h]![c]! > hi) break;
      v[h] = c;
      rec(h + 1, minS + asc[h]![c]!, maxS + desc[h]![c]!);
    }
    v[h] = 0;
  };
  rec(0, 0, 0);
  const total = (x: number[]) => x.reduce((a, b) => a + b, 0);
  const dir = prefer === "many" ? -1 : 1;
  vecs.sort((a, b) => a.spread - b.spread || a.dev - b.dev || dir * (total(a.v) - total(b.v)) || a.v.join().localeCompare(b.v.join()));

  const W = (hi >>> 5) + 1;
  const mid = (lo + hi) / 2;
  const okTail = (s: number) => !isRoundHundred(s + triplicate);
  let tried = 0;
  for (const { v: vec } of vecs) {
    tried++;
    const stages: Uint32Array[] = [];
    let cur: Uint32Array = new Uint32Array(W);
    cur[0] = 1;
    for (let h = 0; h < H; h++) {
      cur = stageReach(cur, halves[h]!, vec[h]!, hi);
      stages.push(cur);
    }
    // 選合計：最接近中間，個位數不是 0 優先
    let best = -1, bestScore = Infinity;
    for (let s = lo; s <= hi; s++) {
      if (!bitGet(cur, s) || !okTail(s)) continue;
      const score = Math.abs(s - mid) + ((s + triplicate) % 10 === 0 ? 1e9 : 0);
      if (score < bestScore) (bestScore = score), (best = s);
    }
    if (best < 0) continue;
    // 由後往前還原各期組合
    const picks: number[][] = new Array(H);
    let rest = best;
    for (let h = H - 1; h >= 0; h--) {
      const prev = h > 0 ? stages[h - 1]! : null;
      const r = pickExact(halves[h]!, vec[h]!, Math.min(rest, desc[h]![vec[h]!]!), (t) => (prev ? bitGet(prev, rest - t) === 1 : rest - t === 0));
      if (!r) throw new Error("內部錯誤：還原組合失敗");
      picks[h] = r.picked;
      rest -= r.t;
    }
    return { ok: true, sum: best, counts: vec, picks, vectorsTried: tried, bound };
  }
  return { ok: false, vectorsTried: tried, bound };
}

export type Mode = "period" | "monthly";

export type Options = {
  mode: Mode;
  /** period 模式：整期總額 */
  total?: number;
  /** monthly 模式：每月總額 */
  monthTotals?: Map<number, number>;
  threshold: number;
  /** 扣代送客戶（石先生）候選金額：淨額／原始金額／不列 */
  deductMode: "net" | "gross" | "exclude";
  includeRetail: boolean;
  /** 固定排除名單（config/exclude.txt） */
  fixedExclude: string[];
  /** 結帳原始檔中有統編的客戶（整期排除）；undefined＝沒提供原始檔 */
  uniformCustomers?: UniformCustomer[];
  /** 三聯只給總額：整期一個（period 模式） */
  triplicateTotal?: number;
  /** 三聯只給總額：每月各一個（monthly 模式必用；period 模式也可，會加總） */
  monthTriplicate?: Map<number, number>;
  /**
   * 超出規則：給了就改為「二聯合計超出差額 overMin～overMax 元、二聯＋三聯合計尾數不為 00、各半月期張數平均」；
   * 沒給＝舊規則（最接近但不超過差額）。monthly 模式每月各自套用。
   */
  overMax?: number;
  /** 超出下限（預設 1）；超出金額落在 overMin～overMax */
  overMin?: number;
  /** 每個半月期最多幾張（使用者工作簿每欄可填格數） */
  maxPerHalf?: number;
  /** 超出規則下，各期張數平均之後的偏好：many＝總張數多（偏小額，預設）；few＝總張數少（偏大客戶） */
  prefer?: Prefer;
};

export type Entry = {
  month: number;
  half: Half;
  name: string;
  code: string;
  amount: number;
  kind: "customer" | "deduct" | "retail";
  note: string;
};

export type Excluded = { entry: Entry; reason: string; uniform?: UniformCustomer };

export type UniformStatus = {
  customer: UniformCustomer;
  /** 用什麼對到對帳檔：代號／名稱；null＝找不到 */
  matchedBy: "代號" | "名稱" | null;
  /** 對到的對帳檔客戶名稱 */
  billNames: string[];
  /** 找不到時的提示 */
  hint: string;
};

export type MonthCalc = {
  label: string; // 「整期」或「9月」
  months: number[];
  total: number;
  triplicateSum: number;
  gap: number;
  selectedSum: number;
  /** 差額 − 二聯合計（超出規則下為負） */
  remaining: number;
  /** 二聯＋三聯合計 */
  grandTotal: number;
  /** 超出規則的上限（沒啟用為 undefined） */
  overMax?: number;
  /** 二聯合計 − 差額 */
  over: number;
  /** 超出規則範圍內無解 */
  noSolution?: { below: { sum: number; count: number } | null; above: { sum: number; count: number } | null };
  overMin?: number;
  /** 超出規則下的張數上限估計（上界） */
  countBound?: CountBound;
  /** 各半月期二聯張數與金額 */
  halfStats: { month: number; half: Half; count: number; sum: number }[];
};

export type SelectionResult = {
  months: number[];
  entries: Entry[];
  candidates: Entry[];
  selected: Entry[];
  unselected: Entry[];
  excluded: Excluded[];
  belowThreshold: number;
  calcs: MonthCalc[];
  triplicateCustomers: { name: string; matched: boolean; hints: string[] }[];
  fixedExcludeStatus: { name: string; matched: boolean; hints: string[] }[];
  uniformStatus: UniformStatus[];
  pending: string[]; // 需人工確認事項
  warnings: string[];
};

function levenshtein(a: string, b: string): number {
  const A = [...a], B = [...b];
  const dp = Array.from({ length: A.length + 1 }, (_, i) => [i, ...Array(B.length).fill(0)] as number[]);
  for (let j = 1; j <= B.length; j++) dp[0]![j] = j;
  for (let i = 1; i <= A.length; i++)
    for (let j = 1; j <= B.length; j++)
      dp[i]![j] = Math.min(dp[i - 1]![j]! + 1, dp[i]![j - 1]! + 1, dp[i - 1]![j - 1]! + (A[i - 1] === B[j - 1] ? 0 : 1));
  return dp[A.length]![B.length]!;
}

/** 名稱找不到時給近似名稱提示（包含關係或編輯距離 ≤ 1） */
export function similarNames(name: string, pool: Iterable<string>): string[] {
  const out = new Set<string>();
  for (const p of pool) {
    if (p === name) continue;
    if ((p.includes(name) || name.includes(p)) && Math.min(p.length, name.length) >= 2) out.add(p);
    else if (levenshtein(p, name) <= 1 && name.length >= 2) out.add(p);
  }
  return [...out].slice(0, 5);
}

export function buildEntries(bills: BillFileParse[], opts: Options): { entries: Entry[]; warnings: string[]; pending: string[] } {
  const entries: Entry[] = [];
  const warnings: string[] = [];
  const pending: string[] = [];
  for (const b of bills) {
    for (const h of b.halves) {
      const byName = new Map<string, Entry>();
      for (const c of h.customers) {
        const name = cleanName(c.name);
        const prev = byName.get(name);
        if (prev) {
          warnings.push(`${b.month}月${h.half}「${name}」出現多次（${prev.note || ""}${c.address}），金額已合併為 ${prev.amount + c.amount}`);
          prev.amount += c.amount;
          prev.note = (prev.note ? prev.note + "、" : "") + c.address;
          if (!prev.code && c.code) prev.code = c.code;
        } else {
          const e: Entry = { month: b.month, half: h.half, name, code: c.code, amount: c.amount, kind: "customer", note: "" };
          byName.set(name, e);
          entries.push(e);
        }
      }
      for (const d of h.deducts) {
        const amount = opts.deductMode === "gross" && d.gross !== null ? d.gross : d.net;
        entries.push({
          month: b.month,
          half: h.half,
          name: cleanName(d.name),
          code: d.code,
          amount,
          kind: "deduct",
          note: `表尾扣代送列：金額 ${d.gross ?? "?"} − 代送 ${d.deduction} = 淨額 ${d.net}`,
        });
      }
      if (opts.includeRetail) {
        for (const r of h.retail) {
          entries.push({ month: b.month, half: h.half, name: `零售（${r.address}）`, code: "", amount: r.amount, kind: "retail", note: "零售列" });
        }
      }
    }
  }
  for (const e of entries) {
    if (!Number.isInteger(e.amount)) {
      warnings.push(`${e.month}月${e.half}「${e.name}」金額 ${e.amount} 不是整數，已四捨五入`);
      e.amount = Math.round(e.amount);
    }
  }
  return { entries, warnings, pending };
}

/** 名稱比對用：去掉所有空白 */
export const nameKey = (s: string) => s.replace(/[\s　 ]+/g, "");

/**
 * 有統編客戶比對：對帳檔該列有代號 → 只用代號比對；沒有代號（例如 7 月格式）→ 用名稱（忽略空白）。
 */
export function makeUniformMatcher(list: UniformCustomer[]) {
  const byCode = new Map<string, UniformCustomer>();
  const byName = new Map<string, UniformCustomer>();
  for (const u of list) {
    if (u.code && !byCode.has(u.code)) byCode.set(u.code, u);
    const k = nameKey(u.name);
    if (k && !byName.has(k)) byName.set(k, u);
  }
  return (name: string, code: string): { u: UniformCustomer; by: "代號" | "名稱" } | null => {
    const c = code.trim();
    if (c) {
      const u = byCode.get(c);
      return u ? { u, by: "代號" } : null;
    }
    const u = byName.get(nameKey(name));
    return u ? { u, by: "名稱" } : null;
  };
}

export function select(bills: BillFileParse[], triplicates: Triplicate[], opts: Options): SelectionResult {
  const months = [...new Set(bills.map((b) => b.month))].sort((a, b) => a - b);
  const { entries, warnings, pending } = buildEntries(bills, opts);
  const allNames = new Set(entries.map((e) => e.name));
  for (const b of bills) for (const h of b.halves) for (const c of h.customers) allNames.add(cleanName(c.name));

  const triNames = new Set(triplicates.map((t) => cleanName(t.name)));
  const fixed = opts.fixedExclude.map(cleanName).filter(Boolean);
  const fixedSet = new Set(fixed);

  const uniList = opts.uniformCustomers ?? [];
  const matchUniform = makeUniformMatcher(uniList);

  const excluded: Excluded[] = [];
  const candidates: Entry[] = [];
  let belowThreshold = 0;
  for (const e of entries) {
    const um = e.kind === "retail" ? null : matchUniform(e.name, e.code);
    if (um) excluded.push({ entry: e, reason: "有統編", uniform: um.u });
    else if (triNames.has(e.name)) excluded.push({ entry: e, reason: "有開三聯式發票" });
    else if (fixedSet.has(e.name)) excluded.push({ entry: e, reason: "固定排除名單" });
    else if (e.kind === "deduct" && opts.deductMode === "exclude") excluded.push({ entry: e, reason: "扣代送客戶（設定不列入）" });
    else if (e.amount > opts.threshold) candidates.push(e);
    else belowThreshold++;
  }

  const triplicateCustomers = [...triNames].map((name) => {
    const matched = allNames.has(name);
    return { name, matched, hints: matched ? [] : similarNames(name, allNames) };
  });
  for (const t of triplicateCustomers.filter((t) => !t.matched)) {
    warnings.push(`三聯客戶「${t.name}」在對帳檔找不到同名客戶${t.hints.length ? `（近似：${t.hints.join("、")}）` : ""}，請確認名稱是否一致，否則該客戶不會被排除`);
  }
  const fixedExcludeStatus = fixed.map((name) => {
    const matched = allNames.has(name);
    return { name, matched, hints: matched ? [] : similarNames(name, allNames) };
  });
  for (const f of fixedExcludeStatus.filter((f) => !f.matched)) {
    warnings.push(`固定排除名單的「${f.name}」在本期對帳檔找不到${f.hints.length ? `（近似：${f.hints.join("、")}，若是同一客戶請改 config/exclude.txt）` : ""}`);
  }

  // 有統編客戶比對狀態（對帳檔所有客戶列，含未達門檻者）
  const billRows: { name: string; code: string }[] = [];
  for (const b of bills)
    for (const h of b.halves) {
      for (const c of h.customers) billRows.push({ name: cleanName(c.name), code: c.code.trim() });
      for (const d of h.deducts) billRows.push({ name: cleanName(d.name), code: d.code.trim() });
    }
  const uniformStatus: UniformStatus[] = uniList.map((u) => {
    const names = new Set<string>();
    let by: "代號" | "名稱" | null = null;
    for (const r of billRows) {
      const m = matchUniform(r.name, r.code);
      if (m && m.u === u) {
        names.add(r.name);
        if (by !== "代號") by = m.by;
      }
    }
    let hint = "";
    if (!by) {
      const same = billRows.filter((r) => r.code && nameKey(r.name) === nameKey(u.name));
      if (same.length) hint = `對帳檔有同名客戶「${same[0]!.name}」但代號是 ${[...new Set(same.map((r) => r.code))].join("/")}，與原始檔 ${u.code} 不同，未排除`;
      else {
        const sim = similarNames(u.name, new Set(billRows.map((r) => r.name)));
        hint = sim.length ? `近似名稱：${sim.join("、")}（未排除）` : "本期對帳檔沒有這位客戶";
      }
    }
    return { customer: u, matchedBy: by, billNames: [...names], hint };
  });
  const unmatchedUni = uniformStatus.filter((s) => !s.matchedBy);
  if (unmatchedUni.length)
    warnings.push(`有統編客戶 ${unmatchedUni.length} 位在對帳檔比對不到：${unmatchedUni.map((s) => `${s.customer.code}「${s.customer.name}」`).join("、")}`);
  for (const s of unmatchedUni.filter((s) => s.hint.includes("代號是"))) pending.push(`有統編客戶 ${s.customer.code}「${s.customer.name}」：${s.hint}，請確認`);

  for (const t of triplicates) {
    if (t.month === null) warnings.push(`三聯發票 ${t.source}「${t.name}」沒有可辨識的日期，無法判斷月份，已算進第一個月（${months[0]}月）`);
    else if (!months.includes(t.month)) warnings.push(`三聯發票 ${t.source}「${t.name}」日期在 ${t.month} 月，不屬於本期（${months.join("、")}月），未計入三聯合計`);
  }
  const triMonth = (t: Triplicate) => t.month ?? months[0]!;
  const listSum = (ms: number[]) => triplicates.filter((t) => ms.includes(triMonth(t))).reduce((a, t) => a + t.amount, 0);
  const triSum = (ms: number[]): number => {
    if (opts.monthTriplicate) {
      return ms.reduce((a, m) => {
        const v = opts.monthTriplicate!.get(m);
        if (v === undefined) throw new Error(`缺少 ${m} 月的三聯總額（--month-triplicate ${m}=金額）`);
        return a + v;
      }, 0);
    }
    if (opts.triplicateTotal !== undefined) {
      if (opts.mode === "monthly") throw new Error("每月模式請用 --month-triplicate 給每月三聯總額，不能用 --triplicate-total");
      return opts.triplicateTotal;
    }
    return listSum(ms);
  };

  const selected: Entry[] = [];
  const calcs: MonthCalc[] = [];
  const groups: { label: string; months: number[]; total: number }[] = [];
  if (opts.mode === "period") {
    if (opts.total === undefined) throw new Error("整期模式需要 --total 總額");
    groups.push({ label: "整期", months, total: opts.total });
  } else {
    for (const m of months) {
      const t = opts.monthTotals?.get(m);
      if (t === undefined) throw new Error(`每月模式缺少 ${m} 月的總額（--month-total ${m}=金額）`);
      groups.push({ label: `${m}月`, months: [m], total: t });
    }
  }
  for (const g of groups) {
    const tri = triSum(g.months);
    const gap = g.total - tri;
    const pool = candidates.filter((c) => g.months.includes(c.month));
    const amounts = pool.map((c) => c.amount);
    const fmt = (n: number) => n.toLocaleString();
    const halfKeys: { month: number; half: Half }[] = g.months.flatMap((m) => (["上期", "下期"] as Half[]).map((half) => ({ month: m, half })));
    let picked: Entry[] = [];
    let noSolution: MonthCalc["noSolution"];
    let countBound: CountBound | undefined;
    const overMin = opts.overMin ?? 1;
    if (gap < 0) {
      warnings.push(`${g.label}：三聯合計 ${tri} 已超過總額 ${g.total}（超過 ${-gap}），不開二聯`);
      if (opts.overMax !== undefined) pending.push(`${g.label}：三聯合計 ${fmt(tri)} 已超過總額 ${fmt(g.total)}，無法套用超出規則，未開二聯`);
    } else if (opts.overMax === undefined) {
      picked = subsetSumMax(amounts, gap).picked.map((i) => pool[i]!);
    } else {
      if (overMin < 1 || overMin > opts.overMax) throw new Error(`超出範圍不合理：--over-min ${overMin}、--over-max ${opts.overMax}`);
      const groupsPool = halfKeys.map((k) => pool.filter((c) => c.month === k.month && c.half === k.half));
      const r = selectBalancedOver(groupsPool.map((p) => p.map((c) => c.amount)), gap + overMin, gap + opts.overMax, tri, opts.maxPerHalf ?? Infinity, opts.prefer ?? "many");
      countBound = r.bound;
      if (r.ok) {
        picked = r.picks.flatMap((idx, h) => idx.map((i) => groupsPool[h]![i]!));
      } else {
        // 範圍內無解：不開二聯，列出最接近的兩個方案（不考慮分散）
        const alt = subsetSumOver(amounts, gap, opts.overMax, tri, overMin);
        const b = !alt.ok ? alt.below : null, ab = !alt.ok ? alt.above : null;
        noSolution = { below: b && { sum: b.sum, count: b.picked.length }, above: ab && { sum: ab.sum, count: ab.picked.length } };
        const desc = (x: { sum: number; count: number } | null, word: string) =>
          x ? `${word}：二聯 ${x.count} 張合計 ${fmt(x.sum)}（${x.sum - gap >= 0 ? "超出" : "不足"} ${fmt(Math.abs(x.sum - gap))}，二聯＋三聯 ${fmt(x.sum + tri)}）` : `${word}：無`;
        const why = alt.ok ? `（不限張數分佈時可湊到 ${fmt(alt.sum)}，但${opts.maxPerHalf !== undefined && opts.maxPerHalf !== Infinity ? `每期最多 ${opts.maxPerHalf} 張的限制下` : ""}無解）` : "";
        const msg = `${g.label}：找不到超出差額 ${fmt(overMin)}～${fmt(opts.overMax)} 元（二聯合計 ${fmt(gap + overMin)}～${fmt(gap + opts.overMax)}）且二聯＋三聯合計尾數不是 00 的組合${why}，未開二聯。${alt.ok ? "" : `最接近的方案 — ${desc(noSolution.below, "低於範圍最接近")}；${desc(noSolution.above, "高於範圍最接近")}。`}可調整 --over-min／--over-max 或總額後重跑`;
        pending.push(msg);
        warnings.push(msg);
      }
    }
    selected.push(...picked);
    const sum = picked.reduce((a, e) => a + e.amount, 0);
    calcs.push({
      label: g.label,
      months: g.months,
      total: g.total,
      triplicateSum: tri,
      gap,
      selectedSum: sum,
      remaining: gap - sum,
      grandTotal: sum + tri,
      overMax: opts.overMax,
      overMin: opts.overMax !== undefined ? overMin : undefined,
      over: sum - gap,
      noSolution,
      countBound,
      halfStats: halfKeys.map((k) => {
        const xs = picked.filter((e) => e.month === k.month && e.half === k.half);
        return { ...k, count: xs.length, sum: xs.reduce((a, e) => a + e.amount, 0) };
      }),
    });
  }
  const selSet = new Set(selected);
  const unselected = candidates.filter((c) => !selSet.has(c));
  const order = (a: Entry, b: Entry) => a.month - b.month || (a.half === b.half ? 0 : a.half === "上期" ? -1 : 1);
  selected.sort(order);
  unselected.sort(order);

  return { months, entries, candidates, selected, unselected, excluded, belowThreshold, calcs, triplicateCustomers, fixedExcludeStatus, uniformStatus, pending, warnings };
}

/** 讀固定排除名單：一行一個客戶名，# 開頭為註解 */
export function parseExcludeList(text: string): string[] {
  return text
    .split(/\r?\n/)
    .map((l) => cleanName(l.replace(/#.*$/, "")))
    .filter(Boolean);
}
