// 移植自 invoice-helper/src/core.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 命令列版與網頁版共用的計算流程（不碰檔案系統，瀏覽器也能跑）：
// 對帳檔警告／需人工確認 → select() 挑選 → 摘要設定列。
import type { BillFileParse } from "./parseBills.ts";
import type { Triplicate } from "./parseTriplicate.ts";
import type { CustomerSourceParse } from "./parseCustomers.ts";
import { select, type Options, type SelectionResult } from "./select.ts";

/** 民國年：有給就用；否則今年，若月份比現在晚則用去年 */
export function inferRocYear(months: number[], given: number | undefined, now = new Date()): number {
  if (given !== undefined) return given;
  let y = now.getFullYear() - 1911;
  if (Math.max(...months) > now.getMonth() + 1) y -= 1;
  return y;
}

export const periodLabelOf = (rocYear: number, months: number[]) => `${rocYear}年${months.map((m) => String(m).padStart(2, "0")).join("-")}月`;
export const modeLabelOf = (mode: Options["mode"]) => (mode === "period" ? "整期一個總額" : "每月各一個總額");

/**
 * 對帳檔的警告與需人工確認事項（附加到 warnings／pending），再跑 select()，
 * 最後把 select 的需人工確認事項接上。bills 需依月份排序。
 */
export function runSelection(bills: BillFileParse[], triplicates: Triplicate[], opts: Options, warnings: string[], pending: string[]): SelectionResult {
  for (const b of bills) {
    const fname = b.file.split(/[\\/]/).pop();
    if (b.skippedSheets.length) warnings.push(`${fname}：略過非上期／下期工作表 ${b.skippedSheets.map((s) => `「${s}」`).join("、")}`);
    warnings.push(...b.warnings.map((w) => `${fname}：${w}`));
    for (const h of b.halves) {
      for (const w of h.warnings) warnings.push(`${b.month}月 ${w}`);
      for (const n of h.notes) pending.push(`${b.month}月${h.half} ${n.address}「${n.name}」${n.amount ?? ""} 判斷為附註／合計列，未當客戶，請確認`);
      for (const u of h.unresolved) pending.push(`${b.month}月${h.half} ${u.address} 不明金額列「${u.name || "(無名稱)"}」${u.amount}，未計入，請確認`);
      const bad = h.checks.filter((c) => c.diff !== 0);
      if (bad.length) {
        const sum = bad.reduce((a, c) => a + c.diff, 0);
        pending.push(`${b.month}月${h.half}：檔案內 ${bad.length} 個「總計」與解析合計不符，合計差 ${sum.toLocaleString()}（${bad.map((c) => `${c.address}${c.formula ? "=" + c.formula : ""} 差 ${c.diff.toLocaleString()}`).join("；")}）。${bad.some((c) => c.note.includes("空白儲存格")) ? "總計公式引用了空白的實收金額欄，所以少算左半。" : ""}工具以逐筆客戶金額為準，請確認對帳檔總計公式`);
      }
    }
  }
  // 兩期資料完全相同 → 很可能複製後未更新
  const halves = bills.flatMap((b) => b.halves.map((h) => ({ key: `${b.month}月${h.half}`, sig: h.customers.map((c) => `${c.name}=${c.amount}`).join("|"), n: h.customers.length })));
  for (let i = 0; i < halves.length; i++)
    for (let j = i + 1; j < halves.length; j++)
      if (halves[i]!.n > 0 && halves[i]!.sig === halves[j]!.sig) pending.push(`${halves[i]!.key} 與 ${halves[j]!.key} 的客戶與金額完全相同，可能是複製後尚未更新`);

  const result = select(bills, triplicates, opts);
  warnings.push(...result.warnings);
  for (const f of result.fixedExcludeStatus.filter((f) => !f.matched)) pending.push(`固定排除名單「${f.name}」在對帳檔找不到${f.hints.length ? `（近似：${f.hints.join("、")}）` : ""}`);
  for (const t of result.triplicateCustomers.filter((t) => !t.matched)) pending.push(`三聯客戶「${t.name}」在對帳檔找不到同名客戶${t.hints.length ? `（近似：${t.hints.join("、")}）` : ""}`);
  pending.push(...result.pending);
  return result;
}

/** 客戶清單「摘要」工作表的「設定」區 */
export function buildSettings(p: {
  opts: Options;
  /** 「使用者工作簿」那一列的說明 */
  workbookDesc: string;
  bills: BillFileParse[];
  /** 「三聯」那一列的說明 */
  triplicateDesc: string;
  sources: CustomerSourceParse[];
  uniformCount: number;
}): [string, string][] {
  const { opts } = p;
  const prefer = opts.prefer ?? "many";
  return [
    ["總額模式", modeLabelOf(opts.mode)],
    ["挑選規則", opts.overMax !== undefined ? `二聯合計超出差額 ${(opts.overMin ?? 1).toLocaleString()}～${opts.overMax.toLocaleString()} 元；二聯＋三聯合計尾數不可為 00；先求各半月期張數最平均${opts.maxPerHalf ? `（每期最多 ${opts.maxPerHalf} 張）` : ""}，再求總張數${prefer === "many" ? "盡量多（偏小額客戶，--prefer many）" : "盡量少（偏大客戶，--prefer few）"}，再取最接近範圍中間、個位數不是 0 的合計` : "最接近差額但不超過"],
    ["使用者工作簿", p.workbookDesc],
    ["候選門檻", `單期帳單金額 > ${opts.threshold.toLocaleString()}`],
    ["扣代送客戶（石先生）", { exclude: "不列候選", net: "用淨額", gross: "用原始金額" }[opts.deductMode]],
    ["零售列", opts.includeRetail ? "列入候選" : "不列入"],
    ["固定排除名單", opts.fixedExclude.join("、") || "（無）"],
    ["對帳檔", p.bills.map((b) => b.file.split(/[\\/]/).pop()).join("、")],
    ["三聯", p.triplicateDesc],
    ["結帳原始檔（統編）", p.sources.length ? p.sources.map((s) => `${s.file}（客戶 ${s.customerCount}，有統編 ${s.uniformCustomers.length}）`).join("、") + `；聯集有統編 ${p.uniformCount} 位` : "（未提供）"],
  ];
}
