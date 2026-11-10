// 移植自 invoice-helper/src/webCalc.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 網頁版計算：輸入是檔案內容（ArrayBuffer）＋網頁上打的數字，不碰檔案系統。
// 在瀏覽器的 Web Worker 裡跑（src/invoice/worker.ts），bun test 也能直接呼叫。
// 規則與命令列版共用 core.ts、select.ts，不另寫一套。
import ExcelJS from "exceljs";
import { parseWorkbook, type BillFileParse } from "./parseBills.ts";
import { parseCustomerSourceWorkbook, mergeUniformCustomers, type CustomerSourceParse } from "./parseCustomers.ts";
import { parseExcludeList, type Options, type Prefer, type SelectionResult } from "./select.ts";
import { runSelection, buildSettings, inferRocYear, periodLabelOf, modeLabelOf } from "./core.ts";
import { buildReportWorkbook, halfVsBlocks, boundNote } from "./writeReport.ts";

export type WebInput = {
  bills: { name: string; month: number; data: ArrayBuffer }[];
  sources: { name: string; data: ArrayBuffer }[];
  /** 每月逐張三聯金額（含稅） */
  triplicate: { month: number; amounts: number[] }[];
  mode: "period" | "monthly";
  total?: number;
  monthTotals?: [number, number][];
  threshold: number;
  overMin: number;
  overMax: number;
  prefer: Prefer;
  /** 每個半月期最多幾張（使用者工作簿每欄可填格數）；undefined＝不限 */
  maxPerHalf?: number;
  /** 民國年；undefined＝自動 */
  year?: number;
  /** config/exclude.txt 內容 */
  excludeText: string;
};

export type HalfRead = { label: string; customers: number; customerSum: number; parsedTotal: number; compare: string };

export type WebOutput = {
  rocYear: number;
  months: number[];
  periodLabel: string;
  modeLabel: string;
  result: SelectionResult;
  pending: string[];
  warnings: string[];
  triplicateSource: string;
  halvesRead: HalfRead[];
  uniformLine: string | null;
  boundNotes: string[];
  reportFileName: string;
  report: ArrayBuffer;
};

async function load(data: ArrayBuffer, name: string): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  try {
    await wb.xlsx.load(data);
  } catch (e) {
    throw new Error(`「${name}」打不開，請確認是 Excel 檔（.xlsx）。${e instanceof Error ? `（${e.message}）` : ""}`);
  }
  return wb;
}

export async function computeFromBuffers(input: WebInput): Promise<WebOutput> {
  if (input.bills.length !== 2) throw new Error(`請選兩個月的對帳檔（目前 ${input.bills.length} 個）`);
  const bills: BillFileParse[] = [];
  for (const b of input.bills) {
    if (!(b.month >= 1 && b.month <= 12)) throw new Error(`「${b.name}」沒有指定月份`);
    bills.push(parseWorkbook(await load(b.data, b.name), b.name, b.month));
  }
  bills.sort((a, b) => a.month - b.month);
  const months = bills.map((b) => b.month);
  if (new Set(months).size !== months.length) throw new Error(`兩個對帳檔的月份相同（${months.join("、")}月），請檢查月份`);
  const rocYear = inferRocYear(months, input.year);

  const warnings: string[] = [];
  const pending: string[] = [];

  // 三聯：網頁逐張輸入的金額，每月一個總額
  const monthTriplicate = new Map<number, number>();
  for (const m of months) {
    const t = input.triplicate.find((x) => x.month === m);
    monthTriplicate.set(m, Math.round((t?.amounts ?? []).reduce((a, v) => a + v, 0)));
  }
  const triCount = (m: number) => input.triplicate.find((x) => x.month === m)?.amounts.length ?? 0;
  const triplicateSource = `網頁逐張輸入：${months.map((m) => `${m}月 ${triCount(m)} 張 ${monthTriplicate.get(m)!.toLocaleString()}`).join("、")}`;
  if (months.every((m) => triCount(m) === 0)) pending.push("三聯金額一張都沒有輸入，三聯合計視為 0");

  // 結帳原始檔 → 有統編客戶
  const sources: CustomerSourceParse[] = [];
  let uniformCustomers: import("./parseCustomers.ts").UniformCustomer[] | undefined;
  if (input.sources.length) {
    for (const s of input.sources) {
      const p = parseCustomerSourceWorkbook(await load(s.data, s.name), s.name);
      sources.push(p);
      warnings.push(...p.warnings);
    }
    const merged = mergeUniformCustomers(sources);
    uniformCustomers = merged.customers;
    warnings.push(...merged.warnings);
  } else {
    pending.push("沒有選結帳原始檔，不會排除有統編的客戶");
  }

  const fixedExclude = parseExcludeList(input.excludeText);
  const opts: Options = {
    mode: input.mode,
    total: input.mode === "period" ? input.total : undefined,
    monthTotals: new Map(input.monthTotals ?? []),
    threshold: input.threshold,
    deductMode: "exclude",
    includeRetail: false,
    fixedExclude,
    uniformCustomers,
    monthTriplicate,
    overMax: input.overMax,
    overMin: input.overMin,
    prefer: input.prefer,
    maxPerHalf: input.maxPerHalf,
  };
  if (opts.mode === "period" && !(opts.total! > 0)) throw new Error("請輸入總額");
  if (opts.mode === "monthly") for (const m of months) if (!(opts.monthTotals!.get(m)! > 0)) throw new Error(`請輸入 ${m} 月的總額`);
  if (!(input.overMin >= 1 && input.overMax >= input.overMin)) throw new Error(`超出範圍不合理：${input.overMin}～${input.overMax} 元`);

  const result = runSelection(bills, [], opts, warnings, pending);

  const periodLabel = periodLabelOf(rocYear, months);
  const modeLabel = modeLabelOf(opts.mode);
  const report = buildReportWorkbook({
    periodLabel,
    modeLabel,
    bills,
    triplicates: [],
    triplicateMode: "total",
    triplicateSource,
    sources,
    result,
    pending,
    warnings,
    settings: buildSettings({
      opts,
      workbookDesc: `（網頁版，不使用工作簿${opts.maxPerHalf ? `；每期最多 ${opts.maxPerHalf} 張` : ""}）`,
      bills,
      triplicateDesc: `每月總額 ${[...monthTriplicate].map(([m, v]) => `${m}月 ${v.toLocaleString()}`).join("、")}（${triplicateSource}）`,
      sources,
      uniformCount: uniformCustomers?.length ?? 0,
    }),
  });
  const buf = await report.xlsx.writeBuffer();
  const reportBuf = buf instanceof ArrayBuffer ? buf : new Uint8Array(buf as ArrayLike<number>).slice().buffer;

  const halvesRead: HalfRead[] = bills.flatMap((b) =>
    b.halves.map((h) => ({
      label: `${b.month}月${h.half}`,
      customers: h.customers.length,
      customerSum: h.customers.reduce((a, x) => a + x.amount, 0),
      parsedTotal: h.parsedTotal,
      compare: h.fileGrandTotal
        ? `表尾總計 ${h.fileGrandTotal.value.toLocaleString()}${h.parsedTotal === h.fileGrandTotal.value ? " 一致" : ` 差 ${(h.parsedTotal - h.fileGrandTotal.value).toLocaleString()}`}`
        : halfVsBlocks(h).replace(/^解析 [\d,]+／/, ""),
    })),
  );
  let uniformLine: string | null = null;
  if (uniformCustomers) {
    const um = result.uniformStatus.filter((s) => s.matchedBy);
    const un = result.uniformStatus.filter((s) => !s.matchedBy);
    uniformLine = `有統編客戶 ${uniformCustomers.length} 位：對帳檔比對到 ${um.length} 位，比對不到 ${un.length} 位${un.length ? `（${un.map((s) => s.customer.name).join("、")}，通常是本期沒叫貨）` : ""}`;
  }
  const boundNotes = result.calcs.filter((c) => c.countBound).map((c) => `${result.calcs.length > 1 ? c.label + "：" : ""}${boundNote(c.countBound!, c.halfStats.reduce((a, h) => a + h.count, 0), c.halfStats.length)}`);

  return {
    rocYear,
    months,
    periodLabel,
    modeLabel,
    result,
    pending,
    warnings,
    triplicateSource,
    halvesRead,
    uniformLine,
    boundNotes,
    reportFileName: `二聯客戶清單_${periodLabel}.xlsx`,
    report: reportBuf,
  };
}
