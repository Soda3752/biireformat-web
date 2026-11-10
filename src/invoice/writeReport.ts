// 移植自 invoice-helper/src/writeReport.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 產出 B：客戶清單 Excel（二聯清單、摘要、未選上候選、排除名單、對帳檔核對、警告）
import ExcelJS from "exceljs";
import type { BillFileParse } from "./parseBills.ts";
import type { SelectionResult } from "./select.ts";
import type { Triplicate } from "./parseTriplicate.ts";
import type { CustomerSourceParse } from "./parseCustomers.ts";

const NUM = "#,##0";

/** 沒有整期總計的期別：解析合計 vs 檔案各區總計加總，並拆解差額 */
export function halfVsBlocks(h: import("./parseBills.ts").HalfParse): string {
  const diff = h.parsedTotal - h.fileBlockTotalSum;
  const blockDiff = h.checks.filter((c) => c.kind === "區塊總計").reduce((a, c) => a + c.diff, 0);
  const ded = h.deducts.reduce((a, x) => a + x.net, 0);
  const parts: string[] = [];
  if (blockDiff) parts.push(`各區總計公式少算 ${blockDiff.toLocaleString()}`);
  if (ded) parts.push(`表尾扣代送淨額 ${ded.toLocaleString()} 不在各區總計內`);
  const other = diff - blockDiff - ded;
  if (other) parts.push(`其他 ${other.toLocaleString()}`);
  return `解析 ${h.parsedTotal.toLocaleString()}／檔案無整期總計，各區總計加總 ${h.fileBlockTotalSum.toLocaleString()}，差 ${diff.toLocaleString()}${parts.length ? `（${parts.join("、")}）` : ""}`;
}

function table(ws: ExcelJS.Worksheet, startRow: number, headers: string[], rows: (string | number | null)[][], numCols: number[] = []): number {
  const hr = ws.getRow(startRow);
  headers.forEach((h, i) => {
    const c = hr.getCell(i + 1);
    c.value = h;
    c.font = { bold: true };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDDEBF7" } };
    c.border = { bottom: { style: "thin" } };
  });
  rows.forEach((r, j) => {
    const row = ws.getRow(startRow + 1 + j);
    r.forEach((v, i) => {
      const c = row.getCell(i + 1);
      c.value = v;
      if (numCols.includes(i)) c.numFmt = NUM;
    });
  });
  return startRow + 1 + rows.length;
}

function section(ws: ExcelJS.Worksheet, row: number, title: string): number {
  const c = ws.getCell(row, 1);
  c.value = title;
  c.font = { bold: true, size: 13 };
  return row + 1;
}

export type ReportParams = {
  periodLabel: string;
  modeLabel: string;
  bills: BillFileParse[];
  triplicates: Triplicate[];
  /** list＝三聯明細清單；total＝只給三聯總額 */
  triplicateMode: "list" | "total";
  /** 三聯總額來源說明（例如「使用者工作簿 115.9.10 S38」） */
  triplicateSource?: string;
  /** 結帳原始檔解析結果（統編） */
  sources: CustomerSourceParse[];
  result: SelectionResult;
  pending: string[];
  warnings: string[];
  settings: [string, string][];
};

/** 產出 B 的工作簿（不寫檔；網頁版用 writeBuffer 下載） */
export function buildReportWorkbook(params: ReportParams): ExcelJS.Workbook {
  const { result: R, bills } = params;
  const wb = new ExcelJS.Workbook();

  // 1) 二聯清單
  const ws1 = wb.addWorksheet("二聯清單", { views: [{ state: "frozen", ySplit: 1 }] });
  const end1 = table(
    ws1,
    1,
    ["月份", "期別", "客戶代號", "客戶名稱", "二聯金額"],
    R.selected.map((e) => [`${e.month}月`, e.half, e.code || null, e.name, e.amount]),
    [4],
  );
  ws1.getCell(end1 + 1, 4).value = "合計";
  ws1.getCell(end1 + 1, 4).font = { bold: true };
  const s = ws1.getCell(end1 + 1, 5);
  s.value = { formula: `SUM(E2:E${end1 - 1})`, result: R.selected.reduce((a, e) => a + e.amount, 0) };
  s.numFmt = NUM;
  s.font = { bold: true };
  [8, 8, 10, 18, 12].forEach((w, i) => (ws1.getColumn(i + 1).width = w));

  // 2) 摘要
  const ws2 = wb.addWorksheet("摘要");
  [26, 16, 16, 16, 16, 40].forEach((w, i) => (ws2.getColumn(i + 1).width = w));
  let r = section(ws2, 1, `二聯式發票補額計算摘要（${params.periodLabel}，${params.modeLabel}）`);
  r++;
  for (const c of R.calcs) {
    if (R.calcs.length > 1) r = section(ws2, r, c.label);
    const triLabel = params.triplicateMode === "total" ? `三聯總額（含稅${params.triplicateSource ? `，${params.triplicateSource}` : "，使用者提供"}）` : "三聯合計（含稅金額）";
    const lines: [string, number, string?][] = [
      ["總額（含稅）", c.total],
      [triLabel, c.triplicateSum],
      ["差額（需用二聯補）", c.gap],
      ["二聯合計", c.selectedSum],
    ];
    if (c.overMax !== undefined) {
      const bad = c.noSolution !== undefined || c.over < (c.overMin ?? 1) || c.over > c.overMax;
      lines.push(["超出差額（元）", c.over, bad ? "bad" : `規則：超出 ${(c.overMin ?? 1).toLocaleString()}～${c.overMax.toLocaleString()} 元`]);
    } else {
      lines.push(["差額剩餘（未補足）", c.remaining, c.remaining > 0 ? "bad" : undefined]);
    }
    lines.push(["二聯＋三聯合計", c.grandTotal, c.overMax !== undefined && c.grandTotal % 100 === 0 ? "bad" : c.overMax !== undefined ? "尾數不是 00" : undefined]);
    for (const [k, v, note] of lines) {
      ws2.getCell(r, 1).value = k;
      const cell = ws2.getCell(r, 2);
      cell.value = v;
      cell.numFmt = NUM;
      if (k.startsWith("差額剩餘") || k.startsWith("超出差額") || k.startsWith("二聯＋三聯")) cell.font = { bold: true, color: { argb: note === "bad" ? "FFC00000" : "FF000000" } };
      if (note && note !== "bad") ws2.getCell(r, 3).value = note;
      r++;
    }
    for (const h of c.halfStats) {
      ws2.getCell(r, 1).value = `　${h.month}月${h.half}二聯`;
      ws2.getCell(r, 2).value = h.sum;
      ws2.getCell(r, 2).numFmt = NUM;
      ws2.getCell(r, 3).value = `${h.count} 張`;
      r++;
    }
    if (c.countBound) {
      const n = c.halfStats.reduce((a, h) => a + h.count, 0);
      ws2.getCell(r, 1).value = "張數上限（說明）";
      ws2.getCell(r, 3).value = boundNote(c.countBound, n, c.halfStats.length);
      r++;
    }
    if (c.noSolution) {
      const d = (x: { sum: number; count: number } | null) => (x ? `二聯 ${x.count} 張 ${x.sum.toLocaleString()}（${x.sum >= c.gap ? "超出" : "不足"} ${Math.abs(x.sum - c.gap).toLocaleString()}；二聯＋三聯 ${(x.sum + c.triplicateSum).toLocaleString()}）` : "無");
      ws2.getCell(r, 1).value = "⚠ 範圍內無解，未開二聯";
      ws2.getCell(r, 1).font = { bold: true, color: { argb: "FFC00000" } };
      r++;
      ws2.getCell(r, 1).value = "低於範圍最接近";
      ws2.getCell(r++, 2).value = d(c.noSolution.below);
      ws2.getCell(r, 1).value = "高於範圍最接近";
      ws2.getCell(r++, 2).value = d(c.noSolution.above);
    }
    r++;
  }
  ws2.getCell(r, 1).value = "二聯張數";
  ws2.getCell(r, 2).value = R.selected.length;
  r++;
  ws2.getCell(r, 1).value = "候選數（單期 > 門檻、未排除）";
  ws2.getCell(r, 2).value = R.candidates.length;
  r += 2;

  r = section(ws2, r, "需人工確認");
  if (params.pending.length === 0) ws2.getCell(r++, 1).value = "（無）";
  for (const p of params.pending) ws2.getCell(r++, 1).value = "・" + p;
  r++;

  r = section(ws2, r, "各期讀到的客戶數與合計");
  r = table(
    ws2,
    r,
    ["期別", "客戶數", "客戶金額合計", "零售合計", "扣代送淨額", "解析合計 vs 檔案總計"],
    bills.flatMap((b) =>
      b.halves.map((h) => {
        const cust = h.customers.reduce((a, x) => a + x.amount, 0);
        const ret = h.retail.reduce((a, x) => a + x.amount, 0);
        const ded = h.deducts.reduce((a, x) => a + x.net, 0);
        let cmp: string;
        if (h.fileGrandTotal) {
          const d = h.parsedTotal - h.fileGrandTotal.value;
          cmp = `解析 ${h.parsedTotal.toLocaleString()}／表尾總計 ${h.fileGrandTotal.value.toLocaleString()}（${h.fileGrandTotal.address}）${d === 0 ? "一致" : `差 ${d.toLocaleString()}`}`;
        } else {
          cmp = halfVsBlocks(h);
        }
        return [`${b.month}月${h.half}`, h.customers.length, cust, ret, ded, cmp];
      }),
    ),
    [2, 3, 4],
  );
  r++;

  r = section(ws2, r, "設定");
  for (const [k, v] of params.settings) {
    ws2.getCell(r, 1).value = k;
    ws2.getCell(r, 2).value = v;
    r++;
  }
  r++;
  r = section(ws2, r, "警告");
  if (!params.warnings.length) ws2.getCell(r++, 1).value = "（無）";
  for (const w of params.warnings) ws2.getCell(r++, 1).value = "・" + w;

  // 3) 未選上候選
  const ws3 = wb.addWorksheet("未選上候選");
  table(
    ws3,
    1,
    ["月份", "期別", "客戶代號", "客戶名稱", "帳單金額"],
    R.unselected.map((e) => [`${e.month}月`, e.half, e.code || null, e.name, e.amount]),
    [4],
  );
  [8, 8, 10, 18, 12].forEach((w, i) => (ws3.getColumn(i + 1).width = w));

  // 4) 排除名單
  const ws4 = wb.addWorksheet("排除名單");
  [10, 8, 12, 18, 12, 14, 12, 22, 40].forEach((w, i) => (ws4.getColumn(i + 1).width = w));
  let r4 = section(ws4, 1, "被排除的客戶（各期帳單）");
  r4 = table(
    ws4,
    r4,
    ["月份", "期別", "客戶代號", "客戶名稱", "帳單金額", "排除原因", "統一編號", "結帳檔客戶編號／名稱", "備註"],
    R.excluded.map((x) => [
      `${x.entry.month}月`,
      x.entry.half,
      x.entry.code || null,
      x.entry.name,
      x.entry.amount,
      x.reason,
      x.uniform?.uniform ?? null,
      x.uniform ? `${x.uniform.code} ${x.uniform.name}` : null,
      x.entry.note || null,
    ]),
    [4],
  );
  r4++;
  if (params.sources.length) {
    r4 = section(ws4, r4, `有統編客戶比對（結帳原始檔：${params.sources.map((s) => s.file).join("、")}）`);
    r4 = table(
      ws4,
      r4,
      ["客戶編號", "原始檔名稱", "統一編號", "比對方式", "對帳檔名稱", "來源檔", "說明"],
      R.uniformStatus.map((u) => [
        u.customer.code || null,
        u.customer.name,
        u.customer.uniform,
        u.matchedBy ?? "找不到",
        u.billNames.join("、") || null,
        u.customer.sources.join("、"),
        u.hint || null,
      ]),
    );
    r4++;
  }
  if (params.triplicateMode === "list") {
    r4 = section(ws4, r4, "三聯客戶名稱比對");
    r4 = table(
      ws4,
      r4,
      ["三聯客戶", "對帳檔有同名？", "近似名稱"],
      R.triplicateCustomers.map((t) => [t.name, t.matched ? "有" : "找不到", t.hints.join("、") || null]),
    );
    r4++;
  }
  r4 = section(ws4, r4, "固定排除名單（config/exclude.txt）比對");
  table(
    ws4,
    r4,
    ["名單客戶", "對帳檔有同名？", "近似名稱"],
    R.fixedExcludeStatus.map((t) => [t.name, t.matched ? "有" : "找不到", t.hints.join("、") || null]),
  );

  // 5) 三聯清單（讀到的內容）；只給總額時不產生這頁
  if (params.triplicateMode === "list") {
  const ws5 = wb.addWorksheet("三聯讀取結果");
  table(
    ws5,
    1,
    ["月份", "日期", "客戶名稱", "(三)發票號碼", "銷售額", "營業稅", "金額", "來源"],
    params.triplicates.map((t) => [t.month ? `${t.month}月` : "?", t.date ? `${t.date.y - 1911}/${t.date.m}/${t.date.d}` : null, t.name, t.invoiceNo || null, t.sales, t.tax, t.amount, t.source]),
    [4, 5, 6],
  );
  [8, 12, 18, 14, 12, 10, 12, 24].forEach((w, i) => (ws5.getColumn(i + 1).width = w));
  }

  // 6) 對帳檔核對
  const ws6 = wb.addWorksheet("對帳檔核對");
  [12, 10, 10, 10, 18, 14, 14, 12, 60].forEach((w, i) => (ws6.getColumn(i + 1).width = w));
  table(
    ws6,
    1,
    ["期別", "類型", "標籤", "儲存格", "公式", "檔案值", "解析值", "差額", "說明"],
    bills.flatMap((b) =>
      b.halves.flatMap((h) => [
        ...h.checks.map((c) => [`${b.month}月${h.half}`, c.kind, c.label, c.address, c.formula, c.fileValue, c.parsedValue, c.diff, c.note || (c.diff === 0 ? "一致" : null)]),
        ...h.deducts.map((d) => [`${b.month}月${h.half}`, "扣代送列", d.name, d.addresses.join("、"), null, d.gross, d.net, null, `金額 ${d.gross ?? "?"} − 代送 ${d.deduction} = 淨額 ${d.net}`]),
        ...h.retail.map((x) => [`${b.month}月${h.half}`, "零售列", "零售", x.address, null, x.amount, null, null, "零售不當客戶"]),
      ]),
    ),
    [5, 6, 7],
  );

  return wb;
}

/** 張數上限說明：每張都 > 門檻，差額有限，所以張數有上界 */
export function boundNote(b: { total: number; equalPerHalf: number }, actual: number, halves: number): string {
  return `本次 ${actual} 張。理論上限（只用金額最小的候選塞到範圍上限）：不管分佈最多 ${b.total} 張；${halves} 期張數相同時每期最多 ${b.equalPerHalf} 張（共 ${b.equalPerHalf * halves} 張）。實際還受合計範圍與尾數限制，上限不一定湊得到`;
}
