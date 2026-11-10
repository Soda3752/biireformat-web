// 移植自 invoice-helper/src/parseBills.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 解析結帳網站匯出的月對帳檔（每檔「上期」「下期」兩張表，每表左右並排、上下多個區塊）。
// 原則：一律用表頭文字找欄位，不寫死欄位位置；合計列、區域小計、附註列都不當客戶。
import ExcelJS from "exceljs";
import { monthFromFilename } from "./dates.ts";
import { readCell, normHeader, isSumFormula, hasCellRef, colName, colNumber, type CellInfo } from "./cell.ts";

export { monthFromFilename };

export type Half = "上期" | "下期";

export type BillRecord = {
  half: Half;
  name: string;
  code: string;
  amount: number;
  address: string; // 金額所在儲存格，例如 C5
};

export type RetailRecord = { half: Half; amount: number; address: string };

/** 表尾「某客戶金額 − 代送 = 淨額」這種扣代送的組合（目前看到的是「石先生」） */
export type DeductGroup = {
  half: Half;
  name: string;
  code: string;
  gross: number | null;
  deduction: number;
  net: number;
  addresses: string[];
};

export type CheckRow = {
  half: Half;
  block: number; // 第幾個區塊（0 起算）
  kind: "小計公式" | "區塊總計" | "整期總計";
  label: string;
  address: string;
  formula: string;
  fileValue: number;
  parsedValue: number;
  diff: number; // 解析 − 檔案
  note: string;
};

export type HalfParse = {
  half: Half;
  sheetName: string;
  customers: BillRecord[];
  retail: RetailRecord[];
  deducts: DeductGroup[];
  /** 表尾不明的金額列（未計入任何地方，需人工確認） */
  unresolved: { address: string; name: string; amount: number; formula: string | null }[];
  /** 表尾附註列（例如「愛買+雲雀」），未計入 */
  notes: { address: string; name: string; amount: number | null; formula: string | null }[];
  checks: CheckRow[];
  /** 解析合計 = 客戶 + 零售 + 扣代送淨額 */
  parsedTotal: number;
  /** 檔案內整期總計（「總營收」或區域總表的「總計」）；沒有就 null */
  fileGrandTotal: { value: number; address: string; formula: string } | null;
  /** 檔案各區「總計」列加總（每區取它的總計；沒有合併總計的區塊取各側標示「總計」的 SUM） */
  fileBlockTotalSum: number;
  warnings: string[];
};

export type BillFileParse = {
  file: string;
  month: number;
  halves: HalfParse[];
  skippedSheets: string[];
  warnings: string[];
};

const TOTAL_RE = /總計|小計|合計|總營收|總額/;
const RETAIL_RE = /零售/;
const DEDUCT_RE = /代送/;

type Side = { nameCol: number; amountCol: number; codeCol: number | null };

type Rec = {
  row: number;
  name: string;
  code: string;
  amt: CellInfo;
  kind: "customer" | "retail" | "deduction" | "subtotal" | "total" | "trailing" | "netRow" | "unnamedData";
};

function sideOf(ws: ExcelJS.Worksheet, headerRow: number, nameCols: number[], idx: number): Side | null {
  const nameCol = nameCols[idx]!;
  const nextName = nameCols[idx + 1] ?? ws.columnCount + 1;
  const prevName = idx > 0 ? nameCols[idx - 1]! : 0;
  let amountCol: number | null = null;
  for (let c = nameCol + 1; c < nextName; c++) {
    if (normHeader(readCell(ws.getRow(headerRow).getCell(c)).text) === "金額") {
      amountCol = c;
      break;
    }
  }
  if (amountCol === null) return null;
  let codeCol: number | null = null;
  // 代號欄：在本塊客戶名稱左邊、上一塊的金額欄之後
  let prevAmount = prevName;
  if (idx > 0) {
    for (let c = prevName + 1; c < nameCol; c++) {
      if (normHeader(readCell(ws.getRow(headerRow).getCell(c)).text) === "金額") {
        prevAmount = c;
        break;
      }
    }
  }
  for (let c = nameCol - 1; c > prevAmount; c--) {
    const h = normHeader(readCell(ws.getRow(headerRow).getCell(c)).text);
    if (h === "代號" || h === "客戶代號" || h === "編號") {
      codeCol = c;
      break;
    }
  }
  return { nameCol, amountCol, codeCol };
}

function codeText(c: CellInfo): string {
  if (c.empty) return "";
  if (c.text) return c.text;
  if (c.num !== null) return String(c.num);
  return "";
}

/** 解析 SUM(C3:C33) → 欄 C、列 3..33；不是單欄範圍回傳 null */
function parseSumRange(f: string): { col: number; from: number; to: number } | null {
  const m = /^\s*SUM\s*\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)\s*$/i.exec(f);
  if (!m || m[1]!.toUpperCase() !== m[3]!.toUpperCase()) return null;
  return { col: colNumber(m[1]!), from: Number(m[2]), to: Number(m[4]) };
}

export function parseSheet(ws: ExcelJS.Worksheet, half: Half): HalfParse {
  const out: HalfParse = {
    half,
    sheetName: ws.name,
    customers: [],
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
  const lastRow = ws.rowCount;
  const lastCol = ws.columnCount;

  // 1) 找表頭列（含「客戶名稱」的列）
  const headers: { row: number; nameCols: number[] }[] = [];
  for (let r = 1; r <= lastRow; r++) {
    const nameCols: number[] = [];
    for (let c = 1; c <= lastCol; c++) {
      if (normHeader(readCell(ws.getRow(r).getCell(c)).text) === "客戶名稱") nameCols.push(c);
    }
    if (nameCols.length) headers.push({ row: r, nameCols });
  }
  if (!headers.length) {
    out.warnings.push(`工作表「${ws.name}」找不到「客戶名稱」表頭，整張略過`);
    return out;
  }

  // 2) 逐區塊、逐側讀取
  for (let h = 0; h < headers.length; h++) {
    const { row: hr, nameCols } = headers[h]!;
    const endRow = h + 1 < headers.length ? headers[h + 1]!.row - 1 : lastRow;
    const blockRecs: { side: Side; recs: Rec[] }[] = [];

    for (let s = 0; s < nameCols.length; s++) {
      const side = sideOf(ws, hr, nameCols, s);
      if (!side) {
        out.warnings.push(`${ws.name}!${colName(nameCols[s]!)}${hr} 的「客戶名稱」旁找不到「金額」欄，該塊略過`);
        continue;
      }
      const recs: Rec[] = [];
      let state: "data" | "afterTotal" | "summary" = "data";
      let lastKind: Rec["kind"] | null = null;

      for (let r = hr + 1; r <= endRow; r++) {
        const row = ws.getRow(r);
        const nameCell = readCell(row.getCell(side.nameCol));
        const amt = readCell(row.getCell(side.amountCol));
        const code = side.codeCol ? codeText(readCell(row.getCell(side.codeCol))) : "";
        const name = nameCell.text && nameCell.num === null ? nameCell.text : "";
        const addr = `${colName(side.amountCol)}${r}`;
        if (!name && amt.empty) continue;

        if (state === "summary") {
          if (name && TOTAL_RE.test(name) && amt.num !== null) {
            out.fileGrandTotal = { value: amt.num, address: addr, formula: amt.formula ?? "" };
          } else if (name && /[+＋]/.test(name)) {
            out.notes.push({ address: `${colName(side.nameCol)}${r}`, name, amount: amt.num, formula: amt.formula });
          }
          continue;
        }

        const push = (kind: Rec["kind"]) => {
          recs.push({ row: r, name, code, amt, kind });
          lastKind = kind;
        };

        if (name && TOTAL_RE.test(name)) {
          push("total");
          state = "afterTotal";
          continue;
        }
        if (isSumFormula(amt.formula)) {
          push("subtotal");
          state = "afterTotal";
          continue;
        }
        if (name && RETAIL_RE.test(name)) {
          if (amt.num !== null) push("retail");
          continue;
        }
        if (name && DEDUCT_RE.test(name)) {
          if (amt.num !== null) push("deduction");
          continue;
        }

        if (state === "data") {
          if (amt.num === null) continue; // 有名字沒金額（或金額是文字）→ 略過
          if (name) push("customer");
          else push("unnamedData");
          continue;
        }

        // state === "afterTotal"：區塊合計之後的表尾
        if (!amt.empty && amt.num === null) {
          // 金額欄出現文字（例如「115年/10月」「上」）→ 進入區域總表
          state = "summary";
          continue;
        }
        if (amt.num === null) continue;
        if (hasCellRef(amt.formula) && !isSumFormula(amt.formula)) {
          if (lastKind === "deduction") {
            push("netRow");
            continue;
          }
          if (name) {
            // 有名字、金額是引用其他格 → 區域小計表（彰化 =J35 …）
            state = "summary";
            continue;
          }
        }
        push("trailing");
      }
      blockRecs.push({ side, recs });
    }

    // 3) 收集結果
    let blockParsed = 0;
    const sideParsed = new Map<number, Rec[]>();
    for (const { side, recs } of blockRecs) {
      const counted: Rec[] = [];
      for (const rec of recs) {
        const addr = `${colName(side.amountCol)}${rec.row}`;
        if (rec.kind === "customer") {
          out.customers.push({ half, name: rec.name, code: rec.code, amount: rec.amt.num!, address: addr });
          counted.push(rec);
        } else if (rec.kind === "retail") {
          out.retail.push({ half, amount: rec.amt.num!, address: addr });
          counted.push(rec);
        } else if (rec.kind === "unnamedData") {
          out.unresolved.push({ address: addr, name: "", amount: rec.amt.num!, formula: rec.amt.formula });
          out.warnings.push(`${ws.name}!${addr} 有金額 ${rec.amt.num} 但沒有客戶名稱，未計入`);
          counted.push(rec); // 檔案的 SUM 會算到它，核對時一併計入
        }
      }
      sideParsed.set(side.amountCol, counted);
      blockParsed += counted.filter((x) => x.kind !== "unnamedData").reduce((a, x) => a + x.amt.num!, 0);

      // 扣代送組合與表尾不明列
      const trailing = recs.filter((x) => x.kind === "trailing" || x.kind === "deduction" || x.kind === "netRow");
      const used = new Set<Rec>();
      for (let i = 0; i < trailing.length; i++) {
        const d = trailing[i]!;
        if (d.kind !== "deduction") continue;
        const gross = i > 0 && trailing[i - 1]!.kind === "trailing" && !used.has(trailing[i - 1]!) ? trailing[i - 1]! : null;
        const netRow = trailing[i + 1]?.kind === "netRow" ? trailing[i + 1]! : null;
        const ded = d.amt.num!;
        const grossV = gross?.amt.num ?? null;
        const net = netRow?.amt.num ?? (grossV !== null ? grossV - ded : null);
        if (net === null) {
          out.warnings.push(`${ws.name}!${colName(side.amountCol)}${d.row} 有「${d.name}」${ded} 但找不到對應的客戶金額，未計入`);
          used.add(d);
          continue;
        }
        const owner = gross?.name || netRow?.name || "(未命名)";
        const code = gross?.code || netRow?.code || "";
        const g: DeductGroup = {
          half,
          name: owner,
          code,
          gross: grossV,
          deduction: ded,
          net,
          addresses: [gross, d, netRow].filter((x): x is Rec => !!x).map((x) => `${colName(side.amountCol)}${x.row}`),
        };
        if (grossV !== null && netRow && grossV - ded !== netRow.amt.num) {
          out.warnings.push(`${ws.name} ${owner}：金額 ${grossV} − 代送 ${ded} ≠ 淨額 ${netRow.amt.num}`);
        }
        out.deducts.push(g);
        if (gross) used.add(gross);
        used.add(d);
        if (netRow) used.add(netRow);
      }
      for (const t of trailing) {
        if (used.has(t)) continue;
        const addr = `${colName(side.amountCol)}${t.row}`;
        out.unresolved.push({ address: addr, name: t.name, amount: t.amt.num!, formula: t.amt.formula });
        out.warnings.push(`${ws.name}!${addr} 表尾有不明金額列「${t.name || "(無名稱)"}」${t.amt.num}，未計入，需人工確認`);
      }

      // 小計公式核對：SUM(範圍) 用解析結果重算
      for (const rec of recs) {
        if (!isSumFormula(rec.amt.formula) || rec.amt.num === null) continue;
        const rng = parseSumRange(rec.amt.formula!);
        if (!rng) continue;
        const list = sideParsed.get(rng.col) ?? [];
        const parsed = list.filter((x) => x.row >= rng.from && x.row <= rng.to).reduce((a, x) => a + x.amt.num!, 0);
        out.checks.push({
          half,
          block: h,
          kind: "小計公式",
          label: rec.name || "（小計）",
          address: `${colName(side.amountCol)}${rec.row}`,
          formula: rec.amt.formula!,
          fileValue: rec.amt.num,
          parsedValue: parsed,
          diff: parsed - rec.amt.num,
          note: parsed === rec.amt.num ? "" : "解析結果與檔案小計不符",
        });
      }
    }

    // 區塊「總計」列（非單純 SUM 的，例如 J35=J34+C34）→ 跟整個區塊的解析合計比
    for (const { side, recs } of blockRecs) {
      for (const rec of recs) {
        if (rec.kind !== "total" || rec.amt.num === null) continue;
        if (isSumFormula(rec.amt.formula)) continue; // 已在小計公式核對
        const addr = `${colName(side.amountCol)}${rec.row}`;
        const f = rec.amt.formula ?? "";
        let note = "";
        const diff = blockParsed - rec.amt.num;
        if (diff !== 0) {
          const emptyRefs = (f.match(/\$?[A-Z]{1,3}\$?\d+/gi) ?? []).filter((ref) => {
            const c = readCell(ws.getCell(ref.replace(/\$/g, "")));
            return c.num === null;
          });
          note = `檔案總計與解析合計差 ${diff}`;
          if (emptyRefs.length) note += `；公式 ${f} 引用了空白儲存格 ${emptyRefs.join("、")}，可能少算`;
        }
        out.checks.push({
          half,
          block: h,
          kind: "區塊總計",
          label: rec.name,
          address: addr,
          formula: f,
          fileValue: rec.amt.num,
          parsedValue: blockParsed,
          diff,
          note,
        });
      }
    }
  }

  // 各區總計加總
  for (let b = 0; b < headers.length; b++) {
    const bc = out.checks.filter((c) => c.block === b);
    const combined = bc.filter((c) => c.kind === "區塊總計");
    const list = combined.length ? combined : bc.filter((c) => c.kind === "小計公式" && TOTAL_RE.test(c.label));
    out.fileBlockTotalSum += list.reduce((a, c) => a + c.fileValue, 0);
  }
  const cust = out.customers.reduce((a, x) => a + x.amount, 0);
  const ret = out.retail.reduce((a, x) => a + x.amount, 0);
  const ded = out.deducts.reduce((a, x) => a + x.net, 0);
  out.parsedTotal = cust + ret + ded;
  if (out.fileGrandTotal) {
    const diff = out.parsedTotal - out.fileGrandTotal.value;
    out.checks.push({
      half,
      block: -1,
      kind: "整期總計",
      label: "表尾總計",
      address: out.fileGrandTotal.address,
      formula: out.fileGrandTotal.formula,
      fileValue: out.fileGrandTotal.value,
      parsedValue: out.parsedTotal,
      diff,
      note: diff === 0 ? "" : `解析合計（客戶＋零售＋扣代送淨額）與表尾總計差 ${diff}`,
    });
  }
  for (const c of out.checks) {
    if (c.diff !== 0) out.warnings.push(`${ws.name}!${c.address}「${c.label}」${c.kind}：檔案 ${c.fileValue}、解析 ${c.parsedValue}、差 ${c.diff}${c.note ? "（" + c.note + "）" : ""}`);
  }
  for (const n of out.notes) {
    out.warnings.push(`${ws.name}!${n.address} 表尾附註列「${n.name}」${n.amount ?? ""}${n.formula ? `（公式 ${n.formula}）` : ""}，判斷為附註／合計，未當成客戶`);
  }
  return out;
}


export function parseWorkbook(wb: ExcelJS.Workbook, file: string, month: number): BillFileParse {
  const res: BillFileParse = { file, month, halves: [], skippedSheets: [], warnings: [] };
  for (const half of ["上期", "下期"] as Half[]) {
    const sheets = wb.worksheets.filter((ws) => ws.name.trim() === half);
    if (sheets.length === 0) {
      res.warnings.push(`${file} 找不到「${half}」工作表`);
      continue;
    }
    if (sheets.length > 1) res.warnings.push(`${file} 有多張「${half}」工作表，只讀第一張`);
    res.halves.push(parseSheet(sheets[0]!, half));
  }
  for (const ws of wb.worksheets) if (!["上期", "下期"].includes(ws.name.trim())) res.skippedSheets.push(ws.name);
  return res;
}

