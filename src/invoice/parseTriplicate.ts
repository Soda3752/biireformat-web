// 移植自 invoice-helper/src/parseTriplicate.ts（二聯發票分頁用；規則與命令列版一致，改動請兩邊同步）。
// 讀三聯式發票清單：依表頭找欄位（至少要有「客戶名稱」＋「銷售額」或「金額」）。
// 支援單一清單，也支援像二聯發票範本那樣左右兩個月並排（左＝前一月、右＝後一月）。
import ExcelJS from "exceljs";
import { readCell, normHeader, isSumFormula, colName } from "./cell.ts";
import { parseDateLoose, type Ymd } from "./dates.ts";

export type Triplicate = {
  date: Ymd | null;
  month: number | null; // 由日期判斷；沒有日期時由並排位置判斷
  name: string;
  invoiceNo: string;
  sales: number; // 銷售額（未稅）
  tax: number; // 營業稅
  amount: number; // 金額（含稅）＝ 用來扣總額
  source: string; // 工作表!儲存格
};

export type TriplicateParse = { items: Triplicate[]; warnings: string[] };

const TOTAL_RE = /總計|小計|合計|總額|共計|三聯共/;

type Group = { nameCol: number; dateCol: number | null; noCol: number | null; salesCol: number | null; taxCol: number | null; amountCol: number | null };

export function parseTriplicateWorkbook(wb: ExcelJS.Workbook, opts: { defaultYear: number; months: number[]; label?: string }): TriplicateParse {
  const res: TriplicateParse = { items: [], warnings: [] };
  const label = opts.label ?? "三聯清單";
  for (const ws of wb.worksheets) {
    const headerRows: { row: number; groups: Group[] }[] = [];
    for (let r = 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const hdr = (c: number) => normHeader(readCell(row.getCell(c)).text);
      const nameCols: number[] = [];
      for (let c = 1; c <= ws.columnCount; c++) if (hdr(c) === "客戶名稱") nameCols.push(c);
      if (!nameCols.length) continue;
      const groups: Group[] = [];
      let leftBound = 0;
      for (let i = 0; i < nameCols.length; i++) {
        const nc = nameCols[i]!;
        const rightBound = nameCols[i + 1] ?? ws.columnCount + 1;
        const g: Group = { nameCol: nc, dateCol: null, noCol: null, salesCol: null, taxCol: null, amountCol: null };
        for (let c = nc - 1; c > leftBound; c--) if (hdr(c) === "日期") { g.dateCol = c; break; }
        for (let c = nc + 1; c < rightBound; c++) {
          const h = hdr(c);
          if (g.noCol === null && h.includes("發票號碼")) g.noCol = c;
          else if (g.salesCol === null && h === "銷售額") g.salesCol = c;
          else if (g.taxCol === null && h === "營業稅") g.taxCol = c;
          else if (g.amountCol === null && h === "金額") { g.amountCol = c; break; }
        }
        if (g.dateCol === null) {
          // 左邊沒有日期 → 在本組欄位範圍內（客戶名稱到最後一個金額欄之間）往右找
          const last = Math.max(g.amountCol ?? 0, g.taxCol ?? 0, g.salesCol ?? 0);
          for (let c = nc + 1; c < last; c++) if (hdr(c) === "日期") { g.dateCol = c; break; }
        }
        if (g.salesCol === null && g.amountCol === null) {
          res.warnings.push(`${label}「${ws.name}」${colName(nc)}${r} 的客戶名稱旁找不到「銷售額」或「金額」欄，略過`);
        } else groups.push(g);
        leftBound = g.amountCol ?? g.salesCol ?? nc;
      }
      if (groups.length) headerRows.push({ row: r, groups });
    }

    for (let h = 0; h < headerRows.length; h++) {
      const { row: hr, groups } = headerRows[h]!;
      const end = h + 1 < headerRows.length ? headerRows[h + 1]!.row - 1 : ws.rowCount;
      groups.forEach((g, gi) => {
        for (let r = hr + 1; r <= end; r++) {
          const row = ws.getRow(r);
          const nameC = readCell(row.getCell(g.nameCol));
          const name = nameC.num === null ? nameC.text : "";
          const salesC = g.salesCol ? readCell(row.getCell(g.salesCol)) : null;
          const taxC = g.taxCol ? readCell(row.getCell(g.taxCol)) : null;
          const amtC = g.amountCol ? readCell(row.getCell(g.amountCol)) : null;
          if (!name) continue; // 合計列通常沒有客戶名稱
          if (TOTAL_RE.test(name) || isSumFormula(amtC?.formula ?? null) || isSumFormula(salesC?.formula ?? null)) continue;
          const sales0 = salesC?.num ?? null;
          const tax0 = taxC?.num ?? null;
          const amt0 = amtC?.num ?? null;
          if (sales0 === null && amt0 === null) {
            res.warnings.push(`${label}「${ws.name}」第 ${r} 列「${name}」沒有銷售額也沒有金額，略過`);
            continue;
          }
          let amount: number, sales: number, tax: number;
          if (amt0 !== null) {
            amount = Math.round(amt0);
            sales = sales0 !== null ? Math.round(sales0) : Math.round(amount / 1.05);
            tax = tax0 !== null ? Math.round(tax0) : amount - sales;
          } else {
            sales = Math.round(sales0!);
            tax = tax0 !== null ? Math.round(tax0) : Math.round(sales * 0.05);
            amount = sales + tax;
          }
          const dateC = g.dateCol ? readCell(row.getCell(g.dateCol)) : null;
          const date = dateC ? parseDateLoose(dateC, opts.defaultYear) : null;
          let month: number | null = date ? date.m : null;
          if (month === null && groups.length === opts.months.length && groups.length > 1) month = opts.months[gi] ?? null;
          if (month === null && opts.months.length === 1) month = opts.months[0]!;
          const noC = g.noCol ? readCell(row.getCell(g.noCol)) : null;
          res.items.push({
            date,
            month,
            name,
            invoiceNo: noC ? noC.text || (noC.num !== null ? String(noC.num) : "") : "",
            sales,
            tax,
            amount,
            source: `${ws.name}!${colName(g.nameCol)}${r}`,
          });
        }
      });
    }
  }
  if (!res.items.length) res.warnings.push(`${label}沒有讀到任何三聯發票`);
  return res;
}

