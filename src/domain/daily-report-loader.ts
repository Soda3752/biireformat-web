import Papa from 'papaparse';

import type { DailyProduct, DailyProductMap } from '@/domain/models/daily-product';

const CSV_PATH = './assets/daily_report_list.csv';

let cached: Promise<DailyProductMap> | null = null;

/**
 * 對應桌面版 DailyCountViewModel.initProductList()
 * 讀 daily_report_list.csv，逐列建立 Product，並依 groupName 分組。
 * - 第一列 header 略過
 * - groupName 在每組第一列出現一次，後續沿用上一個 groupName
 */
export const loadDailyReportTemplate = (): Promise<DailyProductMap> => {
  if (!cached) cached = fetchAndParse();
  return cached;
};

const fetchAndParse = async (): Promise<DailyProductMap> => {
  const url = new URL(CSV_PATH, document.baseURI).toString();
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`載入 daily_report_list.csv 失敗: ${response.status}`);
  }
  const text = await response.text();

  const result = Papa.parse<string[]>(text, {
    header: false,
    skipEmptyLines: 'greedy',
  });

  const rows = (result.data as string[][]).slice(1); // 跳過 header
  const products: DailyProduct[] = [];
  let currentGroup = '';

  for (const row of rows) {
    if (row.length < 3) continue;
    const groupName = (row[0] ?? '').trim();
    const code = (row[1] ?? '').trim();
    const name = (row[2] ?? '').trim();

    if (groupName.length > 0) currentGroup = groupName;
    if (code.length > 0 && name.length > 0) {
      products.push({ code, name, groupName: currentGroup, count: 0 });
    }
  }

  // 依 groupName 分組（保持出現順序）
  const grouped: DailyProductMap = new Map();
  for (const p of products) {
    let arr = grouped.get(p.groupName);
    if (!arr) {
      arr = [];
      grouped.set(p.groupName, arr);
    }
    arr.push(p);
  }
  return grouped;
};
