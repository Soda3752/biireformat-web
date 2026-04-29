import Papa from 'papaparse';

import type {DailyProduct, DailyProductMap} from '@/domain/models/daily-product';
import {localSettings} from '@/infra/local-settings-store';

const CSV_PATH = './assets/daily_report_list.csv';

let cached: Promise<DailyProductMap> | null = null;

/**
 * 對應桌面版 DailyCountViewModel.initProductList()
 * 讀 daily_report_list.csv，逐列建立 Product，並依 groupName 分組。
 * - 第一列 header 略過
 * - groupName 在每組第一列出現一次，後續沿用上一個 groupName
 *
 * 來源優先順序：
 * 1. localStorage（設定頁覆寫）
 * 2. public/assets/daily_report_list.csv（內建預設）
 *
 * 設定頁編輯後呼叫 `invalidateDailyReportTemplate()` 清快取。
 */
export const loadDailyReportTemplate = (): Promise<DailyProductMap> => {
  if (!cached) cached = fetchAndParse();
  return cached;
};

export const invalidateDailyReportTemplate = (): void => {
  cached = null;
};

/**
 * 訂閱 daily_report_list 「外部來源」變更事件。
 *
 * 設計：本機編輯（設定頁 daily 子分頁）只呼叫 `invalidateDailyReportTemplate()`，
 * 不會觸發本事件——避免使用者打字時被自己的 listener 打斷光標。
 * 外部來源（例如數據分析頁的「未分類商品 → 加入分類」彈窗）改呼叫
 * `notifyDailyReportChanged()`，會清快取並通知所有訂閱者重新載入。
 */
const listeners = new Set<() => void>();

export const onDailyReportChanged = (cb: () => void): (() => void) => {
    listeners.add(cb);
    return () => {
        listeners.delete(cb);
    };
};

export const notifyDailyReportChanged = (): void => {
    cached = null;
    for (const cb of listeners) {
        try {
            cb();
        } catch (err) {
            console.error('[daily-report-loader] listener error', err);
        }
    }
};

const fetchAndParse = async (): Promise<DailyProductMap> => {
  const text = await readDailyReportCsvText();
  return parseDailyReportCsv(text);
};

async function readDailyReportCsvText(): Promise<string> {
  const overridden = localSettings.getDailyReportList();
  if (overridden !== null) return overridden;

  const url = new URL(CSV_PATH, document.baseURI).toString();
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`載入 daily_report_list.csv 失敗: ${response.status}`);
  }
  return response.text();
}

export function parseDailyReportCsv(text: string): DailyProductMap {
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
      const costRaw = (row[3] ?? '').trim();

    if (groupName.length > 0) currentGroup = groupName;
      if (name.length > 0) {
          products.push({code, name, groupName: currentGroup, costRaw, count: 0});
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
}
