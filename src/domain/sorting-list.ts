/**
 * 對應桌面版 `billReformat/core/models/SortingList.kt`。
 *
 * 桌面版用 lazy + classpath resource 同步載入；網頁版改為模組層級 async loader：
 * 啟動時 `loadSortingList()` 一次抓取 `/assets/cargo_sort.csv` 並快取，
 * 後續同步使用 `getSortingList()` / `getItemIndex(name)` 等便利函式。
 *
 * 來源優先順序：
 * 1. localStorage（設定頁覆寫）
 * 2. public/assets/cargo_sort.csv（內建預設）
 *
 * 設定頁編輯後呼叫 `invalidateSortingList()` 清快取，再次 `loadSortingList()` 會重新讀。
 */

import Papa from 'papaparse';

import {localSettings} from '@/infra/local-settings-store';

export interface CargoItem {
  id: string;
  name: string;
  deliveryFee: number;
}

export interface SortingList {
  cargoItems: ReadonlyArray<CargoItem>;
  breadItems: ReadonlyArray<string>;
}

const ASSET_URL = `${import.meta.env.BASE_URL}assets/cargo_sort.csv`;

let cache: SortingList | null = null;
let inflight: Promise<SortingList> | null = null;

export function loadSortingList(): Promise<SortingList> {
  if (cache) return Promise.resolve(cache);
  if (inflight) return inflight;

  inflight = (async () => {
      const text = await readCargoSortCsvText();
      const sortingList = parseCargoSortCsv(text);
      cache = sortingList;
      inflight = null;
      return sortingList;
  })().catch((err) => {
      inflight = null;
      throw err;
  });

    return inflight;
}

export function invalidateSortingList(): void {
    cache = null;
    inflight = null;
}

async function readCargoSortCsvText(): Promise<string> {
    const overridden = localSettings.getCargoSort();
    if (overridden !== null) return overridden;

    const res = await fetch(ASSET_URL);
    if (!res.ok) {
        throw new Error(`無法讀取 cargo_sort.csv（HTTP ${res.status}）`);
    }
    return res.text();
}

export function parseCargoSortCsv(text: string): SortingList {
    const parsed = Papa.parse<string[]>(text, {
        header: false,
        skipEmptyLines: true,
    });

    const items: CargoItem[] = [];
    // 跳過第一列標題
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length < 3) continue;
        const fee = Number(String(row[2]).trim());
        if (Number.isNaN(fee)) continue;
        items.push({
            id: String(row[0]).trim(),
            name: String(row[1]).trim(),
            deliveryFee: fee,
        });
    }

    return {
        cargoItems: items,
        breadItems: items.map((it) => it.name),
    };
}

export function getSortingList(): SortingList {
  if (!cache) {
    throw new Error('SortingList 尚未載入，請先呼叫 loadSortingList()');
  }
  return cache;
}

/**
 * 取得品名在排序列表中的索引；找不到時回傳 Number.MAX_SAFE_INTEGER（對齊桌面版的 Int.MAX_VALUE 行為，可直接 sortBy）。
 */
export function getItemIndex(name: string): number {
  const list = getSortingList();
  const idx = list.cargoItems.findIndex((it) => it.name === name);
  return idx < 0 ? Number.MAX_SAFE_INTEGER : idx;
}

export function getCargoById(id: string): CargoItem | undefined {
  return getSortingList().cargoItems.find((it) => it.id === id);
}

export function getCargoByName(name: string): CargoItem | undefined {
  return getSortingList().cargoItems.find((it) => it.name === name);
}

export function getDeliveryFee(name: string): number | undefined {
  return getCargoByName(name)?.deliveryFee;
}
