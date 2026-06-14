/**
 * 一次性資料搬移：把「單日數量（daily_report_list）」裡填的商品成本，
 * 依「商品編號」對應搬到「帳單排序（cargo_sort）」對應的商品上。
 *
 * 背景：成本欄原本掛在 daily_report_list 第 4 欄，現在改由 cargo_sort 第 4 欄保存。
 * 此搬移只跑一次（以 localStorage 旗標記錄），且：
 * - 只把 cargo 端「成本為空」的商品填上 daily 端的成本（不覆寫已填值）。
 * - daily 端的舊成本資料保留封存、不刪除（仍寫回 localStorage）。
 * - daily 有成本但 cargo 找不到同編號者，列為 unmatched 回報，不自動新增。
 */

import Papa from 'papaparse';

import {localSettings} from '@/infra/local-settings-store';

const MIGRATION_FLAG_KEY = 'bii.migration.costToCargo.done';
const CARGO_ASSET_URL = `${import.meta.env.BASE_URL}assets/cargo_sort.csv`;
const DAILY_ASSET_URL = `${import.meta.env.BASE_URL}assets/daily_report_list.csv`;
const CARGO_HEADER = ['貨品編號', '貨品名稱', '代送費', '成本'];

interface CargoRow {
    id: string;
    name: string;
    fee: string;
    cost: string;
}

export interface CostMigrationResult {
    /** 成功搬移（cargo 端成本被填上）的筆數 */
    migratedCount: number;
    /** daily 有成本、但 cargo 找不到同編號的商品 */
    unmatched: Array<{code: string; name: string; cost: string}>;
}

function isMigrationDone(): boolean {
    try {
        return localStorage.getItem(MIGRATION_FLAG_KEY) === '1';
    } catch {
        return false;
    }
}

function markMigrationDone(): void {
    try {
        localStorage.setItem(MIGRATION_FLAG_KEY, '1');
    } catch {
        // ignore
    }
}

async function fetchAssetText(url: string): Promise<string> {
    const res = await fetch(url);
    if (!res.ok) throw new Error(`無法讀取 ${url}（HTTP ${res.status}）`);
    return res.text();
}

/** 讀 daily：商品編號 → 成本（只收非空成本）；另保留品名供回報。 */
function parseDailyCostByCode(text: string): Map<string, {name: string; cost: string}> {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: 'greedy'});
    const map = new Map<string, {name: string; cost: string}>();
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        const code = String(row[1] ?? '').trim();
        const name = String(row[2] ?? '').trim();
        const cost = String(row[3] ?? '').trim();
        if (code.length === 0 || cost.length === 0) continue;
        if (!map.has(code)) map.set(code, {name, cost});
    }
    return map;
}

function parseCargoRows(text: string): CargoRow[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: true});
    const out: CargoRow[] = [];
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        out.push({
            id: String(row[0] ?? '').trim(),
            name: String(row[1] ?? '').trim(),
            fee: String(row[2] ?? '').trim(),
            cost: String(row[3] ?? '').trim(),
        });
    }
    return out;
}

function serializeCargoRows(rows: CargoRow[]): string {
    const data = [CARGO_HEADER, ...rows.map((r) => [r.id, r.name, r.fee, r.cost])];
    return Papa.unparse(data, {newline: '\n'});
}

/**
 * 執行一次性成本搬移。已搬過或無可搬資料時回傳 null。
 * 呼叫端應在 loadSortingList() 之前 await 本函式，確保 cargo 成本已就緒。
 */
export async function migrateDailyCostToCargo(): Promise<CostMigrationResult | null> {
    if (isMigrationDone()) return null;

    const dailyText = localSettings.getDailyReportList() ?? (await fetchAssetText(DAILY_ASSET_URL));
    const dailyCostByCode = parseDailyCostByCode(dailyText);

    // daily 端完全沒有任何成本資料 → 無可搬，直接標記完成
    if (dailyCostByCode.size === 0) {
        markMigrationDone();
        return null;
    }

    const cargoText = localSettings.getCargoSort() ?? (await fetchAssetText(CARGO_ASSET_URL));
    const cargoRows = parseCargoRows(cargoText);
    const cargoIds = new Set(cargoRows.map((r) => r.id).filter((id) => id.length > 0));

    let migratedCount = 0;
    for (const row of cargoRows) {
        if (row.id.length === 0) continue;
        if (row.cost.length > 0) continue; // 不覆寫已填成本
        const daily = dailyCostByCode.get(row.id);
        if (daily) {
            row.cost = daily.cost;
            migratedCount += 1;
        }
    }

    const unmatched: CostMigrationResult['unmatched'] = [];
    for (const [code, info] of dailyCostByCode) {
        if (!cargoIds.has(code)) unmatched.push({code, name: info.name, cost: info.cost});
    }

    // 只有真的有填到成本才寫回，避免在「完全沒對上」時無謂地把 cargo 從內建預設轉成本地覆寫
    if (migratedCount > 0) {
        localSettings.setCargoSort(serializeCargoRows(cargoRows));
    }
    markMigrationDone();

    return {migratedCount, unmatched};
}
