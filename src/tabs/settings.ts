/**
 * 設定分頁：管理 cargo_sort.csv 與 daily_report_list.csv 的本地覆寫。
 *
 * - 內部子分頁：「帳單排序」 / 「品項分類」
 * - 每個子分頁支援：表格內編輯、新增列、刪除列、匯入 CSV、匯出 CSV
 * - 編輯 / 匯入後立即寫入 localStorage 並 invalidate 對應的 loader 快取
 * - 匯出格式與內建 CSV 相同
 */

import ExcelJS from 'exceljs';
import Papa from 'papaparse';
import {saveAs} from 'file-saver';

import {localSettings} from '@/infra/local-settings-store';
import {invalidateSortingList, loadSortingList,} from '@/domain/sorting-list';
import {invalidateDailyReportTemplate, onDailyReportChanged} from '@/domain/daily-report-loader';
import {
    BANK_INFO_HEADER,
    invalidateBankInfos,
    parseBankInfoCsv,
    serializeBankInfoCsv,
} from '@/domain/bank-info-loader';
import {
    collectCustomerEntries,
    CUSTOMER_ORDER_HEADER,
    type CustomerOrderEntry,
    invalidateCustomerOrderBill,
    invalidateCustomerOrderOverview,
    parseCustomerOrderCsv,
    serializeCustomerOrderCsv,
} from '@/domain/customer-order-loader';
import type {BankInfo} from '@/domain/models/bank-info';
import {equalsBankInfo} from '@/domain/models/bank-info';
import {parseBankInfo as parseBankInfoXlsx} from '@/readers/bank-info-reader';
import {readXlsxAsRows} from '@/infra/excel-codec';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';

interface CargoRow {
    id: string;
    name: string;
    fee: string;
}

interface DailyRow {
    group: string;
    code: string;
    name: string;
    cost: string;
}

interface LastFiveRow {
    customerName: string;
    customerLine: string;
    lastFiveDigit: string;
}

const CARGO_HEADER = ['貨品編號', '貨品名稱', '代送費'];
const DAILY_HEADER = ['分類', '貨品編號', '貨品名稱', '成本'];
const LAST_FIVE_HEADER = Array.from(BANK_INFO_HEADER);
const CUSTOMER_HEADER = Array.from(CUSTOMER_ORDER_HEADER);

const CARGO_ASSET_URL = `${import.meta.env.BASE_URL}assets/cargo_sort.csv`;
const DAILY_ASSET_URL = `${import.meta.env.BASE_URL}assets/daily_report_list.csv`;

export function renderSettingsPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">設定</h1>
        <p class="card-subtitle">
          管理「帳單排序（cargo_sort）」與「品項分類（daily_report_list）」兩份資料。
        </p>
      </header>

      <nav class="settings-subnav" role="tablist" aria-label="設定子分頁">
        <button type="button" class="settings-subnav-item is-active" data-subtab="cargo" role="tab" aria-selected="true">帳單排序</button>
        <button type="button" class="settings-subnav-item" data-subtab="daily" role="tab" aria-selected="false">品項分類</button>
        <button type="button" class="settings-subnav-item" data-subtab="bill-customer" role="tab" aria-selected="false">帳單客戶</button>
        <button type="button" class="settings-subnav-item" data-subtab="overview-customer" role="tab" aria-selected="false">明細客戶</button>
        <button type="button" class="settings-subnav-item" data-subtab="lastfive" role="tab" aria-selected="false">末五碼</button>
      </nav>

      <section class="settings-pane is-active" data-pane="cargo" role="tabpanel">
        <div class="settings-toolbar">
          <div class="settings-toolbar-info" data-role="cargo-status">載入中…</div>
          <div class="settings-toolbar-actions">
            <button type="button" class="btn btn-secondary" data-role="cargo-import">
              <span class="btn-icon">${icon('upload', 16)}</span>匯入（CSV / XLSX）
            </button>
            <button type="button" class="btn btn-secondary" data-role="cargo-export-csv">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 CSV
            </button>
            <button type="button" class="btn btn-secondary" data-role="cargo-export-xlsx">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 XLSX
            </button>
            <button type="button" class="btn btn-secondary" data-role="cargo-add">
              <span class="btn-icon">${icon('plus', 16)}</span>新增一列
            </button>
          </div>
          <input type="file" accept=".csv,.xlsx,.xls,text/csv" data-role="cargo-file" hidden />
        </div>
        <div class="settings-table-wrap" data-role="cargo-table-wrap"></div>
      </section>

      <section class="settings-pane" data-pane="daily" role="tabpanel" hidden>
        <div class="settings-toolbar">
          <div class="settings-toolbar-info" data-role="daily-status">載入中…</div>
          <div class="settings-toolbar-actions">
            <button type="button" class="btn btn-secondary" data-role="daily-import">
              <span class="btn-icon">${icon('upload', 16)}</span>匯入（CSV / XLSX）
            </button>
            <button type="button" class="btn btn-secondary" data-role="daily-export-csv">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 CSV
            </button>
            <button type="button" class="btn btn-secondary" data-role="daily-export-xlsx">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 XLSX
            </button>
            <button type="button" class="btn btn-secondary" data-role="daily-add">
              <span class="btn-icon">${icon('plus', 16)}</span>新增一列
            </button>
          </div>
          <input type="file" accept=".csv,.xlsx,.xls,text/csv" data-role="daily-file" hidden />
        </div>
        <div class="settings-table-wrap" data-role="daily-table-wrap"></div>
      </section>

      ${customerPaneHtml('bill-customer', '帳單客戶排序')}
      ${customerPaneHtml('overview-customer', '明細客戶排序')}

      <section class="settings-pane" data-pane="lastfive" role="tabpanel" hidden>
        <div class="settings-toolbar">
          <div class="settings-toolbar-info" data-role="lastfive-status">載入中…</div>
          <div class="settings-toolbar-actions">
            <button type="button" class="btn btn-secondary" data-role="lastfive-import">
              <span class="btn-icon">${icon('upload', 16)}</span>匯入（CSV / XLSX）
            </button>
            <button type="button" class="btn btn-secondary" data-role="lastfive-export-csv">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 CSV
            </button>
            <button type="button" class="btn btn-secondary" data-role="lastfive-export-xlsx">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 XLSX
            </button>
            <button type="button" class="btn btn-secondary" data-role="lastfive-add">
              <span class="btn-icon">${icon('plus', 16)}</span>新增一列
            </button>
          </div>
          <input type="file" accept=".csv,.xlsx,.xls,text/csv" data-role="lastfive-file" hidden />
        </div>
        <div class="settings-table-wrap" data-role="lastfive-table-wrap"></div>
      </section>
    </div>
  `;

    bindSubnav(panel);
    bindCargoPane(panel);
    bindDailyPane(panel);
    bindCustomerOrderPane(panel, 'bill');
    bindCustomerOrderPane(panel, 'overview');
    bindLastFivePane(panel);

    return panel;
}

// ============================================================================
// Sub-tab switching
// ============================================================================

function bindSubnav(panel: HTMLElement): void {
    const items = panel.querySelectorAll<HTMLButtonElement>('.settings-subnav-item');
    items.forEach((btn) => {
        btn.addEventListener('click', () => {
            const target = btn.dataset.subtab!;
            items.forEach((b) => {
                const active = b.dataset.subtab === target;
                b.classList.toggle('is-active', active);
                b.setAttribute('aria-selected', active ? 'true' : 'false');
            });
            panel.querySelectorAll<HTMLElement>('.settings-pane').forEach((pane) => {
                const active = pane.dataset.pane === target;
                pane.classList.toggle('is-active', active);
                pane.hidden = !active;
            });
        });
    });
}

// ============================================================================
// 帳單排序 (cargo_sort)
// ============================================================================

function bindCargoPane(panel: HTMLElement): void {
    const tableWrap = panel.querySelector<HTMLElement>('[data-role="cargo-table-wrap"]')!;
    const status = panel.querySelector<HTMLElement>('[data-role="cargo-status"]')!;
    const importBtn = panel.querySelector<HTMLButtonElement>('[data-role="cargo-import"]')!;
    const exportCsvBtn = panel.querySelector<HTMLButtonElement>('[data-role="cargo-export-csv"]')!;
    const exportXlsxBtn = panel.querySelector<HTMLButtonElement>('[data-role="cargo-export-xlsx"]')!;
    const addBtn = panel.querySelector<HTMLButtonElement>('[data-role="cargo-add"]')!;
    const fileInput = panel.querySelector<HTMLInputElement>('[data-role="cargo-file"]')!;

    let rows: CargoRow[] = [];

    const refreshStatus = () => {
        const overridden = localSettings.hasCargoSort();
        const tag = overridden
            ? '<span class="settings-badge settings-badge-overridden">本地覆寫</span>'
            : '<span class="settings-badge settings-badge-default">內建預設</span>';
        status.innerHTML = `${tag}<span class="settings-status-text">共 ${rows.length} 筆</span>`;
    };

    const persist = async () => {
        const csv = serializeCargoCsv(rows);
        localSettings.setCargoSort(csv);
        invalidateSortingList();
        try {
            await loadSortingList();
        } catch (err) {
            console.error(err);
        }
        refreshStatus();
    };

    const renderTable = () => {
        tableWrap.innerHTML = '';
        tableWrap.appendChild(
            buildEditableTable({
                headers: CARGO_HEADER,
                rows,
                rowToCells: (row) => [row.id, row.name, row.fee],
                onCellChange: (rowIdx, colIdx, value) => {
                    const r = rows[rowIdx];
                    if (!r) return;
                    if (colIdx === 0) r.id = value;
                    else if (colIdx === 1) r.name = value;
                    else if (colIdx === 2) r.fee = value;
                    void persist();
                },
                onDeleteRow: (rowIdx) => {
                    rows.splice(rowIdx, 1);
                    renderTable();
                    void persist();
                },
                onMoveRow: (from, to) => {
                    moveItem(rows, from, to);
                    renderTable();
                    void persist();
                },
                onInsertAbove: (rowIdx) => {
                    rows.splice(rowIdx, 0, {id: '', name: '', fee: '0'});
                    renderTable();
                    void persist();
                },
                onInsertBelow: (rowIdx) => {
                    rows.splice(rowIdx + 1, 0, {id: '', name: '', fee: '0'});
                    renderTable();
                    void persist();
                },
            })
        );
        refreshStatus();
    };

    // 初始載入
    void (async () => {
        rows = await loadCargoRows();
        renderTable();
    })();

    addBtn.addEventListener('click', () => {
        rows.push({id: '', name: '', fee: '0'});
        renderTable();
        void persist();
    });

    exportCsvBtn.addEventListener('click', () => {
        const csv = serializeCargoCsv(rows);
        const blob = new Blob([addBom(csv)], {type: 'text/csv;charset=utf-8'});
        saveAs(blob, 'cargo_sort.csv');
    });

    exportXlsxBtn.addEventListener('click', async () => {
        try {
            const blob = await buildCargoXlsxBlob(rows);
            saveAs(blob, 'cargo_sort.xlsx');
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯出失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        }
    });

    importBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) return;
        try {
            const isXlsx = /\.xlsx?$/i.test(file.name);
            let next: CargoRow[];
            if (isXlsx) {
                next = parseCargoFromXlsxRows(await readXlsxAsRows(file));
            } else {
                const text = stripBom(await file.text());
                next = parseCargoCsv(text);
            }
            rows = next;
            localSettings.setCargoSort(serializeCargoCsv(rows));
            invalidateSortingList();
            await loadSortingList();
            renderTable();
            showToast({
                variant: 'success',
                title: '帳單排序已匯入',
                message: `${file.name}（${rows.length} 筆）`,
            });
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯入失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        } finally {
            fileInput.value = '';
        }
    });
}

async function loadCargoRows(): Promise<CargoRow[]> {
    const overridden = localSettings.getCargoSort();
    if (overridden !== null) return parseCargoCsv(overridden);

    const res = await fetch(CARGO_ASSET_URL);
    if (!res.ok) throw new Error(`無法讀取 cargo_sort.csv（HTTP ${res.status}）`);
    return parseCargoCsv(await res.text());
}

function parseCargoCsv(text: string): CargoRow[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: true});
    const out: CargoRow[] = [];
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        out.push({
            id: String(row[0] ?? '').trim(),
            name: String(row[1] ?? '').trim(),
            fee: String(row[2] ?? '').trim(),
        });
    }
    return out;
}

function serializeCargoCsv(rows: CargoRow[]): string {
    const data = [CARGO_HEADER, ...rows.map((r) => [r.id, r.name, r.fee])];
    return Papa.unparse(data, {newline: '\n'});
}

function parseCargoFromXlsxRows(xlsxRows: string[][]): CargoRow[] {
    const out: CargoRow[] = [];
    // 跳過第一列 header（與 CSV 一致）
    for (let i = 1; i < xlsxRows.length; i++) {
        const row = xlsxRows[i];
        if (!row || row.length === 0) continue;
        out.push({
            id: (row[0] ?? '').trim(),
            name: (row[1] ?? '').trim(),
            fee: (row[2] ?? '').trim(),
        });
    }
    return out;
}

async function buildCargoXlsxBlob(rows: CargoRow[]): Promise<Blob> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('cargo_sort');
    ws.addRow(CARGO_HEADER);
    for (const r of rows) {
        const feeNum = Number(r.fee);
        ws.addRow([r.id, r.name, Number.isFinite(feeNum) ? feeNum : r.fee]);
    }
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 24;
    ws.getColumn(3).width = 10;
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

// ============================================================================
// 品項分類 (daily_report_list)
// ============================================================================

function bindDailyPane(panel: HTMLElement): void {
    const tableWrap = panel.querySelector<HTMLElement>('[data-role="daily-table-wrap"]')!;
    const status = panel.querySelector<HTMLElement>('[data-role="daily-status"]')!;
    const importBtn = panel.querySelector<HTMLButtonElement>('[data-role="daily-import"]')!;
    const exportCsvBtn = panel.querySelector<HTMLButtonElement>('[data-role="daily-export-csv"]')!;
    const exportXlsxBtn = panel.querySelector<HTMLButtonElement>('[data-role="daily-export-xlsx"]')!;
    const addBtn = panel.querySelector<HTMLButtonElement>('[data-role="daily-add"]')!;
    const fileInput = panel.querySelector<HTMLInputElement>('[data-role="daily-file"]')!;

    let rows: DailyRow[] = [];

    const refreshStatus = () => {
        const overridden = localSettings.hasDailyReportList();
        const tag = overridden
            ? '<span class="settings-badge settings-badge-overridden">本地覆寫</span>'
            : '<span class="settings-badge settings-badge-default">內建預設</span>';
        const groupCount = new Set(rows.map((r) => r.group).filter((g) => g.length > 0)).size;
        status.innerHTML = `${tag}<span class="settings-status-text">共 ${rows.length} 筆 / ${groupCount} 個分類</span>`;
    };

    const persist = () => {
        const csv = serializeDailyCsv(rows);
        localSettings.setDailyReportList(csv);
        invalidateDailyReportTemplate();
        refreshStatus();
    };

    const renderTable = () => {
        tableWrap.innerHTML = '';
        tableWrap.appendChild(
            buildEditableTable({
                headers: DAILY_HEADER,
                rows,
                rowToCells: (row) => [row.group, row.code, row.name, row.cost],
                onCellChange: (rowIdx, colIdx, value) => {
                    const r = rows[rowIdx];
                    if (!r) return;
                    if (colIdx === 0) r.group = value;
                    else if (colIdx === 1) r.code = value;
                    else if (colIdx === 2) r.name = value;
                    else if (colIdx === 3) r.cost = value;
                    persist();
                },
                onDeleteRow: (rowIdx) => {
                    rows.splice(rowIdx, 1);
                    renderTable();
                    persist();
                },
                onMoveRow: (from, to) => {
                    moveItem(rows, from, to);
                    renderTable();
                    persist();
                },
                onInsertAbove: (rowIdx) => {
                    // 沿用相鄰列的分類（往上插入時繼承被插入位置那列的 group）
                    const inheritGroup = rows[rowIdx]?.group ?? '';
                    rows.splice(rowIdx, 0, {group: inheritGroup, code: '', name: '', cost: ''});
                    renderTable();
                    persist();
                },
                onInsertBelow: (rowIdx) => {
                    const inheritGroup = rows[rowIdx]?.group ?? '';
                    rows.splice(rowIdx + 1, 0, {group: inheritGroup, code: '', name: '', cost: ''});
                    renderTable();
                    persist();
                },
            })
        );
        refreshStatus();
    };

    void (async () => {
        rows = await loadDailyRows();
        renderTable();
    })();

    // 訂閱外部變更（例：數據分析頁加入未分類商品）→ 重新讀取並重繪表格
    onDailyReportChanged(() => {
        void (async () => {
            try {
                rows = await loadDailyRows();
                renderTable();
            } catch (err) {
                console.error('[settings] daily 重新載入失敗', err);
            }
        })();
    });

    addBtn.addEventListener('click', () => {
        const lastGroup = rows.length > 0 ? rows[rows.length - 1].group : '';
        rows.push({group: lastGroup, code: '', name: '', cost: ''});
        renderTable();
        persist();
    });

    exportCsvBtn.addEventListener('click', () => {
        const csv = serializeDailyCsv(rows);
        const blob = new Blob([addBom(csv)], {type: 'text/csv;charset=utf-8'});
        saveAs(blob, 'daily_report_list.csv');
    });

    exportXlsxBtn.addEventListener('click', async () => {
        try {
            const blob = await buildDailyXlsxBlob(rows);
            saveAs(blob, 'daily_report_list.xlsx');
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯出失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        }
    });

    importBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) return;
        try {
            const isXlsx = /\.xlsx?$/i.test(file.name);
            if (isXlsx) {
                rows = parseDailyFromXlsxRows(await readXlsxAsRows(file));
            } else {
                const text = stripBom(await file.text());
                rows = parseDailyCsv(text);
            }
            localSettings.setDailyReportList(serializeDailyCsv(rows));
            invalidateDailyReportTemplate();
            renderTable();
            showToast({
                variant: 'success',
                title: '品項分類已匯入',
                message: `${file.name}（${rows.length} 筆）`,
            });
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯入失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        } finally {
            fileInput.value = '';
        }
    });
}

async function loadDailyRows(): Promise<DailyRow[]> {
    const overridden = localSettings.getDailyReportList();
    if (overridden !== null) return parseDailyCsv(overridden);

    const res = await fetch(DAILY_ASSET_URL);
    if (!res.ok) throw new Error(`無法讀取 daily_report_list.csv（HTTP ${res.status}）`);
    return parseDailyCsv(await res.text());
}

/**
 * 解析 daily_report_list.csv：
 * 來源 CSV 為稀疏格式（分類欄只有每組第一列填寫），UI 為了方便編輯改成「每列都顯示分類」，
 * 這裡解析時把空白分類自動繼承上一列的值。
 */
function parseDailyCsv(text: string): DailyRow[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: 'greedy'});
    const out: DailyRow[] = [];
    let currentGroup = '';
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        const groupRaw = String(row[0] ?? '').trim();
        const code = String(row[1] ?? '').trim();
        const name = String(row[2] ?? '').trim();
        const cost = String(row[3] ?? '').trim();
        if (groupRaw.length > 0) currentGroup = groupRaw;
        out.push({group: currentGroup, code, name, cost});
    }
    return out;
}

/**
 * 序列化 daily_report_list.csv：
 * UI 上每列都填了分類，輸出時還原為稀疏格式（連續同分類僅首列保留分類，其餘留空），
 * 與內建 daily_report_list.csv 格式一致。
 */
function serializeDailyCsv(rows: DailyRow[]): string {
    const out = sparsifyDailyRows(rows);
    return Papa.unparse([DAILY_HEADER.slice(), ...out], {newline: '\n'});
}

function sparsifyDailyRows(rows: DailyRow[]): string[][] {
    const out: string[][] = [];
    let prevGroup = '';
    for (const r of rows) {
        const groupCol = r.group !== prevGroup ? r.group : '';
        out.push([groupCol, r.code, r.name, r.cost]);
        prevGroup = r.group;
    }
    return out;
}

/**
 * 從 xlsx 讀回的二維字串展開稀疏分類（與 parseDailyCsv 對齊）。
 */
function parseDailyFromXlsxRows(xlsxRows: string[][]): DailyRow[] {
    const out: DailyRow[] = [];
    let currentGroup = '';
    for (let i = 1; i < xlsxRows.length; i++) {
        const row = xlsxRows[i];
        if (!row || row.length === 0) continue;
        const groupRaw = (row[0] ?? '').trim();
        const code = (row[1] ?? '').trim();
        const name = (row[2] ?? '').trim();
        const cost = (row[3] ?? '').trim();
        if (groupRaw.length > 0) currentGroup = groupRaw;
        out.push({group: currentGroup, code, name, cost});
    }
    return out;
}

async function buildDailyXlsxBlob(rows: DailyRow[]): Promise<Blob> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('daily_report_list');
    ws.addRow(DAILY_HEADER);
    for (const r of sparsifyDailyRows(rows)) {
        const [groupCol, code, name, costRaw] = r;
        const costNum = Number(costRaw);
        const costCell = costRaw !== '' && Number.isFinite(costNum) ? costNum : costRaw;
        ws.addRow([groupCol, code, name, costCell]);
    }
    ws.getColumn(1).width = 14;
    ws.getColumn(2).width = 12;
    ws.getColumn(3).width = 24;
    ws.getColumn(4).width = 10;
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

// ============================================================================
// 客戶排序 (customerOrder) — 帳單 / 明細分頁各自一份，互不影響
// ============================================================================

type CustomerVariant = 'bill' | 'overview';

interface CustomerVariantBindings {
    paneId: 'bill-customer' | 'overview-customer';
    filenameBase: string;
    sheetName: string;
    toastTitle: string;

    has(): boolean;

    get(): string | null;

    set(csv: string): void;

    invalidate(): void;
}

const CUSTOMER_VARIANTS: Record<CustomerVariant, CustomerVariantBindings> = {
    bill: {
        paneId: 'bill-customer',
        filenameBase: '帳單客戶排序',
        sheetName: '帳單客戶排序',
        toastTitle: '帳單客戶排序已匯入',
        has: () => localSettings.hasCustomerOrderBill(),
        get: () => localSettings.getCustomerOrderBill(),
        set: (csv) => localSettings.setCustomerOrderBill(csv),
        invalidate: () => invalidateCustomerOrderBill(),
    },
    overview: {
        paneId: 'overview-customer',
        filenameBase: '明細客戶排序',
        sheetName: '明細客戶排序',
        toastTitle: '明細客戶排序已匯入',
        has: () => localSettings.hasCustomerOrderOverview(),
        get: () => localSettings.getCustomerOrderOverview(),
        set: (csv) => localSettings.setCustomerOrderOverview(csv),
        invalidate: () => invalidateCustomerOrderOverview(),
    },
};

function customerPaneHtml(paneId: string, _label: string): string {
    return `
      <section class="settings-pane" data-pane="${paneId}" role="tabpanel" hidden>
        <div class="settings-toolbar">
          <div class="settings-toolbar-info" data-role="${paneId}-status">載入中…</div>
          <div class="settings-toolbar-actions">
            <button type="button" class="btn btn-secondary" data-role="${paneId}-import">
              <span class="btn-icon">${icon('upload', 16)}</span>匯入（CSV / XLSX）
            </button>
            <button type="button" class="btn btn-secondary" data-role="${paneId}-export-csv">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 CSV
            </button>
            <button type="button" class="btn btn-secondary" data-role="${paneId}-export-xlsx">
              <span class="btn-icon">${icon('download', 16)}</span>匯出 XLSX
            </button>
            <button type="button" class="btn btn-secondary" data-role="${paneId}-add">
              <span class="btn-icon">${icon('plus', 16)}</span>新增一列
            </button>
          </div>
          <input type="file" accept=".csv,.xlsx,.xls,text/csv" data-role="${paneId}-file" hidden />
        </div>
        <div class="settings-table-wrap" data-role="${paneId}-table-wrap"></div>
      </section>`;
}

function bindCustomerOrderPane(panel: HTMLElement, variant: CustomerVariant): void {
    const cfg = CUSTOMER_VARIANTS[variant];
    const role = (suffix: string) => `[data-role="${cfg.paneId}-${suffix}"]`;

    const tableWrap = panel.querySelector<HTMLElement>(role('table-wrap'))!;
    const status = panel.querySelector<HTMLElement>(role('status'))!;
    const importBtn = panel.querySelector<HTMLButtonElement>(role('import'))!;
    const exportCsvBtn = panel.querySelector<HTMLButtonElement>(role('export-csv'))!;
    const exportXlsxBtn = panel.querySelector<HTMLButtonElement>(role('export-xlsx'))!;
    const addBtn = panel.querySelector<HTMLButtonElement>(role('add'))!;
    const fileInput = panel.querySelector<HTMLInputElement>(role('file'))!;

    let entries: CustomerOrderEntry[] = loadCustomerEntries(cfg);

    const refreshStatus = () => {
        const overridden = cfg.has();
        const tag = overridden
            ? '<span class="settings-badge settings-badge-overridden">已建立</span>'
            : '<span class="settings-badge settings-badge-default">尚未建立</span>';
        status.innerHTML = `${tag}<span class="settings-status-text">共 ${entries.length} 筆</span>`;
    };

    const persist = () => {
        const csv = serializeCustomerOrderCsv(entries);
        cfg.set(csv);
        cfg.invalidate();
        refreshStatus();
    };

    const renderTable = () => {
        tableWrap.innerHTML = '';
        tableWrap.appendChild(
            buildEditableTable({
                headers: CUSTOMER_HEADER,
                rows: entries,
                rowToCells: (row) => [row.code, row.name],
                onCellChange: (rowIdx, colIdx, value) => {
                    const r = entries[rowIdx];
                    if (!r) return;
                    if (colIdx === 0) r.code = value;
                    else if (colIdx === 1) r.name = value;
                    persist();
                },
                onDeleteRow: (rowIdx) => {
                    entries.splice(rowIdx, 1);
                    renderTable();
                    persist();
                },
                onMoveRow: (from, to) => {
                    moveItem(entries, from, to);
                    renderTable();
                    persist();
                },
                onInsertAbove: (rowIdx) => {
                    entries.splice(rowIdx, 0, {code: '', name: ''});
                    renderTable();
                    persist();
                },
                onInsertBelow: (rowIdx) => {
                    entries.splice(rowIdx + 1, 0, {code: '', name: ''});
                    renderTable();
                    persist();
                },
            })
        );
        refreshStatus();
    };

    renderTable();

    addBtn.addEventListener('click', () => {
        entries.push({code: '', name: ''});
        renderTable();
        persist();
    });

    exportCsvBtn.addEventListener('click', () => {
        const csv = serializeCustomerOrderCsv(entries);
        const blob = new Blob([addBom(csv)], {type: 'text/csv;charset=utf-8'});
        saveAs(blob, `${cfg.filenameBase}.csv`);
    });

    exportXlsxBtn.addEventListener('click', async () => {
        try {
            const blob = await buildCustomerOrderXlsxBlob(entries, cfg.sheetName);
            saveAs(blob, `${cfg.filenameBase}.xlsx`);
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯出失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        }
    });

    importBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) return;
        try {
            const isXlsx = /\.xlsx?$/i.test(file.name);
            let next: CustomerOrderEntry[];
            if (isXlsx) {
                next = collectCustomerEntries(await readXlsxAsRows(file));
            } else {
                const text = stripBom(await file.text());
                next = parseCustomerOrderCsv(text);
            }
            entries = next;
            cfg.set(serializeCustomerOrderCsv(entries));
            cfg.invalidate();
            renderTable();
            showToast({
                variant: 'success',
                title: cfg.toastTitle,
                message: `${file.name}（${entries.length} 筆）`,
            });
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯入失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        } finally {
            fileInput.value = '';
        }
    });
}

function loadCustomerEntries(cfg: CustomerVariantBindings): CustomerOrderEntry[] {
    const overridden = cfg.get();
    return overridden === null ? [] : parseCustomerOrderCsv(overridden);
}

async function buildCustomerOrderXlsxBlob(
    entries: CustomerOrderEntry[],
    sheetName: string
): Promise<Blob> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(sheetName);
    ws.addRow(CUSTOMER_HEADER);
    for (const e of entries) ws.addRow([e.code, e.name]);
    ws.getColumn(1).width = 14;
    ws.getColumn(2).width = 22;
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

// ============================================================================
// 末五碼 (lastFiveDigit) — 對帳分頁的末五碼對照表
// ============================================================================

function bindLastFivePane(panel: HTMLElement): void {
    const tableWrap = panel.querySelector<HTMLElement>('[data-role="lastfive-table-wrap"]')!;
    const status = panel.querySelector<HTMLElement>('[data-role="lastfive-status"]')!;
    const importBtn = panel.querySelector<HTMLButtonElement>('[data-role="lastfive-import"]')!;
    const exportCsvBtn = panel.querySelector<HTMLButtonElement>('[data-role="lastfive-export-csv"]')!;
    const exportXlsxBtn = panel.querySelector<HTMLButtonElement>('[data-role="lastfive-export-xlsx"]')!;
    const addBtn = panel.querySelector<HTMLButtonElement>('[data-role="lastfive-add"]')!;
    const fileInput = panel.querySelector<HTMLInputElement>('[data-role="lastfive-file"]')!;

    let rows: LastFiveRow[] = [];

    const refreshStatus = () => {
        const overridden = localSettings.hasLastFiveDigit();
        const tag = overridden
            ? '<span class="settings-badge settings-badge-overridden">已建立</span>'
            : '<span class="settings-badge settings-badge-default">尚未建立</span>';
        status.innerHTML = `${tag}<span class="settings-status-text">共 ${rows.length} 筆</span>`;
    };

    const persist = () => {
        const csv = serializeLastFiveCsv(rows);
        localSettings.setLastFiveDigit(csv);
        invalidateBankInfos();
        refreshStatus();
    };

    const renderTable = () => {
        tableWrap.innerHTML = '';
        tableWrap.appendChild(
            buildEditableTable({
                headers: LAST_FIVE_HEADER,
                rows,
                rowToCells: (row) => [row.customerName, row.customerLine, row.lastFiveDigit],
                onCellChange: (rowIdx, colIdx, value) => {
                    const r = rows[rowIdx];
                    if (!r) return;
                    if (colIdx === 0) r.customerName = value;
                    else if (colIdx === 1) r.customerLine = value;
                    else if (colIdx === 2) r.lastFiveDigit = value;
                    persist();
                },
                onDeleteRow: (rowIdx) => {
                    rows.splice(rowIdx, 1);
                    renderTable();
                    persist();
                },
                onMoveRow: (from, to) => {
                    moveItem(rows, from, to);
                    renderTable();
                    persist();
                },
                onInsertAbove: (rowIdx) => {
                    rows.splice(rowIdx, 0, {customerName: '', customerLine: '', lastFiveDigit: ''});
                    renderTable();
                    persist();
                },
                onInsertBelow: (rowIdx) => {
                    rows.splice(rowIdx + 1, 0, {customerName: '', customerLine: '', lastFiveDigit: ''});
                    renderTable();
                    persist();
                },
            })
        );
        refreshStatus();
    };

    // 初始載入：localStorage 沒值就空陣列起步（與其他兩個不同，無內建 fallback）
    rows = loadLastFiveRows();
    renderTable();

    addBtn.addEventListener('click', () => {
        rows.push({customerName: '', customerLine: '', lastFiveDigit: ''});
        renderTable();
        persist();
    });

    exportCsvBtn.addEventListener('click', () => {
        const csv = serializeLastFiveCsv(rows);
        const blob = new Blob([addBom(csv)], {type: 'text/csv;charset=utf-8'});
        saveAs(blob, '末五碼對照表.csv');
    });

    exportXlsxBtn.addEventListener('click', async () => {
        try {
            const blob = await buildLastFiveXlsxBlob(rows);
            saveAs(blob, '末五碼對照表.xlsx');
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯出失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        }
    });

    importBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) return;
        try {
            const isXlsx = /\.xlsx?$/i.test(file.name);
            let next: LastFiveRow[];
            if (isXlsx) {
                const infos = await parseBankInfoXlsx(file);
                next = infos.map(bankInfoToRow);
            } else {
                const text = stripBom(await file.text());
                next = bankInfosToRows(parseBankInfoCsv(text));
            }
            rows = next;
            localSettings.setLastFiveDigit(serializeLastFiveCsv(rows));
            invalidateBankInfos();
            renderTable();
            showToast({
                variant: 'success',
                title: '末五碼資料已匯入',
                message: `${file.name}（${rows.length} 筆）`,
            });
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: '匯入失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        } finally {
            fileInput.value = '';
        }
    });
}

function loadLastFiveRows(): LastFiveRow[] {
    const overridden = localSettings.getLastFiveDigit();
    if (overridden === null) return [];
    return bankInfosToRows(parseBankInfoCsv(overridden));
}

function bankInfoToRow(info: BankInfo): LastFiveRow {
    return {
        customerName: info.customerName,
        customerLine: info.customerLine,
        lastFiveDigit: info.lastFiveDigit,
    };
}

function bankInfosToRows(infos: BankInfo[]): LastFiveRow[] {
    // 去重（與 BankInfoReader 對齊）
    const out: LastFiveRow[] = [];
    for (const info of infos) {
        if (info.lastFiveDigit.length === 0) continue;
        const dup = out.some(
            (r) =>
                r.customerName === info.customerName &&
                r.customerLine === info.customerLine &&
                r.lastFiveDigit === info.lastFiveDigit
        );
        if (!dup) out.push(bankInfoToRow(info));
    }
    return out;
}

function serializeLastFiveCsv(rows: LastFiveRow[]): string {
    const infos: BankInfo[] = rows.map((r) => ({
        customerName: r.customerName,
        customerLine: r.customerLine,
        lastFiveDigit: r.lastFiveDigit,
    }));
    // 序列化時不過濾空白，讓使用者編輯中的列也能保留；讀取端 (bank-info-loader) 會略過空末五碼
    return serializeBankInfoCsv(dedupBankInfos(infos));
}

function dedupBankInfos(infos: BankInfo[]): BankInfo[] {
    const out: BankInfo[] = [];
    for (const info of infos) {
        if (!out.some((existing) => equalsBankInfo(existing, info))) out.push(info);
    }
    return out;
}

async function buildLastFiveXlsxBlob(rows: LastFiveRow[]): Promise<Blob> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('末五碼對照表');
    ws.addRow(LAST_FIVE_HEADER);
    for (const r of rows) {
        ws.addRow([r.customerName, r.customerLine, r.lastFiveDigit]);
    }
    ws.getColumn(1).width = 18;
    ws.getColumn(2).width = 12;
    ws.getColumn(3).width = 14;
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}

// ============================================================================
// 共用：可編輯表格
// ============================================================================

interface EditableTableSpec<T> {
    headers: string[];
    rows: T[];
    rowToCells: (row: T) => string[];
    onCellChange: (rowIdx: number, colIdx: number, value: string) => void;
    onDeleteRow: (rowIdx: number) => void;
    onMoveRow: (from: number, to: number) => void;
    onInsertAbove: (rowIdx: number) => void;
    onInsertBelow: (rowIdx: number) => void;
}

function buildEditableTable<T>(spec: EditableTableSpec<T>): HTMLElement {
    const table = document.createElement('table');
    table.className = 'settings-table';

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const thGrip = document.createElement('th');
    thGrip.className = 'settings-table-grip-col';
    thGrip.setAttribute('aria-label', '拖曳排序');
    headerRow.appendChild(thGrip);
    for (const h of spec.headers) {
        const th = document.createElement('th');
        th.textContent = h;
        headerRow.appendChild(th);
    }
    const thAction = document.createElement('th');
    thAction.className = 'settings-table-action-col';
    thAction.textContent = '';
    headerRow.appendChild(thAction);
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    if (spec.rows.length === 0) {
        const tr = document.createElement('tr');
        const td = document.createElement('td');
        td.colSpan = spec.headers.length + 2;
        td.className = 'settings-table-empty';
        td.textContent = '尚無資料，點選「新增一列」或「匯入 CSV」開始。';
        tr.appendChild(td);
        tbody.appendChild(tr);
    } else {
        spec.rows.forEach((row, rowIdx) => {
            const tr = document.createElement('tr');
            tr.dataset.rowIdx = String(rowIdx);

            // grip column —— 拖曳手柄
            const gripTd = document.createElement('td');
            gripTd.className = 'settings-table-grip-col';
            const gripBtn = document.createElement('button');
            gripBtn.type = 'button';
            gripBtn.className = 'settings-table-grip';
            gripBtn.title = '拖曳調整排序';
            gripBtn.setAttribute('aria-label', '拖曳調整排序');
            gripBtn.innerHTML = icon('grip', 16);
            // 預設整列 draggable=false，僅當在 grip 上 mousedown 才暫時打開，避免影響輸入框文字選取
            gripBtn.addEventListener('mousedown', () => {
                tr.setAttribute('draggable', 'true');
            });
            gripBtn.addEventListener('mouseup', () => {
                tr.setAttribute('draggable', 'false');
            });
            gripTd.appendChild(gripBtn);
            tr.appendChild(gripTd);

            // 資料欄位
            const cells = spec.rowToCells(row);
            cells.forEach((value, colIdx) => {
                const td = document.createElement('td');
                const input = document.createElement('input');
                input.type = 'text';
                input.value = value;
                input.className = 'settings-table-input';
                input.addEventListener('input', () => {
                    spec.onCellChange(rowIdx, colIdx, input.value);
                });
                td.appendChild(input);
                tr.appendChild(td);
            });

            // 操作欄：往上插入 / 上移 / 下移 / 往下插入 / 刪除
            const actionTd = document.createElement('td');
            actionTd.className = 'settings-table-action-col';
            const actions = document.createElement('div');
            actions.className = 'settings-table-actions';

            const insertAboveBtn = document.createElement('button');
            insertAboveBtn.type = 'button';
            insertAboveBtn.className = 'settings-table-row-btn';
            insertAboveBtn.title = '往上插入一筆';
            insertAboveBtn.setAttribute('aria-label', '在此列上方插入一筆');
            insertAboveBtn.innerHTML = icon('row-insert-above', 16);
            insertAboveBtn.addEventListener('click', () => spec.onInsertAbove(rowIdx));
            actions.appendChild(insertAboveBtn);

            const upBtn = document.createElement('button');
            upBtn.type = 'button';
            upBtn.className = 'settings-table-row-btn';
            upBtn.title = '上移';
            upBtn.setAttribute('aria-label', '上移此列');
            upBtn.innerHTML = icon('chevron-up', 16);
            upBtn.disabled = rowIdx === 0;
            upBtn.addEventListener('click', () => spec.onMoveRow(rowIdx, rowIdx - 1));
            actions.appendChild(upBtn);

            const downBtn = document.createElement('button');
            downBtn.type = 'button';
            downBtn.className = 'settings-table-row-btn';
            downBtn.title = '下移';
            downBtn.setAttribute('aria-label', '下移此列');
            downBtn.innerHTML = icon('chevron-down', 16);
            downBtn.disabled = rowIdx === spec.rows.length - 1;
            downBtn.addEventListener('click', () => spec.onMoveRow(rowIdx, rowIdx + 1));
            actions.appendChild(downBtn);

            const insertBelowBtn = document.createElement('button');
            insertBelowBtn.type = 'button';
            insertBelowBtn.className = 'settings-table-row-btn';
            insertBelowBtn.title = '往下插入一筆';
            insertBelowBtn.setAttribute('aria-label', '在此列下方插入一筆');
            insertBelowBtn.innerHTML = icon('row-insert-below', 16);
            insertBelowBtn.addEventListener('click', () => spec.onInsertBelow(rowIdx));
            actions.appendChild(insertBelowBtn);

            const delBtn = document.createElement('button');
            delBtn.type = 'button';
            delBtn.className = 'settings-table-row-btn settings-table-row-btn-danger';
            delBtn.title = '刪除此列';
            delBtn.setAttribute('aria-label', '刪除此列');
            delBtn.innerHTML = icon('trash', 16);
            delBtn.addEventListener('click', () => spec.onDeleteRow(rowIdx));
            actions.appendChild(delBtn);

            actionTd.appendChild(actions);
            tr.appendChild(actionTd);

            tbody.appendChild(tr);
        });
    }

    attachDragHandlers(tbody, spec.onMoveRow);

    table.appendChild(tbody);

    const wrap = document.createElement('div');
    wrap.className = 'settings-table-scroll';
    wrap.appendChild(table);
    return wrap;
}

function attachDragHandlers(
    tbody: HTMLElement,
    onMoveRow: (from: number, to: number) => void
): void {
    let sourceIdx: number | null = null;

    const clearDropMarks = () => {
        tbody.querySelectorAll<HTMLElement>('tr').forEach((tr) => {
            tr.classList.remove('is-drop-above', 'is-drop-below');
        });
    };

    tbody.addEventListener('dragstart', (e) => {
        const tr = (e.target as HTMLElement).closest('tr');
        if (!tr || tr.dataset.rowIdx === undefined) return;
        sourceIdx = Number(tr.dataset.rowIdx);
        tr.classList.add('is-dragging');
        e.dataTransfer?.setData('text/plain', String(sourceIdx));
        if (e.dataTransfer) e.dataTransfer.effectAllowed = 'move';
    });

    tbody.addEventListener('dragover', (e) => {
        if (sourceIdx === null) return;
        const tr = (e.target as HTMLElement).closest('tr');
        if (!tr || tr.dataset.rowIdx === undefined) return;
        e.preventDefault();
        if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';

        const rect = tr.getBoundingClientRect();
        const isAbove = e.clientY < rect.top + rect.height / 2;
        clearDropMarks();
        tr.classList.add(isAbove ? 'is-drop-above' : 'is-drop-below');
    });

    tbody.addEventListener('drop', (e) => {
        if (sourceIdx === null) return;
        const tr = (e.target as HTMLElement).closest('tr');
        if (!tr || tr.dataset.rowIdx === undefined) return;
        e.preventDefault();

        const targetRowIdx = Number(tr.dataset.rowIdx);
        const rect = tr.getBoundingClientRect();
        const isAbove = e.clientY < rect.top + rect.height / 2;
        const insertBefore = isAbove ? targetRowIdx : targetRowIdx + 1;

        clearDropMarks();
        const from = sourceIdx;
        sourceIdx = null;
        if (from === insertBefore || from + 1 === insertBefore) return;
        // insertBefore 為「插入到此位置」(0..length) 語意；換算成 onMoveRow 期望的「最終目標索引」
        const finalIdx = from < insertBefore ? insertBefore - 1 : insertBefore;
        onMoveRow(from, finalIdx);
    });

    tbody.addEventListener('dragend', () => {
        sourceIdx = null;
        clearDropMarks();
        tbody.querySelectorAll<HTMLElement>('tr.is-dragging').forEach((tr) => {
            tr.classList.remove('is-dragging');
            tr.setAttribute('draggable', 'false');
        });
    });
}

/**
 * 將陣列中的元素從 `from` 移到「最終目標索引 `to`」（0..length-1），
 * 上/下箭頭與 drag-and-drop 都先換算成此語意再呼叫。
 */
function moveItem<T>(arr: T[], from: number, to: number): void {
    if (from === to) return;
    if (from < 0 || from >= arr.length) return;
    const clamped = Math.max(0, Math.min(arr.length - 1, to));
    const [item] = arr.splice(from, 1);
    arr.splice(clamped, 0, item);
}

// ============================================================================
// CSV 工具
// ============================================================================

function addBom(text: string): string {
    return '﻿' + text;
}

function stripBom(text: string): string {
    return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
