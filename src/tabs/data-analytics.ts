/**
 * 數據分析分頁完整版（P-A2 ~ P-A5）：
 * - 多檔拖曳上傳
 * - KPI 卡片群
 * - 12 張 ECharts 圖表（dynamic import）
 * - 全域篩選器 + 鑽取互動
 * - 明細表（排序/搜尋/分頁）
 * - CSV / PNG 匯出
 */

import {saveAs} from 'file-saver';

import {showToast} from '@/ui/toast';
import {icon} from '@/ui/icons';
import type {TabDefinition} from '@/ui/tabs';
import {parseBillFile} from '@/readers/bill-reader';
import {Bill} from '@/domain/models/bill';
import {ExcelRowType} from '@/domain/excel-row-data';
import {loadCategoryMap} from '@/analytics/category-loader';
import {
    type AnalyticsDataset,
    type AnalyticsRow,
    buildDataset,
    type LoadedFileMeta,
} from '@/analytics/dataset-builder';
import {applyFilter, EMPTY_FILTER, type FilterState} from '@/analytics/filter-engine';
import {createFilterUi} from '@/analytics/filter-ui';
import {createDetailTable, rowsToCsv} from '@/analytics/detail-table';
import {type ChartHandle, createChart, observeChartsResize} from '@/analytics/chart-manager';
import {computeKpi, renderKpiCards} from '@/analytics/kpi';
import {
    anomalyOption,
    categoryTreemapOption,
    customerParetoOption,
    customerTopOption,
    dailyTrendOption,
    detectAnomalies,
    linePieOption,
    monthOverMonthOption,
    productTopOption,
    weekdayOption,
} from '@/analytics/chart-options';

interface LoadedBillRecord {
    fileId: string;
    fileName: string;
    fileSize: number;
    bill: Bill;
}

/** 圖表 slot：對應 panel 內某個圖表卡的容器 + chart instance */
interface ChartSlot {
    id: string;
    title: string;
    container: HTMLElement;
    pngBtn: HTMLButtonElement;
    chart: ChartHandle | null;
    /** 篩選後的 rows + 完整 dataset，產出 EChartsOption。回傳 null 表示無資料。 */
    build: (rows: ReadonlyArray<AnalyticsRow>, dataset: AnalyticsDataset) => unknown | null;
}

export function renderDataAnalyticsPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card analytics-card">
      <header class="card-header">
        <h1 class="card-title">數據分析</h1>
        <p class="card-subtitle">
          拖入一份或多份「明細編號順序 .xlsx」帳單檔，系統會即時聚合並產出
          KPI、營收趨勢、線別佔比、商品/客戶排行、熱力圖等多維度分析。
          可同時拖入多份月份檔做月對月比較。
        </p>
      </header>

      <div class="analytics-uploader">
        <div class="analytics-dropzone" data-role="dropzone" tabindex="0" role="button">
          <span class="analytics-dropzone-icon">${icon('upload', 28)}</span>
          <div class="analytics-dropzone-title">拖入或點擊選擇 .xlsx 帳單檔</div>
          <div class="analytics-dropzone-hint">支援多檔，可一次拖入 1~12 份做月對月比較</div>
          <input type="file" accept=".xlsx,.xls" multiple hidden data-role="file-input">
        </div>
        <div class="analytics-file-list" data-role="file-list" hidden></div>
        <div class="analytics-uploader-actions" hidden data-role="uploader-actions">
          <button type="button" class="btn btn-secondary" data-role="clear-all">清空全部</button>
        </div>
      </div>

      <section class="analytics-filters-section" data-role="filters-section" hidden>
        <div data-role="filters-host"></div>
      </section>

      <section class="analytics-content" data-role="content" hidden>
        <div class="analytics-section-label">總覽 KPI</div>
        <div class="analytics-kpi-grid" data-role="kpi-grid"></div>

        <div class="analytics-section-label">趨勢與分佈</div>
        <div class="analytics-chart-grid" data-role="chart-grid"></div>

        <div class="analytics-section-label">明細表</div>
        <div data-role="detail-host"></div>
        <div class="analytics-detail-actions">
          <button type="button" class="btn btn-secondary" data-role="export-csv">
            ${icon('download', 16)} 匯出明細 CSV
          </button>
        </div>

        <details class="analytics-unmatched" data-role="unmatched" hidden>
          <summary><span data-role="unmatched-summary"></span></summary>
          <ul class="analytics-unmatched-list" data-role="unmatched-list"></ul>
        </details>
      </section>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="status">請拖入或選擇 .xlsx 帳單檔</div>
      </footer>
    </div>
  `;

    // ===== DOM 句柄 =====
    const dropzone = panel.querySelector<HTMLElement>('[data-role="dropzone"]')!;
    const fileInput = panel.querySelector<HTMLInputElement>('[data-role="file-input"]')!;
    const fileListEl = panel.querySelector<HTMLElement>('[data-role="file-list"]')!;
    const uploaderActions = panel.querySelector<HTMLElement>('[data-role="uploader-actions"]')!;
    const clearAllBtn = panel.querySelector<HTMLButtonElement>('[data-role="clear-all"]')!;
    const filtersSection = panel.querySelector<HTMLElement>('[data-role="filters-section"]')!;
    const filtersHost = panel.querySelector<HTMLElement>('[data-role="filters-host"]')!;
    const contentSection = panel.querySelector<HTMLElement>('[data-role="content"]')!;
    const kpiHost = panel.querySelector<HTMLElement>('[data-role="kpi-grid"]')!;
    const chartGridHost = panel.querySelector<HTMLElement>('[data-role="chart-grid"]')!;
    const detailHost = panel.querySelector<HTMLElement>('[data-role="detail-host"]')!;
    const exportCsvBtn = panel.querySelector<HTMLButtonElement>('[data-role="export-csv"]')!;
    const unmatchedHost = panel.querySelector<HTMLDetailsElement>('[data-role="unmatched"]')!;
    const unmatchedSummary = panel.querySelector<HTMLElement>('[data-role="unmatched-summary"]')!;
    const unmatchedList = panel.querySelector<HTMLElement>('[data-role="unmatched-list"]')!;
    const statusEl = panel.querySelector<HTMLElement>('[data-role="status"]')!;

    // ===== 狀態 =====
    const loaded: LoadedBillRecord[] = [];
    let dataset: AnalyticsDataset | null = null;
    let filterState: FilterState = {...EMPTY_FILTER};
    let resizeCleanup: (() => void) | null = null;

    // ===== 圖表切換按鈕的當下狀態（每張圖獨立） =====
    let dailyMetric: 'amount' | 'count' = 'amount';
    let productMetric: 'amount' | 'count' = 'amount';
    let weekdayMetric: 'amount' | 'count' = 'amount';
    let linePieMetric: 'amount' | 'count' = 'amount';
    let customerTopMetric: 'amount' | 'count' = 'amount';
    let categoryMetric: 'amount' | 'count' = 'amount';
    let paretoMetric: 'amount' | 'count' = 'amount';
    let anomalyMetric: 'amount' | 'count' = 'amount';

    // ===== 子元件 =====
    const detailTable = createDetailTable();
    detailHost.appendChild(detailTable.element);

    const filterUi = createFilterUi({
        onChange: (s) => {
            filterState = s;
            void renderAll();
        },
    });
    filtersHost.appendChild(filterUi.element);

    // ===== 圖表 slots 定義 =====
    const slots: ChartSlot[] = buildSlots();

    function buildSlots(): ChartSlot[] {
        interface SlotDef {
            id: string;
            title: string;
            wide?: boolean;
            makeControls?: () => HTMLElement;
            build: ChartSlot['build'];
        }

        const list: SlotDef[] = [
            {
                id: 'daily-trend',
                title: '每日營收趨勢',
                wide: true,
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    dailyMetric = v;
                    void renderAll();
                }, () => dailyMetric),
                build: (rows) => (rows.length === 0 ? null : dailyTrendOption(rows, dailyMetric)),
            },
            {
                id: 'weekday-heat',
                title: '星期銷售熱度',
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    weekdayMetric = v;
                    void renderAll();
                }, () => weekdayMetric),
                build: (rows) => (rows.length === 0 ? null : weekdayOption(rows, weekdayMetric)),
            },
            {
                id: 'product-top',
                title: '商品銷售 Top 10',
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    productMetric = v;
                    void renderAll();
                }, () => productMetric),
                build: (rows) => (rows.length === 0 ? null : productTopOption(rows, productMetric)),
            },
            {
                id: 'customer-top',
                title: '客戶銷售 Top 10',
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    customerTopMetric = v;
                    void renderAll();
                }, () => customerTopMetric),
                build: (rows) => (rows.length === 0 ? null : customerTopOption(rows, customerTopMetric)),
            },
            {
                id: 'category-treemap',
                title: '商品分類佔比',
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    categoryMetric = v;
                    void renderAll();
                }, () => categoryMetric),
                build: (rows) => (rows.length === 0 ? null : categoryTreemapOption(rows, categoryMetric)),
            },
            {
                id: 'customer-pareto',
                title: '客戶 80/20 帕累托',
                wide: true,
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    paretoMetric = v;
                    void renderAll();
                }, () => paretoMetric),
                build: (rows) => (rows.length === 0 ? null : customerParetoOption(rows, paretoMetric)),
            },
            {
                id: 'line-pie',
                title: '線別佔比',
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    linePieMetric = v;
                    void renderAll();
                }, () => linePieMetric),
                build: (rows) => (rows.length === 0 ? null : linePieOption(rows, linePieMetric)),
            },
            {
                id: 'anomaly',
                title: '異常日偵測（|z|>2）',
                wide: true,
                makeControls: () => makeMetricSwitch('amount', (v) => {
                    anomalyMetric = v;
                    void renderAll();
                }, () => anomalyMetric),
                build: (rows) => (rows.length === 0 ? null : anomalyOption(rows, anomalyMetric)),
            },
            {
                id: 'month-over-month',
                title: '月對月成長（多檔時顯示）',
                wide: true,
                build: (_rows, ds) => (ds.files.length < 2 ? null : monthOverMonthOption(ds)),
            },
        ];

        const built: ChartSlot[] = [];
        for (const s of list) {
            const card = document.createElement('div');
            card.className = 'analytics-chart-card' + (s.wide ? ' analytics-chart-card-wide' : '');
            card.dataset.chart = s.id;
            const header = document.createElement('div');
            header.className = 'analytics-chart-header';
            const titleEl = document.createElement('div');
            titleEl.className = 'analytics-chart-title';
            titleEl.textContent = s.title;
            header.appendChild(titleEl);

            const headerRight = document.createElement('div');
            headerRight.className = 'analytics-chart-header-right';
            const ctlEl = s.makeControls ? s.makeControls() : null;
            if (ctlEl) headerRight.appendChild(ctlEl);
            const pngBtn = document.createElement('button');
            pngBtn.type = 'button';
            pngBtn.className = 'analytics-chart-png-btn';
            pngBtn.title = '下載 PNG';
            pngBtn.innerHTML = icon('download', 14);
            headerRight.appendChild(pngBtn);
            header.appendChild(headerRight);

            card.appendChild(header);

            const body = document.createElement('div');
            body.className = 'analytics-chart-body';
            card.appendChild(body);

            chartGridHost.appendChild(card);

            built.push({
                id: s.id,
                title: s.title,
                container: body,
                pngBtn,
                chart: null,
                build: s.build,
            });
        }
        return built;
    }

    // ===== 鑽取（chart click → filterUi.applyPatch） =====
    const wireDrillDown = (slot: ChartSlot) => {
        if (!slot.chart) return;
        slot.chart.off('click');
        slot.chart.on('click', (params) => {
            const p = params as { seriesType?: string; name?: string; data?: unknown };
            if (slot.id === 'line-pie' && p.name) {
                filterUi.applyPatch({lines: new Set([p.name])});
            } else if (slot.id === 'product-top' && p.name) {
                filterUi.applyPatch({productNames: new Set([p.name])});
                detailTable.scrollIntoView();
            } else if (slot.id === 'customer-top' && p.name) {
                // p.name 形如「客戶名(代碼)」，反推回 code
                const m = /\(([^()]+)\)$/.exec(p.name);
                if (m) filterUi.applyPatch({customerCodes: new Set([m[1]])});
            }
        });
    };

    // ===== mount / render =====
    const mountChartsIfNeeded = async () => {
        for (const slot of slots) {
            if (slot.chart) continue;
            try {
                slot.chart = await createChart(slot.container);
                wireDrillDown(slot);
                slot.pngBtn.addEventListener('click', () => exportSlotPng(slot));
            } catch (err) {
                console.error('[analytics] chart mount failed', slot.id, err);
            }
        }
        if (resizeCleanup) resizeCleanup();
        const handles = slots.map((s) => s.chart).filter((c): c is ChartHandle => c !== null);
        resizeCleanup = observeChartsResize(handles);
    };

    const renderAll = async () => {
        if (!dataset) return;
        const filtered = applyFilter(dataset, filterState);
        const filteredRows = filtered.rows;

        // KPI
        renderKpiCards(kpiHost, computeKpi(filtered));

        // 圖表
        await mountChartsIfNeeded();
        for (const slot of slots) {
            if (!slot.chart) continue;
            const opt = slot.build(filteredRows, dataset);
            if (opt) {
                slot.chart.setOption(opt as never);
                slot.container.parentElement!.classList.remove('is-empty');
            } else {
                // 無資料時顯示空狀態，但保留 chart instance
                slot.chart.setOption({
                    title: {
                        text: '無資料',
                        left: 'center',
                        top: 'center',
                        textStyle: {fontSize: 13, color: '#999'}
                    }
                } as never);
            }
            // month-over-month 只在多檔時顯示卡片
            if (slot.id === 'month-over-month') {
                slot.container.parentElement!.hidden = dataset.files.length < 2;
            }
        }

        // 明細表
        detailTable.setRows(filteredRows);

        // 異常清單寫到 status 區（依當前異常圖的 metric）
        const {anomalies} = detectAnomalies(filteredRows, anomalyMetric);
        const anomalyHint = anomalies.length > 0 ? `　|　偵測到 ${anomalies.length} 個異常日（${anomalies.map((a) => `${a.day}日`).join(', ')}）` : '';

        statusEl.textContent = `已載入 ${dataset.files.length} 檔 / 篩選後 ${filteredRows.length.toLocaleString()} 筆 / 原始 ${dataset.rows.length.toLocaleString()} 筆${anomalyHint}`;
    };

    // ===== 上傳/解析 =====
    const fileId = (f: File) => `${f.name}__${f.size}`;

    const renderFileList = () => {
        if (loaded.length === 0) {
            fileListEl.hidden = true;
            uploaderActions.hidden = true;
            fileListEl.innerHTML = '';
            return;
        }
        fileListEl.hidden = false;
        uploaderActions.hidden = false;
        fileListEl.innerHTML = loaded
            .map(
                (rec, idx) => `
          <div class="analytics-file-chip">
            <span class="analytics-file-chip-name">${escapeHtml(rec.fileName)}</span>
            <span class="analytics-file-chip-meta">${formatBytes(rec.fileSize)}</span>
            <button type="button" class="analytics-file-chip-remove" data-idx="${idx}" aria-label="移除">${icon('close', 14)}</button>
          </div>
        `
            )
            .join('');
        fileListEl.querySelectorAll<HTMLButtonElement>('.analytics-file-chip-remove').forEach((btn) => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const idx = Number(btn.dataset.idx);
                loaded.splice(idx, 1);
                renderFileList();
                void rebuildDataset();
            });
        });
    };

    const rebuildDataset = async () => {
        if (loaded.length === 0) {
            dataset = null;
            filtersSection.hidden = true;
            contentSection.hidden = true;
            unmatchedHost.hidden = true;
            statusEl.textContent = '請拖入或選擇 .xlsx 帳單檔';
            return;
        }
        statusEl.textContent = '聚合資料中…';
        try {
            const categoryMap = await loadCategoryMap();
            const fileMetas: LoadedFileMeta[] = loaded.map((rec) => ({
                id: rec.fileId,
                name: rec.fileName,
                bill: rec.bill,
            }));
            dataset = buildDataset(fileMetas, categoryMap);

            filtersSection.hidden = false;
            contentSection.hidden = false;
            filterUi.setDataset(dataset);
            renderUnmatched(dataset);

            console.info('[analytics] dataset built:', {
                files: dataset.files.length,
                rows: dataset.rows.length,
                unmatchedProducts: dataset.unmatchedProducts.length,
                sample: dataset.rows.slice(0, 5),
            });

            statusEl.textContent = '渲染圖表中…';
            await renderAll();
        } catch (err) {
            console.error('[analytics] rebuild failed', err);
            showToast({
                variant: 'error',
                title: '聚合失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            statusEl.textContent = '聚合失敗，請檢查檔案';
        }
    };

    const renderUnmatched = (ds: AnalyticsDataset) => {
        if (ds.unmatchedProducts.length === 0) {
            unmatchedHost.hidden = true;
            return;
        }
        unmatchedHost.hidden = false;
        unmatchedSummary.textContent = `未分類商品（${ds.unmatchedProducts.length}）— 設定頁可補充 daily_report_list 對應`;
        unmatchedList.innerHTML = ds.unmatchedProducts
            .map((name) => `<li>${escapeHtml(name)}</li>`)
            .join('');
    };

    const handleFiles = async (files: FileList | File[]) => {
        const arr = Array.from(files);
        if (arr.length === 0) return;
        const existingIds = new Set(loaded.map((r) => r.fileId));
        const newFiles = arr.filter((f) => !existingIds.has(fileId(f)));
        const skipped = arr.length - newFiles.length;
        if (skipped > 0) {
            showToast({variant: 'warning', title: '已略過重複檔案', message: `${skipped} 份檔案已存在`});
        }
        if (newFiles.length === 0) return;

        statusEl.textContent = `解析中…（${newFiles.length} 檔）`;
        for (const file of newFiles) {
            try {
                const bill = new Bill();
                await parseBillFile(file, ({type, values}) => {
                    if (type === ExcelRowType.BillSettingInfo) bill.setBillDateInfo(values);
                    else if (type === ExcelRowType.CustomerData) bill.newCustomerModel(values);
                    else if (type === ExcelRowType.CustomerSetting) bill.configCustomer(values);
                    else if (type === ExcelRowType.ProductRowSetting) bill.setProductSetting(values);
                    else if (type === ExcelRowType.ProductSellData) bill.addProduct(values);
                });
                bill.appendCustomer();
                loaded.push({
                    fileId: fileId(file),
                    fileName: file.name,
                    fileSize: file.size,
                    bill,
                });
            } catch (err) {
                console.error('[analytics] parse failed', file.name, err);
                showToast({
                    variant: 'error',
                    title: `${file.name} 解析失敗`,
                    message: err instanceof Error ? err.message : String(err),
                });
            }
        }
        renderFileList();
        await rebuildDataset();
    };

    // ===== 拖曳事件 =====
    dropzone.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            fileInput.click();
        }
    });
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('is-dragover');
    });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('is-dragover'));
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('is-dragover');
        if (e.dataTransfer?.files) void handleFiles(e.dataTransfer.files);
    });
    fileInput.addEventListener('change', () => {
        if (fileInput.files) void handleFiles(fileInput.files);
        fileInput.value = '';
    });
    clearAllBtn.addEventListener('click', () => {
        loaded.splice(0, loaded.length);
        filterUi.reset();
        renderFileList();
        void rebuildDataset();
    });

    // ===== CSV 匯出 =====
    exportCsvBtn.addEventListener('click', () => {
        const rows = detailTable.getCurrentRows();
        if (rows.length === 0) {
            showToast({variant: 'warning', title: '沒有資料可匯出', message: '請先載入檔案或調整篩選'});
            return;
        }
        const csv = rowsToCsv(rows);
        const blob = new Blob([csv], {type: 'text/csv;charset=utf-8'});
        const ts = new Date();
        const stamp = `${ts.getFullYear()}${String(ts.getMonth() + 1).padStart(2, '0')}${String(ts.getDate()).padStart(2, '0')}_${String(ts.getHours()).padStart(2, '0')}${String(ts.getMinutes()).padStart(2, '0')}`;
        saveAs(blob, `數據分析明細_${stamp}.csv`);
        showToast({variant: 'success', title: 'CSV 已匯出', message: `${rows.length.toLocaleString()} 筆`});
    });

    // ===== PNG 匯出 =====
    const exportSlotPng = (slot: ChartSlot) => {
        if (!slot.chart) return;
        try {
            const url = slot.chart.getInstance().getDataURL({type: 'png', pixelRatio: 2, backgroundColor: '#fff'});
            const a = document.createElement('a');
            a.href = url;
            a.download = `${slot.title}.png`;
            a.click();
        } catch (err) {
            console.error(err);
            showToast({
                variant: 'error',
                title: 'PNG 匯出失敗',
                message: err instanceof Error ? err.message : String(err)
            });
        }
    };

    return panel;
}

/* ================ helpers ================ */

function makeMetricSwitch(
    initial: 'amount' | 'count',
    onChange: (v: 'amount' | 'count') => void,
    getCurrent: () => 'amount' | 'count'
): HTMLElement {
    const wrap = document.createElement('div');
    wrap.className = 'analytics-metric-switch';
    const renderBtns = () => {
        const cur = getCurrent();
        wrap.innerHTML = `
      <button type="button" class="${cur === 'amount' ? 'is-active' : ''}" data-v="amount">金額</button>
      <button type="button" class="${cur === 'count' ? 'is-active' : ''}" data-v="count">數量</button>
    `;
        wrap.querySelectorAll<HTMLButtonElement>('button').forEach((b) => {
            b.addEventListener('click', () => {
                const v = b.dataset.v as 'amount' | 'count';
                if (v !== getCurrent()) {
                    onChange(v);
                    renderBtns();
                }
            });
        });
    };
    renderBtns();
    void initial;
    return wrap;
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}

function formatBytes(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / 1024 / 1024).toFixed(2)} MB`;
}
