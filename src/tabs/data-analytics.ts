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
import {loadCostMap} from '@/analytics/cost-loader';
import {
    type AnalyticsDataset,
    type AnalyticsRow,
    buildDataset,
    type LoadedFileMeta,
} from '@/analytics/dataset-builder';
import {applyFilter, type FilterState} from '@/analytics/filter-engine';
import {createFilterUi, defaultFilterState} from '@/analytics/filter-ui';
import {createDetailTable, rowsToCsv} from '@/analytics/detail-table';
import {createLeastProfitableTable} from '@/analytics/least-profitable-table';
import {createProductPriceVarianceTable} from '@/analytics/product-price-variance-table';
import {openUncategorizedDialog} from '@/analytics/uncategorized-dialog';
import {openUnsetCostDialog} from '@/analytics/unset-cost-dialog';
import {type ChartHandle, createChart, observeChartsResize} from '@/analytics/chart-manager';
import {computeKpi, renderKpiCards} from '@/analytics/kpi';
import {
    categoryTreemapOption,
    customerParetoOption,
    customerTopOption,
    dailyTrendOption,
    linePieOption,
    type MetricKind,
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

/** 金額/數量/利潤切換按鈕 handle：暴露 element 與一個由外部觸發的 refresh */
interface MetricSwitchHandle {
    element: HTMLElement;
    refresh: () => void;
}

/** 圖表 slot：對應 panel 內某個圖表卡的容器 + chart instance */
interface ChartSlot {
    id: string;
    title: string;
    container: HTMLElement;
    pngBtn: HTMLButtonElement;
    expandBtn: HTMLButtonElement;
    chart: ChartHandle | null;
    modalChart: ChartHandle | null;
    /** 同一個 slot 可能有 header 與 modal 兩個 metric switch，需同步重繪狀態 */
    metricRefreshers: Set<() => void>;
    /** 用於 modal 內動態建立同樣的 metric switch；無 controls 的 slot 為 null */
    makeControls: (() => MetricSwitchHandle) | null;
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
      </div>

      <section class="analytics-content" data-role="content" hidden>
        <div class="analytics-section-label">總覽 KPI</div>
        <div class="analytics-kpi-grid" data-role="kpi-grid"></div>

        <section class="analytics-filters-section" data-role="filters-section" hidden>
          <div data-role="filters-host"></div>
        </section>

        <div class="analytics-section-label">趨勢與分佈</div>
        <div class="analytics-chart-grid" data-role="chart-grid"></div>

        <div class="analytics-section-label">客戶低價值排行（平均售價）</div>
        <p class="analytics-section-hint">以「平均售價（amount ÷ 數量）」由低至高排序，搭配「最低數量」門檻可聚焦在「買很多但平均單價偏低」的客戶。橘底＝平均售價低於整體均價；紅底＝該客戶整體毛利為負。會跟隨上方篩選器即時更新。</p>
        <div data-role="least-profit-host"></div>

        <div class="analytics-section-label">商品價差分析（同商品各客戶售價落差）</div>
        <p class="analytics-section-hint">以「商品」為單位，顯示該商品在所有客戶手中的整體加權平均售價與分散程度，並列出哪些客戶售價低於均價達指定門檻。預設依「均價缺口」（＝把低價客戶補到均價可多收的金額）由大到小排序，最能定位漏血點。點商品列可展開該商品的客戶層級明細。</p>
        <div data-role="price-variance-host"></div>

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

        <details class="analytics-unmatched" data-role="unset-cost" hidden>
          <summary><span data-role="unset-cost-summary"></span></summary>
          <ul class="analytics-unmatched-list" data-role="unset-cost-list"></ul>
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
    const filtersSection = panel.querySelector<HTMLElement>('[data-role="filters-section"]')!;
    const filtersHost = panel.querySelector<HTMLElement>('[data-role="filters-host"]')!;
    const contentSection = panel.querySelector<HTMLElement>('[data-role="content"]')!;
    const kpiHost = panel.querySelector<HTMLElement>('[data-role="kpi-grid"]')!;
    const chartGridHost = panel.querySelector<HTMLElement>('[data-role="chart-grid"]')!;
    const detailHost = panel.querySelector<HTMLElement>('[data-role="detail-host"]')!;
    const leastProfitHost = panel.querySelector<HTMLElement>('[data-role="least-profit-host"]')!;
    const priceVarianceHost = panel.querySelector<HTMLElement>('[data-role="price-variance-host"]')!;
    const exportCsvBtn = panel.querySelector<HTMLButtonElement>('[data-role="export-csv"]')!;
    const unmatchedHost = panel.querySelector<HTMLDetailsElement>('[data-role="unmatched"]')!;
    const unmatchedSummary = panel.querySelector<HTMLElement>('[data-role="unmatched-summary"]')!;
    const unmatchedList = panel.querySelector<HTMLElement>('[data-role="unmatched-list"]')!;
    const unsetCostHost = panel.querySelector<HTMLDetailsElement>('[data-role="unset-cost"]')!;
    const unsetCostSummary = panel.querySelector<HTMLElement>('[data-role="unset-cost-summary"]')!;
    const unsetCostList = panel.querySelector<HTMLElement>('[data-role="unset-cost-list"]')!;
    const statusEl = panel.querySelector<HTMLElement>('[data-role="status"]')!;

    // ===== 狀態 =====
    const loaded: LoadedBillRecord[] = [];
    let dataset: AnalyticsDataset | null = null;
    let filterState: FilterState = defaultFilterState();
    let resizeCleanup: (() => void) | null = null;

    // ===== 圖表切換按鈕的當下狀態（每張圖獨立） =====
    let dailyMetric: MetricKind = 'amount';
    let productMetric: MetricKind = 'amount';
    let weekdayMetric: MetricKind = 'amount';
    let linePieMetric: MetricKind = 'amount';
    let customerTopMetric: MetricKind = 'amount';
    let categoryMetric: MetricKind = 'amount';
    let paretoMetric: MetricKind = 'amount';

    // ===== 子元件 =====
    const detailTable = createDetailTable();
    detailHost.appendChild(detailTable.element);

    const leastProfitableTable = createLeastProfitableTable({
        onCustomerClick: (code) => {
            filterUi.applyPatch({customerCodes: new Set([code])});
            filtersSection.scrollIntoView({behavior: 'smooth', block: 'start'});
        },
    });
    leastProfitHost.appendChild(leastProfitableTable.element);

    const priceVarianceTable = createProductPriceVarianceTable({
        onProductFilter: (productName) => {
            filterUi.applyPatch({productNames: new Set([productName])});
            filtersSection.scrollIntoView({behavior: 'smooth', block: 'start'});
        },
        onCustomerClick: (code) => {
            filterUi.applyPatch({customerCodes: new Set([code])});
            filtersSection.scrollIntoView({behavior: 'smooth', block: 'start'});
        },
    });
    priceVarianceHost.appendChild(priceVarianceTable.element);

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
            makeControls?: (slot: ChartSlot) => MetricSwitchHandle;
            build: ChartSlot['build'];
        }

        // metric switch 共用 helper：onChange 由呼叫端決定如何更新狀態
        const metricSwitchFor = (
            slot: ChartSlot,
            getter: () => MetricKind,
            setter: (v: MetricKind) => void
        ): MetricSwitchHandle =>
            makeMetricSwitch((v) => {
                setter(v);
                slot.metricRefreshers.forEach((fn) => fn());
                void renderAll();
            }, getter);

        const list: SlotDef[] = [
            {
                id: 'daily-trend',
                title: '每日營收趨勢',
                wide: true,
                makeControls: (slot) => metricSwitchFor(slot, () => dailyMetric, (v) => {
                    dailyMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : dailyTrendOption(rows, dailyMetric)),
            },
            {
                id: 'weekday-heat',
                title: '星期銷售熱度',
                makeControls: (slot) => metricSwitchFor(slot, () => weekdayMetric, (v) => {
                    weekdayMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : weekdayOption(rows, weekdayMetric)),
            },
            {
                id: 'product-top',
                title: '商品銷售 Top 10',
                makeControls: (slot) => metricSwitchFor(slot, () => productMetric, (v) => {
                    productMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : productTopOption(rows, productMetric)),
            },
            {
                id: 'customer-top',
                title: '客戶銷售 Top 10',
                makeControls: (slot) => metricSwitchFor(slot, () => customerTopMetric, (v) => {
                    customerTopMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : customerTopOption(rows, customerTopMetric)),
            },
            {
                id: 'line-pie',
                title: '線別佔比',
                makeControls: (slot) => metricSwitchFor(slot, () => linePieMetric, (v) => {
                    linePieMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : linePieOption(rows, linePieMetric)),
            },
            {
                id: 'customer-pareto',
                title: '客戶 80/20 帕累托',
                wide: true,
                makeControls: (slot) => metricSwitchFor(slot, () => paretoMetric, (v) => {
                    paretoMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : customerParetoOption(rows, paretoMetric)),
            },
            {
                id: 'category-treemap',
                title: '商品分類佔比',
                wide: true,
                makeControls: (slot) => metricSwitchFor(slot, () => categoryMetric, (v) => {
                    categoryMetric = v;
                }),
                build: (rows) => (rows.length === 0 ? null : categoryTreemapOption(rows, categoryMetric)),
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

            const pngBtn = document.createElement('button');
            pngBtn.type = 'button';
            pngBtn.className = 'analytics-chart-png-btn';
            pngBtn.title = '下載 PNG';
            pngBtn.innerHTML = icon('download', 14);

            const expandBtn = document.createElement('button');
            expandBtn.type = 'button';
            expandBtn.className = 'analytics-chart-expand-btn';
            expandBtn.title = '放大檢視';
            expandBtn.setAttribute('aria-label', '放大檢視');
            expandBtn.innerHTML = icon('maximize', 14);

            const body = document.createElement('div');
            body.className = 'analytics-chart-body';

            const slot: ChartSlot = {
                id: s.id,
                title: s.title,
                container: body,
                pngBtn,
                expandBtn,
                chart: null,
                modalChart: null,
                metricRefreshers: new Set(),
                makeControls: null,
                build: s.build,
            };
            if (s.makeControls) {
                const factory = s.makeControls;
                slot.makeControls = () => factory(slot);
                const ms = slot.makeControls();
                slot.metricRefreshers.add(ms.refresh);
                headerRight.appendChild(ms.element);
            }

            headerRight.appendChild(pngBtn);
            headerRight.appendChild(expandBtn);
            header.appendChild(headerRight);

            card.appendChild(header);
            card.appendChild(body);
            chartGridHost.appendChild(card);

            expandBtn.addEventListener('click', () => openChartModal(slot));

            built.push(slot);
        }
        return built;
    }

    // ===== 鑽取（chart click → filterUi.applyPatch） =====
    const wireDrillDown = (slot: ChartSlot, handle: ChartHandle, onTrigger?: () => void) => {
        handle.off('click');
        handle.on('click', (params) => {
            const p = params as { seriesType?: string; name?: string; data?: unknown };
            let triggered = false;
            if (slot.id === 'line-pie' && p.name) {
                filterUi.applyPatch({lines: new Set([p.name])});
                triggered = true;
            } else if (slot.id === 'product-top' && p.name) {
                filterUi.applyPatch({productNames: new Set([p.name])});
                detailTable.scrollIntoView();
                triggered = true;
            } else if (slot.id === 'customer-top' && p.name) {
                // p.name 形如「客戶名(代碼)」，反推回 code
                const m = /\(([^()]+)\)$/.exec(p.name);
                if (m) {
                    filterUi.applyPatch({customerCodes: new Set([m[1]])});
                    triggered = true;
                }
            }
            if (triggered) onTrigger?.();
        });
    };

    // ===== mount / render =====
    const mountChartsIfNeeded = async () => {
        for (const slot of slots) {
            if (slot.chart) continue;
            try {
                slot.chart = await createChart(slot.container);
                wireDrillDown(slot, slot.chart);
                slot.pngBtn.addEventListener('click', () => exportSlotPng(slot));
            } catch (err) {
                console.error('[analytics] chart mount failed', slot.id, err);
            }
        }
        if (resizeCleanup) resizeCleanup();
        const handles = slots.map((s) => s.chart).filter((c): c is ChartHandle => c !== null);
        resizeCleanup = observeChartsResize(handles);
    };

    const EMPTY_CHART_OPTION = {
        title: {
            text: '無資料',
            left: 'center',
            top: 'center',
            textStyle: {fontSize: 13, color: '#999'},
        },
    } as const;

    const renderAll = async () => {
        if (!dataset) return;
        const filtered = applyFilter(dataset, filterState);
        const filteredRows = filtered.rows;

        // KPI
        renderKpiCards(kpiHost, computeKpi(filtered));

        // 圖表
        await mountChartsIfNeeded();
        for (const slot of slots) {
            const opt = (slot.chart || slot.modalChart) ? slot.build(filteredRows, dataset) : null;
            const apply = (handle: ChartHandle) => {
                if (opt) handle.setOption(opt as never);
                else handle.setOption(EMPTY_CHART_OPTION as never);
            };
            if (slot.chart) {
                apply(slot.chart);
                slot.container.parentElement!.classList.remove('is-empty');
            }
            if (slot.modalChart) apply(slot.modalChart);
            // month-over-month 只在多檔時顯示卡片
            if (slot.id === 'month-over-month') {
                slot.container.parentElement!.hidden = dataset.files.length < 2;
            }
        }

        // 客戶毛利倒數排行
        leastProfitableTable.setRows(filteredRows);

        // 商品價差分析
        priceVarianceTable.setRows(filteredRows);

        // 明細表
        detailTable.setRows(filteredRows);

        statusEl.textContent = `已載入 ${dataset.files.length} 檔 / 篩選後 ${filteredRows.length.toLocaleString()} 筆 / 原始 ${dataset.rows.length.toLocaleString()} 筆`;
    };

    // ===== 圖表放大彈窗 =====
    const openChartModal = (slot: ChartSlot) => {
        if (!dataset) return;

        const backdrop = document.createElement('div');
        backdrop.className = 'app-modal-backdrop';

        const modal = document.createElement('div');
        modal.className = 'app-modal app-modal--chart';

        const header = document.createElement('div');
        header.className = 'app-modal-header';

        const titleEl = document.createElement('h2');
        titleEl.className = 'app-modal-title';
        titleEl.textContent = slot.title;

        const closeBtn = document.createElement('button');
        closeBtn.type = 'button';
        closeBtn.className = 'app-modal-close';
        closeBtn.setAttribute('aria-label', '關閉');
        closeBtn.innerHTML = icon('close', 18);

        header.appendChild(titleEl);
        header.appendChild(closeBtn);

        const body = document.createElement('div');
        body.className = 'app-modal-body';

        const toolbar = document.createElement('div');
        toolbar.className = 'analytics-chart-modal-toolbar';

        let modalSwitchRefresh: (() => void) | null = null;
        if (slot.makeControls) {
            const ms = slot.makeControls();
            modalSwitchRefresh = ms.refresh;
            slot.metricRefreshers.add(ms.refresh);
            toolbar.appendChild(ms.element);
        }

        const chartHost = document.createElement('div');
        chartHost.className = 'analytics-chart-modal-body';

        body.appendChild(toolbar);
        body.appendChild(chartHost);

        modal.appendChild(header);
        modal.appendChild(body);
        backdrop.appendChild(modal);
        document.body.appendChild(backdrop);

        let isClosed = false;
        let resizeObserver: ResizeObserver | null = null;
        const close = () => {
            if (isClosed) return;
            isClosed = true;
            document.removeEventListener('keydown', onKey);
            if (modalSwitchRefresh) slot.metricRefreshers.delete(modalSwitchRefresh);
            if (resizeObserver) resizeObserver.disconnect();
            if (slot.modalChart) {
                slot.modalChart.dispose();
                slot.modalChart = null;
            }
            backdrop.remove();
        };
        const onKey = (e: KeyboardEvent) => {
            if (e.key === 'Escape') {
                e.preventDefault();
                close();
            }
        };

        closeBtn.addEventListener('click', close);
        document.addEventListener('keydown', onKey);

        void (async () => {
            try {
                const handle = await createChart(chartHost);
                if (isClosed) {
                    handle.dispose();
                    return;
                }
                slot.modalChart = handle;
                wireDrillDown(slot, handle, close);
                const filtered = applyFilter(dataset!, filterState);
                const opt = slot.build(filtered.rows, dataset!);
                handle.setOption((opt ?? EMPTY_CHART_OPTION) as never);
                resizeObserver = new ResizeObserver(() => handle.resize());
                resizeObserver.observe(chartHost);
                requestAnimationFrame(() => handle.resize());
            } catch (err) {
                console.error('[analytics] modal chart mount failed', err);
            }
        })();
    };

    // ===== 上傳/解析 =====
    const fileId = (f: File) => `${f.name}__${f.size}`;

    const renderFileList = () => {
        if (loaded.length === 0) {
            fileListEl.hidden = true;
            fileListEl.innerHTML = '';
            return;
        }
        fileListEl.hidden = false;
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
            unsetCostHost.hidden = true;
            statusEl.textContent = '請拖入或選擇 .xlsx 帳單檔';
            return;
        }
        statusEl.textContent = '聚合資料中…';
        try {
            const [categoryMap, costMap] = await Promise.all([loadCategoryMap(), loadCostMap()]);
            const fileMetas: LoadedFileMeta[] = loaded.map((rec) => ({
                id: rec.fileId,
                name: rec.fileName,
                bill: rec.bill,
            }));
            dataset = buildDataset(fileMetas, categoryMap, costMap);

            filtersSection.hidden = false;
            contentSection.hidden = false;
            filterUi.setDataset(dataset);
            renderUnmatched(dataset);
            renderUnsetCost(dataset);

            console.info('[analytics] dataset built:', {
                files: dataset.files.length,
                rows: dataset.rows.length,
                unmatchedProducts: dataset.unmatchedProducts.length,
                unsetCostProducts: dataset.unsetCostProducts.length,
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
        unmatchedSummary.textContent = `未分類商品（${ds.unmatchedProducts.length}）— 點擊任一項即可加入分類`;
        unmatchedList.innerHTML = ds.unmatchedProducts
            .map(
                (name) =>
                    `<li><button type="button" class="analytics-unmatched-chip" data-name="${escapeHtml(name)}">${escapeHtml(name)}</button></li>`
            )
            .join('');
        unmatchedList.querySelectorAll<HTMLButtonElement>('.analytics-unmatched-chip').forEach((btn) => {
            btn.addEventListener('click', () => {
                const name = btn.dataset.name ?? '';
                if (!name) return;
                void openUncategorizedDialog({
                    productName: name,
                    onSaved: async () => {
                        await rebuildDataset();
                    },
                });
            });
        });
    };

    const renderUnsetCost = (ds: AnalyticsDataset) => {
        if (ds.unsetCostProducts.length === 0) {
            unsetCostHost.hidden = true;
            return;
        }
        unsetCostHost.hidden = false;
        unsetCostSummary.textContent = `未填成本商品（${ds.unsetCostProducts.length}）— 點擊任一項即可填入成本`;
        unsetCostList.innerHTML = ds.unsetCostProducts
            .map(
                (name) =>
                    `<li><button type="button" class="analytics-unmatched-chip" data-name="${escapeHtml(name)}">${escapeHtml(name)}</button></li>`
            )
            .join('');
        unsetCostList.querySelectorAll<HTMLButtonElement>('.analytics-unmatched-chip').forEach((btn) => {
            btn.addEventListener('click', () => {
                const name = btn.dataset.name ?? '';
                if (!name) return;
                void openUnsetCostDialog({
                    productName: name,
                    onSaved: async () => {
                        await rebuildDataset();
                    },
                });
            });
        });
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
    onChange: (v: MetricKind) => void,
    getCurrent: () => MetricKind
): MetricSwitchHandle {
    const wrap = document.createElement('div');
    wrap.className = 'analytics-metric-switch';
    const renderBtns = () => {
        const cur = getCurrent();
        wrap.innerHTML = `
      <button type="button" class="${cur === 'amount' ? 'is-active' : ''}" data-v="amount">金額</button>
      <button type="button" class="${cur === 'count' ? 'is-active' : ''}" data-v="count">數量</button>
      <button type="button" class="${cur === 'profit' ? 'is-active' : ''}" data-v="profit">利潤</button>
    `;
        wrap.querySelectorAll<HTMLButtonElement>('button').forEach((b) => {
            b.addEventListener('click', () => {
                const v = b.dataset.v as MetricKind;
                if (v !== getCurrent()) {
                    // onChange 由呼叫端串接 metricRefreshers，會自動重繪本元件
                    onChange(v);
                }
            });
        });
    };
    renderBtns();
    return {element: wrap, refresh: renderBtns};
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
