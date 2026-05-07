/**
 * 跨月數據分析分頁（首輪 MVP，包含 A1 / A2 / B1 / C1）。
 *
 * 與「數據分析」分頁獨立：本分頁自行管理上傳檔案，不共享狀態。
 * 月份切割以 row.year + row.month 為準，與檔案邊界解耦。
 *
 * 對比策略：所有月份依時間升冪排序，對「每個相鄰月對」渲染一個區塊
 *           （4月 vs 3月、5月 vs 4月、6月 vs 5月 …）。
 *           趨勢折線圖則畫所有月份的連續折線。
 */

import {showToast} from '@/ui/toast';
import {icon} from '@/ui/icons';
import type {TabDefinition} from '@/ui/tabs';
import {parseBillFile} from '@/readers/bill-reader';
import {Bill} from '@/domain/models/bill';
import {ExcelRowType} from '@/domain/excel-row-data';
import {loadCategoryMap} from '@/analytics/category-loader';
import {loadCostMap} from '@/analytics/cost-loader';
import {type AnalyticsDataset, buildDataset, type LoadedFileMeta,} from '@/analytics/dataset-builder';
import {type ChartHandle, createChart, observeChartsResize,} from '@/analytics/chart-manager';
import {
    computeMoM,
    customerSegmentation,
    detectProductPriceChanges,
    type MoMComparison,
    type MonthKey,
    type MonthlyTotals,
    monthlyTotals,
    rowsOfMonth,
} from '@/analytics/cross-month/month-aggregators';
import {monthlyTrendOption, TREND_METRICS, type TrendMetric,} from '@/analytics/cross-month/monthly-trend-chart';
import {createCustomerSegmentTable} from '@/analytics/cross-month/customer-segment-table';
import {createPriceChangeTable} from '@/analytics/cross-month/price-change-table';

interface LoadedBillRecord {
    fileId: string;
    fileName: string;
    fileSize: number;
    bill: Bill;
}

const fmtMoney = (v: number) => v.toLocaleString('zh-TW');
const fmtCount = (v: number) => v.toLocaleString('zh-TW');
const fmtPct = (v: number | null) =>
    v === null ? '—' : `${v >= 0 ? '+' : ''}${v.toFixed(1)}%`;
const fmtDelta = (v: number) => `${v >= 0 ? '+' : ''}${fmtMoney(Math.round(v))}`;
const fmtCountDelta = (v: number) => `${v >= 0 ? '+' : ''}${fmtCount(v)}`;

export function renderCrossMonthAnalyticsPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card analytics-card">
      <header class="card-header">
        <h1 class="card-title">跨月數據分析</h1>
        <p class="card-subtitle">
          拖入兩份以上不同月份的「明細編號順序 .xlsx」帳單檔，系統會把資料依「年/月」拆組，
          產出月度趨勢、相鄰月份的 MoM 環比、客戶新增/流失/留存、商品漲價影響等跨期分析。
        </p>
      </header>

      <div class="analytics-uploader">
        <div class="analytics-dropzone" data-role="dropzone" tabindex="0" role="button">
          <span class="analytics-dropzone-icon">${icon('upload', 28)}</span>
          <div class="analytics-dropzone-title">拖入或點擊選擇 .xlsx 帳單檔（≥ 2 個月份）</div>
          <div class="analytics-dropzone-hint">同一月份多檔會自動合併；不同月份各自獨立計算</div>
          <input type="file" accept=".xlsx,.xls" multiple hidden data-role="file-input">
        </div>
        <div class="analytics-file-list" data-role="file-list" hidden></div>
      </div>

      <div class="cross-month-warning" data-role="warning" hidden></div>

      <section class="analytics-content" data-role="content" hidden>
        <div class="analytics-section-label">月度趨勢</div>
        <p class="analytics-section-hint">所有載入月份依時間順序的折線。可切換指標。</p>
        <div class="analytics-chart-card analytics-chart-card-wide cross-month-trend-card">
          <div class="analytics-chart-header">
            <div class="analytics-chart-title">月度 趨勢</div>
            <div class="analytics-chart-header-right" data-role="trend-controls"></div>
          </div>
          <div class="analytics-chart-body" data-role="trend-body"></div>
        </div>

        <div class="analytics-section-label">相鄰月份對比</div>
        <p class="analytics-section-hint">
          所有月份依時間升冪排序，對每組相鄰月份呈現 MoM 卡片、客戶新增/流失/留存、商品漲價影響。
          每個區塊的左側顏色條：<span style="color:var(--color-success)">綠</span>＝營收成長、<span style="color:var(--color-danger)">紅</span>＝營收衰退。
        </p>
        <div data-role="pairs-host"></div>
      </section>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="status">請拖入或選擇 ≥ 2 份月份的 .xlsx 帳單檔</div>
      </footer>
    </div>
  `;

    const dropzone = panel.querySelector<HTMLElement>('[data-role="dropzone"]')!;
    const fileInput = panel.querySelector<HTMLInputElement>('[data-role="file-input"]')!;
    const fileListEl = panel.querySelector<HTMLElement>('[data-role="file-list"]')!;
    const warningEl = panel.querySelector<HTMLElement>('[data-role="warning"]')!;
    const contentSection = panel.querySelector<HTMLElement>('[data-role="content"]')!;
    const trendBody = panel.querySelector<HTMLElement>('[data-role="trend-body"]')!;
    const trendControls = panel.querySelector<HTMLElement>('[data-role="trend-controls"]')!;
    const pairsHost = panel.querySelector<HTMLElement>('[data-role="pairs-host"]')!;
    const statusEl = panel.querySelector<HTMLElement>('[data-role="status"]')!;

    const loaded: LoadedBillRecord[] = [];
    let dataset: AnalyticsDataset | null = null;
    let monthly: MonthlyTotals[] = [];
    let trendMetric: TrendMetric = 'amount';
    let trendChart: ChartHandle | null = null;
    let resizeCleanup: (() => void) | null = null;

    /* ============== 趨勢圖 metric 切換 ============== */
    const renderTrendControls = () => {
        trendControls.innerHTML = '';
        const wrap = document.createElement('div');
        wrap.className = 'analytics-metric-switch';
        for (const m of TREND_METRICS) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.textContent = m.label;
            btn.className = m.key === trendMetric ? 'is-active' : '';
            btn.addEventListener('click', () => {
                if (m.key === trendMetric) return;
                trendMetric = m.key;
                renderTrendControls();
                renderTrend();
            });
            wrap.appendChild(btn);
        }
        trendControls.appendChild(wrap);
    };

    const renderTrend = async () => {
        if (monthly.length === 0) return;
        if (!trendChart) {
            try {
                trendChart = await createChart(trendBody);
            } catch (err) {
                console.error('[cross-month] chart mount failed', err);
                return;
            }
            if (resizeCleanup) resizeCleanup();
            resizeCleanup = observeChartsResize([trendChart]);
        }
        trendChart.setOption(monthlyTrendOption(monthly, trendMetric) as never);
    };

    /* ============== 相鄰月對區塊 ============== */
    const renderPairs = () => {
        pairsHost.innerHTML = '';
        if (!dataset || monthly.length < 2) return;

        // monthly 已是時間升冪。對每對 (i-1, i) 渲染一個區塊。
        for (let i = 1; i < monthly.length; i++) {
            const prev = monthly[i - 1];
            const cur = monthly[i];
            const mom = computeMoM(cur, prev);
            const curRows = rowsOfMonth(dataset.rows, cur.key.sortKey);
            const prevRows = rowsOfMonth(dataset.rows, prev.key.sortKey);
            const seg = customerSegmentation(curRows, prevRows);
            const priceChanges = detectProductPriceChanges(curRows, prevRows, 0);

            const section = buildPairSection(cur.key, prev.key, mom, seg, priceChanges);
            pairsHost.appendChild(section);
        }
    };

    /* ============== 上傳/解析 ============== */
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

    const showWarning = (message: string | null) => {
        if (!message) {
            warningEl.hidden = true;
            warningEl.textContent = '';
            return;
        }
        warningEl.hidden = false;
        warningEl.textContent = message;
    };

    const rebuildDataset = async () => {
        if (loaded.length === 0) {
            dataset = null;
            monthly = [];
            contentSection.hidden = true;
            showWarning(null);
            statusEl.textContent = '請拖入或選擇 ≥ 2 份月份的 .xlsx 帳單檔';
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
            monthly = monthlyTotals(dataset.rows);

            if (monthly.length < 2) {
                showWarning('目前只載入 1 個月份，需要 ≥ 2 個月份才能做跨月對比。');
                contentSection.hidden = true;
                statusEl.textContent = `已載入 ${dataset.files.length} 檔，跨月分析需要 ≥ 2 個月份`;
                return;
            }

            showWarning(null);
            contentSection.hidden = false;
            renderTrendControls();
            await renderTrend();
            renderPairs();

            statusEl.textContent = `已載入 ${dataset.files.length} 檔 / ${monthly.length} 個月份 / ${monthly.length - 1} 組相鄰對比 / 共 ${dataset.rows.length.toLocaleString()} 筆`;
        } catch (err) {
            console.error('[cross-month] rebuild failed', err);
            showToast({
                variant: 'error',
                title: '聚合失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            statusEl.textContent = '聚合失敗，請檢查檔案';
        }
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
                console.error('[cross-month] parse failed', file.name, err);
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

    /* ============== 拖曳事件 ============== */
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

    return panel;
}

// 重新匯出供外部 import
export type {MonthKey};

/* ====================== 區塊組裝 ====================== */

function buildPairSection(
    cur: MonthKey,
    prev: MonthKey,
    mom: MoMComparison,
    seg: ReturnType<typeof customerSegmentation>,
    priceChanges: ReturnType<typeof detectProductPriceChanges>
): HTMLElement {
    const amountDelta = mom.diffs.amount.delta;
    const tone = amountDelta > 0 ? 'is-positive' : amountDelta < 0 ? 'is-negative' : 'is-neutral';
    const headerArrow = amountDelta > 0 ? '▲' : amountDelta < 0 ? '▼' : '·';

    const section = document.createElement('section');
    section.className = `cross-month-pair ${tone}`;
    section.innerHTML = `
    <header class="cross-month-pair-header">
      <button type="button" class="cross-month-pair-toggle" data-role="toggle" aria-expanded="true">
        <span class="cross-month-pair-chev">▾</span>
        <h3 class="cross-month-pair-title">${cur.label} <span class="cross-month-pair-arrow">←</span> ${prev.label}</h3>
      </button>
      <span class="cross-month-pair-summary ${tone}">
        <span class="cross-month-pair-summary-arrow">${headerArrow}</span>
        營收 ${fmtDelta(amountDelta)}
        ${mom.diffs.amount.pct !== null ? `（${fmtPct(mom.diffs.amount.pct)}）` : ''}
      </span>
    </header>
    <div class="cross-month-pair-body" data-role="body">
      <div class="cross-month-mom-grid" data-role="mom-grid"></div>
      <div class="cross-month-pair-subhead">客戶新增 / 流失 / 留存</div>
      <div data-role="segment-host"></div>
      <div class="cross-month-pair-subhead">商品漲價影響</div>
      <div data-role="price-host"></div>
    </div>
  `;

    const momGrid = section.querySelector<HTMLElement>('[data-role="mom-grid"]')!;
    renderMomCards(momGrid, mom);

    const segmentHost = section.querySelector<HTMLElement>('[data-role="segment-host"]')!;
    const segmentTable = createCustomerSegmentTable();
    segmentHost.appendChild(segmentTable.element);
    segmentTable.setData(seg, cur.label, prev.label);

    const priceHost = section.querySelector<HTMLElement>('[data-role="price-host"]')!;
    const priceChangeTable = createPriceChangeTable();
    priceHost.appendChild(priceChangeTable.element);
    priceChangeTable.setData(priceChanges, cur.label, prev.label);

    // 收合/展開
    const toggle = section.querySelector<HTMLButtonElement>('[data-role="toggle"]')!;
    const body = section.querySelector<HTMLElement>('[data-role="body"]')!;
    const chev = section.querySelector<HTMLElement>('.cross-month-pair-chev')!;
    toggle.addEventListener('click', () => {
        const expanded = toggle.getAttribute('aria-expanded') === 'true';
        toggle.setAttribute('aria-expanded', String(!expanded));
        body.hidden = expanded;
        chev.textContent = expanded ? '▸' : '▾';
    });

    return section;
}

function renderMomCards(host: HTMLElement, mom: MoMComparison): void {
    const cards: Array<{
        label: string;
        currValue: string;
        delta: number;
        deltaLabel: string;
        pct: number | null;
        higherIsBetter: boolean;
    }> = [
        {
            label: '營收',
            currValue: fmtMoney(Math.round(mom.current.amount)),
            delta: mom.diffs.amount.delta,
            deltaLabel: fmtDelta(mom.diffs.amount.delta),
            pct: mom.diffs.amount.pct,
            higherIsBetter: true,
        },
        {
            label: '毛利',
            currValue: fmtMoney(Math.round(mom.current.profit)),
            delta: mom.diffs.profit.delta,
            deltaLabel: fmtDelta(mom.diffs.profit.delta),
            pct: mom.diffs.profit.pct,
            higherIsBetter: true,
        },
        {
            label: '銷售數量',
            currValue: fmtCount(mom.current.count),
            delta: mom.diffs.count.delta,
            deltaLabel: fmtCountDelta(mom.diffs.count.delta),
            pct: mom.diffs.count.pct,
            higherIsBetter: true,
        },
        {
            label: '客戶數',
            currValue: fmtCount(mom.current.customerCount),
            delta: mom.diffs.customerCount.delta,
            deltaLabel: fmtCountDelta(mom.diffs.customerCount.delta),
            pct: mom.diffs.customerCount.pct,
            higherIsBetter: true,
        },
        {
            label: '商品種類',
            currValue: fmtCount(mom.current.productCount),
            delta: mom.diffs.productCount.delta,
            deltaLabel: fmtCountDelta(mom.diffs.productCount.delta),
            pct: mom.diffs.productCount.pct,
            higherIsBetter: true,
        },
        {
            label: '毛利率',
            currValue: mom.current.amount > 0
                ? `${((mom.current.profit / mom.current.amount) * 100).toFixed(1)}%`
                : '—',
            delta: mom.diffs.marginPct.delta,
            deltaLabel: `${mom.diffs.marginPct.delta >= 0 ? '+' : ''}${mom.diffs.marginPct.delta.toFixed(1)}pp`,
            pct: null,
            higherIsBetter: true,
        },
    ];

    host.innerHTML = cards
        .map((c) => {
            const tone = (c.higherIsBetter ? c.delta >= 0 : c.delta <= 0) ? 'is-positive' : 'is-negative';
            const arrow = c.delta > 0 ? '▲' : c.delta < 0 ? '▼' : '·';
            return `
        <div class="cross-month-mom-card ${tone}">
          <div class="cross-month-mom-label">${c.label}</div>
          <div class="cross-month-mom-value">${c.currValue}</div>
          <div class="cross-month-mom-delta ${tone}">
            <span class="cross-month-mom-arrow">${arrow}</span>
            ${c.deltaLabel}${c.pct !== null ? `（${fmtPct(c.pct)}）` : ''}
          </div>
        </div>
      `;
        })
        .join('');
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
