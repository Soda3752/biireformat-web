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
    type PeriodTotals,
    periodTotals,
    rowsOfMonth,
    DEFAULT_TREND_DAYS_PER_PERIOD,
    MIN_TREND_DAYS_PER_PERIOD,
    MAX_TREND_DAYS_PER_PERIOD,
    clampTrendDaysPerPeriod,
} from '@/analytics/cross-month/month-aggregators';
import {monthlyTrendOption, TREND_METRICS, type TrendMetric,} from '@/analytics/cross-month/monthly-trend-chart';
import {createCustomerSegmentTable, type MonthTotalsRef} from '@/analytics/cross-month/customer-segment-table';
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
        <div class="analytics-section-label" data-role="trend-section-label">每 ${DEFAULT_TREND_DAYS_PER_PERIOD} 日趨勢</div>
        <p class="analytics-section-hint" data-role="trend-section-hint">所有載入月份依時間順序的折線，每 ${DEFAULT_TREND_DAYS_PER_PERIOD} 日為一個基準點（連續切段、跨月不重置）。可切換指標、調整每段天數（${MIN_TREND_DAYS_PER_PERIOD}~${MAX_TREND_DAYS_PER_PERIOD} 日）。</p>
        <div class="analytics-chart-card analytics-chart-card-wide cross-month-trend-card">
          <div class="analytics-chart-header">
            <div class="analytics-chart-title" data-role="trend-title">每 ${DEFAULT_TREND_DAYS_PER_PERIOD} 日 趨勢</div>
            <div class="analytics-chart-header-right" data-role="trend-controls"></div>
          </div>
          <div class="analytics-chart-body" data-role="trend-body"></div>
        </div>

        <div class="analytics-section-label">相鄰月份對比</div>
        <p class="analytics-section-hint">
          所有月份依時間升冪排序，對每組相鄰月份呈現 MoM 卡片、客戶新增/流失/留存、商品漲價影響。
          每個區塊的左側顏色條：<span style="color:var(--color-success)">綠</span>＝日平均營收成長、<span style="color:var(--color-danger)">紅</span>＝日平均營收衰退。
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
    const trendSectionLabel = panel.querySelector<HTMLElement>('[data-role="trend-section-label"]')!;
    const trendSectionHint = panel.querySelector<HTMLElement>('[data-role="trend-section-hint"]')!;
    const trendTitle = panel.querySelector<HTMLElement>('[data-role="trend-title"]')!;
    const pairsHost = panel.querySelector<HTMLElement>('[data-role="pairs-host"]')!;
    const statusEl = panel.querySelector<HTMLElement>('[data-role="status"]')!;

    const loaded: LoadedBillRecord[] = [];
    let dataset: AnalyticsDataset | null = null;
    let monthly: MonthlyTotals[] = [];
    let periods: PeriodTotals[] = [];
    let trendMetric: TrendMetric = 'amount';
    let daysPerPeriod: number = DEFAULT_TREND_DAYS_PER_PERIOD;
    let trendChart: ChartHandle | null = null;
    let resizeCleanup: (() => void) | null = null;

    const updateTrendLabels = () => {
        trendSectionLabel.textContent = `每 ${daysPerPeriod} 日趨勢`;
        trendTitle.textContent = `每 ${daysPerPeriod} 日 趨勢`;
        trendSectionHint.textContent =
            `所有載入月份依時間順序的折線，每 ${daysPerPeriod} 日為一個基準點（連續切段、跨月不重置）。可切換指標、調整每段天數（${MIN_TREND_DAYS_PER_PERIOD}~${MAX_TREND_DAYS_PER_PERIOD} 日）。`;
    };

    /* ============== 趨勢圖 metric 切換 ＋ 每段天數 ============== */
    const renderTrendControls = () => {
        trendControls.innerHTML = '';

        // 每段天數選擇器
        const daysWrap = document.createElement('label');
        daysWrap.className = 'cross-month-trend-days';
        daysWrap.innerHTML = `
            <span class="cross-month-trend-days-label">每段天數</span>
            <input
              type="number"
              class="cross-month-trend-days-input"
              min="${MIN_TREND_DAYS_PER_PERIOD}"
              max="${MAX_TREND_DAYS_PER_PERIOD}"
              step="1"
              inputmode="numeric"
              value="${daysPerPeriod}"
              aria-label="每段天數（${MIN_TREND_DAYS_PER_PERIOD}~${MAX_TREND_DAYS_PER_PERIOD} 日）"
            >
            <span class="cross-month-trend-days-suffix">日</span>
        `;
        const input = daysWrap.querySelector<HTMLInputElement>('input')!;
        const commit = () => {
            const next = clampTrendDaysPerPeriod(Number(input.value));
            if (next !== daysPerPeriod) {
                daysPerPeriod = next;
                updateTrendLabels();
                recomputePeriods();
                void renderTrend();
            }
            // 即使值不變也回填，避免使用者輸入越界值殘留在 UI
            input.value = String(daysPerPeriod);
        };
        input.addEventListener('change', commit);
        input.addEventListener('blur', commit);
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit();
                input.blur();
            }
        });
        trendControls.appendChild(daysWrap);

        // 指標切換
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

    const recomputePeriods = () => {
        if (!dataset) {
            periods = [];
            return;
        }
        periods = periodTotals(dataset.rows, daysPerPeriod);
    };

    const renderTrend = async () => {
        if (periods.length === 0) return;
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
        trendChart.setOption(monthlyTrendOption(periods, trendMetric) as never);
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
            periods = [];
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
            periods = periodTotals(dataset.rows, daysPerPeriod);

            if (monthly.length < 2) {
                showWarning('目前只載入 1 個月份，需要 ≥ 2 個月份才能做跨月對比。');
                contentSection.hidden = true;
                statusEl.textContent = `已載入 ${dataset.files.length} 檔，跨月分析需要 ≥ 2 個月份`;
                return;
            }

            showWarning(null);
            contentSection.hidden = false;
            updateTrendLabels();
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
    const dailyAmountDelta = mom.diffs.dailyAmount.delta;
    const tone = dailyAmountDelta > 0 ? 'is-positive' : dailyAmountDelta < 0 ? 'is-negative' : 'is-neutral';
    const headerArrow = dailyAmountDelta > 0 ? '▲' : dailyAmountDelta < 0 ? '▼' : '·';

    const section = document.createElement('section');
    section.className = `cross-month-pair ${tone}`;
    section.innerHTML = `
    <header class="cross-month-pair-header">
      <button type="button" class="cross-month-pair-toggle" data-role="toggle" aria-expanded="true">
        <span class="cross-month-pair-chev">▾</span>
        <h3 class="cross-month-pair-title">${prev.label} <span class="cross-month-pair-arrow">→</span> ${cur.label}</h3>
      </button>
      <span class="cross-month-pair-summary ${tone}">
        <span class="cross-month-pair-summary-arrow">${headerArrow}</span>
        每日平均營收 ${fmtDelta(dailyAmountDelta)}
        ${mom.diffs.dailyAmount.pct !== null ? `（${fmtPct(mom.diffs.dailyAmount.pct)}）` : ''}
      </span>
    </header>
    <div class="cross-month-pair-body" data-role="body">
      <div class="cross-month-mom-stack" data-role="mom-grid"></div>
      <div class="cross-month-pair-subhead">客戶分群（新增 / 零售新增 / 流失 / 零售流失 / 留存）</div>
      <div data-role="segment-host"></div>
      <div class="cross-month-pair-subhead">商品漲價影響</div>
      <div data-role="price-host"></div>
    </div>
  `;

    const momGrid = section.querySelector<HTMLElement>('[data-role="mom-grid"]')!;
    renderMomCards(momGrid, mom);

    const zeroMedian = {amount: 0, count: 0, profit: 0};
    const curRef: MonthTotalsRef = {
        label: cur.label,
        amount: mom.current.amount,
        profit: mom.current.profit,
        count: mom.current.count,
        median: mom.current.customerMedian,
    };
    const prevRef: MonthTotalsRef = mom.previous
        ? {
            label: prev.label,
            amount: mom.previous.amount,
            profit: mom.previous.profit,
            count: mom.previous.count,
            median: mom.previous.customerMedian,
        }
        : {label: prev.label, amount: 0, profit: 0, count: 0, median: zeroMedian};

    const segmentHost = section.querySelector<HTMLElement>('[data-role="segment-host"]')!;
    const segmentTable = createCustomerSegmentTable();
    segmentHost.appendChild(segmentTable.element);
    segmentTable.setData(seg, curRef, prevRef);

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

/** 兩月日曆天數合併標籤，例：「3月31天 / 4月30天」（時間順序：上月 → 本月） */
function dailyDaysSuffix(mom: MoMComparison): string {
    const cur = `${mom.current.key.month}月${mom.current.daysInMonth}天`;
    if (!mom.previous) return cur;
    const prev = `${mom.previous.key.month}月${mom.previous.daysInMonth}天`;
    return `${prev} / ${cur}`;
}

/** 兩月客戶數合併標籤，例：「409 客戶 / 411 客戶」（樣本數說明，跨客戶中位數用） */
function customerSampleSuffix(mom: MoMComparison): string {
    const cur = `${mom.current.customerCount} 客戶`;
    if (!mom.previous) return cur;
    const prev = `${mom.previous.customerCount} 客戶`;
    return `${prev} / ${cur}`;
}

function renderMomCards(host: HTMLElement, mom: MoMComparison): void {
    const dailyMeta = dailyDaysSuffix(mom);
    const medianMeta = customerSampleSuffix(mom);
    interface CardSpec {
        label: string;
        meta?: string;
        currValue: string;
        delta: number;
        deltaLabel: string;
        pct: number | null;
        higherIsBetter: boolean;
        /** 是否為「每日平均」衍生指標（樣式略微內凹） */
        isDaily?: boolean;
        /** 是否為「每日中位數」衍生指標 */
        isMedian?: boolean;
    }

    // 主要指標（家族配置）：源序為「總量 → 每日平均 → 客戶中位數」三聯，CSS 用 grid-auto-flow:column
    //   ≥3 欄：每欄一個家族（col1=營收家族、col2=毛利家族、col3=銷售數量家族）
    //   ≤960px：3 欄 row-flow，每列一個指標層級（總量列 / 每日平均列 / 客戶中位數列）
    //   ≤560px：單欄堆疊，源序保持「總量 → 每日平均 → 客戶中位數」家族相鄰
    const primaryCards: CardSpec[] = [
        // === 營收家族 ===
        {
            label: '營收',
            currValue: fmtMoney(Math.round(mom.current.amount)),
            delta: mom.diffs.amount.delta,
            deltaLabel: fmtDelta(mom.diffs.amount.delta),
            pct: mom.diffs.amount.pct,
            higherIsBetter: true,
        },
        {
            label: '每日平均營收',
            meta: dailyMeta,
            currValue: fmtMoney(Math.round(mom.currentDaily.amount)),
            delta: mom.diffs.dailyAmount.delta,
            deltaLabel: fmtDelta(mom.diffs.dailyAmount.delta),
            pct: mom.diffs.dailyAmount.pct,
            higherIsBetter: true,
            isDaily: true,
        },
        {
            label: '客戶中位數營收',
            meta: medianMeta,
            currValue: fmtMoney(Math.round(mom.currentMedian.amount)),
            delta: mom.diffs.medianAmount.delta,
            deltaLabel: fmtDelta(mom.diffs.medianAmount.delta),
            pct: mom.diffs.medianAmount.pct,
            higherIsBetter: true,
            isMedian: true,
        },
        // === 毛利家族 ===
        {
            label: '毛利',
            currValue: fmtMoney(Math.round(mom.current.profit)),
            delta: mom.diffs.profit.delta,
            deltaLabel: fmtDelta(mom.diffs.profit.delta),
            pct: mom.diffs.profit.pct,
            higherIsBetter: true,
        },
        {
            label: '每日平均毛利',
            meta: dailyMeta,
            currValue: fmtMoney(Math.round(mom.currentDaily.profit)),
            delta: mom.diffs.dailyProfit.delta,
            deltaLabel: fmtDelta(mom.diffs.dailyProfit.delta),
            pct: mom.diffs.dailyProfit.pct,
            higherIsBetter: true,
            isDaily: true,
        },
        {
            label: '客戶中位數毛利',
            meta: medianMeta,
            currValue: fmtMoney(Math.round(mom.currentMedian.profit)),
            delta: mom.diffs.medianProfit.delta,
            deltaLabel: fmtDelta(mom.diffs.medianProfit.delta),
            pct: mom.diffs.medianProfit.pct,
            higherIsBetter: true,
            isMedian: true,
        },
        // === 銷售數量家族 ===
        {
            label: '銷售數量',
            currValue: fmtCount(mom.current.count),
            delta: mom.diffs.count.delta,
            deltaLabel: fmtCountDelta(mom.diffs.count.delta),
            pct: mom.diffs.count.pct,
            higherIsBetter: true,
        },
        {
            label: '每日平均數量',
            meta: dailyMeta,
            currValue: fmtCount(Math.round(mom.currentDaily.count)),
            delta: mom.diffs.dailyCount.delta,
            deltaLabel: fmtCountDelta(Math.round(mom.diffs.dailyCount.delta)),
            pct: mom.diffs.dailyCount.pct,
            higherIsBetter: true,
            isDaily: true,
        },
        {
            label: '客戶中位數數量',
            meta: medianMeta,
            currValue: fmtCount(Math.round(mom.currentMedian.count)),
            delta: mom.diffs.medianCount.delta,
            deltaLabel: fmtCountDelta(Math.round(mom.diffs.medianCount.delta)),
            pct: mom.diffs.medianCount.pct,
            higherIsBetter: true,
            isMedian: true,
        },
    ];

    // 附屬指標（1×3）：與主要指標家族無一對一關係，獨立一列
    const auxiliaryCards: CardSpec[] = [
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
    ];

    const renderCard = (c: CardSpec): string => {
        const tone = (c.higherIsBetter ? c.delta >= 0 : c.delta <= 0) ? 'is-positive' : 'is-negative';
        const arrow = c.delta > 0 ? '▲' : c.delta < 0 ? '▼' : '·';
        const metaLine = c.meta ? `<div class="cross-month-mom-meta">${c.meta}</div>` : '';
        const variantClass = c.isDaily ? ' is-daily' : c.isMedian ? ' is-median' : '';
        return `
        <div class="cross-month-mom-card ${tone}${variantClass}">
          <div class="cross-month-mom-label">${c.label}</div>
          ${metaLine}
          <div class="cross-month-mom-value">${c.currValue}</div>
          <div class="cross-month-mom-delta ${tone}">
            <span class="cross-month-mom-arrow">${arrow}</span>
            ${c.deltaLabel}${c.pct !== null ? `（${fmtPct(c.pct)}）` : ''}
          </div>
        </div>
      `;
    };

    host.innerHTML = `
      <div class="cross-month-mom-grid cross-month-mom-grid-primary">
        ${primaryCards.map(renderCard).join('')}
      </div>
      <div class="cross-month-mom-grid cross-month-mom-grid-auxiliary">
        ${auxiliaryCards.map(renderCard).join('')}
      </div>
    `;
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
