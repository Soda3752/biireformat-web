/**
 * 「支出」分頁：上傳收支表 Excel（每月支出）＋ 帳單檔（網站營收），
 * 算出扣原物料後的毛利、扣全部支出後的結餘，並可看年度累計。
 *
 * - 收支表每月重新上傳即覆蓋該月；帳單檔以檔名＋大小為鍵，同檔重傳即覆蓋。
 * - 「其他」明細預設全部勾為家用，可逐項取消勾選改算店務（以項目名稱記住被取消的，下個月同名自動套用），
 *   「銀行」是還貸款本金，不算費用，畫面上改稱「還款」；結餘分「營業結餘」「還款後結餘」「實際剩下的錢」三層。
 * - 下方「分析」區塊：各項占營收 %、近 3 個月平均毛利率、支出異常、每月至少要賣多少才不虧、家用占還款後結餘 %。
 * - 資料皆存在瀏覽器 localStorage（見 expense-store）。
 */

import './expense.css';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';
import {type ChartHandle, createChart, observeChartsResize} from '@/analytics/chart-manager';
import {EXPENSE_CATEGORIES, type ExpenseCategory, type ExpenseItem} from '@/expense/expense-model';
import {parseExpenseWorkbook} from '@/expense/expense-parser';
import {readExpenseGrids} from '@/expense/expense-workbook';
import {
    isHouseholdName,
    loadExpenseMonths,
    loadRevenueFiles,
    loadShopNames,
    removeExpenseMonth,
    removeRevenueFile,
    saveExpenseMonths,
    saveRevenueFiles,
    setHousehold,
} from '@/expense/expense-store';
import {computeMonthRevenue, type MonthRevenue, parseBillToRevenueFile} from '@/expense/revenue-snapshot';
import {categoryLabel, type MonthSummary, summarizeMonth, summarizeYear, type YearSummary} from '@/expense/expense-summary';
import {balanceOption, compositionOption, marginOption, materialOption, shareOption} from '@/expense/expense-charts';
import {
    ANOMALY_KEYS,
    ANOMALY_THRESHOLD,
    anomalyAmount,
    breakeven,
    detectAnomalies,
    HEALTHY_MATERIAL_LABOR_RATIO,
    householdShares,
    revenueRows,
    revenueShares,
    rollingMargins,
    ROLLING_MONTHS,
    SHARE_KEYS,
} from '@/expense/expense-analysis';

const fmtMoney = (v: number) => Math.round(v).toLocaleString('zh-TW');
const fmtOpt = (v: number | null) => (v === null ? '—' : fmtMoney(v));
const fmtSigned = (v: number | null) => (v === null ? '—' : `${v > 0 ? '+' : ''}${fmtMoney(v)}`);
const fmtPct = (num: number | null, den: number | null) =>
    num === null || !den ? '—' : `${((num / den) * 100).toFixed(1)}%`;
const toneClass = (v: number | null) => (v === null ? '' : v < 0 ? 'is-negative' : v > 0 ? 'is-positive' : '');

export function renderExpensePanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card analytics-card expense-card">
      <header class="card-header">
        <h1 class="card-title">支出</h1>
        <p class="card-subtitle">
          上傳收支表 Excel（每月各類支出）與當月帳單檔（網站營收），算出扣掉原物料後的毛利，以及營業結餘、還款後結餘、實際剩下的錢。
          資料存在這台電腦的瀏覽器，同一個月重新上傳會覆蓋。
        </p>
      </header>

      <div class="expense-uploaders">
        <div class="expense-uploader">
          <div class="analytics-dropzone" data-role="expense-drop" tabindex="0" role="button">
            <span class="analytics-dropzone-icon">${icon('upload', 28)}</span>
            <div class="analytics-dropzone-title">① 收支表 Excel</div>
            <div class="analytics-dropzone-hint">例：115-收支表.xlsx，會讀取每個「X月」工作表</div>
            <input type="file" accept=".xlsx" hidden data-role="expense-input">
          </div>
          <div class="analytics-file-list" data-role="expense-chips"></div>
        </div>
        <div class="expense-uploader">
          <div class="analytics-dropzone" data-role="bill-drop" tabindex="0" role="button">
            <span class="analytics-dropzone-icon">${icon('upload', 28)}</span>
            <div class="analytics-dropzone-title">② 帳單檔（網站營收）</div>
            <div class="analytics-dropzone-hint">「明細編號順序 .xlsx」帳單檔，可一次選多個月份</div>
            <input type="file" accept=".xlsx,.xls" multiple hidden data-role="bill-input">
          </div>
          <div class="analytics-file-list" data-role="bill-chips"></div>
        </div>
      </div>

      <div data-role="notices"></div>

      <section class="analytics-content" data-role="content" hidden>
        <div class="expense-toolbar">
          <div class="analytics-section-label" data-role="year-label">年度累計</div>
          <div class="analytics-metric-switch" data-role="year-switch"></div>
        </div>
        <div class="analytics-kpi-grid" data-role="kpi"></div>

        <div class="analytics-chart-grid">
          <div class="analytics-chart-card analytics-chart-card-wide">
            <div class="analytics-chart-title">每月營收、營業費用與結餘</div>
            <div class="analytics-chart-body" data-role="chart-balance"></div>
          </div>
          <div class="analytics-chart-card">
            <div class="analytics-chart-title">每月支出組成</div>
            <div class="analytics-chart-body" data-role="chart-composition"></div>
          </div>
          <div class="analytics-chart-card">
            <div class="analytics-chart-title">每月原物料（Excel 實付）</div>
            <div class="analytics-chart-body" data-role="chart-material"></div>
          </div>
        </div>

        <div class="analytics-section-label">每月明細</div>
        <p class="analytics-section-hint">
          毛利＝網站營收 − 原物料（Excel 實付）。<br>
          營業費用＝原物料＋營運＋車輛＋人事＋其他（店務部分）；「還款」是還貸款本金，不算費用，「家用」也不算。<br>
          營業結餘＝網站營收 − 營業費用 → 還款後結餘＝營業結餘 − 還款 → 實際剩下的錢＝還款後結餘 − 家用。<br>
          點月份可展開該月支出明細；「其他」預設全部算家用，取消勾選的項目改算店務。
        </p>
        <div class="analytics-detail-table-wrap expense-table-wrap">
          <table class="analytics-detail-table-el expense-table" data-role="table"></table>
        </div>
        <div data-role="detail"></div>

        <div class="expense-analysis" data-role="analysis"></div>
      </section>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="status"></div>
      </footer>
    </div>
  `;

    const $ = <T extends HTMLElement>(role: string) => panel.querySelector<T>(`[data-role="${role}"]`)!;
    const expenseDrop = $('expense-drop');
    const expenseInput = $<HTMLInputElement>('expense-input');
    const billDrop = $('bill-drop');
    const billInput = $<HTMLInputElement>('bill-input');
    const expenseChips = $('expense-chips');
    const billChips = $('bill-chips');
    const noticesEl = $('notices');
    const contentEl = $('content');
    const yearSwitch = $('year-switch');
    const yearLabel = $('year-label');
    const kpiEl = $('kpi');
    const tableEl = $('table');
    const detailEl = $('detail');
    const analysisEl = $('analysis');
    const statusEl = $('status');

    let selectedYear: number | null = null;
    let selectedMonth: string | null = null;
    /** 解析時產生的一次性提醒（重新整理後不保留） */
    let parseNotices: string[] = [];
    const charts: {
        balance: ChartHandle | null;
        composition: ChartHandle | null;
        material: ChartHandle | null;
        share: ChartHandle | null;
        margin: ChartHandle | null;
    } = {balance: null, composition: null, material: null, share: null, margin: null};
    let resizeCleanup: (() => void) | null = null;
    let analysisResizeCleanup: (() => void) | null = null;
    const disposeAnalysisCharts = () => {
        analysisResizeCleanup?.();
        analysisResizeCleanup = null;
        charts.share?.dispose();
        charts.margin?.dispose();
        charts.share = null;
        charts.margin = null;
    };
    let renderSeq = 0;

    /* ============== 上傳 ============== */

    const handleExpenseFile = async (file: File) => {
        statusEl.textContent = `解析 ${file.name} 中…`;
        try {
            const grids = await readExpenseGrids(await file.arrayBuffer());
            const currentRoc = new Date().getFullYear() - 1911;
            const res = parseExpenseWorkbook(grids, file.name, currentRoc);
            if (res.months.length === 0) {
                throw new Error(res.errors[0] ?? '找不到任何月份工作表（工作表 A1 應為如「114/1月」，並有「原物料／營運／車輛／人事／其他」標題列）');
            }
            saveExpenseMonths(res.months);
            parseNotices = [
                ...res.errors,
                ...res.months.flatMap((m) => m.warnings),
                ...res.summaryMismatches.map(
                    (x) => `${x.month}月「${x.category}」：年度總表 ${fmtMoney(x.summary)}，解析結果 ${fmtMoney(x.parsed)}，請確認`,
                ),
            ];
            const verified = res.hasSummary && res.summaryMismatches.length === 0;
            showToast({
                variant: parseNotices.length ? 'warning' : 'success',
                title: `已匯入 ${res.months.length} 個月份支出`,
                message:
                    `${res.months.map((m) => `${m.month}月`).join('、')}` +
                    (verified ? '；各類金額與「年度總表」核對一致' : '') +
                    (parseNotices.length ? `；有 ${parseNotices.length} 則提醒` : ''),
            });
            selectedYear = res.months[res.months.length - 1].year;
        } catch (err) {
            console.error('[expense] parse failed', err);
            showToast({variant: 'error', title: `${file.name} 解析失敗`, message: err instanceof Error ? err.message : String(err)});
        }
        await render();
    };

    const handleBillFiles = async (files: File[]) => {
        if (files.length === 0) return;
        const parsed = [];
        for (const file of files) {
            statusEl.textContent = `解析帳單 ${file.name} 中…`;
            try {
                parsed.push(await parseBillToRevenueFile(file));
            } catch (err) {
                console.error('[expense] bill parse failed', file.name, err);
                showToast({variant: 'error', title: `${file.name} 解析失敗`, message: err instanceof Error ? err.message : String(err)});
            }
        }
        if (parsed.length > 0) {
            try {
                saveRevenueFiles(parsed);
                const months = [...new Set(parsed.flatMap((f) => Object.keys(f.months)))].sort();
                showToast({variant: 'success', title: `已載入 ${parsed.length} 份帳單`, message: `月份：${months.map(sortKeyLabel).join('、')}`});
            } catch (err) {
                showToast({variant: 'error', title: '帳單儲存失敗', message: err instanceof Error ? err.message : String(err)});
            }
        }
        await render();
    };

    bindDropzone(expenseDrop, expenseInput, (files) => {
        if (files[0]) void handleExpenseFile(files[0]);
    });
    bindDropzone(billDrop, billInput, (files) => void handleBillFiles(files));

    /* ============== 渲染 ============== */

    const renderChips = () => {
        const months = Object.values(loadExpenseMonths()).sort((a, b) => a.sortKey.localeCompare(b.sortKey));
        expenseChips.innerHTML = months
            .map(
                (m) => `<div class="analytics-file-chip" title="${escapeHtml(m.sourceFile)}／${escapeHtml(m.sheetName)}">
              <span class="analytics-file-chip-name">${m.year}年${m.month}月</span>
              <button type="button" class="analytics-file-chip-remove" data-remove-month="${m.sortKey}" aria-label="移除">${icon('close', 14)}</button>
            </div>`,
            )
            .join('');
        const files = loadRevenueFiles().sort((a, b) => a.fileName.localeCompare(b.fileName));
        billChips.innerHTML = files
            .map(
                (f) => `<div class="analytics-file-chip">
              <span class="analytics-file-chip-name">${escapeHtml(f.fileName)}</span>
              <span class="analytics-file-chip-meta">${Object.keys(f.months).sort().map(sortKeyLabel).join('、')}</span>
              <button type="button" class="analytics-file-chip-remove" data-remove-bill="${escapeHtml(f.id)}" aria-label="移除">${icon('close', 14)}</button>
            </div>`,
            )
            .join('');
    };

    panel.addEventListener('click', (e) => {
        const target = e.target as HTMLElement;
        const monthBtn = target.closest<HTMLElement>('[data-remove-month]');
        if (monthBtn) {
            e.stopPropagation();
            removeExpenseMonth(monthBtn.dataset.removeMonth!);
            void render();
            return;
        }
        const billBtn = target.closest<HTMLElement>('[data-remove-bill]');
        if (billBtn) {
            e.stopPropagation();
            removeRevenueFile(billBtn.dataset.removeBill!);
            void render();
            return;
        }
        const row = target.closest<HTMLElement>('tr[data-month]');
        if (row) {
            selectedMonth = selectedMonth === row.dataset.month ? null : row.dataset.month!;
            void render();
        }
    });

    panel.addEventListener('change', (e) => {
        const box = (e.target as HTMLElement).closest<HTMLInputElement>('input[data-household]');
        if (!box) return;
        setHousehold(box.dataset.household!, box.checked);
        void render();
    });

    const render = async () => {
        const seq = ++renderSeq;
        renderChips();
        const months = Object.values(loadExpenseMonths());
        const revenueFiles = loadRevenueFiles();
        const revenue = computeMonthRevenue(revenueFiles);
        if (seq !== renderSeq) return;
        const shopNames = loadShopNames();

        const years = [...new Set(months.map((m) => m.year))].sort((a, b) => a - b);
        if (selectedYear === null || !years.includes(selectedYear)) selectedYear = years[years.length - 1] ?? null;

        const rows: MonthSummary[] = months
            .filter((m) => m.year === selectedYear)
            .sort((a, b) => a.sortKey.localeCompare(b.sortKey))
            .map((m) => summarizeMonth(m, revenue.get(m.sortKey), shopNames));

        renderNotices(rows, revenue, new Set(months.map((m) => m.sortKey)));

        if (rows.length === 0) {
            contentEl.hidden = true;
            statusEl.textContent = months.length === 0 ? '請先上傳收支表 Excel' : '';
            return;
        }
        contentEl.hidden = false;

        yearSwitch.innerHTML = years.length > 1
            ? years.map((y) => `<button type="button" data-year="${y}" class="${y === selectedYear ? 'is-active' : ''}">${y}年</button>`).join('')
            : '';
        yearSwitch.hidden = years.length <= 1;
        yearSwitch.querySelectorAll<HTMLButtonElement>('button[data-year]').forEach((btn) =>
            btn.addEventListener('click', () => {
                selectedYear = Number(btn.dataset.year);
                selectedMonth = null;
                void render();
            }),
        );

        const year = summarizeYear(rows);
        yearLabel.textContent = `${selectedYear}年 累計（${rows.map((r) => r.month).join('、')}月）`;
        kpiEl.innerHTML = kpiHtml(year);
        tableEl.innerHTML = tableHtml(rows, year, selectedMonth);
        const selected = rows.find((r) => r.sortKey === selectedMonth);
        detailEl.innerHTML = selected ? detailHtml(selected, shopNames) : '';
        const hasAnalysis = revenueRows(rows).length > 0;
        if (!hasAnalysis) disposeAnalysisCharts();
        analysisEl.innerHTML = analysisHtml(rows);

        await renderCharts(rows);
        statusEl.textContent =
            `${selectedYear}年：支出 ${rows.length} 個月份，其中 ${year.revenueMonths} 個月份已載入帳單營收` +
            (year.revenueMonths < rows.length ? '（缺帳單的月份不計入毛利、結餘累計與分析）' : '');
    };

    const renderNotices = (rows: MonthSummary[], revenue: Map<string, MonthRevenue>, expenseKeys: Set<string>) => {
        const notes: string[] = [...parseNotices];
        const orphan = [...revenue.keys()].filter((k) => !expenseKeys.has(k)).sort();
        if (orphan.length > 0) {
            notes.push(`已載入帳單但還沒有支出資料的月份：${orphan.map(sortKeyLabel).join('、')}（上傳該月收支表後才會出現在下方）。`);
        }
        const multi = [...revenue.entries()].filter(([, v]) => v.fileNames.length > 1);
        for (const [k, v] of multi) {
            notes.push(`${sortKeyLabel(k)} 有 ${v.fileNames.length} 份帳單（${v.fileNames.join('、')}）營收會相加，若是同一份資料重複請移除多的。`);
        }
        const missing = rows.filter((r) => r.revenue === null).map((r) => `${r.month}月`);
        if (rows.length > 0 && missing.length > 0) {
            notes.push(`${missing.join('、')} 還沒有載入帳單檔，這些月份的營收、毛利、結餘先顯示「—」。`);
        }
        noticesEl.innerHTML = notes.map((n) => `<div class="cross-month-warning">${escapeHtml(n)}</div>`).join('');
    };

    const renderCharts = async (rows: MonthSummary[]) => {
        try {
            if (!charts.balance) {
                charts.balance = await createChart($('chart-balance'));
                charts.composition = await createChart($('chart-composition'));
                charts.material = await createChart($('chart-material'));
                resizeCleanup?.();
                resizeCleanup = observeChartsResize([charts.balance, charts.composition, charts.material]);
            }
            charts.balance.setOption(balanceOption(rows));
            charts.composition?.setOption(compositionOption(rows));
            charts.material?.setOption(materialOption(rows));

            const shareEl = panel.querySelector<HTMLElement>('[data-role="chart-share"]');
            const marginEl = panel.querySelector<HTMLElement>('[data-role="chart-margin"]');
            if (shareEl && marginEl) {
                // 分析區塊每次重繪都會換新 DOM，圖表跟著重建
                disposeAnalysisCharts();
                charts.share = await createChart(shareEl);
                charts.margin = await createChart(marginEl);
                analysisResizeCleanup = observeChartsResize([charts.share, charts.margin]);
                const shares = revenueShares(rows);
                charts.share.setOption(shareOption(shares.year ? [...shares.months, shares.year] : shares.months));
                charts.margin.setOption(marginOption(rollingMargins(rows)));
            }
        } catch (err) {
            console.error('[expense] chart failed', err);
        }
    };

    // 切到本分頁時重算（其他分頁改了資料都會反映）
    window.addEventListener('hashchange', () => {
        if (window.location.hash === tab.hash) void render();
    });
    void render();

    return panel;
}

/* ====================== HTML 片段 ====================== */

function kpiHtml(y: YearSummary): string {
    const hasRev = y.revenueMonths > 0;
    const opt = (v: number) => (hasRev ? fmtMoney(v) : '—');
    const tone = (v: number) => toneClass(hasRev ? v : null);
    const items: { label: string; value: string; hint?: string; tone?: string }[] = [
        {label: '網站營收', value: fmtMoney(y.revenue), hint: `${y.revenueMonths} 個月`},
        {label: '原物料（Excel 實付）', value: fmtMoney(y.materialPaid)},
        {
            label: '扣原物料毛利',
            value: opt(y.grossProfit),
            hint: hasRev ? `毛利率 ${fmtPct(y.grossProfit, y.revenue)}` : undefined,
            tone: tone(y.grossProfit),
        },
        {label: '營業費用', value: fmtMoney(y.operatingExpense), hint: '不含還款、家用'},
        {label: '營業結餘', value: opt(y.operatingBalance), hint: '營收 − 營業費用', tone: tone(y.operatingBalance)},
        {label: '還款', value: fmtMoney(y.repayment), hint: '還貸款本金，不算費用'},
        {label: '還款後結餘', value: opt(y.afterRepayBalance), hint: '營業結餘 − 還款', tone: tone(y.afterRepayBalance)},
        {
            label: '實際剩下的錢',
            value: opt(y.netBalance),
            hint: `再扣家用 ${fmtMoney(y.household)}；全部支出（含還款、家用）${fmtMoney(y.totalExpense)}`,
            tone: tone(y.netBalance),
        },
    ];
    return items
        .map(
            (it) => `<div class="analytics-kpi-card">
          <div class="analytics-kpi-label">${it.label}</div>
          <div class="analytics-kpi-value ${it.tone ?? ''}">${it.value}</div>
          ${it.hint ? `<div class="analytics-kpi-hint" title="${escapeHtml(it.hint)}">${escapeHtml(it.hint)}</div>` : ''}
        </div>`,
        )
        .join('');
}

const FEE_CATS: ExpenseCategory[] = ['營運', '車輛', '人事', '其他'];

function tableHtml(rows: MonthSummary[], y: YearSummary, selected: string | null): string {
    const head = `<thead><tr>
      <th>月份</th>
      <th class="is-right">網站營收</th>
      <th class="is-right">原物料<br><small>Excel 實付</small></th>
      <th class="is-right">扣原物料毛利</th>
      <th class="is-right">毛利率</th>
      ${FEE_CATS.map((c) => `<th class="is-right">${c}</th>`).join('')}
      <th class="is-right">其中家用<br><small>不算營業費用</small></th>
      <th class="is-right">營業費用<br><small>不含還款、家用</small></th>
      <th class="is-right">營業結餘<br><small>營收 − 營業費用</small></th>
      <th class="is-right">還款<br><small>還貸款本金</small></th>
      <th class="is-right">還款後結餘</th>
      <th class="is-right">實際剩下的錢<br><small>再扣家用</small></th>
    </tr></thead>`;

    const body = rows
        .map((r) => {
            return `<tr data-month="${r.sortKey}" class="expense-row ${r.sortKey === selected ? 'is-selected' : ''}">
          <td><span class="expense-row-caret">${icon(r.sortKey === selected ? 'chevron-up' : 'chevron-down', 12)}</span>${r.month}月</td>
          <td class="is-right">${fmtOpt(r.revenue)}</td>
          <td class="is-right">${fmtMoney(r.materialPaid)}</td>
          <td class="is-right ${toneClass(r.grossProfit)}">${fmtOpt(r.grossProfit)}</td>
          <td class="is-right">${fmtPct(r.grossProfit, r.revenue)}</td>
          ${FEE_CATS.map((c) => `<td class="is-right">${fmtMoney(r.totals[c])}</td>`).join('')}
          <td class="is-right expense-muted">${r.household ? fmtMoney(r.household) : '—'}</td>
          <td class="is-right">${fmtMoney(r.operatingExpense)}</td>
          <td class="is-right ${toneClass(r.operatingBalance)}"><strong>${fmtOpt(r.operatingBalance)}</strong></td>
          <td class="is-right">${fmtMoney(r.repayment)}</td>
          <td class="is-right ${toneClass(r.afterRepayBalance)}">${fmtOpt(r.afterRepayBalance)}</td>
          <td class="is-right ${toneClass(r.netBalance)}"><strong>${fmtOpt(r.netBalance)}</strong></td>
        </tr>`;
        })
        .join('');

    const hasRev = y.revenueMonths > 0;
    const opt = (v: number) => (hasRev ? fmtMoney(v) : '—');
    const tone = (v: number) => toneClass(hasRev ? v : null);
    const foot = `<tfoot><tr class="expense-total-row">
      <td>年度累計</td>
      <td class="is-right">${opt(y.revenue)}</td>
      <td class="is-right">${fmtMoney(y.materialPaid)}</td>
      <td class="is-right ${tone(y.grossProfit)}">${opt(y.grossProfit)}</td>
      <td class="is-right">${hasRev ? fmtPct(y.grossProfit, y.revenue) : '—'}</td>
      ${FEE_CATS.map((c) => `<td class="is-right">${fmtMoney(y.totals[c])}</td>`).join('')}
      <td class="is-right expense-muted">${y.household ? fmtMoney(y.household) : '—'}</td>
      <td class="is-right">${fmtMoney(y.operatingExpense)}</td>
      <td class="is-right ${tone(y.operatingBalance)}"><strong>${opt(y.operatingBalance)}</strong></td>
      <td class="is-right">${fmtMoney(y.repayment)}</td>
      <td class="is-right ${tone(y.afterRepayBalance)}">${opt(y.afterRepayBalance)}</td>
      <td class="is-right ${tone(y.netBalance)}"><strong>${opt(y.netBalance)}</strong></td>
    </tr></tfoot>`;
    return head + `<tbody>${body}</tbody>` + foot;
}

/* ---------- 分析區塊 ---------- */

const ratioText = (v: number | null) => (v === null ? '—' : `${(v * 100).toFixed(1)}%`);

function analysisHtml(rows: MonthSummary[]): string {
    const list = revenueRows(rows);
    const header = `<div class="analytics-section-label">分析</div>`;
    if (list.length === 0) {
        return `${header}<p class="analytics-section-hint">還沒有任何月份載入帳單營收，分析需要網站營收才能計算。</p>`;
    }
    const scope = `<p class="analytics-section-hint">以下只用已載入帳單營收的月份計算（${list.map((r) => `${r.month}月`).join('、')}，共 ${list.length} 個月）。</p>`;
    return [header, scope, shareSectionHtml(rows), marginSectionHtml(rows), anomalySectionHtml(rows), breakevenSectionHtml(rows)].join('');
}

function shareSectionHtml(rows: MonthSummary[]): string {
    const {months, year} = revenueShares(rows);
    const limit = HEALTHY_MATERIAL_LABOR_RATIO;
    const line = (label: string, r: (typeof months)[number], cls = '') => `<tr class="${cls}">
        <td>${label}</td>
        ${SHARE_KEYS.map((k) => `<td class="is-right ${k === '營業結餘' ? toneClass(r.amounts[k]) : ''}">${ratioText(r.pct[k])}</td>`).join('')}
        <td class="is-right ${r.materialLabor !== null && r.materialLabor > limit ? 'is-negative' : ''}"><strong>${ratioText(r.materialLabor)}</strong></td>
      </tr>`;
    const body = months.map((r) => line(`${r.month}月`, r)).join('');
    const foot = year ? line('年度累計', year, 'expense-total-row') : '';
    return `<div class="expense-analysis-block">
      <div class="expense-analysis-title">① 各項占營收比例</div>
      <p class="analytics-section-hint">
        每 100 元營收花在哪裡。「原物料＋人事」超過 ${Math.round(limit * 100)}% 標紅——食品業一般認為 60～65% 以內算健康，這只是一般參考，不是硬性標準。
      </p>
      <div class="analytics-chart-card analytics-chart-card-wide">
        <div class="analytics-chart-title">各項占營收 %（虛線為一般參考 ${Math.round(limit * 100)}%）</div>
        <div class="analytics-chart-body" data-role="chart-share"></div>
      </div>
      <div class="analytics-detail-table-wrap expense-table-wrap">
        <table class="analytics-detail-table-el expense-table">
          <thead><tr><th>月份</th>${SHARE_KEYS.map((k) => `<th class="is-right">${k}</th>`).join('')}<th class="is-right">原物料＋人事</th></tr></thead>
          <tbody>${body}</tbody>
          <tfoot>${foot}</tfoot>
        </table>
      </div>
    </div>`;
}

function marginSectionHtml(rows: MonthSummary[]): string {
    const margins = rollingMargins(rows);
    const hh = householdShares(rows);
    const hhBy = new Map(hh.months.map((h) => [h.sortKey, h]));
    const hhCell = (h: (typeof hh.months)[number] | undefined | null) => {
        if (!h) return '<td class="is-right">—</td>';
        if (h.ratio === null) {
            return `<td class="is-right is-negative" title="還款後結餘是 ${fmtMoney(h.afterRepayBalance)}，家用沒有結餘可以付">還款後已沒剩（家用靠老本）</td>`;
        }
        return `<td class="is-right">${ratioText(h.ratio)}</td>`;
    };
    const body = margins
        .map((m) => {
            const h = hhBy.get(m.sortKey);
            return `<tr>
          <td>${m.month}月</td>
          <td class="is-right expense-muted">${ratioText(m.margin)}</td>
          <td class="is-right"><strong>${ratioText(m.average)}</strong>${m.monthsUsed < ROLLING_MONTHS ? ` <small class="expense-muted">（${m.monthsUsed} 個月）</small>` : ''}</td>
          <td class="is-right ${toneClass(h?.afterRepayBalance ?? null)}">${h ? fmtMoney(h.afterRepayBalance) : '—'}</td>
          <td class="is-right">${h ? fmtMoney(h.household) : '—'}</td>
          ${hhCell(h)}
        </tr>`;
        })
        .join('');
    const y = hh.year;
    const foot = y
        ? `<tr class="expense-total-row"><td>年度累計</td><td class="is-right" colspan="2"></td>
            <td class="is-right ${toneClass(y.afterRepayBalance)}">${fmtMoney(y.afterRepayBalance)}</td>
            <td class="is-right">${fmtMoney(y.household)}</td>${hhCell(y)}</tr>`
        : '';
    return `<div class="expense-analysis-block">
      <div class="expense-analysis-title">② 近 ${ROLLING_MONTHS} 個月平均毛利率、⑤ 家用占還款後結餘</div>
      <p class="analytics-section-hint">
        麵粉等原物料是壓 3 個月才付款，單月毛利率會忽高忽低；看「近 ${ROLLING_MONTHS} 個月平均」（含當月，合計毛利 ÷ 合計營收）比較準。<br>
        家用占比的分母是「還款後結餘」（營業結餘扣掉還款後剩的錢）：100% 代表還款後剩的錢剛好全拿去家用，超過 100% 代表家用比剩的還多；還款後已是負的月份顯示「還款後已沒剩」。
      </p>
      <div class="analytics-chart-card analytics-chart-card-wide">
        <div class="analytics-chart-title">單月毛利率與近 ${ROLLING_MONTHS} 個月平均</div>
        <div class="analytics-chart-body" data-role="chart-margin"></div>
      </div>
      <div class="analytics-detail-table-wrap expense-table-wrap">
        <table class="analytics-detail-table-el expense-table">
          <thead><tr>
            <th>月份</th><th class="is-right">單月毛利率</th><th class="is-right">近 ${ROLLING_MONTHS} 個月平均毛利率</th>
            <th class="is-right">還款後結餘</th><th class="is-right">家用</th><th class="is-right">家用占還款後結餘</th>
          </tr></thead>
          <tbody>${body}</tbody>
          <tfoot>${foot}</tfoot>
        </table>
      </div>
    </div>`;
}

function anomalySectionHtml(rows: MonthSummary[]): string {
    const list = revenueRows(rows);
    const anomalies = detectAnomalies(rows);
    const flagged = new Set(anomalies.map((a) => `${a.sortKey}|${a.key}`));
    const pct = Math.round(ANOMALY_THRESHOLD * 100);
    const body = list
        .map(
            (r) => `<tr><td>${r.month}月</td>${ANOMALY_KEYS.map((k) => {
                const hit = flagged.has(`${r.sortKey}|${k}`);
                return `<td class="is-right ${hit ? 'expense-anomaly' : ''}">${fmtMoney(anomalyAmount(r, k))}</td>`;
            }).join('')}</tr>`,
        )
        .join('');
    const notes = list.length < 2
        ? '<li class="expense-muted">至少要有 2 個月營收資料才能比較。</li>'
        : anomalies.length === 0
            ? `<li class="expense-muted">沒有哪個月的支出比平常高 ${pct}% 以上。</li>`
            : anomalies
                .map(
                    (a) => `<li>${a.month}月 ${a.key} ${fmtMoney(a.amount)}，比其他月份平均（${fmtMoney(a.average)}）高 ${Math.round(a.over * 100)}%</li>`,
                )
                .join('');
    return `<div class="expense-analysis-block">
      <div class="expense-analysis-title">③ 支出比平常高的月份</div>
      <p class="analytics-section-hint">某個月的金額比「同一類其他月份的平均」高 ${pct}% 以上就標紅，提醒去看那個月是不是有特別支出。</p>
      <ul class="expense-anomaly-list">${notes}</ul>
      <div class="analytics-detail-table-wrap expense-table-wrap">
        <table class="analytics-detail-table-el expense-table">
          <thead><tr><th>月份</th>${ANOMALY_KEYS.map((k) => `<th class="is-right">${k}</th>`).join('')}</tr></thead>
          <tbody>${body}</tbody>
        </table>
      </div>
    </div>`;
}

function breakevenSectionHtml(rows: MonthSummary[]): string {
    const b = breakeven(rows);
    if (!b) return '';
    const marginTxt = ratioText(b.grossMargin);
    const card = (label: string, need: number | null, cost: string) => {
        const gap = need === null ? null : b.avgRevenue - need;
        const gapTxt = gap === null
            ? '毛利率不是正的，怎麼賣都打不平'
            : gap >= 0
                ? `實際平均每月營收 ${fmtMoney(b.avgRevenue)}，多出 ${fmtMoney(gap)}`
                : `實際平均每月營收 ${fmtMoney(b.avgRevenue)}，還差 ${fmtMoney(-gap)}`;
        return `<div class="analytics-kpi-card">
          <div class="analytics-kpi-label">${label}</div>
          <div class="analytics-kpi-value ${gap === null ? 'is-negative' : toneClass(gap)}">${need === null ? '—' : fmtMoney(need)}</div>
          <div class="analytics-kpi-hint" title="${escapeHtml(gapTxt)}">${escapeHtml(gapTxt)}</div>
          <div class="analytics-kpi-hint" title="${escapeHtml(cost)}">${escapeHtml(cost)}</div>
        </div>`;
    };
    return `<div class="expense-analysis-block">
      <div class="expense-analysis-title">④ 每月至少要賣多少才不虧</div>
      <p class="analytics-section-hint">
        每賣 100 元扣掉原物料剩 ${b.grossMargin === null ? '—' : (b.grossMargin * 100).toFixed(1)} 元（年度累計毛利率 ${marginTxt}），
        要用這些錢付掉每月的固定支出。至少要賣的營收 ＝ 每月平均要付的錢 ÷ 毛利率。
      </p>
      <div class="analytics-kpi-grid expense-breakeven-grid">
        ${card('營業不虧', b.operating, `固定支出（營運＋車輛＋人事＋其他店務）平均每月 ${fmtMoney(b.avgFixed)}`)}
        ${card('要還款也不虧', b.withRepayment, `再加還款平均每月 ${fmtMoney(b.avgRepayment)}`)}
        ${card('連家用也打平', b.withHousehold, `再加家用平均每月 ${fmtMoney(b.avgHousehold)}`)}
      </div>
    </div>`;
}

function detailHtml(r: MonthSummary, shopNames: ReadonlySet<string>): string {
    const byCat = new Map<ExpenseCategory, ExpenseItem[]>();
    for (const it of r.expense.items) {
        const list = byCat.get(it.category) ?? [];
        list.push(it);
        byCat.set(it.category, list);
    }
    const blocks = EXPENSE_CATEGORIES.map((cat) => {
        const items = byCat.get(cat) ?? [];
        const isOther = cat === '其他';
        const list = items
            .map((it) => {
                const isHousehold = isOther && isHouseholdName(it.name, shopNames);
                return `<li class="expense-item ${isHousehold ? 'is-household' : ''}">
              ${isOther ? `<label class="expense-household-toggle" title="預設為家用；取消勾選改算店務（同名項目之後每月自動套用）">
                <input type="checkbox" data-household="${escapeHtml(it.name)}" ${isHousehold ? 'checked' : ''}>家用</label>` : ''}
              <span class="expense-item-name">${escapeHtml(it.name)}</span>
              ${it.note ? `<span class="expense-item-note">${escapeHtml(it.note)}</span>` : ''}
              <span class="expense-item-amount">${fmtMoney(it.amount)}</span>
            </li>`;
            })
            .join('');
        const extra = isOther && r.household
            ? `<div class="expense-block-sub">店務 ${fmtMoney(r.totals[cat] - r.household)}／家用 ${fmtMoney(r.household)}</div>`
            : '';
        return `<div class="expense-block">
          <div class="expense-block-head"><span>${categoryLabel(cat)}</span><span>${fmtMoney(r.totals[cat])}</span></div>
          ${extra}
          <ul class="expense-item-list">${list || '<li class="expense-muted">（無明細）</li>'}</ul>
        </div>`;
    }).join('');

    const income = r.expense.income.length
        ? `<p class="analytics-section-hint">Excel 收入欄（僅供參考，營收以網站帳單為準）：合計 ${fmtMoney(r.expense.incomeTotal)}
            ${r.revenue !== null ? `，與網站營收差 ${fmtSigned(r.expense.incomeTotal - r.revenue)}` : ''}
            ——${r.expense.income.map((i) => `${escapeHtml(i.name)} ${fmtMoney(i.amount)}`).join('、')}</p>`
        : '';

    return `<div class="expense-detail">
      <div class="analytics-section-label">${r.year}年${r.month}月 支出明細（來源：${escapeHtml(r.expense.sourceFile)}／${escapeHtml(r.expense.sheetName)}）</div>
      ${income}
      <div class="expense-blocks">${blocks}</div>
    </div>`;
}

/* ====================== 小工具 ====================== */

function sortKeyLabel(key: string): string {
    const [y, m] = key.split('-');
    return `${y}年${Number(m)}月`;
}

function bindDropzone(zone: HTMLElement, input: HTMLInputElement, onFiles: (files: File[]) => void): void {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            input.click();
        }
    });
    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('is-dragover');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('is-dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('is-dragover');
        if (e.dataTransfer?.files) onFiles(Array.from(e.dataTransfer.files));
    });
    input.addEventListener('change', () => {
        if (input.files) onFiles(Array.from(input.files));
        input.value = '';
    });
}

function escapeHtml(value: string): string {
    return value
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}
