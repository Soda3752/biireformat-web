/**
 * 「支出」分頁：上傳收支表 Excel（每月支出）＋ 帳單檔（網站營收），
 * 算出扣原物料後的毛利、扣全部支出後的結餘，並可看年度累計。
 *
 * - 收支表每月重新上傳即覆蓋該月；帳單檔以檔名＋大小為鍵，同檔重傳即覆蓋。
 * - 「其他」明細可逐項勾為家用（以項目名稱記住，下個月同名自動套用），
 *   結餘分「店務結餘」（不含家用）與「含家用結餘」兩種。
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
    householdKey,
    loadExpenseMonths,
    loadHouseholdNames,
    loadRevenueFiles,
    removeExpenseMonth,
    removeRevenueFile,
    saveExpenseMonths,
    saveRevenueFiles,
    setHousehold,
} from '@/expense/expense-store';
import {computeMonthRevenue, type MonthRevenue, parseBillToRevenueFile} from '@/expense/revenue-snapshot';
import {type MonthSummary, summarizeMonth, summarizeYear, type YearSummary} from '@/expense/expense-summary';
import {balanceOption, compositionOption, materialOption} from '@/expense/expense-charts';

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
          上傳收支表 Excel（每月各類支出）與當月帳單檔（網站營收），算出扣掉原物料後的毛利，以及扣掉全部支出後的結餘。
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
            <div class="analytics-chart-title">每月營收、店務支出與結餘</div>
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
          毛利＝網站營收 − 原物料（Excel 實付）；店務結餘＝網站營收 − 全部支出（不含勾為家用的項目）；含家用結餘＝網站營收 − 全部支出。
          點月份可展開該月支出明細、勾選家用項目。
        </p>
        <div class="analytics-detail-table-wrap expense-table-wrap">
          <table class="analytics-detail-table-el expense-table" data-role="table"></table>
        </div>
        <div data-role="detail"></div>
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
    const statusEl = $('status');

    let selectedYear: number | null = null;
    let selectedMonth: string | null = null;
    /** 解析時產生的一次性提醒（重新整理後不保留） */
    let parseNotices: string[] = [];
    const charts: { balance: ChartHandle | null; composition: ChartHandle | null; material: ChartHandle | null } = {
        balance: null,
        composition: null,
        material: null,
    };
    let resizeCleanup: (() => void) | null = null;
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
        const household = loadHouseholdNames();

        const years = [...new Set(months.map((m) => m.year))].sort((a, b) => a - b);
        if (selectedYear === null || !years.includes(selectedYear)) selectedYear = years[years.length - 1] ?? null;

        const rows: MonthSummary[] = months
            .filter((m) => m.year === selectedYear)
            .sort((a, b) => a.sortKey.localeCompare(b.sortKey))
            .map((m) => summarizeMonth(m, revenue.get(m.sortKey), household));

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
        detailEl.innerHTML = selected ? detailHtml(selected, household) : '';

        await renderCharts(rows);
        statusEl.textContent =
            `${selectedYear}年：支出 ${rows.length} 個月份，其中 ${year.revenueMonths} 個月份已載入帳單營收` +
            (year.revenueMonths < rows.length ? '（缺帳單的月份不計入毛利與結餘累計）' : '');
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
    const items: { label: string; value: string; hint?: string; tone?: string }[] = [
        {label: '網站營收', value: fmtMoney(y.revenue), hint: `${y.revenueMonths} 個月`},
        {label: '原物料（Excel 實付）', value: fmtMoney(y.materialPaid)},
        {
            label: '扣原物料毛利',
            value: y.revenueMonths ? fmtMoney(y.grossProfit) : '—',
            hint: y.revenueMonths ? `毛利率 ${fmtPct(y.grossProfit, y.revenue)}` : undefined,
            tone: toneClass(y.revenueMonths ? y.grossProfit : null),
        },
        {label: '店務支出', value: fmtMoney(y.shopExpense), hint: `家用 ${fmtMoney(y.household)}`},
        {
            label: '店務結餘',
            value: y.revenueMonths ? fmtMoney(y.shopBalance) : '—',
            tone: toneClass(y.revenueMonths ? y.shopBalance : null),
        },
        {
            label: '含家用結餘',
            value: y.revenueMonths ? fmtMoney(y.allBalance) : '—',
            hint: `全部支出 ${fmtMoney(y.totalExpense)}`,
            tone: toneClass(y.revenueMonths ? y.allBalance : null),
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

const OTHER_CATS: ExpenseCategory[] = ['營運', '車輛', '人事', '其他', '銀行'];

function tableHtml(rows: MonthSummary[], y: YearSummary, selected: string | null): string {
    const head = `<thead><tr>
      <th>月份</th>
      <th class="is-right">網站營收</th>
      <th class="is-right">原物料<br><small>Excel 實付</small></th>
      <th class="is-right">扣原物料毛利</th>
      <th class="is-right">毛利率</th>
      ${OTHER_CATS.map((c) => `<th class="is-right">${c}</th>`).join('')}
      <th class="is-right">其中家用</th>
      <th class="is-right">全部支出</th>
      <th class="is-right">店務結餘</th>
      <th class="is-right">含家用結餘</th>
    </tr></thead>`;

    const body = rows
        .map((r) => {
            return `<tr data-month="${r.sortKey}" class="expense-row ${r.sortKey === selected ? 'is-selected' : ''}">
          <td><span class="expense-row-caret">${icon(r.sortKey === selected ? 'chevron-up' : 'chevron-down', 12)}</span>${r.month}月</td>
          <td class="is-right">${fmtOpt(r.revenue)}</td>
          <td class="is-right">${fmtMoney(r.materialPaid)}</td>
          <td class="is-right ${toneClass(r.grossProfit)}">${fmtOpt(r.grossProfit)}</td>
          <td class="is-right">${fmtPct(r.grossProfit, r.revenue)}</td>
          ${OTHER_CATS.map((c) => `<td class="is-right">${fmtMoney(r.totals[c])}</td>`).join('')}
          <td class="is-right expense-muted">${r.household ? fmtMoney(r.household) : '—'}</td>
          <td class="is-right">${fmtMoney(r.totalExpense)}</td>
          <td class="is-right ${toneClass(r.shopBalance)}"><strong>${fmtOpt(r.shopBalance)}</strong></td>
          <td class="is-right ${toneClass(r.allBalance)}">${fmtOpt(r.allBalance)}</td>
        </tr>`;
        })
        .join('');

    const hasRev = y.revenueMonths > 0;
    const foot = `<tfoot><tr class="expense-total-row">
      <td>年度累計</td>
      <td class="is-right">${hasRev ? fmtMoney(y.revenue) : '—'}</td>
      <td class="is-right">${fmtMoney(y.materialPaid)}</td>
      <td class="is-right ${toneClass(hasRev ? y.grossProfit : null)}">${hasRev ? fmtMoney(y.grossProfit) : '—'}</td>
      <td class="is-right">${hasRev ? fmtPct(y.grossProfit, y.revenue) : '—'}</td>
      ${OTHER_CATS.map((c) => `<td class="is-right">${fmtMoney(y.totals[c])}</td>`).join('')}
      <td class="is-right expense-muted">${y.household ? fmtMoney(y.household) : '—'}</td>
      <td class="is-right">${fmtMoney(y.totalExpense)}</td>
      <td class="is-right ${toneClass(hasRev ? y.shopBalance : null)}"><strong>${hasRev ? fmtMoney(y.shopBalance) : '—'}</strong></td>
      <td class="is-right ${toneClass(hasRev ? y.allBalance : null)}">${hasRev ? fmtMoney(y.allBalance) : '—'}</td>
    </tr></tfoot>`;
    return head + `<tbody>${body}</tbody>` + foot;
}

function detailHtml(r: MonthSummary, household: ReadonlySet<string>): string {
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
                const isHousehold = isOther && household.has(householdKey(it.name));
                return `<li class="expense-item ${isHousehold ? 'is-household' : ''}">
              ${isOther ? `<label class="expense-household-toggle" title="勾選為家用（同名項目之後每月自動套用）">
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
          <div class="expense-block-head"><span>${cat}</span><span>${fmtMoney(r.totals[cat])}</span></div>
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
