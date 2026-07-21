/**
 * 路線分析：針對 1~9 線的互動式比較區，取代舊的「各路線營收排行」。
 * - 指標切換：營收 / 毛利 / 數量
 * - 路線勾選：自由選擇要顯示哪幾條線（預設全選），同步套用到比較表與三張圖
 * - 各線比較表：每條線一列（營收、毛利、毛利率、數量、佔比），可點欄位排序，附總計列
 * - 三張圖同時呈現：每日折線、佔比圓餅、總計條狀
 * 跟隨上方全域篩選器（setRows 傳入篩選後 rows）。
 */

import type {AnalyticsRow} from './dataset-builder';
import {groupBy, marginPct, type GroupSum} from './aggregators';
import {CATEGORY_PALETTE} from './chart-colors';
import {type ChartHandle, createChart, observeChartsResize} from './chart-manager';
import {type MetricKind, routeBarOption, routeDailyLineOption, routePieOption} from './chart-options';

type SortKey = 'line' | 'amount' | 'profit' | 'margin' | 'count' | 'share';

const METRICS: Array<{value: MetricKind; label: string}> = [
    {value: 'amount', label: '營收'},
    {value: 'profit', label: '毛利'},
    {value: 'count', label: '數量'},
];

export interface RouteAnalyticsController {
    element: HTMLElement;

    setRows(rows: ReadonlyArray<AnalyticsRow>): void;
}

/** 從線別 key（如「第3線」）取出路線號碼供排序與配色；無數字者排最後。 */
function routeNumber(key: string): number {
    const m = /\d+/.exec(key);
    return m ? parseInt(m[0], 10) : Number.MAX_SAFE_INTEGER;
}

/** 依路線號碼給固定顏色，讓折線／圓餅／條狀三張圖同一條線同色。 */
function routeColor(line: string): string {
    const n = routeNumber(line);
    const idx = Number.isFinite(n) && n !== Number.MAX_SAFE_INTEGER ? n - 1 : 0;
    return CATEGORY_PALETTE[((idx % CATEGORY_PALETTE.length) + CATEGORY_PALETTE.length) % CATEGORY_PALETTE.length];
}

const EMPTY_OPTION = {
    title: {text: '無資料', left: 'center', top: 'center', textStyle: {fontSize: 13, color: '#999'}},
} as const;

export function createRouteAnalyticsSection(): RouteAnalyticsController {
    let rows: ReadonlyArray<AnalyticsRow> = [];
    let metric: MetricKind = 'amount';
    // null = 尚未自訂（全選）；一旦使用者動過勾選就materialize成明確集合。
    let selected: Set<string> | null = null;
    let sortKey: SortKey = 'amount';
    let sortDesc = true;

    let lineChart: ChartHandle | null = null;
    let pieChart: ChartHandle | null = null;
    let barChart: ChartHandle | null = null;
    let mountPromise: Promise<void> | null = null;

    const root = document.createElement('section');
    root.className = 'analytics-route-analytics';
    root.innerHTML = `
      <div class="analytics-route-analytics-head">
        <div class="analytics-section-label">路線分析</div>
        <div class="analytics-metric-switch" data-role="metric"></div>
      </div>
      <div class="analytics-route-lines" data-role="lines"></div>
      <div class="analytics-route-table-wrap" data-role="table"></div>
      <div class="analytics-chart-grid analytics-route-charts">
        <div class="analytics-chart-card analytics-chart-card-wide">
          <div class="analytics-chart-header"><div class="analytics-chart-title">每日走勢</div></div>
          <div class="analytics-chart-body" data-role="chart-line"></div>
        </div>
        <div class="analytics-chart-card">
          <div class="analytics-chart-header"><div class="analytics-chart-title">路線佔比</div></div>
          <div class="analytics-chart-body" data-role="chart-pie"></div>
        </div>
        <div class="analytics-chart-card">
          <div class="analytics-chart-header"><div class="analytics-chart-title">路線總計比較</div></div>
          <div class="analytics-chart-body" data-role="chart-bar"></div>
        </div>
      </div>
    `;

    const metricHost = root.querySelector<HTMLElement>('[data-role="metric"]')!;
    const linesHost = root.querySelector<HTMLElement>('[data-role="lines"]')!;
    const tableHost = root.querySelector<HTMLElement>('[data-role="table"]')!;
    const lineContainer = root.querySelector<HTMLElement>('[data-role="chart-line"]')!;
    const pieContainer = root.querySelector<HTMLElement>('[data-role="chart-pie"]')!;
    const barContainer = root.querySelector<HTMLElement>('[data-role="chart-bar"]')!;

    const fmtNum = (v: number) => Math.round(v).toLocaleString('zh-TW');
    const fmtPct = (v: number | null) => (v === null ? '—' : `${v.toFixed(1)}%`);

    const availableLines = (): string[] => {
        const s = new Set<string>();
        for (const r of rows) s.add(r.line);
        return [...s].sort((a, b) => routeNumber(a) - routeNumber(b) || a.localeCompare(b));
    };

    const selectedLines = (available: ReadonlyArray<string>): string[] =>
        selected === null ? [...available] : available.filter((l) => selected!.has(l));

    // ===== 指標切換 =====
    const renderMetricSwitch = () => {
        metricHost.innerHTML = METRICS.map(
            (m) => `<button type="button" class="${m.value === metric ? 'is-active' : ''}" data-v="${m.value}">${m.label}</button>`
        ).join('');
        metricHost.querySelectorAll<HTMLButtonElement>('button').forEach((b) => {
            b.addEventListener('click', () => {
                const v = b.dataset.v as MetricKind;
                if (v === metric) return;
                metric = v;
                // 佔比欄與預設排序跟著指標走；若目前依「佔比」排序則沿用。
                render();
            });
        });
    };

    // ===== 路線勾選 =====
    const renderLineChips = (available: ReadonlyArray<string>) => {
        const sel = new Set(selectedLines(available));
        linesHost.innerHTML = available
            .map(
                (line) => `
                  <button type="button" class="analytics-route-line-chip${sel.has(line) ? ' is-active' : ''}" data-line="${line}">
                    <span class="analytics-route-line-dot" style="background:${routeColor(line)}"></span>${line}
                  </button>`
            )
            .join('');
        linesHost.querySelectorAll<HTMLButtonElement>('.analytics-route-line-chip').forEach((btn) => {
            btn.addEventListener('click', () => {
                const line = btn.dataset.line!;
                if (selected === null) selected = new Set(available); // 首次動作先materialize全選
                if (selected.has(line)) selected.delete(line);
                else selected.add(line);
                render();
            });
        });
    };

    // ===== 各線比較表 =====
    const renderTable = (lines: ReadonlyArray<string>) => {
        if (rows.length === 0 || lines.length === 0) {
            tableHost.innerHTML = '<div class="analytics-route-analytics-empty">無資料（請確認已勾選路線）</div>';
            return;
        }
        const gmap = new Map(groupBy(rows, 'line', 'amount', true).map((g) => [g.key, g]));
        // profit 取整，與圓餅的 metricValueFromGroup 一致，佔比才會和圓餅 {d}% 完全對得上。
        const metricVal = (g: GroupSum) => (metric === 'amount' ? g.amount : metric === 'profit' ? Math.round(g.profit) : g.count);
        // 佔比＝該線佔「已勾選各線」的比重，與圓餅一致：分母為勾選各線的指標值（負值取 0），
        // 故各列佔比加總等於總計列的佔比（正常為 100%，全部非正時為 0%）。
        const selectedGroups = lines
            .map((line) => gmap.get(line))
            .filter((g): g is GroupSum => g !== undefined);
        const shareDenom = selectedGroups.reduce((s, g) => s + Math.max(0, metricVal(g)), 0);

        interface Row {
            line: string;
            amount: number;
            profit: number;
            margin: number | null;
            count: number;
            share: number;
        }
        const data: Row[] = selectedGroups.map((g) => ({
            line: g.key,
            amount: g.amount,
            profit: g.profit,
            margin: marginPct(g.profit, g.amount),
            count: g.count,
            share: shareDenom > 0 ? (Math.max(0, metricVal(g)) / shareDenom) * 100 : 0,
        }));

        const dir = sortDesc ? -1 : 1;
        data.sort((a, b) => {
            if (sortKey === 'line') return dir * (routeNumber(a.line) - routeNumber(b.line));
            if (sortKey === 'margin') return dir * ((a.margin ?? -Infinity) - (b.margin ?? -Infinity));
            return dir * (a[sortKey] - b[sortKey]);
        });

        const totalAmount = data.reduce((s, r) => s + r.amount, 0);
        const totalProfit = data.reduce((s, r) => s + r.profit, 0);
        const totalCount = data.reduce((s, r) => s + r.count, 0);
        const totalShare = data.reduce((s, r) => s + r.share, 0);

        const cols: Array<{key: SortKey; label: string}> = [
            {key: 'line', label: '路線'},
            {key: 'amount', label: '營收'},
            {key: 'profit', label: '毛利'},
            {key: 'margin', label: '毛利率'},
            {key: 'count', label: '數量'},
            {key: 'share', label: '佔比'},
        ];
        const arrow = (key: SortKey) => (key === sortKey ? (sortDesc ? ' ▼' : ' ▲') : '');
        const neg = (v: number) => (v < 0 ? ' is-negative' : '');

        tableHost.innerHTML = `
          <table class="analytics-route-table">
            <thead>
              <tr>${cols
                  .map(
                      (c) =>
                          `<th data-sort="${c.key}" class="${c.key === 'line' ? 'is-text' : 'is-num'}${c.key === sortKey ? ' is-sorted' : ''}">${c.label}${arrow(c.key)}</th>`
                  )
                  .join('')}</tr>
            </thead>
            <tbody>
              ${data
                  .map(
                      (r) => `
                    <tr>
                      <td class="is-text"><span class="analytics-route-line-dot" style="background:${routeColor(r.line)}"></span>${r.line}</td>
                      <td class="is-num">${fmtNum(r.amount)}</td>
                      <td class="is-num${neg(r.profit)}">${fmtNum(r.profit)}</td>
                      <td class="is-num${neg(r.margin ?? 0)}">${fmtPct(r.margin)}</td>
                      <td class="is-num">${r.count.toLocaleString('zh-TW')}</td>
                      <td class="is-num">${r.share.toFixed(1)}%</td>
                    </tr>`
                  )
                  .join('')}
            </tbody>
            <tfoot>
              <tr>
                <td class="is-text">總計</td>
                <td class="is-num">${fmtNum(totalAmount)}</td>
                <td class="is-num${neg(totalProfit)}">${fmtNum(totalProfit)}</td>
                <td class="is-num${neg(marginPct(totalProfit, totalAmount) ?? 0)}">${fmtPct(marginPct(totalProfit, totalAmount))}</td>
                <td class="is-num">${totalCount.toLocaleString('zh-TW')}</td>
                <td class="is-num">${totalShare.toFixed(1)}%</td>
              </tr>
            </tfoot>
          </table>
        `;
        tableHost.querySelectorAll<HTMLTableCellElement>('th[data-sort]').forEach((th) => {
            th.addEventListener('click', () => {
                const key = th.dataset.sort as SortKey;
                if (key === sortKey) sortDesc = !sortDesc;
                else {
                    sortKey = key;
                    sortDesc = key !== 'line'; // 文字欄預設升冪，數值欄預設降冪
                }
                renderTable(selectedLines(availableLines()));
            });
        });
    };

    // ===== 圖表 =====
    const mountIfNeeded = (): Promise<void> => {
        if (!mountPromise) {
            mountPromise = (async () => {
                lineChart = await createChart(lineContainer);
                pieChart = await createChart(pieContainer);
                barChart = await createChart(barContainer);
                observeChartsResize([lineChart, pieChart, barChart]);
            })();
        }
        return mountPromise;
    };

    const renderCharts = async (lines: ReadonlyArray<string>) => {
        await mountIfNeeded();
        const empty = rows.length === 0 || lines.length === 0;
        lineChart!.setOption((empty ? EMPTY_OPTION : routeDailyLineOption(rows, metric, lines, routeColor)) as never);
        pieChart!.setOption((empty ? EMPTY_OPTION : routePieOption(rows, metric, lines, routeColor)) as never);
        barChart!.setOption((empty ? EMPTY_OPTION : routeBarOption(rows, metric, lines, routeColor)) as never);
        requestAnimationFrame(() => {
            lineChart?.resize();
            pieChart?.resize();
            barChart?.resize();
        });
    };

    const render = () => {
        const available = availableLines();
        renderMetricSwitch();
        renderLineChips(available);
        const lines = selectedLines(available);
        renderTable(lines);
        void renderCharts(lines);
    };

    return {
        element: root,
        setRows(next) {
            rows = next;
            render();
        },
    };
}
