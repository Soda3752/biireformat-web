import {saveAs} from 'file-saver';

import {getTotalCount} from '@/domain/models/daily-product';
import {
    buildDailyResultForDate,
    type DailyCountParseResult,
    type DailyCountResult,
    parseDailyCount,
} from '@/readers/daily-count-reader';
import {buildDailyCountFilename, formatDateDisplay, writeDailyCount} from '@/writers/daily-count-writer';
import {createDateCalendar} from '@/ui/date-picker-calendar';
import {createDropZone} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';

export function renderDailyCountPanel(tab: TabDefinition): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = tab.id;
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">單日數量統計</h1>
        <p class="card-subtitle">
          上傳「應收帳款對帳單明細表」xlsx，可在月曆上選擇有資料的日期，
          系統依
          <code>daily_report_list.csv</code>
          將該日商品分組統計。未對應到的商品會自動歸入「其他」分類。
        </p>
      </header>

      <div data-role="zone-host"></div>

      <section data-role="result" hidden>
        <div class="metric-grid">
          <div class="metric metric-date">
            <div class="metric-label">輸出日期</div>
            <div data-role="calendar-host"></div>
            <div class="metric-date-hint" data-role="m-date-hint"></div>
          </div>
          <div class="metric">
            <div class="metric-label">已對應</div>
            <div class="metric-value" data-role="m-matched"></div>
          </div>
          <div class="metric">
            <div class="metric-label">「其他」筆數</div>
            <div class="metric-value" data-role="m-other"></div>
          </div>
          <div class="metric">
            <div class="metric-label">總分組</div>
            <div class="metric-value" data-role="m-groups"></div>
          </div>
        </div>

        <div class="card-section-label">分組總量</div>
        <div class="summary-list" data-role="group-summary"></div>
      </section>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="overall-status">請選擇對帳單明細 .xlsx</div>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-secondary" data-role="reset">重設</button>
          <button type="button" class="btn btn-primary btn-lg" data-role="export" disabled>下載統計檔</button>
        </div>
      </footer>
    </div>
  `;

  const zoneHost = panel.querySelector<HTMLElement>('[data-role="zone-host"]')!;
  const resultSection = panel.querySelector<HTMLElement>('[data-role="result"]')!;
    const calendarHost = panel.querySelector<HTMLElement>('[data-role="calendar-host"]')!;
    const dateHint = panel.querySelector<HTMLElement>('[data-role="m-date-hint"]')!;
  const metricMatched = panel.querySelector<HTMLElement>('[data-role="m-matched"]')!;
  const metricOther = panel.querySelector<HTMLElement>('[data-role="m-other"]')!;
  const metricGroups = panel.querySelector<HTMLElement>('[data-role="m-groups"]')!;
  const groupSummary = panel.querySelector<HTMLElement>('[data-role="group-summary"]')!;
  const overallStatus = panel.querySelector<HTMLElement>('[data-role="overall-status"]')!;
  const exportBtn = panel.querySelector<HTMLButtonElement>('[data-role="export"]')!;
  const resetBtn = panel.querySelector<HTMLButtonElement>('[data-role="reset"]')!;

    let parsed: DailyCountParseResult | null = null;
    let dailyResult: DailyCountResult | null = null;
    let selectedDateKey: string | null = null;

    const dateKeyToDate = (key: string): Date => {
        const [y, m, d] = key.split('-').map((v) => Number.parseInt(v, 10));
        return new Date(y, m - 1, d);
    };

    const calendar = createDateCalendar({
        availableDates: [],
        initialDate: null,
        emptyHint: '請先上傳檔案',
        onSelect: async (key) => {
            if (!parsed) return;
            selectedDateKey = key;
            try {
                dailyResult = await buildDailyResultForDate(parsed.byDate, key);
                renderDailyResult(dailyResult, key);
                exportBtn.disabled = false;
                overallStatus.textContent = `已選擇 ${formatDateDisplay(dateKeyToDate(key))}`;
            } catch (err) {
                console.error(err);
                showToast({
                    variant: 'error',
                    title: '統計失敗',
                    message: err instanceof Error ? err.message : String(err),
                });
            }
        },
    });
    calendarHost.appendChild(calendar.element);

  const zone = createDropZone({
      title: '對帳單明細表',
    hint: '拖曳或點擊上傳 .xlsx',
    accept: '.xlsx,.xls',
    onFile: async (file) => {
      try {
          const result = await parseDailyCount(file);
          parsed = result;

          if (result.availableDates.length === 0) {
              dailyResult = null;
              selectedDateKey = null;
              calendar.setAvailableDates([]);
              calendar.setSelected(null);
              clearResult();
              resultSection.hidden = true;
              exportBtn.disabled = true;
              zone.setStatus('error', '檔案內未找到任何日期資料');
              overallStatus.textContent = '檔案內無有效銷售資料';
              return;
          }

          // 預設選擇第一筆有資料的日期
          const firstDate = result.availableDates[0];
          calendar.setAvailableDates(result.availableDates);
          calendar.setSelected(firstDate);
          selectedDateKey = firstDate;
          dailyResult = await buildDailyResultForDate(result.byDate, firstDate);

          resultSection.hidden = false;
          renderDailyResult(dailyResult, firstDate);
        exportBtn.disabled = false;

          zone.setStatus(
              'loaded',
              `${file.name}　含 ${result.availableDates.length} 個日期`
          );
        overallStatus.textContent = `已就緒　${file.name}`;
      } catch (err) {
          parsed = null;
          dailyResult = null;
          selectedDateKey = null;
          calendar.setAvailableDates([]);
          calendar.setSelected(null);
        const message = err instanceof Error ? err.message : '讀取失敗';
        zone.setStatus('error', message);
        clearResult();
          resultSection.hidden = true;
        exportBtn.disabled = true;
        overallStatus.textContent = '讀取失敗，請檢查檔案';
        throw err;
      }
    },
  });
  zoneHost.appendChild(zone.element);

    const renderDailyResult = (r: DailyCountResult, dateKey: string) => {
    resultSection.hidden = false;
        dateHint.textContent = `xlsx 第 1 列將寫入 ${formatDateDisplay(dateKeyToDate(dateKey))}`;
    metricMatched.textContent = String(r.matched);
    metricOther.textContent = String(r.otherCount);

        const nonEmptyGroups = [...r.map.entries()].filter(([, ps]) =>
            ps.some((p) => getTotalCount(p) > 0)
        );
    metricGroups.textContent = String(nonEmptyGroups.length);

    groupSummary.innerHTML = '';
    for (const [groupName, products] of nonEmptyGroups) {
      const total = products.reduce((sum, p) => sum + getTotalCount(p), 0);
      const row = document.createElement('div');
      row.className = 'summary-list-row';
      if (groupName === '其他') row.classList.add('is-other');
      const name = document.createElement('span');
      name.className = 'summary-list-name';
        const visibleCount = products.filter((p) => getTotalCount(p) > 0).length;
        name.textContent = `${groupName}（${visibleCount} 項）`;
      const value = document.createElement('span');
      value.className = 'summary-list-value';
      value.textContent = `總量 ${total}`;
      row.appendChild(name);
      row.appendChild(value);
      groupSummary.appendChild(row);
    }
  };

  const clearResult = () => {
    groupSummary.innerHTML = '';
      metricMatched.textContent = '';
      metricOther.textContent = '';
      metricGroups.textContent = '';
      dateHint.textContent = '';
  };

  resetBtn.addEventListener('click', () => {
      parsed = null;
      dailyResult = null;
      selectedDateKey = null;
    zone.reset();
      calendar.setAvailableDates([]);
      calendar.setSelected(null);
    clearResult();
      resultSection.hidden = true;
    exportBtn.disabled = true;
      overallStatus.textContent = '請選擇對帳單明細 .xlsx';
  });

  exportBtn.addEventListener('click', async () => {
      if (!dailyResult || !selectedDateKey) return;
    exportBtn.disabled = true;
    const original = exportBtn.textContent;
    exportBtn.textContent = '處理中...';
    try {
        const exportDate = dateKeyToDate(selectedDateKey);
        const blob = await writeDailyCount(dailyResult.map, exportDate);
        const filename = buildDailyCountFilename(exportDate);
      saveAs(blob, filename);
      showToast({
        variant: 'success',
        title: '統計檔已輸出',
        message: filename,
      });
    } catch (err) {
      console.error(err);
      showToast({
        variant: 'error',
        title: '輸出失敗',
        message: err instanceof Error ? err.message : String(err),
      });
    } finally {
      exportBtn.textContent = original;
      exportBtn.disabled = false;
    }
  });

  return panel;
}
