import { saveAs } from 'file-saver';

import { getTotalCount } from '@/domain/models/daily-product';
import { parseDailyCount, type DailyCountResult } from '@/readers/daily-count-reader';
import { buildDailyCountFilename, formatDateDisplay, writeDailyCount } from '@/writers/daily-count-writer';
import { createDropZone } from '@/ui/drop-zone';
import { showToast } from '@/ui/toast';
import type { TabDefinition } from '@/ui/tabs';

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
          上傳當日銷售 .xlsx，系統依
          <code>daily_report_list.csv</code>
          將商品分組統計。未對應到的商品會自動歸入「其他」分類。
        </p>
      </header>

      <div data-role="zone-host"></div>

      <section data-role="result" hidden>
        <div class="metric-grid">
          <div class="metric">
            <div class="metric-label">輸出日期</div>
            <div class="metric-value" data-role="m-date"></div>
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
        <div class="action-bar-status" data-role="overall-status">請選擇單日銷售 .xlsx</div>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-secondary" data-role="reset">重設</button>
          <button type="button" class="btn btn-primary btn-lg" data-role="export" disabled>下載統計檔</button>
        </div>
      </footer>
    </div>
  `;

  const zoneHost = panel.querySelector<HTMLElement>('[data-role="zone-host"]')!;
  const resultSection = panel.querySelector<HTMLElement>('[data-role="result"]')!;
  const metricDate = panel.querySelector<HTMLElement>('[data-role="m-date"]')!;
  const metricMatched = panel.querySelector<HTMLElement>('[data-role="m-matched"]')!;
  const metricOther = panel.querySelector<HTMLElement>('[data-role="m-other"]')!;
  const metricGroups = panel.querySelector<HTMLElement>('[data-role="m-groups"]')!;
  const groupSummary = panel.querySelector<HTMLElement>('[data-role="group-summary"]')!;
  const overallStatus = panel.querySelector<HTMLElement>('[data-role="overall-status"]')!;
  const exportBtn = panel.querySelector<HTMLButtonElement>('[data-role="export"]')!;
  const resetBtn = panel.querySelector<HTMLButtonElement>('[data-role="reset"]')!;

  let result: DailyCountResult | null = null;

  const zone = createDropZone({
    title: '單日銷售統計',
    hint: '拖曳或點擊上傳 .xlsx',
    accept: '.xlsx,.xls',
    onFile: async (file) => {
      try {
        const parsed = await parseDailyCount(file);
        result = parsed;
        zone.setStatus('loaded', `${file.name}　已對應 ${parsed.matched} 筆`);
        renderResult(parsed);
        exportBtn.disabled = false;
        overallStatus.textContent = `已就緒　${file.name}`;
      } catch (err) {
        result = null;
        const message = err instanceof Error ? err.message : '讀取失敗';
        zone.setStatus('error', message);
        clearResult();
        exportBtn.disabled = true;
        overallStatus.textContent = '讀取失敗，請檢查檔案';
        throw err;
      }
    },
  });
  zoneHost.appendChild(zone.element);

  const renderResult = (r: DailyCountResult) => {
    resultSection.hidden = false;
    metricDate.textContent = formatDateDisplay(new Date());
    metricMatched.textContent = String(r.matched);
    metricOther.textContent = String(r.otherCount);

    const nonEmptyGroups = [...r.map.entries()].filter(([, ps]) => ps.length > 0);
    metricGroups.textContent = String(nonEmptyGroups.length);

    groupSummary.innerHTML = '';
    for (const [groupName, products] of nonEmptyGroups) {
      const total = products.reduce((sum, p) => sum + getTotalCount(p), 0);
      const row = document.createElement('div');
      row.className = 'summary-list-row';
      if (groupName === '其他') row.classList.add('is-other');
      const name = document.createElement('span');
      name.className = 'summary-list-name';
      name.textContent = `${groupName}（${products.length} 項）`;
      const value = document.createElement('span');
      value.className = 'summary-list-value';
      value.textContent = `總量 ${total}`;
      row.appendChild(name);
      row.appendChild(value);
      groupSummary.appendChild(row);
    }
  };

  const clearResult = () => {
    resultSection.hidden = true;
    groupSummary.innerHTML = '';
  };

  resetBtn.addEventListener('click', () => {
    result = null;
    zone.reset();
    clearResult();
    exportBtn.disabled = true;
    overallStatus.textContent = '請選擇單日銷售 .xlsx';
  });

  exportBtn.addEventListener('click', async () => {
    if (!result) return;
    exportBtn.disabled = true;
    const original = exportBtn.textContent;
    exportBtn.textContent = '處理中...';
    try {
      const blob = await writeDailyCount(result.map);
      const filename = buildDailyCountFilename();
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
