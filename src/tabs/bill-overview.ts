/**
 * 明細分頁（P3.1.2）。
 *
 * 對應桌面版 `billOverView/BillOverViewTab.kt`：
 *  1) DropZone：帳單 .xlsx
 *  2) 客戶排序：改由「設定 → 客戶排序」管理（不再每次上傳 .xlsx）
 *  3) 統計區：依線別、月結、現金各列計
 *  4) 開始處理按鈕：呼叫 OverViewWriter，輸出單一 .xlsx 並下載
 */

import {saveAs} from 'file-saver';

import {createDropZone, type DropZoneController} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';

import {processBillFile} from '@/domain/process-bill';
import {getCustomerOrderOverview} from '@/domain/customer-order-loader';
import {localSettings} from '@/infra/local-settings-store';
import {OverViewWriter} from '@/writers/overview-writer';
import {getCashCustomer, getMonthlyCustomer} from '@/writers/sheet-extension';

import type {Bill} from '@/domain/models/bill';
import type {TabDefinition} from '@/ui/tabs';

interface State {
  bill: Bill | null;
  billFileName: string | null;
  isProcessing: boolean;
}

export function renderBillOverviewPanel(tab: TabDefinition): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = 'overview';
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">帳單明細總覽產生工具</h1>
        <p class="card-subtitle">上傳帳單 .xlsx，依設定頁的「客戶排序」輸出依線別分頁的明細總覽</p>
      </header>

      <div class="notice-banner" data-role="customer-order-banner" hidden></div>

      <div class="card-section">
        <div class="card-section-label">檔案上傳</div>
        <div id="overview-dropzones"></div>
      </div>

      <div class="card-section" id="overview-stats-section" hidden>
        <div class="card-section-label">統計</div>
        <div class="metric-grid" id="overview-stats"></div>
      </div>

      <div class="action-bar">
        <div class="action-bar-actions">
          <button type="button" class="btn btn-primary btn-lg" id="overview-process" disabled>
            開始處理
          </button>
        </div>
      </div>
    </div>
  `;

  const state: State = {
    bill: null,
    billFileName: null,
    isProcessing: false,
  };

    const banner = panel.querySelector<HTMLElement>('[data-role="customer-order-banner"]')!;
  const dropzonesHost = panel.querySelector<HTMLElement>('#overview-dropzones')!;
  const statsSection = panel.querySelector<HTMLElement>('#overview-stats-section')!;
  const statsHost = panel.querySelector<HTMLElement>('#overview-stats')!;
  const processBtn = panel.querySelector<HTMLButtonElement>('#overview-process')!;

  let billDz: DropZoneController;

  billDz = createDropZone({
    title: '選擇帳單 Excel 檔案',
    hint: '拖曳或點擊上傳 .xlsx',
    accept: '.xlsx',
    onFile: async (file) => {
      try {
        const bill = await processBillFile(file);
        state.bill = bill;
        state.billFileName = file.name;
        billDz.setStatus('loaded', `${file.name}（${bill.customerModels.length} 位客戶）`);
        renderStats();
        refreshButton();
      } catch (err) {
        state.bill = null;
        state.billFileName = null;
        billDz.setStatus('error', err instanceof Error ? err.message : '解析失敗');
        renderStats();
        refreshButton();
        throw err;
      }
    },
  });

  dropzonesHost.appendChild(billDz.element);

  processBtn.addEventListener('click', async () => {
    if (!state.bill || state.isProcessing) return;
    state.isProcessing = true;
    refreshButton();

    try {
        const orderList = getCustomerOrderOverview().map((e) => e.code);
        const writer = new OverViewWriter(state.bill, orderList);
      const file = await writer.write();
      saveAs(file.blob, file.filename);
      showToast({
        variant: 'success',
        title: '處理完成',
        message: `已輸出 ${file.filename}`,
      });
    } catch (err) {
      console.error(err);
      showToast({
        variant: 'error',
        title: '處理失敗',
        message: err instanceof Error ? err.message : String(err),
      });
    } finally {
      state.isProcessing = false;
      refreshButton();
    }
  });

    function refreshBanner(): void {
        const has = localSettings.hasCustomerOrderOverview();
        const count = getCustomerOrderOverview().length;
        if (has && count > 0) {
            banner.hidden = true;
            banner.innerHTML = '';
            return;
        }
        banner.hidden = false;
        banner.innerHTML = has
            ? '明細客戶排序為空。請到 <a href="#settings" class="notice-banner-link">設定 → 明細客戶</a> 至少新增一筆，否則輸出將以「找不到排在最後」順序產生。'
            : '尚未建立明細客戶排序。請到 <a href="#settings" class="notice-banner-link">設定 → 明細客戶</a> 匯入或新增資料；目前仍可處理，但客戶順序會缺乏依據。';
    }

  function renderStats(): void {
    if (!state.bill) {
      statsSection.hidden = true;
      statsHost.innerHTML = '';
      return;
    }
    statsSection.hidden = false;

    const lines = state.bill.groupedCustomerByLine();
    const monthly = getMonthlyCustomer(state.bill.customerModels);
    const cash = getCashCustomer(state.bill.customerModels);

    const items: Array<{ label: string; value: string }> = [];
    for (const [name, list] of [...lines].sort(([a], [b]) => a.localeCompare(b))) {
      items.push({ label: name, value: `${list.length} 間` });
    }
    if (monthly.length > 0) items.push({ label: '月結客戶', value: `${monthly.length} 間` });
    if (cash.length > 0) items.push({ label: '現金客戶', value: `${cash.length} 間` });

    statsHost.innerHTML = items
      .map(
        (it) => `
          <div class="metric">
            <div class="metric-label">${escapeHtml(it.label)}</div>
            <div class="metric-value">${escapeHtml(it.value)}</div>
          </div>
        `
      )
      .join('');
  }

  function refreshButton(): void {
    const ready = state.bill !== null && !state.isProcessing;
    processBtn.disabled = !ready;
    processBtn.textContent = state.isProcessing ? '處理中…' : '開始處理';
  }

    window.addEventListener('hashchange', () => {
        if (window.location.hash === tab.hash) refreshBanner();
    });

    refreshBanner();

  return panel;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
