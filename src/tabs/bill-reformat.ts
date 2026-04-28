/**
 * 帳單分頁（P2.15）。
 *
 * 對應桌面版 `billReformat/BillReformatTab.kt`：
 *  1) 兩個 DropZone：帳單 .xlsx + 排序 .xlsx
 *  2) Switch：半月結 / 全月結
 *  3) 統計區：依線別、月結、現金各列計
 *  4) 開始處理按鈕：呼叫 BillWriter 產出多檔，逐檔下載
 *     （ZIP 整合排在 P5.1，目前以 saveAs 多次觸發）
 */

import { saveAs } from 'file-saver';

import { createDropZone, type DropZoneController } from '@/ui/drop-zone';
import { showToast } from '@/ui/toast';

import { processBillFile } from '@/domain/process-bill';
import { parseOrderList } from '@/readers/order-reader';
import { BillWriter } from '@/writers/bill-writer';
import {
  getCashCustomer,
  getMonthlyCustomer,
  getHalfMonthlyCustomer,
} from '@/writers/sheet-extension';

import type { Bill } from '@/domain/models/bill';

interface State {
  bill: Bill | null;
  billFileName: string | null;
  orderList: string[];
  orderFileName: string | null;
  isFullMonth: boolean;
  isProcessing: boolean;
}

export function renderBillReformatPanel(): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = 'bill';
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">帳單產生工具</h1>
        <p class="card-subtitle">上傳帳單與排序檔，選擇結算模式後輸出</p>
      </header>

      <div class="card-section">
        <div class="card-section-label">檔案上傳</div>
        <div class="dual-dropzone" id="bill-dropzones"></div>
      </div>

      <div class="card-section" id="bill-stats-section" hidden>
        <div class="card-section-label">統計</div>
        <div class="metric-grid" id="bill-stats"></div>
      </div>

      <div class="action-bar">
        <label class="toggle">
          <input type="checkbox" id="bill-full-month" />
          <span class="toggle-track"><span class="toggle-thumb"></span></span>
          <span class="toggle-label" id="bill-mode-label">輸出半月結</span>
        </label>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-primary btn-lg" id="bill-process" disabled>
            開始處理
          </button>
        </div>
      </div>
    </div>
  `;

  const state: State = {
    bill: null,
    billFileName: null,
    orderList: [],
    orderFileName: null,
    isFullMonth: false,
    isProcessing: false,
  };

  const dropzonesHost = panel.querySelector<HTMLElement>('#bill-dropzones')!;
  const statsSection = panel.querySelector<HTMLElement>('#bill-stats-section')!;
  const statsHost = panel.querySelector<HTMLElement>('#bill-stats')!;
  const fullMonthToggle = panel.querySelector<HTMLInputElement>('#bill-full-month')!;
  const modeLabel = panel.querySelector<HTMLElement>('#bill-mode-label')!;
  const processBtn = panel.querySelector<HTMLButtonElement>('#bill-process')!;

  /* ------ DropZones ------ */
  let billDz: DropZoneController;
  let orderDz: DropZoneController;

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

  orderDz = createDropZone({
    title: '選擇排序 Excel 檔案',
    hint: '第一欄為客戶代碼',
    accept: '.xlsx',
    onFile: async (file) => {
      try {
        const list = await parseOrderList(file);
        state.orderList = list;
        state.orderFileName = file.name;
        orderDz.setStatus('loaded', `${file.name}（${list.length} 筆）`);
        refreshButton();
      } catch (err) {
        state.orderList = [];
        state.orderFileName = null;
        orderDz.setStatus('error', err instanceof Error ? err.message : '讀取排序檔失敗');
        refreshButton();
        throw err;
      }
    },
  });

  dropzonesHost.appendChild(billDz.element);
  dropzonesHost.appendChild(orderDz.element);

  /* ------ Toggle ------ */
  fullMonthToggle.addEventListener('change', () => {
    state.isFullMonth = fullMonthToggle.checked;
    modeLabel.textContent = state.isFullMonth ? '輸出全月結' : '輸出半月結';
  });

  /* ------ Process ------ */
  processBtn.addEventListener('click', async () => {
    if (!state.bill || state.isProcessing) return;
    state.isProcessing = true;
    refreshButton();

    try {
      const writer = new BillWriter(state.bill, state.orderList);
      const files = await writer.write(state.isFullMonth);

      if (files.length === 0) {
        showToast({
          variant: 'warning',
          title: '沒有產出檔案',
          message: '帳單中找不到對應結算模式的客戶',
        });
        return;
      }

      for (const f of files) {
        saveAs(f.blob, f.filename);
      }

      showToast({
        variant: 'success',
        title: '處理完成',
        message: `已輸出 ${files.length} 份檔案`,
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

  /* ------ Helpers ------ */

  function renderStats(): void {
    if (!state.bill) {
      statsSection.hidden = true;
      statsHost.innerHTML = '';
      return;
    }
    statsSection.hidden = false;

    const lines = state.bill.groupedCustomerByLine();
    const monthly = getMonthlyCustomer(state.bill.customerModels);
    const halfMonthly = getHalfMonthlyCustomer(state.bill.customerModels);
    const cash = getCashCustomer(state.bill.customerModels);

    const items: Array<{ label: string; value: string }> = [];
    for (const [name, list] of [...lines].sort(([a], [b]) => a.localeCompare(b))) {
      items.push({ label: name, value: `${list.length} 間` });
    }
    if (halfMonthly.length > 0) items.push({ label: '半月結', value: `${halfMonthly.length} 間` });
    if (monthly.length > 0) items.push({ label: '月結', value: `${monthly.length} 間` });
    if (cash.length > 0) items.push({ label: '現金', value: `${cash.length} 間` });

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
