/**
 * 明細分頁（P3.1.2）。
 *
 * 對應桌面版 `billOverView/BillOverViewTab.kt`：
 *  1) 兩個 DropZone：帳單 .xlsx + 排序 .xlsx
 *  2) 統計區：依線別、月結、現金各列計
 *  3) 開始處理按鈕：呼叫 OverViewWriter，輸出單一 .xlsx 並下載
 */

import { saveAs } from 'file-saver';

import { createDropZone, type DropZoneController } from '@/ui/drop-zone';
import { showToast } from '@/ui/toast';

import { processBillFile } from '@/domain/process-bill';
import { parseOrderList } from '@/readers/order-reader';
import { OverViewWriter } from '@/writers/overview-writer';
import { getCashCustomer, getMonthlyCustomer } from '@/writers/sheet-extension';

import type { Bill } from '@/domain/models/bill';

interface State {
  bill: Bill | null;
  billFileName: string | null;
  orderList: string[];
  orderFileName: string | null;
  isProcessing: boolean;
}

export function renderBillOverviewPanel(): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = 'overview';
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">帳單明細總覽產生工具</h1>
        <p class="card-subtitle">上傳帳單與排序檔，輸出依線別分頁的明細總覽</p>
      </header>

      <div class="card-section">
        <div class="card-section-label">檔案上傳</div>
        <div class="dual-dropzone" id="overview-dropzones"></div>
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
    orderList: [],
    orderFileName: null,
    isProcessing: false,
  };

  const dropzonesHost = panel.querySelector<HTMLElement>('#overview-dropzones')!;
  const statsSection = panel.querySelector<HTMLElement>('#overview-stats-section')!;
  const statsHost = panel.querySelector<HTMLElement>('#overview-stats')!;
  const processBtn = panel.querySelector<HTMLButtonElement>('#overview-process')!;

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

  processBtn.addEventListener('click', async () => {
    if (!state.bill || state.isProcessing) return;
    state.isProcessing = true;
    refreshButton();

    try {
      const writer = new OverViewWriter(state.bill, state.orderList);
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
