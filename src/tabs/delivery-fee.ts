/**
 * 代送費分頁（P3.2.7）。
 *
 * 對應桌面版 `deliverFee/DeliveryFeeTab.kt`：
 *  1) 單一 DropZone：帳單 .xlsx
 *  2) 統計區：總客戶數 + 線別分布
 *  3) 開始處理按鈕：呼叫 DeliveryFeeWriter，輸出單一 .xlsx 並下載
 */

import { saveAs } from 'file-saver';

import { createDropZone, type DropZoneController } from '@/ui/drop-zone';
import { showToast } from '@/ui/toast';

import { processBillFile } from '@/domain/process-bill';
import { DeliveryFeeWriter } from '@/writers/delivery-fee-writer';

import type { Bill } from '@/domain/models/bill';

interface State {
  bill: Bill | null;
  billFileName: string | null;
  isProcessing: boolean;
}

export function renderDeliveryFeePanel(): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = 'delivery';
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">代送費計算工具</h1>
        <p class="card-subtitle">上傳帳單，依商品代送費單價計算每客戶與全部總計</p>
      </header>

      <div class="card-section">
        <div class="card-section-label">檔案上傳</div>
        <div id="delivery-dropzone"></div>
      </div>

      <div class="card-section" id="delivery-stats-section" hidden>
        <div class="card-section-label">統計</div>
        <div class="metric-grid" id="delivery-stats"></div>
      </div>

      <div class="action-bar">
        <div class="action-bar-actions">
          <button type="button" class="btn btn-primary btn-lg" id="delivery-process" disabled>
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

  const dropzoneHost = panel.querySelector<HTMLElement>('#delivery-dropzone')!;
  const statsSection = panel.querySelector<HTMLElement>('#delivery-stats-section')!;
  const statsHost = panel.querySelector<HTMLElement>('#delivery-stats')!;
  const processBtn = panel.querySelector<HTMLButtonElement>('#delivery-process')!;

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

  dropzoneHost.appendChild(billDz.element);

  processBtn.addEventListener('click', async () => {
    if (!state.bill || state.isProcessing) return;
    state.isProcessing = true;
    refreshButton();

    try {
      const writer = new DeliveryFeeWriter(state.bill);
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

    const items: Array<{ label: string; value: string }> = [
      { label: '總客戶數', value: `${state.bill.customerModels.length} 間` },
    ];
    for (const [name, list] of [...lines].sort(([a], [b]) => a.localeCompare(b))) {
      items.push({ label: name, value: `${list.length} 間` });
    }

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
