/**
 * 代送費分頁（P3.2.7）。
 *
 * 對應桌面版 `deliverFee/DeliveryFeeTab.kt`：
 *  1) 單一 DropZone：帳單 .xlsx
 *  2) 統計區：總客戶數 + 線別分布
 *  3) 開始處理按鈕：呼叫 DeliveryFeeWriter，輸出單一 .xlsx 並下載
 *
 * 加值：上傳成功後立即掃描 bill 內所有商品名，比對 cargo_sort.csv，
 * 列出「找不到」或「代送費欄位空白」的商品，引導使用者點擊補填。
 */

import {saveAs} from 'file-saver';

import {createDropZone, type DropZoneController} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';

import {processBillFile} from '@/domain/process-bill';
import {DeliveryFeeWriter} from '@/writers/delivery-fee-writer';
import {findDeliveryFeeStatus} from '@/domain/sorting-list';
import {openUnsetDeliveryFeeDialog} from '@/tabs/unset-delivery-fee-dialog';
import {settingsJumpButtonHtml} from '@/ui/tabs';

import type {Bill} from '@/domain/models/bill';

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
        ${settingsJumpButtonHtml('cargo', '帳單排序')}
      </header>

      <div class="card-section">
        <div class="card-section-label">檔案上傳</div>
        <div id="delivery-dropzone"></div>
      </div>

      <div class="card-section" id="delivery-stats-section" hidden>
        <div class="card-section-label">統計</div>
        <div class="metric-grid" id="delivery-stats"></div>
      </div>

      <details class="analytics-unmatched" data-role="unset-fee" hidden>
        <summary><span data-role="unset-fee-summary"></span></summary>
        <ul class="analytics-unmatched-list" data-role="unset-fee-list"></ul>
      </details>

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
    const unsetFeeHost = panel.querySelector<HTMLDetailsElement>('[data-role="unset-fee"]')!;
    const unsetFeeSummary = panel.querySelector<HTMLElement>('[data-role="unset-fee-summary"]')!;
    const unsetFeeList = panel.querySelector<HTMLElement>('[data-role="unset-fee-list"]')!;

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
          renderUnsetFee();
        refreshButton();
      } catch (err) {
        state.bill = null;
        state.billFileName = null;
        billDz.setStatus('error', err instanceof Error ? err.message : '解析失敗');
        renderStats();
          renderUnsetFee();
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

    /**
     * 掃描帳單中所有不重複商品名，列出 cargo_sort.csv 找不到（missing）
     * 或代送費欄位空白（unfilled）的清單，引導使用者點擊補填。
     * 同名商品保留最先出現的判定結果，順序依商品在帳單中首次出現。
     */
    function renderUnsetFee(): void {
        if (!state.bill) {
            unsetFeeHost.hidden = true;
            unsetFeeList.innerHTML = '';
            return;
        }

        const seen = new Set<string>();
        const unset: Array<{ name: string; kind: 'missing' | 'unfilled' }> = [];
        for (const customer of state.bill.customerModels) {
            for (const product of customer.productList) {
                const name = product.name;
                if (!name || seen.has(name)) continue;
                seen.add(name);
                const status = findDeliveryFeeStatus(name);
                if (status.kind === 'missing' || status.kind === 'unfilled') {
                    unset.push({name, kind: status.kind});
                }
            }
        }

        if (unset.length === 0) {
            unsetFeeHost.hidden = true;
            unsetFeeList.innerHTML = '';
            return;
        }

        unsetFeeHost.hidden = false;
        unsetFeeHost.open = true;
        unsetFeeSummary.textContent = `未填代送費商品（${unset.length}）— 點擊任一項即可填入代送費（未填者輸出時將以 0 計）`;
        unsetFeeList.innerHTML = unset
            .map(
                (it) =>
                    `<li><button type="button" class="analytics-unmatched-chip" data-name="${escapeAttr(it.name)}" title="${it.kind === 'missing' ? '未列於排序表' : '排序表中代送費欄位為空白'}">${escapeHtml(it.name)}</button></li>`
            )
            .join('');
        unsetFeeList.querySelectorAll<HTMLButtonElement>('.analytics-unmatched-chip').forEach((btn) => {
            btn.addEventListener('click', () => {
                const name = btn.dataset.name ?? '';
                if (!name) return;
                void openUnsetDeliveryFeeDialog({
                    productName: name,
                    onSaved: () => {
                        renderUnsetFee();
                    },
                });
            });
        });
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

function escapeAttr(value: string): string {
    return escapeHtml(value);
}
