/**
 * 帳單分頁（P2.15）。
 *
 * 對應桌面版 `billReformat/BillReformatTab.kt`：
 *  1) DropZone：帳單 .xlsx
 *  2) 客戶排序：改由「設定 → 客戶排序」管理（不再每次上傳 .xlsx）
 *  3) Switch：半月結 / 全月結
 *  4) 統計區：依線別、月結、現金各列計
 *  5) 開始處理按鈕：呼叫 BillWriter 產出多檔，多檔包成 ZIP 後下載；單檔直接下載
 */

import {saveAs} from 'file-saver';
import JSZip from 'jszip';

import {createDropZone, type DropZoneController} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';

import {processBillFile} from '@/domain/process-bill';
import {getCustomerOrderBill} from '@/domain/customer-order-loader';
import {localSettings} from '@/infra/local-settings-store';
import {BillWriter} from '@/writers/bill-writer';
import {getCashCustomer, getHalfMonthlyCustomer, getMonthlyCustomer,} from '@/writers/sheet-extension';

import type {Bill} from '@/domain/models/bill';
import type {TabDefinition} from '@/ui/tabs';

const DATE_SHIFT_MIN = -30;
const DATE_SHIFT_MAX = 30;

interface State {
  bill: Bill | null;
  billFileName: string | null;
  isFullMonth: boolean;
  isProcessing: boolean;
  dateShiftDays: number;
}

export function renderBillReformatPanel(tab: TabDefinition): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = 'bill';
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">帳單產生工具</h1>
        <p class="card-subtitle">上傳帳單 .xlsx，依設定頁的「客戶排序」輸出</p>
      </header>

      <div class="notice-banner" data-role="customer-order-banner" hidden></div>

      <div class="card-section">
        <div class="card-section-label">檔案上傳</div>
        <div id="bill-dropzones"></div>
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
        <div class="date-shift" id="bill-date-shift">
          <span class="date-shift-label">日期校正</span>
          <div class="date-shift-control">
            <button type="button" class="date-shift-step" id="bill-date-shift-dec" aria-label="減一天">−</button>
            <input
              type="number"
              class="date-shift-input"
              id="bill-date-shift-input"
              value="0"
              step="1"
              min="${DATE_SHIFT_MIN}"
              max="${DATE_SHIFT_MAX}"
            />
            <button type="button" class="date-shift-step" id="bill-date-shift-inc" aria-label="加一天">+</button>
          </div>
          <span class="date-shift-hint" id="bill-date-shift-hint"></span>
        </div>
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
    isFullMonth: false,
    isProcessing: false,
    dateShiftDays: 0,
  };

    const banner = panel.querySelector<HTMLElement>('[data-role="customer-order-banner"]')!;
  const dropzonesHost = panel.querySelector<HTMLElement>('#bill-dropzones')!;
  const statsSection = panel.querySelector<HTMLElement>('#bill-stats-section')!;
  const statsHost = panel.querySelector<HTMLElement>('#bill-stats')!;
  const fullMonthToggle = panel.querySelector<HTMLInputElement>('#bill-full-month')!;
  const modeLabel = panel.querySelector<HTMLElement>('#bill-mode-label')!;
  const processBtn = panel.querySelector<HTMLButtonElement>('#bill-process')!;
  const dateShiftWrap = panel.querySelector<HTMLElement>('#bill-date-shift')!;
  const dateShiftInput = panel.querySelector<HTMLInputElement>('#bill-date-shift-input')!;
  const dateShiftDec = panel.querySelector<HTMLButtonElement>('#bill-date-shift-dec')!;
  const dateShiftInc = panel.querySelector<HTMLButtonElement>('#bill-date-shift-inc')!;
  const dateShiftHint = panel.querySelector<HTMLElement>('#bill-date-shift-hint')!;

    /* ------ DropZone ------ */
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
        resetDateShift();
        billDz.setStatus('loaded', `${file.name}（${bill.customerModels.length} 位客戶）`);
        renderStats();
        refreshDateShiftHint();
        refreshButton();
      } catch (err) {
        state.bill = null;
        state.billFileName = null;
        resetDateShift();
        billDz.setStatus('error', err instanceof Error ? err.message : '解析失敗');
        renderStats();
        refreshDateShiftHint();
        refreshButton();
        throw err;
      }
    },
  });

  dropzonesHost.appendChild(billDz.element);

  /* ------ Toggle ------ */
  fullMonthToggle.addEventListener('change', () => {
    state.isFullMonth = fullMonthToggle.checked;
    modeLabel.textContent = state.isFullMonth ? '輸出全月結' : '輸出半月結';
    if (state.isFullMonth) resetDateShift();
    refreshDateShiftDisabled();
    refreshDateShiftHint();
  });

  /* ------ Date shift ------ */
  dateShiftDec.addEventListener('click', () => applyDateShiftDelta(-1));
  dateShiftInc.addEventListener('click', () => applyDateShiftDelta(1));
  dateShiftInput.addEventListener('input', () => {
    const raw = parseInt(dateShiftInput.value, 10);
    setDateShift(Number.isFinite(raw) ? raw : 0, false);
  });
  dateShiftInput.addEventListener('blur', () => {
    // 失焦時將輸入值正規化（空白或非數字 → 0；超出範圍夾住）
    setDateShift(state.dateShiftDays, true);
  });

  /* ------ Process ------ */
  processBtn.addEventListener('click', async () => {
    if (!state.bill || state.isProcessing) return;
    state.isProcessing = true;
    refreshButton();

    try {
        const orderList = getCustomerOrderBill().map((e) => e.code);
      const shift = state.isFullMonth ? 0 : state.dateShiftDays;
      const writer = new BillWriter(state.bill, orderList, shift);
      const files = await writer.write(state.isFullMonth);

      if (files.length === 0) {
        showToast({
          variant: 'warning',
          title: '沒有產出檔案',
          message: '帳單中找不到對應結算模式的客戶',
        });
        return;
      }

      if (files.length === 1) {
        const only = files[0];
        saveAs(only.blob, only.filename);
        showToast({
          variant: 'success',
          title: '處理完成',
          message: `已輸出 ${only.filename}`,
        });
      } else {
        const zip = new JSZip();
        for (const f of files) {
          zip.file(f.filename, f.blob);
        }
        const zipBlob = await zip.generateAsync({type: 'blob'});
        const zipName = `帳單_${state.bill.billDateInfo.month}月.zip`;
        saveAs(zipBlob, zipName);
        showToast({
          variant: 'success',
          title: '處理完成',
          message: `已輸出 ${zipName}（含 ${files.length} 份檔案）`,
        });
      }
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

    function refreshBanner(): void {
        const has = localSettings.hasCustomerOrderBill();
        const count = getCustomerOrderBill().length;
        if (has && count > 0) {
            banner.hidden = true;
            banner.innerHTML = '';
            return;
        }
        banner.hidden = false;
        banner.innerHTML = has
            ? '帳單客戶排序為空。請到 <a href="#settings" class="notice-banner-link">設定 → 帳單客戶</a> 至少新增一筆，否則輸出將以「找不到排在最後」順序產生。'
            : '尚未建立帳單客戶排序。請到 <a href="#settings" class="notice-banner-link">設定 → 帳單客戶</a> 匯入或新增資料；目前仍可處理，但客戶順序會缺乏依據。';
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

  function clamp(value: number): number {
    return Math.max(DATE_SHIFT_MIN, Math.min(DATE_SHIFT_MAX, value));
  }

  function setDateShift(value: number, syncInput: boolean): void {
    const next = clamp(Math.trunc(value || 0));
    state.dateShiftDays = next;
    if (syncInput) dateShiftInput.value = String(next);
    refreshDateShiftHint();
  }

  function applyDateShiftDelta(delta: number): void {
    if (state.isFullMonth) return;
    setDateShift(state.dateShiftDays + delta, true);
  }

  function resetDateShift(): void {
    setDateShift(0, true);
  }

  function refreshDateShiftDisabled(): void {
    const disabled = state.isFullMonth;
    dateShiftInput.disabled = disabled;
    dateShiftDec.disabled = disabled;
    dateShiftInc.disabled = disabled;
    dateShiftWrap.classList.toggle('is-disabled', disabled);
  }

  function refreshDateShiftHint(): void {
    if (!state.bill || state.isFullMonth || state.dateShiftDays === 0) {
      dateShiftHint.textContent = '';
      return;
    }
    const {year, month, dateRange} = state.bill.billDateInfo;
    if (dateRange.length === 0) {
      dateShiftHint.textContent = '';
      return;
    }
    const wYear = parseInt(year, 10) + 1911;
    const mIdx = parseInt(month, 10) - 1;
    const start = new Date(wYear, mIdx, dateRange[0] + state.dateShiftDays);
    const end = new Date(wYear, mIdx, dateRange[dateRange.length - 1] + state.dateShiftDays);
    const fmt = (d: Date) => `${d.getMonth() + 1}/${d.getDate()}`;
    dateShiftHint.textContent = `→ ${fmt(start)}–${fmt(end)}`;
  }

    // 使用者從設定頁回來時 refresh banner
    window.addEventListener('hashchange', () => {
        if (window.location.hash === tab.hash) refreshBanner();
    });

    refreshBanner();
  refreshDateShiftDisabled();
  refreshDateShiftHint();

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
