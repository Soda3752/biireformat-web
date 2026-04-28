import { saveAs } from 'file-saver';

import type { BankInfo } from '@/domain/models/bank-info';
import { parseBankInfo } from '@/readers/bank-info-reader';
import { parseTransRecord } from '@/readers/trans-record-reader';
import { buildBankResultFilename, writeBankNameMerged } from '@/writers/bank-name-writer';
import { createDropZone, type DropZoneController } from '@/ui/drop-zone';
import { showToast } from '@/ui/toast';
import type { TabDefinition } from '@/ui/tabs';

interface BankState {
  bankInfos: BankInfo[] | null;
  transRecord: string[][] | null;
}

export function renderBankNameFormatPanel(tab: TabDefinition): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = tab.id;
  panel.setAttribute('role', 'tabpanel');

  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">銀行對帳格式化</h1>
        <p class="card-subtitle">
          上傳「末五碼對照表 (.xlsx)」與「銀行對帳單 (.csv)」，系統會自動將每筆交易末五碼配對到對應的客戶名稱與線別，產出合併後的對帳結果。
        </p>
      </header>

      <div class="dual-dropzone" data-role="zones"></div>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="overall-status">請先選擇兩個必要檔案</div>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-secondary" data-role="reset">重設</button>
          <button type="button" class="btn btn-primary btn-lg" data-role="export" disabled>輸出對帳結果</button>
        </div>
      </footer>
    </div>
  `;

  const state: BankState = {
    bankInfos: null,
    transRecord: null,
  };

  const zonesHost = panel.querySelector<HTMLElement>('[data-role="zones"]')!;
  const overallStatus = panel.querySelector<HTMLElement>('[data-role="overall-status"]')!;
  const exportBtn = panel.querySelector<HTMLButtonElement>('[data-role="export"]')!;
  const resetBtn = panel.querySelector<HTMLButtonElement>('[data-role="reset"]')!;

  const bankZone = createDropZone({
    title: '末五碼對照表',
    hint: '拖曳或點擊上傳 .xlsx',
    accept: '.xlsx,.xls',
    onFile: async (file) => {
      try {
        const parsed = await parseBankInfo(file);
        state.bankInfos = parsed;
        bankZone.setStatus('loaded', `${file.name}　共 ${parsed.length} 筆`);
        refreshOverall();
      } catch (err) {
        state.bankInfos = null;
        const message = err instanceof Error ? err.message : '讀取失敗';
        bankZone.setStatus('error', message);
        refreshOverall();
        throw err;
      }
    },
  });

  const transZone = createDropZone({
    title: '銀行對帳單',
    hint: '拖曳或點擊上傳 .csv（支援 Big5 / UTF-8）',
    accept: '.csv,text/csv',
    onFile: async (file) => {
      try {
        const parsed = await parseTransRecord(file);
        state.transRecord = parsed;
        transZone.setStatus('loaded', `${file.name}　共 ${parsed.length} 筆`);
        refreshOverall();
      } catch (err) {
        state.transRecord = null;
        const message = err instanceof Error ? err.message : '讀取失敗';
        transZone.setStatus('error', message);
        refreshOverall();
        throw err;
      }
    },
  });

  zonesHost.appendChild(wrapZone('末五碼對照表 (.xlsx)', bankZone));
  zonesHost.appendChild(wrapZone('銀行對帳單 (.csv)', transZone));

  const refreshOverall = () => {
    const ready = state.bankInfos !== null && state.transRecord !== null;
    exportBtn.disabled = !ready;
    if (ready) {
      overallStatus.textContent = `已就緒　末五碼 ${state.bankInfos!.length} 筆 ／ 對帳單 ${state.transRecord!.length} 筆`;
    } else if (state.bankInfos === null && state.transRecord === null) {
      overallStatus.textContent = '請先選擇兩個必要檔案';
    } else if (state.bankInfos === null) {
      overallStatus.textContent = '尚缺末五碼對照表';
    } else {
      overallStatus.textContent = '尚缺銀行對帳單';
    }
  };

  const reset = () => {
    state.bankInfos = null;
    state.transRecord = null;
    bankZone.reset();
    transZone.reset();
    refreshOverall();
  };

  resetBtn.addEventListener('click', reset);

  exportBtn.addEventListener('click', async () => {
    if (state.bankInfos === null || state.transRecord === null) return;
    exportBtn.disabled = true;
    const original = exportBtn.textContent;
    exportBtn.textContent = '處理中...';
    try {
      const blob = await writeBankNameMerged(state.bankInfos, state.transRecord);
      const filename = buildBankResultFilename();
      saveAs(blob, filename);
      showToast({
        variant: 'success',
        title: '對帳結果已輸出',
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
      refreshOverall();
    }
  });

  return panel;
}

const wrapZone = (label: string, ctrl: DropZoneController): HTMLElement => {
  const wrapper = document.createElement('div');
  wrapper.className = 'dual-dropzone-item';
  const labelEl = document.createElement('div');
  labelEl.className = 'card-section-label';
  labelEl.textContent = label;
  wrapper.appendChild(labelEl);
  wrapper.appendChild(ctrl.element);
  return wrapper;
};
