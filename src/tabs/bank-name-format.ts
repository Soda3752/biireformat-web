import {saveAs} from 'file-saver';

import {getBankInfos, loadBankInfos} from '@/domain/bank-info-loader';
import type {BankMatchResult, BankRowMatch} from '@/domain/bank-match-service';
import {matchTransRecord} from '@/domain/bank-match-service';
import {localSettings} from '@/infra/local-settings-store';
import {parseTransRecord} from '@/readers/trans-record-reader';
import {buildBankResultFilename, writeBankNameMerged} from '@/writers/bank-name-writer';
import {openBankAddDialog} from '@/ui/bank-add-dialog';
import {createBankPreviewTable} from '@/ui/bank-preview-table';
import {createDropZone} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';

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
          上傳「銀行對帳單 (.csv)」，系統會自動依設定頁的「末五碼對照表」將每筆交易配對到對應的客戶名稱與線別，產出合併後的對帳結果。
        </p>
      </header>

      <div class="notice-banner" data-role="notice-banner" hidden></div>

      <div data-role="zone-host"></div>

      <div data-role="preview-host"></div>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="overall-status">請上傳銀行對帳單 .csv</div>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-secondary" data-role="reset">重設</button>
          <button type="button" class="btn btn-primary btn-lg" data-role="export" disabled>輸出對帳結果</button>
        </div>
      </footer>
    </div>
  `;

    let transRecord: string[][] | null = null;
    let lastResult: BankMatchResult | null = null;

    const banner = panel.querySelector<HTMLElement>('[data-role="notice-banner"]')!;
    const zoneHost = panel.querySelector<HTMLElement>('[data-role="zone-host"]')!;
    const previewHost = panel.querySelector<HTMLElement>('[data-role="preview-host"]')!;
    const overallStatus = panel.querySelector<HTMLElement>('[data-role="overall-status"]')!;
    const exportBtn = panel.querySelector<HTMLButtonElement>('[data-role="export"]')!;
    const resetBtn = panel.querySelector<HTMLButtonElement>('[data-role="reset"]')!;

    const previewTable = createBankPreviewTable({
        onAddClick: (row: BankRowMatch) => {
            openBankAddDialog({
                presetLastFiveDigit: row.account,
                presetSummary: row.summary,
                presetDate: row.date,
                onSaved: () => {
                    rebuildMatch();
                },
            });
        },
    });
    previewHost.appendChild(previewTable.element);

  const transZone = createDropZone({
    title: '銀行對帳單',
    hint: '拖曳或點擊上傳 .csv（支援 Big5 / UTF-8）',
    accept: '.csv,text/csv',
    onFile: async (file) => {
      try {
        const parsed = await parseTransRecord(file);
          transRecord = parsed;
        transZone.setStatus('loaded', `${file.name}　共 ${parsed.length} 筆`);
          rebuildMatch();
      } catch (err) {
          transRecord = null;
          lastResult = null;
          previewTable.clear();
        const message = err instanceof Error ? err.message : '讀取失敗';
        transZone.setStatus('error', message);
        refreshOverall();
        throw err;
      }
    },
  });

    zoneHost.appendChild(transZone.element);

    // 確保 cache 已載入（同步函式但保險起見呼叫一次）
    loadBankInfos();

    const refreshBanner = () => {
        const has = localSettings.hasLastFiveDigit();
        const count = getBankInfos().length;
        if (has && count > 0) {
            banner.hidden = true;
            banner.innerHTML = '';
            return;
        }
        banner.hidden = false;
        if (!has) {
            banner.innerHTML =
                '尚未建立末五碼對照表。請先到 <a href="#settings" class="notice-banner-link">設定 → 末五碼</a> 匯入或新增資料。';
        } else {
            banner.innerHTML =
                '末五碼對照表為空。請到 <a href="#settings" class="notice-banner-link">設定 → 末五碼</a> 至少新增一筆資料。';
        }
    };

    const rebuildMatch = () => {
        if (transRecord === null) {
            lastResult = null;
            previewTable.clear();
            refreshOverall();
            return;
        }
        const infos = getBankInfos();
        lastResult = matchTransRecord(transRecord, infos);
        previewTable.setData(lastResult);
        refreshOverall();
    };

  const refreshOverall = () => {
      refreshBanner();
      const infos = getBankInfos();
      const ready = transRecord !== null && infos.length > 0;
    exportBtn.disabled = !ready;
      if (ready && lastResult) {
          const {matchedCount, unmatchedCount, rows} = lastResult;
          const warn = unmatchedCount > 0 ? `（其中 ${unmatchedCount} 筆未配對）` : '';
          overallStatus.textContent = `已就緒　共 ${rows.length} 筆 ／ 已配對 ${matchedCount}${warn}`;
      } else if (ready) {
        overallStatus.textContent = `已就緒　末五碼 ${infos.length} 筆 ／ 對帳單 ${transRecord!.length} 筆`;
    } else if (infos.length === 0 && transRecord === null) {
        overallStatus.textContent = '請建立末五碼對照表並上傳對帳單';
    } else if (infos.length === 0) {
        overallStatus.textContent = '尚缺末五碼對照表（請至設定頁建立）';
    } else {
      overallStatus.textContent = '尚缺銀行對帳單';
    }
  };

  const reset = () => {
      transRecord = null;
      lastResult = null;
      previewTable.clear();
    transZone.reset();
    refreshOverall();
  };

  resetBtn.addEventListener('click', reset);

  exportBtn.addEventListener('click', async () => {
      const infos = getBankInfos();
      if (transRecord === null || infos.length === 0) return;
    exportBtn.disabled = true;
    const original = exportBtn.textContent;
    exportBtn.textContent = '處理中...';
    try {
        const blob = await writeBankNameMerged(infos, transRecord);
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

    // hashchange 時若回到對帳分頁，重新確認末五碼狀態（使用者可能剛在設定頁編輯）
    window.addEventListener('hashchange', () => {
        if (window.location.hash === tab.hash) rebuildMatch();
    });

    refreshOverall();

  return panel;
}
