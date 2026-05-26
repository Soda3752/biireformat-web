/**
 * 對帳 2.0 分頁
 *
 * 整合「明細（帳單 .xlsx 應收）」與「對帳（銀行對帳單 已收）」：
 *  - 兩個 dropzone 並列上傳
 *  - 兩份檔案皆就緒 + 末五碼對照表非空 → 計算 ReconcileResult、渲染主表 + 警示區
 *  - 輸出單一 .xlsx，含 主表 / 需人工複核 / 原始交易 三個區塊
 *
 * 配對策略採嚴格 storeCode：未設店家編號的末五碼不會被歸到任何客戶。
 */

import {saveAs} from 'file-saver';

import {getBankInfos, loadBankInfos} from '@/domain/bank-info-loader';
import type {BankRowMatch} from '@/domain/bank-match-service';
import {matchTransRecord} from '@/domain/bank-match-service';
import {processBillFile} from '@/domain/process-bill';
import type {ReconcileResult} from '@/domain/bank-reconcile-service';
import {reconcileByCustomer} from '@/domain/bank-reconcile-service';
import type {Bill} from '@/domain/models/bill';
import {localSettings} from '@/infra/local-settings-store';
import {parseTransRecord} from '@/readers/trans-record-reader';
import {writeReconcileWorkbook} from '@/writers/bank-reconcile-writer-v2';
import {openBankAddDialog} from '@/ui/bank-add-dialog';
import {openBankEditDialog} from '@/ui/bank-edit-dialog';
import {createBankReconcileTable} from '@/ui/bank-reconcile-table';
import {createDropZone} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';

export function renderBankNameFormatV2Panel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">對帳匯總 2.0</h1>
        <p class="card-subtitle">
          同時上傳「帳單 .xlsx」與「銀行對帳單 .csv / .xlsx」，依設定頁的「末五碼對照表（含店家編號）」交叉核對每位客戶的應收與已收。
        </p>
      </header>

      <div class="notice-banner" data-role="notice-banner" hidden></div>

      <div class="reconcile-dropzones" data-role="dropzones"></div>

      <div data-role="preview-host"></div>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="overall-status">請上傳帳單與銀行對帳單</div>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-secondary" data-role="reset">重設</button>
          <button type="button" class="btn btn-primary btn-lg" data-role="export" disabled>輸出對帳匯總</button>
        </div>
      </footer>
    </div>
  `;

    let bill: Bill | null = null;
    let billFileName: string | null = null;
    let transRecord: string[][] | null = null;
    let transFileName: string | null = null;
    let lastResult: ReconcileResult | null = null;

    const banner = panel.querySelector<HTMLElement>('[data-role="notice-banner"]')!;
    const dropzonesHost = panel.querySelector<HTMLElement>('[data-role="dropzones"]')!;
    const previewHost = panel.querySelector<HTMLElement>('[data-role="preview-host"]')!;
    const overallStatus = panel.querySelector<HTMLElement>('[data-role="overall-status"]')!;
    const exportBtn = panel.querySelector<HTMLButtonElement>('[data-role="export"]')!;
    const resetBtn = panel.querySelector<HTMLButtonElement>('[data-role="reset"]')!;

    const previewTable = createBankReconcileTable({
        onAddClick: (row: BankRowMatch) => {
            openBankAddDialog({
                presetLastFiveDigit: row.account,
                presetSummary: row.summary,
                presetDate: row.date,
                onSaved: () => {
                    rebuildReconcile();
                },
            });
        },
        onCandidateClick: (candidate, item) => {
            openBankEditDialog({
                targetInfo: {
                    customerName: candidate.customerName,
                    storeCode: candidate.storeCode,
                    customerLine: candidate.customerLine,
                    lastFiveDigit: candidate.lastFiveDigit,
                },
                contextSummary: item.row.summary,
                contextDate: item.row.date,
                onSaved: () => {
                    rebuildReconcile();
                },
            });
        },
    });
    previewHost.appendChild(previewTable.element);

    const billZone = createDropZone({
        title: '帳單 .xlsx',
        hint: '拖曳或點擊上傳明細用的帳單檔',
        accept: '.xlsx',
        onFile: async (file) => {
            try {
                const parsed = await processBillFile(file);
                bill = parsed;
                billFileName = file.name;
                billZone.setStatus('loaded', `${file.name}　共 ${parsed.customerModels.length} 位客戶`);
                rebuildReconcile();
            } catch (err) {
                bill = null;
                billFileName = null;
                lastResult = null;
                previewTable.clear();
                const message = err instanceof Error ? err.message : '讀取失敗';
                billZone.setStatus('error', message);
                refreshOverall();
                throw err;
            }
        },
    });

    const transZone = createDropZone({
        title: '銀行對帳單',
        hint: '拖曳或點擊上傳 .csv（支援 Big5 / UTF-8）或 .xlsx',
        accept: '.csv,text/csv,.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        onFile: async (file) => {
            try {
                const parsed = await parseTransRecord(file);
                transRecord = parsed;
                transFileName = file.name;
                transZone.setStatus('loaded', `${file.name}　共 ${parsed.length} 筆`);
                rebuildReconcile();
            } catch (err) {
                transRecord = null;
                transFileName = null;
                lastResult = null;
                previewTable.clear();
                const message = err instanceof Error ? err.message : '讀取失敗';
                transZone.setStatus('error', message);
                refreshOverall();
                throw err;
            }
        },
    });

    dropzonesHost.appendChild(billZone.element);
    dropzonesHost.appendChild(transZone.element);

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
                '尚未建立末五碼對照表。請先到 <a href="#settings" class="notice-banner-link">設定 → 末五碼</a> 匯入或新增資料（記得補上店家編號）。';
        } else {
            banner.innerHTML =
                '末五碼對照表為空。請到 <a href="#settings" class="notice-banner-link">設定 → 末五碼</a> 至少新增一筆資料。';
        }
    };

    const rebuildReconcile = () => {
        if (bill === null || transRecord === null) {
            lastResult = null;
            previewTable.clear();
            refreshOverall();
            return;
        }
        const infos = getBankInfos();
        const bankResult = matchTransRecord(transRecord, infos);
        lastResult = reconcileByCustomer(bill, bankResult, infos);
        previewTable.setData(lastResult);
        refreshOverall();
    };

    const refreshOverall = () => {
        refreshBanner();
        const infos = getBankInfos();
        const ready = bill !== null && transRecord !== null && infos.length > 0;
        exportBtn.disabled = !ready;

        if (ready && lastResult) {
            const s = lastResult.summary;
            const warn = s.manualReviewCount > 0 ? `　／　待覆核 ${s.manualReviewCount} 筆` : '';
            overallStatus.textContent = `已就緒　客戶 ${s.customerCount} 間（已配對 ${s.matchedCount}）${warn}`;
        } else if (bill === null && transRecord === null) {
            overallStatus.textContent = '請上傳帳單與銀行對帳單';
        } else if (bill === null) {
            overallStatus.textContent = `尚缺帳單　已上傳對帳單：${transFileName ?? ''}`;
        } else if (transRecord === null) {
            overallStatus.textContent = `尚缺銀行對帳單　已上傳帳單：${billFileName ?? ''}`;
        } else if (infos.length === 0) {
            overallStatus.textContent = '尚缺末五碼對照表（請至設定頁建立）';
        }
    };

    const reset = () => {
        bill = null;
        billFileName = null;
        transRecord = null;
        transFileName = null;
        lastResult = null;
        previewTable.clear();
        billZone.reset();
        transZone.reset();
        refreshOverall();
    };

    resetBtn.addEventListener('click', reset);

    exportBtn.addEventListener('click', async () => {
        if (!lastResult) return;
        exportBtn.disabled = true;
        const original = exportBtn.textContent;
        exportBtn.textContent = '處理中...';
        try {
            const file = await writeReconcileWorkbook(lastResult);
            saveAs(file.blob, file.filename);
            showToast({
                variant: 'success',
                title: '對帳匯總已輸出',
                message: file.filename,
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

    window.addEventListener('hashchange', () => {
        if (window.location.hash === tab.hash) rebuildReconcile();
    });

    refreshOverall();

    return panel;
}
