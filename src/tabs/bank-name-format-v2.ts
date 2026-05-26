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
import type {BankMatchResult, BankRowMatch} from '@/domain/bank-match-service';
import {matchTransRecord} from '@/domain/bank-match-service';
import {processBillFile} from '@/domain/process-bill';
import type {ReconcileResult} from '@/domain/bank-reconcile-service';
import {DEFAULT_FEE_TOLERANCE, reconcileByCustomer} from '@/domain/bank-reconcile-service';
import type {BankInfo} from '@/domain/models/bank-info';
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
          同時上傳「帳單 .xlsx」與「銀行對帳單 .csv / .xlsx（可多檔）」，依設定頁的「末五碼對照表（含店家編號）」交叉核對每位客戶的應收與已收。
        </p>
      </header>

      <div class="notice-banner" data-role="notice-banner" hidden></div>

      <div class="reconcile-options">
        <label class="reconcile-options-field">
          <span class="reconcile-options-label">手續費容差 (元)</span>
          <input
            type="number"
            min="0"
            step="1"
            class="reconcile-options-input"
            data-role="fee-tolerance"
          />
          <span class="reconcile-options-hint">用於濾掉跨月匯款：每位客戶取「金額最接近應收且差額 ≤ 此值」那筆當本月款，其餘列標記為「跨月」。</span>
        </label>
      </div>

      <div class="reconcile-dropzones" data-role="dropzones">
        <div data-role="bill-slot"></div>
        <div class="reconcile-trans-slot">
          <div data-role="trans-slot"></div>
          <div class="reconcile-trans-list" data-role="trans-list" hidden></div>
        </div>
      </div>

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

    interface TransFileEntry {
        name: string;
        rows: string[][];
    }

    const FEE_TOLERANCE_KEY = 'bii.reconcile.feeTolerance';

    const readStoredFeeTolerance = (): number => {
        const raw = localStorage.getItem(FEE_TOLERANCE_KEY);
        if (raw === null) return DEFAULT_FEE_TOLERANCE;
        const n = Number(raw);
        return Number.isFinite(n) && n >= 0 ? n : DEFAULT_FEE_TOLERANCE;
    };

    let bill: Bill | null = null;
    let billFileName: string | null = null;
    const transFiles: TransFileEntry[] = [];
    let lastResult: ReconcileResult | null = null;
    let feeTolerance = readStoredFeeTolerance();

    const banner = panel.querySelector<HTMLElement>('[data-role="notice-banner"]')!;
    const feeToleranceInput = panel.querySelector<HTMLInputElement>('[data-role="fee-tolerance"]')!;
    const billSlot = panel.querySelector<HTMLElement>('[data-role="bill-slot"]')!;
    const transSlot = panel.querySelector<HTMLElement>('[data-role="trans-slot"]')!;
    const transListEl = panel.querySelector<HTMLElement>('[data-role="trans-list"]')!;
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
        hint: '拖曳或點擊上傳（可多檔）.csv（Big5 / UTF-8）或 .xlsx',
        accept: '.csv,text/csv,.xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        multiple: true,
        onFile: async (file) => {
            if (transFiles.some((f) => f.name === file.name)) {
                showToast({
                    variant: 'info',
                    title: '已存在相同檔名',
                    message: `${file.name} 已在清單中，跳過。`,
                });
                return;
            }
            try {
                const parsed = await parseTransRecord(file);
                transFiles.push({name: file.name, rows: parsed});
                renderTransList();
                rebuildReconcile();
            } catch (err) {
                const message = err instanceof Error ? err.message : '讀取失敗';
                showToast({
                    variant: 'error',
                    title: `${file.name} 讀取失敗`,
                    message,
                });
                console.error(err);
            }
        },
    });

    billSlot.appendChild(billZone.element);
    transSlot.appendChild(transZone.element);

    const renderTransList = () => {
        if (transFiles.length === 0) {
            transListEl.hidden = true;
            transListEl.innerHTML = '';
            return;
        }
        transListEl.hidden = false;
        transListEl.innerHTML = '';

        const titleEl = document.createElement('div');
        titleEl.className = 'reconcile-trans-list-title';
        titleEl.textContent = `已上傳 ${transFiles.length} 個`;
        transListEl.appendChild(titleEl);

        for (const entry of transFiles) {
            const item = document.createElement('div');
            item.className = 'reconcile-trans-list-item';

            const nameEl = document.createElement('span');
            nameEl.className = 'reconcile-trans-list-name';
            nameEl.textContent = entry.name;

            const metaEl = document.createElement('span');
            metaEl.className = 'reconcile-trans-list-meta';
            metaEl.textContent = `${entry.rows.length} 筆`;

            const removeBtn = document.createElement('button');
            removeBtn.type = 'button';
            removeBtn.className = 'reconcile-trans-list-remove';
            removeBtn.setAttribute('aria-label', `移除 ${entry.name}`);
            removeBtn.textContent = '✕';
            removeBtn.addEventListener('click', () => {
                const idx = transFiles.indexOf(entry);
                if (idx >= 0) transFiles.splice(idx, 1);
                renderTransList();
                rebuildReconcile();
            });

            item.appendChild(nameEl);
            item.appendChild(metaEl);
            item.appendChild(removeBtn);
            transListEl.appendChild(item);
        }
    };

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
        if (bill === null || transFiles.length === 0) {
            lastResult = null;
            previewTable.clear();
            refreshOverall();
            return;
        }
        const infos = getBankInfos();
        const bankResult = matchAllTransFiles(transFiles, infos);
        lastResult = reconcileByCustomer(bill, bankResult, infos, {feeTolerance});
        previewTable.setData(lastResult);
        refreshOverall();
    };

    const refreshOverall = () => {
        refreshBanner();
        const infos = getBankInfos();
        const hasTrans = transFiles.length > 0;
        const ready = bill !== null && hasTrans && infos.length > 0;
        exportBtn.disabled = !ready;

        if (ready && lastResult) {
            const s = lastResult.summary;
            const warn = s.manualReviewCount > 0 ? `　／　待覆核 ${s.manualReviewCount} 筆` : '';
            const transInfo = transFiles.length > 1 ? `（對帳單 ${transFiles.length} 份）` : '';
            overallStatus.textContent = `已就緒${transInfo}　客戶 ${s.customerCount} 間（已配對 ${s.matchedCount}）${warn}`;
        } else if (bill === null && !hasTrans) {
            overallStatus.textContent = '請上傳帳單與銀行對帳單';
        } else if (bill === null) {
            overallStatus.textContent = `尚缺帳單　已上傳對帳單 ${transFiles.length} 份`;
        } else if (!hasTrans) {
            overallStatus.textContent = `尚缺銀行對帳單　已上傳帳單：${billFileName ?? ''}`;
        } else if (infos.length === 0) {
            overallStatus.textContent = '尚缺末五碼對照表（請至設定頁建立）';
        }
    };

    const reset = () => {
        bill = null;
        billFileName = null;
        transFiles.length = 0;
        lastResult = null;
        previewTable.clear();
        billZone.reset();
        transZone.reset();
        renderTransList();
        refreshOverall();
    };

    resetBtn.addEventListener('click', reset);

    feeToleranceInput.value = String(feeTolerance);
    feeToleranceInput.addEventListener('change', () => {
        const raw = feeToleranceInput.value.trim();
        const n = raw === '' ? DEFAULT_FEE_TOLERANCE : Number(raw);
        const next = Number.isFinite(n) && n >= 0 ? Math.floor(n) : DEFAULT_FEE_TOLERANCE;
        feeTolerance = next;
        feeToleranceInput.value = String(next);
        localStorage.setItem(FEE_TOLERANCE_KEY, String(next));
        rebuildReconcile();
    });

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

/**
 * 將多份對帳單分別比對後合併為單一 BankMatchResult：
 *  - 每筆 row 標記 sourceFile（供 tooltip 顯示來源）
 *  - rowIndex 重新編號為合併後的順序索引
 *  - fileLineNumber 維持各自檔案內的真實行號
 *  - header 取第一份非空者，作為 popover 欄位名稱備援
 */
function matchAllTransFiles(
    files: ReadonlyArray<{ name: string; rows: string[][] }>,
    infos: ReadonlyArray<BankInfo>,
): BankMatchResult {
    const mergedRows: BankRowMatch[] = [];
    let header: string[] | null = null;
    let matched = 0;
    let unmatched = 0;

    for (const file of files) {
        const result = matchTransRecord(file.rows, infos);
        if (header === null && result.header) header = result.header;
        matched += result.matchedCount;
        unmatched += result.unmatchedCount;
        for (const row of result.rows) {
            mergedRows.push({
                ...row,
                rowIndex: mergedRows.length,
                sourceFile: file.name,
            });
        }
    }

    return {header, rows: mergedRows, matchedCount: matched, unmatchedCount: unmatched};
}
