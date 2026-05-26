/**
 * 編輯既有末五碼設定的 modal。
 *
 * 用於對帳 2.0「需人工複核」表的配對候選 chip 點擊事件：
 *  - multi-match：使用者修正其中一筆 BankInfo（改末五碼避免衝突 / 改客戶名稱）
 *  - no-store-code：補上店家編號
 *
 * 與 bank-add-dialog 的差異：
 *  - 多了「店家編號」欄位（必填可空，由使用者控制）
 *  - 儲存時用 (customerName, customerLine, lastFiveDigit) 三元組定位原始 row 並 replace
 *  - 若編輯後內容與其他既有 row 完全相同 → 拒絕儲存，避免造成重複
 */

import {invalidateBankInfos, loadBankInfos, serializeBankInfoCsv} from '@/domain/bank-info-loader';
import type {BankInfo} from '@/domain/models/bank-info';
import {equalsBankInfo} from '@/domain/models/bank-info';
import {localSettings} from '@/infra/local-settings-store';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';

export interface BankEditDialogOptions {
    /** 待編輯的原始 BankInfo。由 (customerName, customerLine, lastFiveDigit) 定位至 localSettings 中對應的 row。 */
    targetInfo: BankInfo;
    /** 該筆銀行交易摘要（純顯示用），協助使用者辨識正在編輯哪筆設定的觸發來源。 */
    contextSummary?: string;
    /** 該筆銀行交易日期（純顯示用）。 */
    contextDate?: string;
    /** 儲存成功後的 callback。 */
    onSaved?: () => void | Promise<void>;
}

export function openBankEditDialog(options: BankEditDialogOptions): void {
    const existing = loadBankInfos();
    const uniqueCustomers = uniqueByCustomerName(existing);
    const uniqueLines = unique(existing.map((i) => i.customerLine).filter((s) => s.length > 0));

    const customerListId = 'bank-edit-dialog-customer-list';
    const lineListId = 'bank-edit-dialog-line-list';

    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'bank-edit-dialog-title');

    dialog.innerHTML = `
      <header class="app-modal-header">
        <h2 id="bank-edit-dialog-title" class="app-modal-title">編輯末五碼設定</h2>
        <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
      </header>
      <div class="app-modal-body">
        <div class="app-form-row">
          <label class="app-form-label">編輯來源</label>
          <div class="app-form-readonly" data-role="hint"></div>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-edit-dialog-customer">客戶名稱</label>
          <input type="text" id="bank-edit-dialog-customer" class="app-form-input" data-role="customer"
                 list="${customerListId}" autocomplete="off">
          <datalist id="${customerListId}">
            ${uniqueCustomers.map((info) => `<option value="${escapeAttr(info.customerName)}"></option>`).join('')}
          </datalist>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-edit-dialog-storecode">店家編號</label>
          <input type="text" id="bank-edit-dialog-storecode" class="app-form-input" data-role="storecode"
                 autocomplete="off" placeholder="對應帳單客戶編號，例：1001">
          <div class="app-form-hint">對應「明細客戶編號」；空白代表此設定不參與對帳客戶聚合。</div>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-edit-dialog-line">線別</label>
          <input type="text" id="bank-edit-dialog-line" class="app-form-input" data-role="line"
                 list="${lineListId}" autocomplete="off">
          <datalist id="${lineListId}">
            ${uniqueLines.map((line) => `<option value="${escapeAttr(line)}"></option>`).join('')}
          </datalist>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-edit-dialog-digit">末五碼 / 比對字串</label>
          <input type="text" id="bank-edit-dialog-digit" class="app-form-input" data-role="digit"
                 autocomplete="off">
          <div class="app-form-hint">純數字會反向比對隱碼（含 *）；中文則前綴比對。</div>
        </div>
      </div>
      <footer class="app-modal-footer bank-edit-footer">
        <button type="button" class="btn btn-danger-outline" data-role="delete">刪除這筆</button>
        <div class="bank-edit-footer-actions">
          <button type="button" class="btn btn-secondary" data-role="cancel">取消</button>
          <button type="button" class="btn btn-primary" data-role="save">儲存並重比對</button>
        </div>
      </footer>
    `;

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    const hintEl = dialog.querySelector<HTMLElement>('[data-role="hint"]')!;
    const customerInput = dialog.querySelector<HTMLInputElement>('[data-role="customer"]')!;
    const storeCodeInput = dialog.querySelector<HTMLInputElement>('[data-role="storecode"]')!;
    const lineInput = dialog.querySelector<HTMLInputElement>('[data-role="line"]')!;
    const digitInput = dialog.querySelector<HTMLInputElement>('[data-role="digit"]')!;
    const closeBtn = dialog.querySelector<HTMLButtonElement>('[data-role="close"]')!;
    const cancelBtn = dialog.querySelector<HTMLButtonElement>('[data-role="cancel"]')!;
    const saveBtn = dialog.querySelector<HTMLButtonElement>('[data-role="save"]')!;
    const deleteBtn = dialog.querySelector<HTMLButtonElement>('[data-role="delete"]')!;

    hintEl.textContent = formatHint(options);
    customerInput.value = options.targetInfo.customerName;
    storeCodeInput.value = options.targetInfo.storeCode;
    lineInput.value = options.targetInfo.customerLine;
    digitInput.value = options.targetInfo.lastFiveDigit;

    const previouslyFocused = document.activeElement as HTMLElement | null;

    const close = () => {
        backdrop.remove();
        document.removeEventListener('keydown', onKey);
        previouslyFocused?.focus?.();
    };

    const onKey = (e: KeyboardEvent) => {
        if (e.key === 'Escape') {
            e.preventDefault();
            close();
        }
    };
    document.addEventListener('keydown', onKey);

    backdrop.addEventListener('click', (e) => {
        if (e.target === backdrop) close();
    });
    closeBtn.addEventListener('click', close);
    cancelBtn.addEventListener('click', close);

    deleteBtn.addEventListener('click', async () => {
        const target = options.targetInfo;
        const label = `${target.customerName}${target.customerLine ? ` / ${target.customerLine}` : ''}　#${target.lastFiveDigit}`;
        const ok = window.confirm(`確定要刪除這筆末五碼設定嗎？\n\n${label}\n\n此操作無法復原。`);
        if (!ok) return;

        try {
            deleteBtn.disabled = true;
            saveBtn.disabled = true;
            const result = persistDelete(target);
            if (result === 'not-found') {
                showToast({
                    variant: 'error',
                    title: '找不到原始紀錄',
                    message: '末五碼對照表可能已被其他流程改動，請重新開啟對話框',
                });
                close();
                return;
            }
            showToast({
                variant: 'success',
                title: '已刪除末五碼',
                message: label,
            });
            close();
            await options.onSaved?.();
        } catch (err) {
            console.error('[bank-edit-dialog] delete failed', err);
            showToast({
                variant: 'error',
                title: '刪除失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            deleteBtn.disabled = false;
            saveBtn.disabled = false;
        }
    });

    saveBtn.addEventListener('click', async () => {
        const customerName = customerInput.value.trim();
        const storeCode = storeCodeInput.value.trim();
        const customerLine = lineInput.value.trim();
        const lastFiveDigit = digitInput.value.trim();

        if (customerName.length === 0) {
            showToast({variant: 'warning', title: '請輸入客戶名稱', message: '客戶名稱不可空白'});
            customerInput.focus();
            return;
        }
        if (lastFiveDigit.length === 0) {
            showToast({variant: 'warning', title: '請輸入末五碼', message: '末五碼不可空白'});
            digitInput.focus();
            return;
        }

        const next: BankInfo = {customerName, storeCode, customerLine, lastFiveDigit};

        try {
            saveBtn.disabled = true;
            const result = persistReplace(options.targetInfo, next);
            if (result === 'duplicate') {
                showToast({
                    variant: 'warning',
                    title: '已存在相同設定',
                    message: '此 (客戶 / 線別 / 末五碼) 已存在於其他列，請改其他欄位後再儲存',
                });
                saveBtn.disabled = false;
                return;
            }
            if (result === 'not-found') {
                showToast({
                    variant: 'error',
                    title: '找不到原始紀錄',
                    message: '末五碼對照表可能已被其他流程改動，請重新開啟對話框',
                });
                close();
                return;
            }
            showToast({
                variant: 'success',
                title: '已更新末五碼',
                message: `${customerName}${customerLine ? ` / ${customerLine}` : ''}${storeCode ? ` #${storeCode}` : ''}`,
            });
            close();
            await options.onSaved?.();
        } catch (err) {
            console.error('[bank-edit-dialog] save failed', err);
            showToast({
                variant: 'error',
                title: '儲存失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            saveBtn.disabled = false;
        }
    });

    queueMicrotask(() => storeCodeInput.focus());
}

type PersistResult = 'ok' | 'duplicate' | 'not-found';
type DeleteResult = 'ok' | 'not-found';

function persistDelete(target: BankInfo): DeleteResult {
    const current = loadBankInfos();
    const idx = current.findIndex((info) => equalsBankInfo(info, target));
    if (idx < 0) return 'not-found';
    const next = current.slice();
    next.splice(idx, 1);
    const csv = serializeBankInfoCsv(next);
    localSettings.setLastFiveDigit(csv);
    invalidateBankInfos();
    return 'ok';
}

function persistReplace(target: BankInfo, next: BankInfo): PersistResult {
    const current = loadBankInfos();
    const targetIdx = current.findIndex((info) => equalsBankInfo(info, target));
    if (targetIdx < 0) return 'not-found';

    // 若 next 已存在於其他列（不算 target 自己） → 拒絕，避免重複
    const conflictIdx = current.findIndex((info, idx) => idx !== targetIdx && equalsBankInfo(info, next));
    if (conflictIdx >= 0) return 'duplicate';

    const replaced = current.slice();
    replaced[targetIdx] = next;
    const csv = serializeBankInfoCsv(replaced);
    localSettings.setLastFiveDigit(csv);
    invalidateBankInfos();
    return 'ok';
}

function uniqueByCustomerName(infos: ReadonlyArray<BankInfo>): BankInfo[] {
    const seen = new Set<string>();
    const out: BankInfo[] = [];
    for (const info of infos) {
        if (info.customerName.length === 0) continue;
        if (seen.has(info.customerName)) continue;
        seen.add(info.customerName);
        out.push(info);
    }
    return out;
}

function unique(values: ReadonlyArray<string>): string[] {
    const seen = new Set<string>();
    const out: string[] = [];
    for (const v of values) {
        if (seen.has(v)) continue;
        seen.add(v);
        out.push(v);
    }
    return out;
}

function formatHint(options: BankEditDialogOptions): string {
    const segs: string[] = [];
    if (options.contextDate) segs.push(options.contextDate);
    if (options.contextSummary) segs.push(options.contextSummary);
    segs.push(`原 #${options.targetInfo.lastFiveDigit}`);
    return segs.join('　·　');
}

function escapeAttr(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
