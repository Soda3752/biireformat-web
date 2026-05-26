/**
 * 銀行對帳分頁的「補建末五碼」對話框。
 *
 * 觸發情境：對帳預覽表中某列未配對 → 點「補建」→ 開啟此 modal。
 * 欄位：
 *   - 客戶名稱：datalist 從現有 bankInfos 唯一客戶建議，可自由輸入
 *   - 線別：datalist 從現有 bankInfos 唯一線別建議；選既有客戶名稱時自動帶入但仍可改
 *   - 末五碼：預填觸發列的隱碼字串，可改
 * 儲存時：append 至 localSettings.lastFiveDigit、invalidateBankInfos、回呼 onSaved。
 */

import {invalidateBankInfos, loadBankInfos, serializeBankInfoCsv} from '@/domain/bank-info-loader';
import type {BankInfo} from '@/domain/models/bank-info';
import {equalsBankInfo} from '@/domain/models/bank-info';
import {localSettings} from '@/infra/local-settings-store';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';

export interface BankAddDialogOptions {
    /** 觸發補建的隱碼字串（CSV 該列尾欄），會預填到末五碼欄位。 */
    presetLastFiveDigit: string;
    /** 該列摘要（電匯/跨行轉/轉帳存…），純顯示輔助。 */
    presetSummary?: string;
    /** 該列日期，純顯示輔助。 */
    presetDate?: string;
    /** 儲存成功後的 callback，呼叫端可重新載入 bankInfos 並重比對。 */
    onSaved?: () => void | Promise<void>;
}

export function openBankAddDialog(options: BankAddDialogOptions): void {
    const existing = loadBankInfos();
    const uniqueCustomers = uniqueByCustomerName(existing);
    const uniqueLines = unique(existing.map((i) => i.customerLine).filter((s) => s.length > 0));

    const customerListId = 'bank-add-dialog-customer-list';
    const lineListId = 'bank-add-dialog-line-list';

    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'bank-add-dialog-title');

    dialog.innerHTML = `
      <header class="app-modal-header">
        <h2 id="bank-add-dialog-title" class="app-modal-title">補建末五碼</h2>
        <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
      </header>
      <div class="app-modal-body">
        <div class="app-form-row">
          <label class="app-form-label">該筆交易</label>
          <div class="app-form-readonly" data-role="hint"></div>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-add-dialog-customer">客戶名稱</label>
          <input type="text" id="bank-add-dialog-customer" class="app-form-input" data-role="customer"
                 list="${customerListId}" autocomplete="off" placeholder="從建議清單選擇或自行輸入">
          <datalist id="${customerListId}">
            ${uniqueCustomers.map((info) => `<option value="${escapeAttr(info.customerName)}"></option>`).join('')}
          </datalist>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-add-dialog-line">線別</label>
          <input type="text" id="bank-add-dialog-line" class="app-form-input" data-role="line"
                 list="${lineListId}" autocomplete="off" placeholder="例：A 線 / 配送一">
          <datalist id="${lineListId}">
            ${uniqueLines.map((line) => `<option value="${escapeAttr(line)}"></option>`).join('')}
          </datalist>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="bank-add-dialog-digit">末五碼 / 比對字串</label>
          <input type="text" id="bank-add-dialog-digit" class="app-form-input" data-role="digit"
                 autocomplete="off" placeholder="預填當前列的最後一欄；可手動修改">
          <div class="app-form-hint">純數字會反向比對隱碼（含 *）；中文則前綴比對。</div>
        </div>
      </div>
      <footer class="app-modal-footer">
        <button type="button" class="btn btn-secondary" data-role="cancel">取消</button>
        <button type="button" class="btn btn-primary" data-role="save">儲存並重比對</button>
      </footer>
    `;

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    const hintEl = dialog.querySelector<HTMLElement>('[data-role="hint"]')!;
    const customerInput = dialog.querySelector<HTMLInputElement>('[data-role="customer"]')!;
    const lineInput = dialog.querySelector<HTMLInputElement>('[data-role="line"]')!;
    const digitInput = dialog.querySelector<HTMLInputElement>('[data-role="digit"]')!;
    const closeBtn = dialog.querySelector<HTMLButtonElement>('[data-role="close"]')!;
    const cancelBtn = dialog.querySelector<HTMLButtonElement>('[data-role="cancel"]')!;
    const saveBtn = dialog.querySelector<HTMLButtonElement>('[data-role="save"]')!;

    hintEl.textContent = formatHint(options);
    digitInput.value = options.presetLastFiveDigit ?? '';

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

    // 客戶名稱與既有客戶相符 → 自動帶入線別（若使用者尚未自行填過）
    customerInput.addEventListener('change', () => {
        const name = customerInput.value.trim();
        if (name.length === 0) return;
        if (lineInput.dataset.userEdited === '1') return;
        const matched = uniqueCustomers.find((c) => c.customerName === name);
        if (matched) lineInput.value = matched.customerLine;
    });
    lineInput.addEventListener('input', () => {
        lineInput.dataset.userEdited = '1';
    });

    saveBtn.addEventListener('click', async () => {
        const customerName = customerInput.value.trim();
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

        const newInfo: BankInfo = {customerName, storeCode: '', customerLine, lastFiveDigit};

        try {
            saveBtn.disabled = true;
            persistAppend(newInfo);
            showToast({
                variant: 'success',
                title: '已新增末五碼',
                message: `${customerName}${customerLine ? ` / ${customerLine}` : ''}`,
            });
            close();
            await options.onSaved?.();
        } catch (err) {
            console.error('[bank-add-dialog] save failed', err);
            showToast({
                variant: 'error',
                title: '儲存失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            saveBtn.disabled = false;
        }
    });

    queueMicrotask(() => customerInput.focus());
}

function persistAppend(newInfo: BankInfo): void {
    const current = loadBankInfos();
    const exists = current.some((info) => equalsBankInfo(info, newInfo));
    const next = exists ? current.slice() : [...current, newInfo];
    const csv = serializeBankInfoCsv(next);
    localSettings.setLastFiveDigit(csv);
    invalidateBankInfos();
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

function formatHint(options: BankAddDialogOptions): string {
    const segs: string[] = [];
    if (options.presetDate) segs.push(options.presetDate);
    if (options.presetSummary) segs.push(options.presetSummary);
    segs.push(options.presetLastFiveDigit || '(空)');
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
