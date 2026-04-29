/**
 * 未填代送費商品 → 填入代送費的對話框。
 *
 * 使用情境：使用者上傳帳單到代送費分頁後，系統掃描出在 cargo_sort.csv 找不到、
 * 或代送費欄位為空白的商品；使用者點擊任一筆即可填入代送費單價。
 *
 * 儲存後寫回 localStorage 的 cargo_sort.csv：
 *  - 若商品已存在於 csv：直接更新該列的代送費欄
 *  - 若商品不存在：append 到末尾（使用者可後續至設定頁微調排序）
 *
 * invalidateSortingList + loadSortingList 重建快取，呼叫方可即時 rescan。
 */

import Papa from 'papaparse';

import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import {localSettings} from '@/infra/local-settings-store';
import {invalidateSortingList, loadSortingList} from '@/domain/sorting-list';

const CARGO_HEADER = ['貨品編號', '貨品名稱', '代送費'] as const;
const CARGO_ASSET_URL = `${import.meta.env.BASE_URL}assets/cargo_sort.csv`;

interface CargoRow {
    id: string;
    name: string;
    fee: string;
}

export interface UnsetDeliveryFeeDialogOptions {
    productName: string;
    onSaved?: () => void | Promise<void>;
}

export async function openUnsetDeliveryFeeDialog(
    options: UnsetDeliveryFeeDialogOptions
): Promise<void> {
    const rows = await readCargoRows();

    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'unset-fee-dialog-title');

    dialog.innerHTML = `
      <header class="app-modal-header">
        <h2 id="unset-fee-dialog-title" class="app-modal-title">填入商品代送費</h2>
        <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
      </header>
      <div class="app-modal-body">
        <div class="app-form-row">
          <label class="app-form-label">商品名稱</label>
          <div class="app-form-readonly" data-role="product-name"></div>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="unset-fee-dialog-input">代送費</label>
          <input type="number" id="unset-fee-dialog-input" class="app-form-input" data-role="fee"
                 placeholder="請輸入代送費單價" min="0" step="0.01" autocomplete="off">
        </div>
      </div>
      <footer class="app-modal-footer">
        <button type="button" class="btn btn-secondary" data-role="cancel">取消</button>
        <button type="button" class="btn btn-primary" data-role="save">儲存</button>
      </footer>
    `;

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    const productNameEl = dialog.querySelector<HTMLElement>('[data-role="product-name"]')!;
    const feeInput = dialog.querySelector<HTMLInputElement>('[data-role="fee"]')!;
    const closeBtn = dialog.querySelector<HTMLButtonElement>('[data-role="close"]')!;
    const cancelBtn = dialog.querySelector<HTMLButtonElement>('[data-role="cancel"]')!;
    const saveBtn = dialog.querySelector<HTMLButtonElement>('[data-role="save"]')!;

    productNameEl.textContent = options.productName;

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

    const submit = async () => {
        const raw = feeInput.value.trim();
        if (raw.length === 0) {
            showToast({variant: 'warning', title: '請輸入代送費', message: '代送費欄位不可空白'});
            feeInput.focus();
            return;
        }
        const num = Number(raw);
        if (!Number.isFinite(num) || num < 0) {
            showToast({variant: 'warning', title: '代送費格式錯誤', message: '請輸入大於等於 0 的數字'});
            feeInput.focus();
            return;
        }

        try {
            saveBtn.disabled = true;
            await persistFee(rows, options.productName, raw);
            showToast({
                variant: 'success',
                title: '已更新代送費',
                message: `「${options.productName}」代送費 = ${raw}`,
            });
            close();
            await options.onSaved?.();
        } catch (err) {
            console.error('[unset-delivery-fee-dialog] save failed', err);
            showToast({
                variant: 'error',
                title: '儲存失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            saveBtn.disabled = false;
        }
    };

    saveBtn.addEventListener('click', () => void submit());
    feeInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            void submit();
        }
    });

    queueMicrotask(() => feeInput.focus());
}

async function persistFee(existing: CargoRow[], productName: string, feeRaw: string): Promise<void> {
    const idx = existing.findIndex((r) => r.name === productName);
    let next: CargoRow[];
    if (idx >= 0) {
        next = existing.map((r, i) => (i === idx ? {...r, fee: feeRaw} : r));
    } else {
        next = [...existing, {id: '', name: productName, fee: feeRaw}];
    }
    const csv = serializeCargoCsv(next);
    localSettings.setCargoSort(csv);
    invalidateSortingList();
    await loadSortingList();
}

async function readCargoRows(): Promise<CargoRow[]> {
    const overridden = localSettings.getCargoSort();
    if (overridden !== null) return parseCargoCsv(overridden);

    const url = new URL(CARGO_ASSET_URL, document.baseURI).toString();
    const res = await fetch(url);
    if (!res.ok) throw new Error(`無法讀取 cargo_sort.csv（HTTP ${res.status}）`);
    return parseCargoCsv(await res.text());
}

function parseCargoCsv(text: string): CargoRow[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: 'greedy'});
    const out: CargoRow[] = [];
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        const id = String(row[0] ?? '').trim();
        const name = String(row[1] ?? '').trim();
        const fee = String(row[2] ?? '').trim();
        if (name.length === 0) continue;
        out.push({id, name, fee});
    }
    return out;
}

function serializeCargoCsv(rows: CargoRow[]): string {
    const body = rows.map((r) => [r.id, r.name, r.fee]);
    return Papa.unparse([CARGO_HEADER.slice(), ...body], {newline: '\n'});
}
