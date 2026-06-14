/**
 * 未填成本商品 → 填入成本的對話框。
 *
 * 使用情境：使用者在數據分析頁的「未填成本商品」清單上點擊任一筆，
 * 顯示彈窗讓他輸入成本（單品成本，正數，允許小數）。
 * 成本已搬移到 cargo_sort（帳單排序），故儲存後寫回 localStorage 的 cargo_sort
 * （同名商品所有列同步更新；若清單裡沒有這個商品則新增一列），
 * invalidate sorting list 與 cost map 快取，讓呼叫方可即時 rebuildDataset。
 */

import Papa from 'papaparse';

import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import {localSettings} from '@/infra/local-settings-store';
import {invalidateSortingList} from '@/domain/sorting-list';
import {invalidateCostMap} from '@/analytics/cost-loader';

const CARGO_HEADER = ['貨品編號', '貨品名稱', '代送費', '成本'] as const;
const CARGO_ASSET_URL = `${import.meta.env.BASE_URL}assets/cargo_sort.csv`;

interface CargoRow {
    id: string;
    name: string;
    fee: string;
    cost: string;
}

export interface UnsetCostDialogOptions {
    productName: string;
    onSaved?: () => void | Promise<void>;
}

export async function openUnsetCostDialog(options: UnsetCostDialogOptions): Promise<void> {
    const rows = await readCargoRows();

    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'unset-cost-dialog-title');

    dialog.innerHTML = `
      <header class="app-modal-header">
        <h2 id="unset-cost-dialog-title" class="app-modal-title">填入商品成本</h2>
        <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
      </header>
      <div class="app-modal-body">
        <div class="app-form-row">
          <label class="app-form-label">商品名稱</label>
          <div class="app-form-readonly" data-role="product-name"></div>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="unset-cost-dialog-input">成本</label>
          <input type="number" id="unset-cost-dialog-input" class="app-form-input" data-role="cost"
                 placeholder="請輸入單品成本" min="0" step="0.01" autocomplete="off">
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
    const costInput = dialog.querySelector<HTMLInputElement>('[data-role="cost"]')!;
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
        const raw = costInput.value.trim();
        if (raw.length === 0) {
            showToast({variant: 'warning', title: '請輸入成本', message: '成本欄位不可空白'});
            costInput.focus();
            return;
        }
        const num = Number(raw);
        if (!Number.isFinite(num) || num < 0) {
            showToast({variant: 'warning', title: '成本格式錯誤', message: '請輸入大於等於 0 的數字'});
            costInput.focus();
            return;
        }

        try {
            saveBtn.disabled = true;
            await persistCost(rows, options.productName, raw);
            showToast({
                variant: 'success',
                title: '已更新成本',
                message: `「${options.productName}」成本 = ${raw}`,
            });
            close();
            await options.onSaved?.();
        } catch (err) {
            console.error('[unset-cost-dialog] save failed', err);
            showToast({
                variant: 'error',
                title: '儲存失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            saveBtn.disabled = false;
        }
    };

    saveBtn.addEventListener('click', () => void submit());
    costInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            void submit();
        }
    });

    queueMicrotask(() => costInput.focus());
}

async function persistCost(existing: CargoRow[], productName: string, costRaw: string): Promise<void> {
    let matched = false;
    const next = existing.map((r) => {
        if (r.name === productName) {
            matched = true;
            return {...r, cost: costRaw};
        }
        return r;
    });
    // 帳單排序清單裡沒有這個商品 → 新增一列（編號／代送費留空，使用者可再到設定補齊）
    if (!matched) {
        next.push({id: '', name: productName, fee: '', cost: costRaw});
    }
    const csv = serializeCargoCsv(next);
    localSettings.setCargoSort(csv);
    invalidateSortingList();
    invalidateCostMap();
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
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: true});
    const out: CargoRow[] = [];
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        out.push({
            id: String(row[0] ?? '').trim(),
            name: String(row[1] ?? '').trim(),
            fee: String(row[2] ?? '').trim(),
            cost: String(row[3] ?? '').trim(),
        });
    }
    return out;
}

function serializeCargoCsv(rows: CargoRow[]): string {
    const data = [CARGO_HEADER.slice(), ...rows.map((r) => [r.id, r.name, r.fee, r.cost])];
    return Papa.unparse(data, {newline: '\n'});
}
