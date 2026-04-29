/**
 * 未分類商品 → 加入分類的對話框。
 *
 * 使用情境：使用者在數據分析頁的「未分類商品」清單上點擊任一筆，
 * 顯示彈窗讓他選擇分類（既有 / 自訂新增）並可選擇填入商品編號。
 * 儲存後寫回 localStorage 的 daily_report_list，並 invalidate 兩個快取，
 * 讓呼叫方可即時 rebuildDataset。
 */

import Papa from 'papaparse';

import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import {localSettings} from '@/infra/local-settings-store';
import {notifyDailyReportChanged} from '@/domain/daily-report-loader';
import {invalidateCategoryMap} from '@/analytics/category-loader';

const DAILY_HEADER = ['分類', '編號', '品名', '成本'] as const;
const DAILY_ASSET_URL = `${import.meta.env.BASE_URL}assets/daily_report_list.csv`;
const NEW_GROUP_TOKEN = '__new__';

interface DailyRow {
    group: string;
    code: string;
    name: string;
    cost: string;
}

export interface UncategorizedDialogOptions {
    productName: string;
    onSaved?: () => void | Promise<void>;
}

export async function openUncategorizedDialog(options: UncategorizedDialogOptions): Promise<void> {
    const rows = await readDailyRows();
    const groups = uniqueGroupsInOrder(rows);

    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'uncat-dialog-title');

    dialog.innerHTML = `
      <header class="app-modal-header">
        <h2 id="uncat-dialog-title" class="app-modal-title">加入未分類商品至分類</h2>
        <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
      </header>
      <div class="app-modal-body">
        <div class="app-form-row">
          <label class="app-form-label">商品名稱</label>
          <div class="app-form-readonly" data-role="product-name"></div>
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="uncat-dialog-group">分類</label>
          <select id="uncat-dialog-group" class="app-form-input" data-role="group">
            <option value="">— 請選擇 —</option>
            ${groups.map((g) => `<option value="${escapeAttr(g)}">${escapeHtml(g)}</option>`).join('')}
            <option value="${NEW_GROUP_TOKEN}">＋ 新增分類…</option>
          </select>
        </div>

        <div class="app-form-row" data-role="new-group-row" hidden>
          <label class="app-form-label" for="uncat-dialog-new-group">新分類名稱</label>
          <input type="text" id="uncat-dialog-new-group" class="app-form-input" data-role="new-group" placeholder="例：水果類" autocomplete="off">
        </div>

        <div class="app-form-row">
          <label class="app-form-label" for="uncat-dialog-code">商品編號 <span class="app-form-hint">（選填）</span></label>
          <input type="text" id="uncat-dialog-code" class="app-form-input" data-role="code" placeholder="若不需要可留空" autocomplete="off">
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
    const groupSelect = dialog.querySelector<HTMLSelectElement>('[data-role="group"]')!;
    const newGroupRow = dialog.querySelector<HTMLElement>('[data-role="new-group-row"]')!;
    const newGroupInput = dialog.querySelector<HTMLInputElement>('[data-role="new-group"]')!;
    const codeInput = dialog.querySelector<HTMLInputElement>('[data-role="code"]')!;
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

    groupSelect.addEventListener('change', () => {
        const isNew = groupSelect.value === NEW_GROUP_TOKEN;
        newGroupRow.hidden = !isNew;
        if (isNew) newGroupInput.focus();
    });

    saveBtn.addEventListener('click', async () => {
        const selected = groupSelect.value;
        let groupName = '';
        if (selected === NEW_GROUP_TOKEN) {
            groupName = newGroupInput.value.trim();
            if (groupName.length === 0) {
                showToast({variant: 'warning', title: '請輸入分類', message: '新分類名稱不可空白'});
                newGroupInput.focus();
                return;
            }
        } else if (selected.length > 0) {
            groupName = selected;
        } else {
            showToast({variant: 'warning', title: '請選擇分類', message: '請從下拉選單選擇或新增分類'});
            groupSelect.focus();
            return;
        }

        const code = codeInput.value.trim();
        const productName = options.productName;

        try {
            saveBtn.disabled = true;
            await persistNewMapping(rows, {group: groupName, code, name: productName, cost: ''});
            showToast({variant: 'success', title: '已加入分類', message: `「${productName}」→ ${groupName}`});
            close();
            await options.onSaved?.();
        } catch (err) {
            console.error('[uncategorized-dialog] save failed', err);
            showToast({
                variant: 'error',
                title: '儲存失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            saveBtn.disabled = false;
        }
    });

    queueMicrotask(() => groupSelect.focus());
}

async function persistNewMapping(existing: DailyRow[], entry: DailyRow): Promise<void> {
    const next = appendIntoGroup(existing, entry);
    const csv = serializeDailyCsv(next);
    localSettings.setDailyReportList(csv);
    invalidateCategoryMap();
    // 同時清 daily-report 快取並通知訂閱者（例：設定頁品項分類分頁即時同步）
    notifyDailyReportChanged();
}

/**
 * 將新項目插入既有 group 的最後一筆之後；若 group 不存在則附加在整體尾端。
 */
function appendIntoGroup(rows: DailyRow[], entry: DailyRow): DailyRow[] {
    const next = rows.slice();
    let lastIdx = -1;
    for (let i = 0; i < next.length; i++) {
        if (next[i].group === entry.group) lastIdx = i;
    }
    if (lastIdx >= 0) {
        next.splice(lastIdx + 1, 0, entry);
    } else {
        next.push(entry);
    }
    return next;
}

async function readDailyRows(): Promise<DailyRow[]> {
    const overridden = localSettings.getDailyReportList();
    if (overridden !== null) return parseDailyCsv(overridden);

    const url = new URL(DAILY_ASSET_URL, document.baseURI).toString();
    const res = await fetch(url);
    if (!res.ok) throw new Error(`無法讀取 daily_report_list.csv（HTTP ${res.status}）`);
    return parseDailyCsv(await res.text());
}

function parseDailyCsv(text: string): DailyRow[] {
    const parsed = Papa.parse<string[]>(text, {header: false, skipEmptyLines: 'greedy'});
    const out: DailyRow[] = [];
    let currentGroup = '';
    for (let i = 1; i < parsed.data.length; i++) {
        const row = parsed.data[i];
        if (!row || row.length === 0) continue;
        const groupRaw = String(row[0] ?? '').trim();
        const code = String(row[1] ?? '').trim();
        const name = String(row[2] ?? '').trim();
        const cost = String(row[3] ?? '').trim();
        if (groupRaw.length > 0) currentGroup = groupRaw;
        out.push({group: currentGroup, code, name, cost});
    }
    return out;
}

function serializeDailyCsv(rows: DailyRow[]): string {
    const out: string[][] = [];
    let prevGroup = '';
    for (const r of rows) {
        const groupCol = r.group !== prevGroup ? r.group : '';
        out.push([groupCol, r.code, r.name, r.cost]);
        prevGroup = r.group;
    }
    return Papa.unparse([DAILY_HEADER.slice(), ...out], {newline: '\n'});
}

function uniqueGroupsInOrder(rows: DailyRow[]): string[] {
    const seen = new Set<string>();
    const out: string[] = [];
    for (const r of rows) {
        if (r.group.length > 0 && !seen.has(r.group)) {
            seen.add(r.group);
            out.push(r.group);
        }
    }
    return out;
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
