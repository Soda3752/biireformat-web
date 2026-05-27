/**
 * 「生成手填本」主面板。
 *
 * 流程：新建 / 匯入 .xlsx / 從歷史紀錄載入 → 客戶卡片編輯 → 輸出 .xlsx
 *
 * 自動儲存：每次輸入後 debounce 500ms 寫入 localStorage。
 */

import {saveAs} from 'file-saver';

import {icon} from '@/ui/icons';
import type {TabDefinition} from '@/ui/tabs';
import {showToast} from '@/ui/toast';
import {renderHandfillCustomerCard} from '@/ui/handfill-customer-card';
import {openHandfillHistoryDialog} from '@/ui/handfill-history-dialog';
import {openHandfillJsonImportDialog} from '@/ui/handfill-json-import-dialog';

import {
    createEmptyBook,
    createEmptyCustomer,
    type HandfillBook,
    isCustomerEmpty,
    lineLabel,
    sortProductsBlanksFirst,
} from '@/domain/models/handfill-book';
import {deleteBook as deleteBookFromStore, getActiveId, loadBook, saveBook, setActiveId,} from '@/infra/handfill-store';
import {readHandfillBook} from '@/readers/handfill-reader';
import {buildHandfillWorkbook} from '@/writers/handfill-writer';
import {getSortingList, loadSortingList} from '@/domain/sorting-list';

const SAVE_DEBOUNCE_MS = 500;
type ViewMode = 'card' | 'list';

interface PanelState {
    book: HandfillBook;
    cursor: number;          // 當前卡片索引
    viewMode: ViewMode;
    saveTimer: number | null;
    saveStatusEl: HTMLElement | null;
    draggingIndex: number | null;
    cargoNames: ReadonlyArray<string>;
}

function readCargoNamesFromCache(): ReadonlyArray<string> {
    try {
        return getSortingList().cargoItems.map((it) => it.name);
    } catch {
        return [];
    }
}

export function renderHandfillPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel handfill-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    // 還原上次編輯或建立新本
    const activeId = getActiveId();
    const initialBook = (activeId && loadBook(activeId)) || createEmptyBook();
    if (initialBook.customers.length === 0) {
        initialBook.customers.push(createEmptyCustomer());
    }

    const state: PanelState = {
        book: initialBook,
        cursor: 0,
        viewMode: 'list',
        saveTimer: null,
        saveStatusEl: null,
        draggingIndex: null,
        cargoNames: readCargoNamesFromCache(),
    };

    setActiveId(state.book.id);

    panel.innerHTML = `
    <div class="card handfill-card">
      <header class="card-header handfill-header">
        <div>
          <h1 class="card-title">生成手填本</h1>
          <p class="card-subtitle">
            為每月手填本產生空白範本 .xlsx，支援匯入既有檔案編輯、自動儲存與歷史紀錄。
          </p>
        </div>
      </header>

      <section class="handfill-toolbar">
        <div class="handfill-toolbar-row">
          <div class="app-form-row handfill-field-line-no">
            <label class="app-form-label" for="hf-line-no">線別編號</label>
            <input id="hf-line-no" class="app-form-input" type="number" min="1" max="99" step="1">
          </div>
          <div class="handfill-line-label" data-role="line-label">(一)</div>
          <div class="app-form-row handfill-field-line-name">
            <label class="app-form-label" for="hf-line-name">線別名稱</label>
            <input id="hf-line-name" class="app-form-input" type="text" placeholder="例：彰化">
          </div>
          <div class="app-form-row handfill-field-year">
            <label class="app-form-label" for="hf-year">民國年</label>
            <input id="hf-year" class="app-form-input" type="number" min="100" max="200" step="1">
          </div>
          <div class="app-form-row handfill-field-month">
            <label class="app-form-label" for="hf-month">月份</label>
            <select id="hf-month" class="app-form-input">
              ${Array.from({length: 12}, (_, i) => `<option value="${i + 1}">${i + 1}</option>`).join('')}
            </select>
          </div>
        </div>
        <div class="handfill-toolbar-row">
          <div class="handfill-toolbar-actions">
            <button type="button" class="btn btn-secondary btn-sm" data-role="new-book">
              ${icon('plus', 14)}<span>新建</span>
            </button>
            <button type="button" class="btn btn-secondary btn-sm" data-role="import-xlsx">
              ${icon('upload', 14)}<span>匯入 .xlsx / .xls</span>
            </button>
            <input type="file" data-role="file-input" accept=".xlsx,.xls" hidden>
            <button type="button" class="btn btn-secondary btn-sm" data-role="import-json">
              ${icon('clipboard-list', 14)}<span>從 JSON 匯入</span>
            </button>
            <button type="button" class="btn btn-secondary btn-sm" data-role="open-history">
              ${icon('list', 14)}<span>歷史紀錄</span>
            </button>
          </div>
          <div class="handfill-toolbar-status">
            <span class="handfill-save-status" data-role="save-status">已儲存</span>
            <button type="button" class="btn btn-primary btn-sm" data-role="export-xlsx">
              ${icon('download', 14)}<span>匯出 .xlsx</span>
            </button>
          </div>
        </div>
      </section>

      <section class="handfill-viewer">
        <div class="handfill-cust-toolbar" data-view-mode="list">
          <div class="handfill-cust-toolbar-left">
            <div class="handfill-cust-pager" data-mode="card-only">
              <button type="button" class="btn btn-secondary btn-sm" data-role="back-to-list">
                ${icon('chevron-left', 14)}<span>返回列表</span>
              </button>
              <span class="handfill-cust-toolbar-divider" aria-hidden="true"></span>
              <button type="button" class="btn btn-secondary btn-icon" data-role="prev-cust" aria-label="上一個客戶">${icon('chevron-up', 14)}</button>
              <span class="handfill-cust-counter" data-role="cust-counter">1 / 1</span>
              <button type="button" class="btn btn-secondary btn-icon" data-role="next-cust" aria-label="下一個客戶">${icon('chevron-down', 14)}</button>
            </div>
            <div class="handfill-cust-list-hint" data-mode="list-only" data-role="list-hint">
              共 0 個客戶 · 拖曳左側 <span class="handfill-cust-list-hint-grip">${icon('grip', 12)}</span> 可重新排序
            </div>
          </div>
          <div class="handfill-cust-actions">
            <button type="button" class="btn btn-secondary btn-sm" data-role="sort-all" data-mode="list-only" title="把每一家客戶的品名清單都排序（空白置頂，其餘依帳單排序表）">
              ${icon('chevron-down', 14)}<span>全部排序</span>
            </button>
            <button type="button" class="btn btn-secondary btn-sm" data-role="add-cust">
              ${icon('plus', 14)}<span>新增客戶</span>
            </button>
            <button type="button" class="btn btn-danger btn-sm" data-role="delete-cust" data-mode="card-only">
              ${icon('trash', 14)}<span>刪除此客戶</span>
            </button>
          </div>
        </div>
        <div class="handfill-cust-body" data-role="cust-body"></div>
      </section>
    </div>
  `;

    // ====== DOM 參照 ======
    const lineNoInput = panel.querySelector<HTMLInputElement>('#hf-line-no')!;
    const lineLabelEl = panel.querySelector<HTMLElement>('[data-role="line-label"]')!;
    const lineNameInput = panel.querySelector<HTMLInputElement>('#hf-line-name')!;
    const yearInput = panel.querySelector<HTMLInputElement>('#hf-year')!;
    const monthSelect = panel.querySelector<HTMLSelectElement>('#hf-month')!;
    const newBtn = panel.querySelector<HTMLButtonElement>('[data-role="new-book"]')!;
    const importBtn = panel.querySelector<HTMLButtonElement>('[data-role="import-xlsx"]')!;
    const fileInput = panel.querySelector<HTMLInputElement>('[data-role="file-input"]')!;
    const importJsonBtn = panel.querySelector<HTMLButtonElement>('[data-role="import-json"]')!;
    const historyBtn = panel.querySelector<HTMLButtonElement>('[data-role="open-history"]')!;
    const exportBtn = panel.querySelector<HTMLButtonElement>('[data-role="export-xlsx"]')!;
    state.saveStatusEl = panel.querySelector<HTMLElement>('[data-role="save-status"]');
    const prevBtn = panel.querySelector<HTMLButtonElement>('[data-role="prev-cust"]')!;
    const nextBtn = panel.querySelector<HTMLButtonElement>('[data-role="next-cust"]')!;
    const counterEl = panel.querySelector<HTMLElement>('[data-role="cust-counter"]')!;
    const backToListBtn = panel.querySelector<HTMLButtonElement>('[data-role="back-to-list"]')!;
    const listHintEl = panel.querySelector<HTMLElement>('[data-role="list-hint"]')!;
    const custToolbar = panel.querySelector<HTMLElement>('.handfill-cust-toolbar')!;
    const sortAllBtn = panel.querySelector<HTMLButtonElement>('[data-role="sort-all"]')!;
    const addCustBtn = panel.querySelector<HTMLButtonElement>('[data-role="add-cust"]')!;
    const deleteCustBtn = panel.querySelector<HTMLButtonElement>('[data-role="delete-cust"]')!;
    const custBody = panel.querySelector<HTMLElement>('[data-role="cust-body"]')!;

    // ====== 同步狀態 → UI ======
    function syncToolbar(): void {
        lineNoInput.value = String(state.book.lineNo);
        lineLabelEl.textContent = lineLabel(state.book.lineNo);
        lineNameInput.value = state.book.lineName;
        yearInput.value = String(state.book.year);
        monthSelect.value = String(state.book.month);
    }

    function setSaveStatus(text: string, variant: 'saving' | 'saved' | 'error' = 'saved'): void {
        if (!state.saveStatusEl) return;
        state.saveStatusEl.textContent = text;
        state.saveStatusEl.dataset.variant = variant;
    }

    function scheduleSave(): void {
        if (state.saveTimer !== null) {
            window.clearTimeout(state.saveTimer);
        }
        setSaveStatus('儲存中…', 'saving');
        state.saveTimer = window.setTimeout(() => {
            try {
                saveBook(state.book);
                setActiveId(state.book.id);
                setSaveStatus('已儲存', 'saved');
            } catch (err) {
                setSaveStatus('儲存失敗', 'error');
                showToast({
                    variant: 'error',
                    title: '自動儲存失敗',
                    message: err instanceof Error ? err.message : String(err),
                });
            } finally {
                state.saveTimer = null;
            }
        }, SAVE_DEBOUNCE_MS);
    }

    function applyViewMode(): void {
        custToolbar.dataset.viewMode = state.viewMode;
        const total = state.book.customers.length;
        listHintEl.innerHTML = `共 ${total} 個客戶 · 拖曳左側 <span class="handfill-cust-list-hint-grip">${icon('grip', 12)}</span> 可重新排序`;
    }

    function renderCustBody(): void {
        applyViewMode();
        custBody.innerHTML = '';
        const total = state.book.customers.length;
        if (total === 0) {
            custBody.innerHTML = `<div class="handfill-empty">尚無客戶，請點擊「新增客戶」開始。</div>`;
            counterEl.textContent = '0 / 0';
            return;
        }
        state.cursor = Math.max(0, Math.min(state.cursor, total - 1));

        if (state.viewMode === 'list') {
            renderListView();
        } else {
            renderCardView();
        }
        counterEl.textContent = `${state.cursor + 1} / ${total}`;
    }

    function renderCardView(): void {
        const customer = state.book.customers[state.cursor];
        const cardEl = renderHandfillCustomerCard({
            customer,
            cargoNames: state.cargoNames,
            onChange: (updated) => {
                state.book.customers[state.cursor] = updated;
                scheduleSave();
            },
        });
        custBody.appendChild(cardEl);
    }

    function renderListView(): void {
        const list = document.createElement('div');
        list.className = 'handfill-cust-list';
        state.book.customers.forEach((c, i) => {
            const item = document.createElement('div');
            item.className = 'handfill-cust-list-item';
            item.setAttribute('role', 'button');
            item.tabIndex = 0;
            item.draggable = true;
            item.dataset.index = String(i);
            if (i === state.cursor) item.classList.add('is-active');
            const productSummary = c.products
                .map((p) => p.name)
                .filter((s) => s.trim())
                .slice(0, 4)
                .join('、');
            item.innerHTML = `
        <span class="handfill-cust-list-handle" aria-hidden="true" title="拖曳以排序">${icon('grip', 14)}</span>
        <span class="handfill-cust-list-seq">${i + 1}</span>
        <span class="handfill-cust-list-id">${escapeHtml(c.customerId) || '─'}</span>
        <span class="handfill-cust-list-name">${escapeHtml(c.customerName) || '(未命名)'}</span>
        <span class="handfill-cust-list-products">${escapeHtml(productSummary)}</span>
        <span class="handfill-cust-list-meta">${c.products.filter((p) => p.name.trim()).length} 品</span>
      `;
            // 點擊（非拖曳） → 進入卡片編輯
            item.addEventListener('click', () => {
                state.cursor = i;
                state.viewMode = 'card';
                renderCustBody();
            });
            item.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    item.click();
                }
            });
            // 拖曳事件
            item.addEventListener('dragstart', (e) => {
                state.draggingIndex = i;
                item.classList.add('is-dragging');
                if (e.dataTransfer) {
                    e.dataTransfer.effectAllowed = 'move';
                    // Safari 需設定 data 才會啟動拖曳
                    e.dataTransfer.setData('text/plain', String(i));
                }
            });
            item.addEventListener('dragend', () => {
                state.draggingIndex = null;
                item.classList.remove('is-dragging');
                clearDropHints(list);
            });
            item.addEventListener('dragover', (e) => {
                if (state.draggingIndex === null || state.draggingIndex === i) return;
                e.preventDefault();
                if (e.dataTransfer) e.dataTransfer.dropEffect = 'move';
                const rect = item.getBoundingClientRect();
                const isAbove = e.clientY < rect.top + rect.height / 2;
                clearDropHints(list);
                item.classList.add(isAbove ? 'is-drop-above' : 'is-drop-below');
            });
            item.addEventListener('dragleave', (e) => {
                // 只在真的離開（不是進到子元素）時清掉
                if (!item.contains(e.relatedTarget as Node | null)) {
                    item.classList.remove('is-drop-above', 'is-drop-below');
                }
            });
            item.addEventListener('drop', (e) => {
                if (state.draggingIndex === null || state.draggingIndex === i) return;
                e.preventDefault();
                const from = state.draggingIndex;
                const rect = item.getBoundingClientRect();
                const isAbove = e.clientY < rect.top + rect.height / 2;
                let to = isAbove ? i : i + 1;
                state.draggingIndex = null;
                clearDropHints(list);
                reorderCustomers(from, to);
            });
            list.appendChild(item);
        });
        custBody.appendChild(list);
    }

    function clearDropHints(list: HTMLElement): void {
        list.querySelectorAll('.is-drop-above, .is-drop-below').forEach((el) => {
            el.classList.remove('is-drop-above', 'is-drop-below');
        });
    }

    function reorderCustomers(from: number, to: number): void {
        if (from === to || from === to - 1) return; // 同位置或不變
        const customers = state.book.customers;
        const [moved] = customers.splice(from, 1);
        // 移除來源後，若目標 index 原本在來源之後，需要 -1
        const adjustedTo = from < to ? to - 1 : to;
        customers.splice(adjustedTo, 0, moved);

        // 更新 cursor
        if (state.cursor === from) {
            state.cursor = adjustedTo;
        } else if (from < state.cursor && state.cursor <= adjustedTo) {
            state.cursor -= 1;
        } else if (adjustedTo <= state.cursor && state.cursor < from) {
            state.cursor += 1;
        }

        scheduleSave();
        renderCustBody();
    }

    // ====== 工具列事件 ======
    lineNoInput.addEventListener('input', () => {
        const v = parseInt(lineNoInput.value, 10);
        state.book.lineNo = Number.isFinite(v) && v > 0 ? v : 1;
        lineLabelEl.textContent = lineLabel(state.book.lineNo);
        scheduleSave();
    });
    lineNameInput.addEventListener('input', () => {
        state.book.lineName = lineNameInput.value;
        scheduleSave();
    });
    yearInput.addEventListener('input', () => {
        const v = parseInt(yearInput.value, 10);
        state.book.year = Number.isFinite(v) ? v : state.book.year;
        scheduleSave();
    });
    monthSelect.addEventListener('change', () => {
        state.book.month = parseInt(monthSelect.value, 10);
        scheduleSave();
    });

    newBtn.addEventListener('click', () => {
        // 先確保當前已儲存
        if (state.saveTimer !== null) {
            window.clearTimeout(state.saveTimer);
            saveBook(state.book);
            state.saveTimer = null;
        }
        state.book = createEmptyBook();
        state.book.customers.push(createEmptyCustomer());
        state.cursor = 0;
        saveBook(state.book);
        setActiveId(state.book.id);
        syncToolbar();
        renderCustBody();
        setSaveStatus('已儲存', 'saved');
        showToast({variant: 'success', title: '已建立新手填本', message: '請填寫線別與客戶資料'});
    });

    // 匯入成功後的共用處理：.xlsx 與「從 JSON 匯入」皆走此路徑。
    // 沿用匯入來源帶來的 id（與 readHandfillBook 行為一致），存檔並刷新 UI。
    // 不自動重排：匯入的順序即視為真相，由使用者自行決定是否按「排序」。
    function applyImportedBook(imported: HandfillBook): void {
        state.book = imported;
        state.cursor = 0;
        saveBook(state.book);
        setActiveId(state.book.id);
        syncToolbar();
        renderCustBody();
        setSaveStatus('已儲存', 'saved');
        showToast({
            variant: 'success',
            title: '匯入完成',
            message: `共解析 ${state.book.customers.length} 個客戶`,
        });
    }

    importBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', async () => {
        const file = fileInput.files?.[0];
        if (!file) return;
        try {
            setSaveStatus('解析中…', 'saving');
            applyImportedBook(await readHandfillBook(file));
        } catch (err) {
            setSaveStatus('已儲存', 'saved');
            showToast({
                variant: 'error',
                title: '匯入失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        } finally {
            fileInput.value = '';
        }
    });

    importJsonBtn.addEventListener('click', () => {
        openHandfillJsonImportDialog({
            onImport: (book) => applyImportedBook(book),
        });
    });

    historyBtn.addEventListener('click', () => {
        openHandfillHistoryDialog({
            currentId: state.book.id,
            onSelect: (id) => {
                if (id === state.book.id) return;
                if (state.saveTimer !== null) {
                    window.clearTimeout(state.saveTimer);
                    saveBook(state.book);
                    state.saveTimer = null;
                }
                const loaded = loadBook(id);
                if (!loaded) {
                    showToast({variant: 'error', title: '載入失敗', message: '找不到該紀錄'});
                    return;
                }
                state.book = loaded;
                if (state.book.customers.length === 0) {
                    state.book.customers.push(createEmptyCustomer());
                }
                state.cursor = 0;
                setActiveId(state.book.id);
                syncToolbar();
                renderCustBody();
                setSaveStatus('已儲存', 'saved');
            },
            onDelete: (id) => {
                deleteBookFromStore(id);
                if (id === state.book.id) {
                    state.book = createEmptyBook();
                    state.book.customers.push(createEmptyCustomer());
                    state.cursor = 0;
                    saveBook(state.book);
                    setActiveId(state.book.id);
                    syncToolbar();
                    renderCustBody();
                }
            },
        });
    });

    exportBtn.addEventListener('click', async () => {
        if (state.saveTimer !== null) {
            window.clearTimeout(state.saveTimer);
            saveBook(state.book);
            state.saveTimer = null;
            setSaveStatus('已儲存', 'saved');
        }
        try {
            const validCustomers = state.book.customers.filter((c) => !isCustomerEmpty(c));
            if (validCustomers.length === 0) {
                showToast({variant: 'error', title: '無法匯出', message: '請至少填寫一個客戶'});
                return;
            }
            const wb = await buildHandfillWorkbook({...state.book, customers: validCustomers});
            const blob = await workbookToBlob(wb);
            const fname = `${state.book.lineNo}.${state.book.lineName || '未命名'}_${state.book.year}年${state.book.month}月.xlsx`;
            saveAs(blob, fname);
            showToast({variant: 'success', title: '匯出完成', message: fname});
        } catch (err) {
            showToast({
                variant: 'error',
                title: '匯出失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        }
    });

    // ====== 客戶切換事件 ======
    prevBtn.addEventListener('click', () => {
        if (state.viewMode !== 'card') return;
        if (state.cursor > 0) {
            state.cursor--;
            renderCustBody();
        }
    });
    nextBtn.addEventListener('click', () => {
        if (state.viewMode !== 'card') return;
        if (state.cursor < state.book.customers.length - 1) {
            state.cursor++;
            renderCustBody();
        }
    });
    backToListBtn.addEventListener('click', () => {
        state.viewMode = 'list';
        renderCustBody();
    });
    sortAllBtn.addEventListener('click', () => {
        if (state.book.customers.length === 0) return;
        // 對每一家客戶套用與單店「排序」鈕相同的規則：空白置頂 → 帳單排序表順序 → 未知殿後。
        for (const cust of state.book.customers) {
            sortProductsBlanksFirst(cust.products, state.cargoNames);
        }
        renderCustBody();
        scheduleSave();
        showToast({
            variant: 'success',
            title: '已全部排序',
            message: `已將 ${state.book.customers.length} 個客戶的品名清單排序`,
        });
    });
    addCustBtn.addEventListener('click', () => {
        const newCust = createEmptyCustomer();
        state.book.customers.push(newCust);
        state.cursor = state.book.customers.length - 1;
        state.viewMode = 'card';
        renderCustBody();
        scheduleSave();
    });
    deleteCustBtn.addEventListener('click', () => {
        if (state.book.customers.length === 0) return;
        const confirmed = window.confirm('確定要刪除此客戶？');
        if (!confirmed) return;
        state.book.customers.splice(state.cursor, 1);
        if (state.book.customers.length === 0) {
            state.book.customers.push(createEmptyCustomer());
        }
        state.cursor = Math.max(0, state.cursor - 1);
        renderCustBody();
        scheduleSave();
    });

    // ====== 初始 render ======
    syncToolbar();
    renderCustBody();

    // 若 sortingList 尚未準備好（首次啟動 race），等載入後更新品名建議並重新繪製。
    // 不自動重排：cargoNames 僅供「排序」按鈕與品名輸入建議使用。
    if (state.cargoNames.length === 0) {
        loadSortingList()
            .then((list) => {
                state.cargoNames = list.cargoItems.map((it) => it.name);
                renderCustBody();
            })
            .catch(() => {
                // 載入失敗的 toast 已由 main.ts 統一處理，這裡不重複提示
            });
    }

    return panel;
}

function escapeHtml(s: string): string {
    return s.replace(/[&<>"']/g, (ch) =>
        ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '"' ? '&quot;' : '&#39;'
    );
}

// blob helper（避免在 export 時 import excel-service —— 統一從 writer 模組 re-export）
async function workbookToBlob(wb: import('exceljs').Workbook): Promise<Blob> {
    const buffer = await wb.xlsx.writeBuffer();
    return new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
}
