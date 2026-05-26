/**
 * 「生成手填本」單一客戶卡片元件。
 *
 * 結構（對應範本檔 col 0、col 1、col 2 邏輯）：
 *   - 代號 / 名稱
 *   - 休息日清單（動態增刪）
 *   - 電話清單（預設 2 個輸入框，可增刪）
 *   - 品名清單（每筆含單價，動態增刪，無上限）
 *
 * 任何輸入變動透過 onChange 回傳更新後的 customer，由外層決定何時 debounce 寫入。
 */

import {icon} from '@/ui/icons';
import {type HandfillCustomer, type HandfillProduct, sortProductsByCargoOrder,} from '@/domain/models/handfill-book';

export interface HandfillCardOptions {
    customer: HandfillCustomer;
    /** 帳單排序品項（cargo_sort）名稱列表，按 csv 順序排列 */
    cargoNames: ReadonlyArray<string>;
    onChange: (next: HandfillCustomer) => void;
}

export function renderHandfillCustomerCard(opts: HandfillCardOptions): HTMLElement {
    let cust: HandfillCustomer = cloneCustomer(opts.customer);

    const card = document.createElement('div');
    card.className = 'handfill-card-inner';

    card.innerHTML = `
    <div class="handfill-card-grid">
      <div class="app-form-row">
        <label class="app-form-label" for="hf-cust-id">客戶代號</label>
        <input id="hf-cust-id" class="app-form-input" type="text" autocomplete="off" placeholder="例：1037">
      </div>
      <div class="app-form-row handfill-card-name">
        <label class="app-form-label" for="hf-cust-name">客戶名稱</label>
        <input id="hf-cust-name" class="app-form-input" type="text" autocomplete="off" placeholder="例：早船長">
      </div>
    </div>

    <div class="handfill-card-section">
      <header class="handfill-card-section-header">
        <h3 class="handfill-card-section-title">休息日 / 備註</h3>
        <button type="button" class="btn btn-secondary btn-sm" data-role="add-rest">
          ${icon('plus', 12)}<span>新增</span>
        </button>
      </header>
      <div class="handfill-card-list" data-role="rest-list"></div>
    </div>

    <div class="handfill-card-section">
      <header class="handfill-card-section-header">
        <h3 class="handfill-card-section-title">電話</h3>
        <button type="button" class="btn btn-secondary btn-sm" data-role="add-phone">
          ${icon('plus', 12)}<span>新增</span>
        </button>
      </header>
      <div class="handfill-card-list" data-role="phone-list"></div>
    </div>

    <div class="handfill-card-section">
      <header class="handfill-card-section-header">
        <h3 class="handfill-card-section-title">
          品名清單
          <span class="handfill-card-section-meta" data-role="product-count">(0)</span>
          <span class="handfill-manual-sort-badge" data-role="manual-sort-badge" hidden>手動排序</span>
        </h3>
        <div class="handfill-card-section-actions">
          <button type="button" class="btn btn-secondary btn-sm" data-role="restore-auto-sort" hidden>
            ${icon('chevron-down', 12)}<span>還原自動排序</span>
          </button>
          <button type="button" class="btn btn-secondary btn-sm" data-role="insert-blank-product">
            ${icon('row-insert-above', 12)}<span>插入空白行</span>
          </button>
          <button type="button" class="btn btn-secondary btn-sm" data-role="add-product">
            ${icon('plus', 12)}<span>新增品名</span>
          </button>
        </div>
      </header>
      <div class="handfill-card-product-list" data-role="product-list"></div>
    </div>
  `;

    const idInput = card.querySelector<HTMLInputElement>('#hf-cust-id')!;
    const nameInput = card.querySelector<HTMLInputElement>('#hf-cust-name')!;
    const restList = card.querySelector<HTMLElement>('[data-role="rest-list"]')!;
    const phoneList = card.querySelector<HTMLElement>('[data-role="phone-list"]')!;
    const productList = card.querySelector<HTMLElement>('[data-role="product-list"]')!;
    const productCountEl = card.querySelector<HTMLElement>('[data-role="product-count"]')!;
    const addRestBtn = card.querySelector<HTMLButtonElement>('[data-role="add-rest"]')!;
    const addPhoneBtn = card.querySelector<HTMLButtonElement>('[data-role="add-phone"]')!;
    const addProductBtn = card.querySelector<HTMLButtonElement>('[data-role="add-product"]')!;
    const insertBlankProductBtn = card.querySelector<HTMLButtonElement>('[data-role="insert-blank-product"]')!;
    const restoreAutoSortBtn = card.querySelector<HTMLButtonElement>('[data-role="restore-auto-sort"]')!;
    const manualSortBadge = card.querySelector<HTMLElement>('[data-role="manual-sort-badge"]')!;

    function notify(): void {
        opts.onChange(cloneCustomer(cust));
    }

    // ====== 基本欄位 ======
    idInput.value = cust.customerId;
    nameInput.value = cust.customerName;
    idInput.addEventListener('input', () => {
        cust.customerId = idInput.value;
        notify();
    });
    nameInput.addEventListener('input', () => {
        cust.customerName = nameInput.value;
        notify();
    });

    // ====== 休息日 ======
    function renderRestList(): void {
        restList.innerHTML = '';
        if (cust.restNotes.length === 0) {
            restList.innerHTML = `<div class="handfill-card-empty">尚無休息日 / 備註</div>`;
            return;
        }
        cust.restNotes.forEach((note, idx) => {
            const row = document.createElement('div');
            row.className = 'handfill-card-list-row';
            row.innerHTML = `
        <input type="text" class="app-form-input" value="${escapeAttr(note)}" placeholder="例：一休、收現、10:30">
        <button type="button" class="btn btn-icon btn-secondary" aria-label="刪除">${icon('close', 14)}</button>
      `;
            const input = row.querySelector<HTMLInputElement>('input')!;
            const delBtn = row.querySelector<HTMLButtonElement>('button')!;
            input.addEventListener('input', () => {
                cust.restNotes[idx] = input.value;
                notify();
            });
            delBtn.addEventListener('click', () => {
                cust.restNotes.splice(idx, 1);
                renderRestList();
                notify();
            });
            restList.appendChild(row);
        });
    }

    addRestBtn.addEventListener('click', () => {
        cust.restNotes.push('');
        renderRestList();
        notify();
        // focus 最後一個
        const last = restList.querySelector<HTMLInputElement>('.handfill-card-list-row:last-child input');
        last?.focus();
    });

    // ====== 電話 ======
    function renderPhoneList(): void {
        phoneList.innerHTML = '';
        if (cust.phones.length === 0) {
            phoneList.innerHTML = `<div class="handfill-card-empty">尚無電話</div>`;
            return;
        }
        cust.phones.forEach((phone, idx) => {
            const row = document.createElement('div');
            row.className = 'handfill-card-list-row';
            const placeholder = idx === 0 ? '例：0937-751896（手機）' : idx === 1 ? '例：7224873（市話）' : '其他電話';
            row.innerHTML = `
        <input type="text" class="app-form-input" value="${escapeAttr(phone)}" placeholder="${placeholder}">
        <button type="button" class="btn btn-icon btn-secondary" aria-label="刪除">${icon('close', 14)}</button>
      `;
            const input = row.querySelector<HTMLInputElement>('input')!;
            const delBtn = row.querySelector<HTMLButtonElement>('button')!;
            input.addEventListener('input', () => {
                cust.phones[idx] = input.value;
                notify();
            });
            delBtn.addEventListener('click', () => {
                cust.phones.splice(idx, 1);
                renderPhoneList();
                notify();
            });
            phoneList.appendChild(row);
        });
    }

    addPhoneBtn.addEventListener('click', () => {
        cust.phones.push('');
        renderPhoneList();
        notify();
        const last = phoneList.querySelector<HTMLInputElement>('.handfill-card-list-row:last-child input');
        last?.focus();
    });

    // ====== 品名 ======
    // 共用 datalist：提供帳單排序品項作為輸入建議；不限制使用者手填新項目
    const datalistId = `hf-cargo-${cust.id}`;
    const datalistEl = document.createElement('datalist');
    datalistEl.id = datalistId;
    datalistEl.innerHTML = opts.cargoNames
        .map((name) => `<option value="${escapeAttr(name)}"></option>`)
        .join('');
    card.appendChild(datalistEl);

    function syncSortUI(): void {
        const isManual = cust.manualSort === true;
        restoreAutoSortBtn.hidden = !isManual;
        manualSortBadge.hidden = !isManual;
    }

    // 拖曳排序狀態：來源列索引（null 表示沒有正在拖曳）
    let dragSrcIdx: number | null = null;

    function clearDropIndicators(): void {
        productList.querySelectorAll('.drop-above, .drop-below').forEach((el) => {
            el.classList.remove('drop-above', 'drop-below');
        });
    }

    function renderProductList(): void {
        productList.innerHTML = '';
        productCountEl.textContent = `(${cust.products.filter((p) => p.name.trim()).length})`;
        if (cust.products.length === 0) {
            productList.innerHTML = `<div class="handfill-card-empty">尚無品名，點擊「新增品名」開始</div>`;
            return;
        }
        cust.products.forEach((prod, idx) => {
            const row = document.createElement('div');
            row.className = 'handfill-card-product-row';
            row.dataset.idx = String(idx);
            row.innerHTML = `
        <button type="button" class="handfill-product-handle" data-role="drag-handle" aria-label="拖曳排序" title="拖曳排序">${icon('grip', 14)}</button>
        <span class="handfill-product-idx">${idx + 1}.</span>
        <input type="text" class="app-form-input handfill-product-name" list="${datalistId}" value="${escapeAttr(prod.name)}" placeholder="品名（可從建議選或手填）" autocomplete="off">
        <input type="number" class="app-form-input handfill-product-price" value="${prod.unitPrice ?? ''}" placeholder="單價" step="0.01">
        <button type="button" class="btn btn-icon btn-secondary handfill-product-delete" aria-label="刪除">${icon('close', 14)}</button>
      `;
            const handleBtn = row.querySelector<HTMLButtonElement>('[data-role="drag-handle"]')!;
            const nameInp = row.querySelector<HTMLInputElement>('.handfill-product-name')!;
            const priceInp = row.querySelector<HTMLInputElement>('.handfill-product-price')!;
            const delBtn = row.querySelector<HTMLButtonElement>('.handfill-product-delete')!;

            // 拖曳把手：mousedown 才把 row 變為 draggable，避免使用者選取 input 文字時誤觸
            handleBtn.addEventListener('mousedown', () => {
                row.draggable = true;
            });
            // 把手只負責觸發 drag，禁止 click 動作（避免 button default 行為）
            handleBtn.addEventListener('click', (e) => e.preventDefault());

            row.addEventListener('dragstart', (e) => {
                if (!row.draggable) return;
                dragSrcIdx = idx;
                row.classList.add('dragging');
                if (e.dataTransfer) {
                    e.dataTransfer.effectAllowed = 'move';
                    e.dataTransfer.setData('text/plain', String(idx));
                }
            });
            row.addEventListener('dragend', () => {
                row.classList.remove('dragging');
                row.draggable = false;
                dragSrcIdx = null;
                clearDropIndicators();
            });

            // input：每次鍵入皆同步資料，確保自動儲存不漏；不重繪也不重排
            nameInp.addEventListener('input', () => {
                cust.products[idx].name = nameInp.value;
                productCountEl.textContent = `(${cust.products.filter((p) => p.name.trim()).length})`;
                notify();
            });
            // change：使用者「確認」輸入（blur 或選 datalist 建議）時依排序表重排
            // 但若已進入手動排序模式（cust.manualSort=true），則保留手動順序，不再自動重排
            nameInp.addEventListener('change', () => {
                cust.products[idx].name = nameInp.value;
                if (!cust.manualSort && opts.cargoNames.length > 0) {
                    sortProductsByCargoOrder(cust.products, opts.cargoNames);
                    renderProductList();
                }
                notify();
            });
            priceInp.addEventListener('input', () => {
                const v = parseFloat(priceInp.value);
                cust.products[idx].unitPrice = Number.isFinite(v) ? v : undefined;
                notify();
            });
            delBtn.addEventListener('click', () => {
                cust.products.splice(idx, 1);
                renderProductList();
                notify();
            });
            productList.appendChild(row);
        });
    }

    addProductBtn.addEventListener('click', () => {
        cust.products.push({name: '', unitPrice: undefined});
        renderProductList();
        notify();
        const last = productList.querySelector<HTMLInputElement>('.handfill-card-product-row:last-child .handfill-product-name');
        last?.focus();
    });

    insertBlankProductBtn.addEventListener('click', () => {
        cust.products.unshift({name: '', unitPrice: undefined});
        // 「插入空白行」是排版意圖，視同手動排序，避免後續 nameInput change 把空白行沖到末端
        cust.manualSort = true;
        renderProductList();
        syncSortUI();
        notify();
        const first = productList.querySelector<HTMLInputElement>('.handfill-card-product-row:first-child .handfill-product-name');
        first?.focus();
    });

    restoreAutoSortBtn.addEventListener('click', () => {
        cust.manualSort = false;
        if (opts.cargoNames.length > 0) {
            sortProductsByCargoOrder(cust.products, opts.cargoNames);
        }
        renderProductList();
        syncSortUI();
        notify();
    });

    // productList 層級的 D&D listener：只綁一次，不隨 renderProductList 重複綁定
    productList.addEventListener('dragover', (e) => {
        if (dragSrcIdx === null) return;
        e.preventDefault();
        const target = (e.target as HTMLElement | null)?.closest<HTMLElement>('.handfill-card-product-row');
        if (!target || !productList.contains(target)) return;
        clearDropIndicators();
        const rect = target.getBoundingClientRect();
        const midY = rect.top + rect.height / 2;
        target.classList.add(e.clientY < midY ? 'drop-above' : 'drop-below');
    });

    productList.addEventListener('dragleave', (e) => {
        // 真正離開整個 list 容器時才清掉指示
        if (e.target === productList) {
            clearDropIndicators();
        }
    });

    productList.addEventListener('drop', (e) => {
        if (dragSrcIdx === null) return;
        e.preventDefault();
        const target = (e.target as HTMLElement | null)?.closest<HTMLElement>('.handfill-card-product-row');
        if (!target || !productList.contains(target)) {
            clearDropIndicators();
            return;
        }
        const targetIdx = parseInt(target.dataset.idx ?? '-1', 10);
        if (targetIdx < 0) {
            clearDropIndicators();
            return;
        }
        const rect = target.getBoundingClientRect();
        const midY = rect.top + rect.height / 2;
        const insertAfter = e.clientY >= midY;
        const fromIdx = dragSrcIdx;
        let toIdx = insertAfter ? targetIdx + 1 : targetIdx;
        // 從原位置移除後，若 fromIdx < toIdx，目標位置會往前縮 1
        if (fromIdx < toIdx) toIdx -= 1;
        clearDropIndicators();
        if (fromIdx === toIdx) return;
        const [moved] = cust.products.splice(fromIdx, 1);
        cust.products.splice(toIdx, 0, moved);
        cust.manualSort = true;
        renderProductList();
        syncSortUI();
        notify();
    });

    renderRestList();
    renderPhoneList();
    renderProductList();
    syncSortUI();

    return card;
}

function cloneCustomer(c: HandfillCustomer): HandfillCustomer {
    return {
        id: c.id,
        customerId: c.customerId,
        customerName: c.customerName,
        products: c.products.map((p): HandfillProduct => ({name: p.name, unitPrice: p.unitPrice})),
        restNotes: [...c.restNotes],
        phones: [...c.phones],
        manualSort: c.manualSort === true,
    };
}

function escapeAttr(s: string): string {
    return s.replace(/[&<>"']/g, (ch) =>
        ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '"' ? '&quot;' : '&#39;'
    );
}
