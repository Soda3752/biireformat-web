/**
 * 「手填本搬移」主面板。
 *
 * 把一份手填本裡的商家複製 / 搬移到另一份手填本。兩份都必須已存在於歷史紀錄
 * （在「生成手填本」按過「匯入 .xlsx」或「新建」），本頁只在 localStorage 之間搬資料，
 * 不解析也不產生 .xlsx —— 搬完仍由「生成手填本」匯出。
 *
 * 與「生成手填本」的協調：本頁直接寫 localStorage，該頁改為切離時寫回、切回時重讀
 * （見 handfill.ts 的 MutationObserver），避免那邊的舊記憶體狀態覆蓋搬移結果。
 */

import {icon} from '@/ui/icons';
import {type TabDefinition} from '@/ui/tabs';
import {showToast} from '@/ui/toast';

import {
    genId,
    type HandfillCustomer,
    createEmptyCustomer,
    isCustomerEmpty,
} from '@/domain/models/handfill-book';
import {type HandfillBookSummary, listBooks, loadBook, saveBook, setActiveId} from '@/infra/handfill-store';

/** 搬移方式：複製（來源保留）或搬移（來源刪掉） */
type TransferMode = 'copy' | 'move';
/** 目標本已有同代號商家時的處理方式 */
type ConflictMode = 'skip' | 'overwrite' | 'both';

interface TransferState {
    books: HandfillBookSummary[];
    sourceId: string;
    targetId: string;
    /** 已勾選的來源商家（HandfillCustomer.id） */
    selected: Set<string>;
    mode: TransferMode;
    conflict: ConflictMode;
}

export function renderHandfillTransferPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel hf-transfer-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    const state: TransferState = {
        books: [],
        sourceId: '',
        targetId: '',
        selected: new Set(),
        mode: 'copy',
        conflict: 'skip',
    };

    panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <div>
          <h1 class="card-title">手填本搬移</h1>
          <p class="card-subtitle">
            把一份手填本裡的商家複製或搬移到另一份手填本。兩份都要先在「生成手填本」用「匯入 .xlsx」匯入過（或是在那裡新建的），才會出現在下方清單。
          </p>
        </div>
      </header>

      <section class="hf-transfer-books">
        <div class="app-form-row hf-transfer-book-field">
          <label class="app-form-label" for="hf-tr-source">來源手填本</label>
          <select id="hf-tr-source" class="app-form-input" data-role="source"></select>
        </div>
        <span class="hf-transfer-arrow" aria-hidden="true">${icon('transfer', 20)}</span>
        <div class="app-form-row hf-transfer-book-field">
          <label class="app-form-label" for="hf-tr-target">目標手填本</label>
          <select id="hf-tr-target" class="app-form-input" data-role="target"></select>
        </div>
        <button type="button" class="btn btn-secondary btn-sm hf-transfer-reload" data-role="reload">
          ${icon('upload', 14)}<span>重新整理清單</span>
        </button>
      </section>
      <p class="hf-transfer-hint" data-role="hint" hidden></p>

      <section class="hf-transfer-list-toolbar">
        <label class="hf-transfer-checkall">
          <input type="checkbox" data-role="check-all">
          <span>全選</span>
        </label>
        <span class="hf-transfer-count" data-role="sel-count">已選 0 家</span>
        <span class="hf-transfer-target-info" data-role="target-info"></span>
      </section>
      <div class="hf-transfer-list" data-role="list"></div>

      <section class="hf-transfer-options">
        <fieldset class="hf-transfer-optgroup">
          <legend>搬移方式</legend>
          <label><input type="radio" name="hf-tr-mode" value="copy" checked><span>複製過去（來源保留）</span></label>
          <label><input type="radio" name="hf-tr-mode" value="move"><span>搬過去（來源刪掉這幾家）</span></label>
        </fieldset>
        <fieldset class="hf-transfer-optgroup">
          <legend>目標本已有同代號商家時</legend>
          <label><input type="radio" name="hf-tr-conflict" value="skip" checked><span>跳過不搬</span></label>
          <label><input type="radio" name="hf-tr-conflict" value="overwrite"><span>覆蓋目標那一家</span></label>
          <label><input type="radio" name="hf-tr-conflict" value="both"><span>兩筆都留</span></label>
        </fieldset>
      </section>

      <footer class="hf-transfer-footer">
        <button type="button" class="btn btn-primary" data-role="run">
          ${icon('check', 16)}<span>開始搬移</span>
        </button>
        <p class="hf-transfer-result" data-role="result" hidden></p>
      </footer>
    </div>
  `;

    const sourceSel = panel.querySelector<HTMLSelectElement>('[data-role="source"]')!;
    const targetSel = panel.querySelector<HTMLSelectElement>('[data-role="target"]')!;
    const reloadBtn = panel.querySelector<HTMLButtonElement>('[data-role="reload"]')!;
    const hintEl = panel.querySelector<HTMLElement>('[data-role="hint"]')!;
    const checkAll = panel.querySelector<HTMLInputElement>('[data-role="check-all"]')!;
    const selCountEl = panel.querySelector<HTMLElement>('[data-role="sel-count"]')!;
    const targetInfoEl = panel.querySelector<HTMLElement>('[data-role="target-info"]')!;
    const listEl = panel.querySelector<HTMLElement>('[data-role="list"]')!;
    const runBtn = panel.querySelector<HTMLButtonElement>('[data-role="run"]')!;
    const resultEl = panel.querySelector<HTMLElement>('[data-role="result"]')!;

    /** 目前來源本可搬移的商家（略過空殼），每次重繪時重讀 localStorage */
    let sourceCustomers: HandfillCustomer[] = [];
    /** 目標本已存在的客戶代號，用來標示「目標已有」 */
    let targetCodes = new Set<string>();

    /**
     * 每本手填本「實際有資料的商家數」。listBooks 的 customerCount 會把空白列也算進去，
     * 與畫面上「目標本現有 N 家」的算法不同，兩個數字並排會對不起來，所以這裡自己算。
     */
    let realCounts = new Map<string, number>();

    function bookLabel(b: HandfillBookSummary): string {
        const name = b.fullName.trim() || '(未命名線別)';
        return `${name} ${b.year}年${b.month}月 · ${realCounts.get(b.id) ?? b.customerCount} 家`;
    }

    function fillSelect(sel: HTMLSelectElement, selectedId: string): void {
        sel.innerHTML = '';
        for (const b of state.books) {
            const opt = document.createElement('option');
            opt.value = b.id;
            opt.textContent = bookLabel(b);
            sel.appendChild(opt);
        }
        sel.value = selectedId;
    }

    /** 重讀歷史紀錄清單；盡量沿用使用者已選的來源／目標 */
    function refreshBooks(): void {
        state.books = listBooks();
        realCounts = new Map(
            state.books.map((b) => [
                b.id,
                (loadBook(b.id)?.customers ?? []).filter((c) => !isCustomerEmpty(c)).length,
            ])
        );
        const ids = state.books.map((b) => b.id);
        if (!ids.includes(state.sourceId)) state.sourceId = ids[0] ?? '';
        // 目標預設挑一本跟來源不同的
        if (!ids.includes(state.targetId) || state.targetId === state.sourceId) {
            state.targetId = ids.find((id) => id !== state.sourceId) ?? '';
        }
        fillSelect(sourceSel, state.sourceId);
        fillSelect(targetSel, state.targetId);
        renderList();
    }

    function renderList(): void {
        const source = state.sourceId ? loadBook(state.sourceId) : null;
        const target = state.targetId ? loadBook(state.targetId) : null;
        sourceCustomers = (source?.customers ?? []).filter((c) => !isCustomerEmpty(c));
        targetCodes = new Set(
            (target?.customers ?? [])
                .map((c) => c.customerId.trim())
                .filter((code) => code.length > 0)
        );

        // 只保留仍存在於來源本的勾選
        const validIds = new Set(sourceCustomers.map((c) => c.id));
        for (const id of [...state.selected]) {
            if (!validIds.has(id)) state.selected.delete(id);
        }

        // 頁面狀態提示
        const hints: string[] = [];
        if (state.books.length === 0) {
            hints.push('目前沒有任何手填本。請先到「生成手填本」用「匯入 .xlsx」把檔案匯入，或在那裡新建一本。');
        } else if (state.books.length === 1) {
            hints.push('目前只有一份手填本，還沒有可以搬過去的目標。請再匯入另一份檔案。');
        } else if (source && target && source.id === target.id) {
            hints.push('來源與目標是同一份手填本，請改選其中一邊。');
        }
        hintEl.textContent = hints.join('');
        hintEl.hidden = hints.length === 0;

        targetInfoEl.textContent = target
            ? `目標本現有 ${target.customers.filter((c) => !isCustomerEmpty(c)).length} 家`
            : '';

        listEl.innerHTML = '';
        if (sourceCustomers.length === 0) {
            const empty = document.createElement('div');
            empty.className = 'hf-transfer-empty';
            empty.textContent = source
                ? '這份手填本沒有可搬移的商家。'
                : '請先選擇來源手填本。';
            listEl.appendChild(empty);
            syncSelectionUi();
            return;
        }

        for (const c of sourceCustomers) {
            const row = document.createElement('label');
            row.className = 'hf-transfer-item';

            const box = document.createElement('input');
            box.type = 'checkbox';
            box.className = 'hf-transfer-item-check';
            box.checked = state.selected.has(c.id);
            box.addEventListener('change', () => {
                if (box.checked) state.selected.add(c.id);
                else state.selected.delete(c.id);
                syncSelectionUi();
            });
            row.appendChild(box);

            const code = document.createElement('span');
            code.className = 'hf-transfer-item-id';
            code.textContent = c.customerId.trim() || '─';
            row.appendChild(code);

            const name = document.createElement('span');
            name.className = 'hf-transfer-item-name';
            name.textContent = c.customerName.trim() || '(未命名)';
            row.appendChild(name);

            const products = document.createElement('span');
            products.className = 'hf-transfer-item-products';
            products.textContent = c.products
                .map((p) => p.name)
                .filter((s) => s.trim())
                .slice(0, 4)
                .join('、');
            row.appendChild(products);

            const meta = document.createElement('span');
            meta.className = 'hf-transfer-item-meta';
            meta.textContent = `${c.products.filter((p) => p.name.trim()).length} 品`;
            row.appendChild(meta);

            const flag = document.createElement('span');
            flag.className = 'hf-transfer-item-flag';
            if (c.customerId.trim() && targetCodes.has(c.customerId.trim())) {
                flag.textContent = '目標已有';
                flag.classList.add('is-dup');
            }
            row.appendChild(flag);

            listEl.appendChild(row);
        }
        syncSelectionUi();
    }

    function syncSelectionUi(): void {
        // 上一次的搬移結果只描述當時那組來源／目標與勾選，一有變動就收掉
        resultEl.hidden = true;
        resultEl.textContent = '';
        const total = sourceCustomers.length;
        const picked = state.selected.size;
        selCountEl.textContent = `已選 ${picked} 家`;
        checkAll.checked = total > 0 && picked === total;
        checkAll.indeterminate = picked > 0 && picked < total;
        checkAll.disabled = total === 0;
        runBtn.disabled =
            picked === 0 || !state.sourceId || !state.targetId || state.sourceId === state.targetId;
    }

    function transferCustomer(c: HandfillCustomer): HandfillCustomer {
        return {
            id: genId('cust-'),
            customerId: c.customerId,
            customerName: c.customerName,
            products: c.products.map((p) => ({name: p.name, unitPrice: p.unitPrice})),
            restNotes: [...c.restNotes],
            phones: [...c.phones],
            manualSort: c.manualSort,
        };
    }

    function runTransfer(): void {
        // 重新讀一次，避免用畫面上的舊資料寫入
        const source = loadBook(state.sourceId);
        const target = loadBook(state.targetId);
        if (!source || !target) {
            showToast({variant: 'error', title: '搬移失敗', message: '找不到來源或目標手填本，請重新整理清單'});
            refreshBooks();
            return;
        }

        const picked = source.customers.filter((c) => state.selected.has(c.id));
        if (picked.length === 0) {
            showToast({variant: 'error', title: '無法搬移', message: '請先勾選要搬移的商家'});
            renderList();
            return;
        }

        let added = 0;
        let overwritten = 0;
        let skipped = 0;
        /** 真正搬成功的來源 id，「搬過去」模式只刪這些（跳過的要留著） */
        const doneIds = new Set<string>();
        /**
         * 本次已被覆蓋掉的目標列。來源本若有兩家同代號，第二家不能再蓋同一列
         * （否則第一家的資料會被吃掉，「搬過去」模式還會把兩家都從來源刪除）。
         */
        const usedTargetIdx = new Set<number>();

        // 目標本尾端的空白列（生成手填本刪到最後一家時留下的）先剔除，
        // 免得搬過去的商家排在空白列後面，看起來中間夾了一列空的
        while (
            target.customers.length > 0 &&
            isCustomerEmpty(target.customers[target.customers.length - 1])
            ) {
            target.customers.pop();
        }

        for (const c of picked) {
            const code = c.customerId.trim();
            const dupIdx = code
                ? target.customers.findIndex(
                    (t, i) => !usedTargetIdx.has(i) && t.customerId.trim() === code
                )
                : -1;
            if (dupIdx >= 0 && state.conflict === 'skip') {
                skipped++;
                continue;
            }
            if (dupIdx >= 0 && state.conflict === 'overwrite') {
                // 沿用目標那一家原本的 id，維持它在目標本裡的身分
                target.customers[dupIdx] = {...transferCustomer(c), id: target.customers[dupIdx].id};
                usedTargetIdx.add(dupIdx);
                overwritten++;
            } else {
                target.customers.push(transferCustomer(c));
                // 新增的列同樣不能再被後面的同代號來源蓋掉
                usedTargetIdx.add(target.customers.length - 1);
                added++;
            }
            doneIds.add(c.id);
        }

        if (doneIds.size === 0) {
            resultEl.hidden = false;
            resultEl.textContent = `沒有任何商家被搬移（跳過 ${skipped} 家：目標本已有相同代號）。`;
            showToast({variant: 'warning', title: '沒有搬移任何商家', message: `跳過 ${skipped} 家同代號商家`});
            return;
        }

        try {
            saveBook(target);
            if (state.mode === 'move') {
                source.customers = source.customers.filter((c) => !doneIds.has(c.id));
                // 與「生成手填本」刪到最後一家的行為一致：留一張空白客戶
                if (source.customers.length === 0) source.customers.push(createEmptyCustomer());
                saveBook(source);
            }
        } catch (err) {
            showToast({
                variant: 'error',
                title: '搬移失敗',
                message: err instanceof Error ? err.message : String(err),
            });
            return;
        }

        // 動作放前面（例：「搬移：覆蓋 3 家」），全部都是覆蓋時也看得出來源那幾家已被移除；
        // 數量為 0 的項目不列出，免得出現「新增 0 家」這種怪字樣
        const verb = state.mode === 'move' ? '搬移' : '複製';
        const parts: string[] = [];
        if (added > 0) parts.push(`新增 ${added} 家`);
        if (overwritten > 0) parts.push(`覆蓋 ${overwritten} 家`);
        if (skipped > 0) parts.push(`跳過 ${skipped} 家`);
        const summary = `${verb}：${parts.join('、')}`;

        state.selected.clear();
        refreshBooks();
        resultEl.hidden = false;
        resultEl.textContent = `已${summary}。到「生成手填本」核對後即可匯出 .xlsx。`;
        showToast({variant: 'success', title: '搬移完成', message: summary});

        const targetId = target.id;
        if (window.confirm(`已${summary}。要現在切到目標手填本檢查嗎？`)) {
            // 設為當前工作中的手填本；切過去時「生成手填本」會自己重讀
            setActiveId(targetId);
            window.location.hash = '#handfill';
        }
    }

    // ====== 事件 ======
    sourceSel.addEventListener('change', () => {
        state.sourceId = sourceSel.value;
        if (state.targetId === state.sourceId) {
            const other = state.books.find((b) => b.id !== state.sourceId);
            state.targetId = other?.id ?? '';
            targetSel.value = state.targetId;
        }
        state.selected.clear();
        renderList();
    });

    targetSel.addEventListener('change', () => {
        state.targetId = targetSel.value;
        renderList();
    });

    reloadBtn.addEventListener('click', () => {
        refreshBooks();
        showToast({variant: 'success', title: '清單已更新', message: `共 ${state.books.length} 份手填本`});
    });

    checkAll.addEventListener('change', () => {
        state.selected.clear();
        if (checkAll.checked) {
            for (const c of sourceCustomers) state.selected.add(c.id);
        }
        renderList();
    });

    panel.querySelectorAll<HTMLInputElement>('input[name="hf-tr-mode"]').forEach((el) => {
        el.addEventListener('change', () => {
            if (el.checked) state.mode = el.value as TransferMode;
        });
    });

    panel.querySelectorAll<HTMLInputElement>('input[name="hf-tr-conflict"]').forEach((el) => {
        el.addEventListener('change', () => {
            if (el.checked) state.conflict = el.value as ConflictMode;
        });
    });

    runBtn.addEventListener('click', runTransfer);

    // 使用者可能剛在「生成手填本」匯入新檔案；每次切回本頁就重讀清單
    const observer = new MutationObserver(() => {
        if (panel.classList.contains('is-active')) refreshBooks();
    });
    observer.observe(panel, {attributeFilter: ['class']});

    refreshBooks();
    return panel;
}
