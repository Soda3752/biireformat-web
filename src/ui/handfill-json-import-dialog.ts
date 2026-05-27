/**
 * 「生成手填本」從 JSON 文字匯入對話框。
 *
 * 讓使用者直接貼上 manifest JSON（{ version, book, layoutHash } 或頂層 HandfillBook），
 * 跳過 .xlsx 檔案解析。解析與驗證在 dialog 內進行：
 *   - 成功 → 關窗並以解析出的 book 呼叫 onImport
 *   - 失敗 → 不關窗，行內顯示錯誤訊息讓使用者修正後再試
 */

import {icon} from '@/ui/icons';
import type {HandfillBook} from '@/domain/models/handfill-book';
import {parseHandfillBookFromJson} from '@/readers/handfill-reader';

export interface HandfillJsonImportDialogOptions {
    onImport: (book: HandfillBook) => void;
}

const PLACEHOLDER_EXAMPLE = `{
  "version": 1,
  "book": {
    "lineNo": 1,
    "lineName": "彰化",
    "year": 113,
    "month": 5,
    "customers": [
      { "customerId": "001", "customerName": "範例客戶", "products": [{ "name": "品名A", "unitPrice": 100 }] }
    ]
  }
}`;

export function openHandfillJsonImportDialog(opts: HandfillJsonImportDialogOptions): void {
    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal handfill-json-import-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');

    dialog.innerHTML = `
    <header class="app-modal-header">
      <h2 class="app-modal-title">從 JSON 匯入</h2>
      <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
    </header>
    <div class="app-modal-body handfill-json-import-body">
      <label class="app-form-label" for="hf-json-input">貼上 manifest JSON</label>
      <textarea
        id="hf-json-input"
        class="app-form-input handfill-json-import-textarea"
        spellcheck="false"
        autocomplete="off"
        data-role="json-input"></textarea>
      <p class="handfill-json-import-error" data-role="error" hidden></p>
    </div>
    <footer class="app-modal-footer">
      <button type="button" class="btn btn-secondary" data-role="cancel">取消</button>
      <button type="button" class="btn btn-primary" data-role="import">匯入</button>
    </footer>
  `;

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    const textarea = dialog.querySelector<HTMLTextAreaElement>('[data-role="json-input"]')!;
    const errorEl = dialog.querySelector<HTMLElement>('[data-role="error"]')!;
    textarea.placeholder = PLACEHOLDER_EXAMPLE;

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
    dialog.querySelector<HTMLButtonElement>('[data-role="close"]')!.addEventListener('click', close);
    dialog.querySelector<HTMLButtonElement>('[data-role="cancel"]')!.addEventListener('click', close);

    const showError = (msg: string) => {
        errorEl.textContent = msg;
        errorEl.hidden = false;
    };
    // 使用者再次輸入時清掉舊錯誤
    textarea.addEventListener('input', () => {
        if (!errorEl.hidden) errorEl.hidden = true;
    });

    dialog.querySelector<HTMLButtonElement>('[data-role="import"]')!.addEventListener('click', () => {
        let book: HandfillBook;
        try {
            book = parseHandfillBookFromJson(textarea.value);
        } catch (err) {
            showError(err instanceof Error ? err.message : String(err));
            return;
        }
        close();
        opts.onImport(book);
    });

    // 開窗即聚焦輸入框，方便直接貼上
    textarea.focus();
}
