/**
 * 「生成手填本」歷史紀錄對話框。
 *
 * 列出 localStorage 中所有手填本，依線別 + 年月分組，
 * 支援載入與刪除。
 */

import {icon} from '@/ui/icons';
import {type HandfillBookSummary, listBooks} from '@/infra/handfill-store';

export interface HandfillHistoryDialogOptions {
    currentId?: string;
    onSelect: (id: string) => void;
    onDelete: (id: string) => void;
}

export function openHandfillHistoryDialog(opts: HandfillHistoryDialogOptions): void {
    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal handfill-history-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');

    dialog.innerHTML = `
    <header class="app-modal-header">
      <h2 class="app-modal-title">歷史紀錄</h2>
      <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
    </header>
    <div class="app-modal-body handfill-history-body">
      <div class="handfill-history-toolbar">
        <input class="app-form-input" type="search" placeholder="搜尋線別 / 年月" data-role="search" autocomplete="off">
      </div>
      <div class="handfill-history-list" data-role="list"></div>
      <div class="handfill-history-empty" data-role="empty" hidden>尚無歷史紀錄</div>
    </div>
    <footer class="app-modal-footer">
      <button type="button" class="btn btn-secondary" data-role="cancel">關閉</button>
    </footer>
  `;

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    const listEl = dialog.querySelector<HTMLElement>('[data-role="list"]')!;
    const emptyEl = dialog.querySelector<HTMLElement>('[data-role="empty"]')!;
    const searchInput = dialog.querySelector<HTMLInputElement>('[data-role="search"]')!;

    let keyword = '';
    let allBooks: HandfillBookSummary[] = listBooks();

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

    searchInput.addEventListener('input', () => {
        keyword = searchInput.value.trim().toLowerCase();
        renderList();
    });

    function renderList(): void {
        const filtered = allBooks.filter((b) => {
            if (!keyword) return true;
            const target = `${b.fullName} ${b.year}年${b.month}月`.toLowerCase();
            return target.includes(keyword);
        });

        listEl.innerHTML = '';
        if (filtered.length === 0) {
            emptyEl.hidden = false;
            return;
        }
        emptyEl.hidden = true;

        // 依線別分組
        const groups = new Map<string, HandfillBookSummary[]>();
        for (const b of filtered) {
            const key = b.fullName || '(未命名線別)';
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key)!.push(b);
        }

        for (const [groupName, items] of groups) {
            const group = document.createElement('div');
            group.className = 'handfill-history-group';
            group.innerHTML = `<h3 class="handfill-history-group-title">${escapeHtml(groupName)}</h3>`;
            for (const b of items) {
                const row = document.createElement('div');
                row.className = 'handfill-history-row';
                if (b.id === opts.currentId) row.classList.add('is-active');
                row.innerHTML = `
          <div class="handfill-history-row-main">
            <span class="handfill-history-row-title">${b.year} 年 ${b.month} 月</span>
            <span class="handfill-history-row-meta">${b.customerCount} 個客戶 · ${relativeTime(b.updatedAt)}</span>
          </div>
          <div class="handfill-history-row-actions">
            <button type="button" class="btn btn-secondary btn-sm" data-action="load">載入</button>
            <button type="button" class="btn btn-icon btn-secondary" data-action="delete" aria-label="刪除">${icon('trash', 14)}</button>
          </div>
        `;
                row.querySelector<HTMLButtonElement>('[data-action="load"]')!.addEventListener('click', () => {
                    opts.onSelect(b.id);
                    close();
                });
                row.querySelector<HTMLButtonElement>('[data-action="delete"]')!.addEventListener('click', () => {
                    if (!window.confirm(`確定要刪除「${groupName} ${b.year}/${b.month}」這筆紀錄？`)) return;
                    opts.onDelete(b.id);
                    allBooks = allBooks.filter((x) => x.id !== b.id);
                    renderList();
                });
                group.appendChild(row);
            }
            listEl.appendChild(group);
        }
    }

    renderList();
}

function relativeTime(ts: number): string {
    if (!ts) return '';
    const diff = Date.now() - ts;
    const minute = 60_000;
    const hour = 60 * minute;
    const day = 24 * hour;
    if (diff < minute) return '剛剛';
    if (diff < hour) return `${Math.floor(diff / minute)} 分鐘前`;
    if (diff < day) return `${Math.floor(diff / hour)} 小時前`;
    if (diff < 30 * day) return `${Math.floor(diff / day)} 天前`;
    const d = new Date(ts);
    return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

function escapeHtml(s: string): string {
    return s.replace(/[&<>"']/g, (ch) =>
        ch === '&' ? '&amp;' : ch === '<' ? '&lt;' : ch === '>' ? '&gt;' : ch === '"' ? '&quot;' : '&#39;'
    );
}
