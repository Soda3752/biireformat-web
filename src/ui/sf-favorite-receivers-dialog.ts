import {
    deleteFavoriteReceiver,
    type FavoriteReceiver,
    listFavoriteReceivers,
    updateFavoriteReceiver,
} from '@/infra/sf-favorite-receivers';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';

export interface FavoriteReceiversDialogOptions {
    onApply?: (receiver: FavoriteReceiver) => void;
}

export function openFavoriteReceiversDialog(
    options: FavoriteReceiversDialogOptions = {}
): void {
    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal sf-fav-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'sf-fav-dialog-title');

    dialog.innerHTML = `
    <header class="app-modal-header">
      <h2 id="sf-fav-dialog-title" class="app-modal-title">收件人最愛</h2>
      <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
    </header>
    <div class="app-modal-body sf-fav-body">
      <div class="sf-fav-toolbar">
        <input class="app-form-input sf-fav-search" type="search" placeholder="搜尋姓名 / 手機 / 地址" data-role="search" autocomplete="off">
      </div>
      <div class="sf-fav-list" data-role="list"></div>
      <div class="sf-fav-empty" data-role="empty" hidden>尚未收藏任何收件人</div>
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

    let editingId: string | null = null;
    let keyword = '';

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
    dialog
        .querySelector<HTMLButtonElement>('[data-role="close"]')!
        .addEventListener('click', close);
    dialog
        .querySelector<HTMLButtonElement>('[data-role="cancel"]')!
        .addEventListener('click', close);

    searchInput.addEventListener('input', () => {
        keyword = searchInput.value.trim().toLowerCase();
        rerender();
    });

    function rerender(): void {
        const all = listFavoriteReceivers();
        const filtered = keyword
            ? all.filter(
                (r) =>
                    r.name.toLowerCase().includes(keyword) ||
                    r.phone.toLowerCase().includes(keyword) ||
                    r.address.toLowerCase().includes(keyword)
            )
            : all;

        if (all.length === 0) {
            listEl.innerHTML = '';
            emptyEl.hidden = false;
            emptyEl.textContent = '尚未收藏任何收件人';
            return;
        }
        if (filtered.length === 0) {
            listEl.innerHTML = '';
            emptyEl.hidden = false;
            emptyEl.textContent = `沒有符合「${keyword}」的收件人`;
            return;
        }
        emptyEl.hidden = true;
        listEl.innerHTML = filtered.map((r) => renderRow(r, r.id === editingId)).join('');

        listEl.querySelectorAll<HTMLElement>('[data-row]').forEach((rowEl) => {
            const id = rowEl.dataset.row!;
            rowEl
                .querySelector<HTMLButtonElement>('[data-role="apply"]')
                ?.addEventListener('click', () => {
                    const target = listFavoriteReceivers().find((x) => x.id === id);
                    if (!target) return;
                    options.onApply?.(target);
                    showToast({
                        variant: 'success',
                        title: '已套用',
                        message: `${target.name}　${target.phone}`,
                    });
                    close();
                });
            rowEl
                .querySelector<HTMLButtonElement>('[data-role="edit"]')
                ?.addEventListener('click', () => {
                    editingId = id;
                    rerender();
                });
            rowEl
                .querySelector<HTMLButtonElement>('[data-role="delete"]')
                ?.addEventListener('click', () => {
                    const target = listFavoriteReceivers().find((x) => x.id === id);
                    if (!target) return;
                    if (!confirm(`確定要刪除「${target.name} ${target.phone}」？`)) return;
                    deleteFavoriteReceiver(id);
                    showToast({
                        variant: 'success',
                        title: '已刪除',
                        message: `${target.name}　${target.phone}`,
                    });
                    rerender();
                });
            rowEl
                .querySelector<HTMLButtonElement>('[data-role="save-edit"]')
                ?.addEventListener('click', () => {
                    const nameEl = rowEl.querySelector<HTMLInputElement>('[data-edit="name"]')!;
                    const phoneEl = rowEl.querySelector<HTMLInputElement>('[data-edit="phone"]')!;
                    const addressEl = rowEl.querySelector<HTMLInputElement>('[data-edit="address"]')!;
                    const name = nameEl.value.trim();
                    const phone = phoneEl.value.trim();
                    const address = addressEl.value.trim();
                    if (!name || !phone || !address) {
                        showToast({
                            variant: 'warning',
                            title: '欄位不可空白',
                            message: '姓名、手機、地址皆需填寫',
                        });
                        return;
                    }
                    updateFavoriteReceiver(id, {name, phone, address});
                    editingId = null;
                    showToast({
                        variant: 'success',
                        title: '已儲存',
                        message: `${name}　${phone}`,
                    });
                    rerender();
                });
            rowEl
                .querySelector<HTMLButtonElement>('[data-role="cancel-edit"]')
                ?.addEventListener('click', () => {
                    editingId = null;
                    rerender();
                });
        });
    }

    function renderRow(r: FavoriteReceiver, editing: boolean): string {
        if (editing) {
            return `
        <div class="sf-fav-row sf-fav-row--editing" data-row="${r.id}">
          <div class="sf-fav-edit-grid">
            <input class="app-form-input" type="text" data-edit="name" value="${escapeAttr(r.name)}" placeholder="姓名">
            <input class="app-form-input" type="text" data-edit="phone" value="${escapeAttr(r.phone)}" placeholder="手機">
            <input class="app-form-input sf-fav-edit-address" type="text" data-edit="address" value="${escapeAttr(r.address)}" placeholder="地址">
          </div>
          <div class="sf-fav-row-actions">
            <button type="button" class="btn btn-secondary btn-sm" data-role="cancel-edit">取消</button>
            <button type="button" class="btn btn-primary btn-sm" data-role="save-edit">儲存</button>
          </div>
        </div>
      `;
        }
        return `
      <div class="sf-fav-row" data-row="${r.id}">
        <div class="sf-fav-row-info">
          <div class="sf-fav-row-name">${escapeHtml(r.name)}</div>
          <div class="sf-fav-row-meta">${escapeHtml(r.phone)}　·　${escapeHtml(r.address)}</div>
        </div>
        <div class="sf-fav-row-actions">
          <button type="button" class="btn btn-primary btn-sm" data-role="apply">套用</button>
          <button type="button" class="icon-btn" data-role="edit" title="編輯">${icon('document', 14)}</button>
          <button type="button" class="icon-btn sf-fav-row-delete" data-role="delete" title="刪除">${icon('trash', 14)}</button>
        </div>
      </div>
    `;
    }

    rerender();
    queueMicrotask(() => searchInput.focus());
}

function escapeAttr(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}

function escapeHtml(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;');
}
