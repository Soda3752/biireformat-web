import {saveAs} from 'file-saver';

import {addFavoriteReceiver, findFavoriteByKey,} from '@/infra/sf-favorite-receivers';
import {loadSfSettings, type SfShippingSettings,} from '@/infra/sf-shipping-settings';
import {buildOrderNumbers, buildSfShippingWorkbook, type SfShippingOrder,} from '@/writers/sf-shipping-writer';
import {openFavoriteReceiversDialog} from '@/ui/sf-favorite-receivers-dialog';
import {icon} from '@/ui/icons';
import {openSfShippingSettingsDialog} from '@/ui/sf-shipping-settings-dialog';
import type {TabDefinition} from '@/ui/tabs';
import {showToast} from '@/ui/toast';

interface OrderRow {
    productName: string;
    productQty: number;
    productPrice: number;
    parcelCount: number;
    totalWeight: number;
}

export function renderSfShippingPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel sf-shipping-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card">
      <header class="card-header sf-card-header">
        <div>
          <h1 class="card-title">順豐託運單</h1>
          <p class="card-subtitle">
            為單一收件人產生一份順豐託運單 .xlsx，保留模板的格式（標題列、下拉、隱藏設定表）。
          </p>
        </div>
        <button type="button" class="btn btn-secondary" data-role="settings-btn">
          ${icon('settings', 16)}<span>設定預設值</span>
        </button>
      </header>

      <section class="sf-section">
        <div class="sf-section-toolbar">
          <h2 class="sf-section-title">收件人資訊</h2>
          <div class="sf-receiver-actions">
            <button type="button" class="btn btn-secondary btn-sm" data-role="open-favorites">
              ${icon('list', 14)}<span>從最愛帶入</span>
            </button>
            <button type="button" class="btn btn-secondary btn-sm" data-role="toggle-favorite" disabled>
              <span data-role="fav-star">☆</span><span data-role="fav-label">加入最愛</span>
            </button>
          </div>
        </div>
        <div class="sf-receiver-grid">
          <div class="app-form-row">
            <label class="app-form-label" for="sf-rcv-name">姓名</label>
            <input id="sf-rcv-name" class="app-form-input" type="text" autocomplete="off" placeholder="例：白藝群">
          </div>
          <div class="app-form-row">
            <label class="app-form-label" for="sf-rcv-phone">手機</label>
            <input id="sf-rcv-phone" class="app-form-input" type="text" autocomplete="off" placeholder="例：0937247653">
          </div>
          <div class="app-form-row sf-receiver-address">
            <label class="app-form-label" for="sf-rcv-address">詳細地址</label>
            <input id="sf-rcv-address" class="app-form-input" type="text" autocomplete="off" placeholder="例：台南市中西區和緯路五段89巷116號">
          </div>
        </div>
      </section>

      <section class="sf-section">
        <div class="sf-section-toolbar">
          <h2 class="sf-section-title">訂單清單 <span class="sf-order-count" data-role="order-count">(0)</span></h2>
          <div class="sf-section-toolbar-actions">
            <div class="app-form-row sf-date-row">
              <label class="app-form-label" for="sf-order-date">訂單日期</label>
              <input id="sf-order-date" class="app-form-input" type="date">
            </div>
            <div class="app-form-row sf-seq-row">
              <label class="app-form-label" for="sf-start-seq">起始流水</label>
              <input id="sf-start-seq" class="app-form-input" type="number" min="1" value="1">
            </div>
          </div>
        </div>
        <div class="sf-orders-table-wrap">
          <table class="sf-orders-table">
            <thead>
              <tr>
                <th>#</th>
                <th>訂單號</th>
                <th>商品名稱</th>
                <th>數量</th>
                <th>單價</th>
                <th>件數</th>
                <th>重量</th>
                <th></th>
              </tr>
            </thead>
            <tbody data-role="orders-tbody"></tbody>
          </table>
        </div>
        <button type="button" class="btn btn-secondary sf-add-order" data-role="add-order">
          ${icon('plus', 14)}<span>新增一筆</span>
        </button>
      </section>

      <footer class="action-bar sf-action-bar">
        <div class="sf-action-filename">
          <label class="app-form-label" for="sf-filename">輸出檔名</label>
          <div class="sf-filename-row">
            <input id="sf-filename" class="app-form-input" type="text" placeholder="例：1150504台南白先生">
            <span class="sf-filename-suffix">.xlsx</span>
          </div>
        </div>
        <div class="action-bar-actions">
          <button type="button" class="btn btn-primary btn-lg" data-role="download">
            ${icon('download', 16)}<span>下載託運單</span>
          </button>
        </div>
      </footer>
    </div>
  `;

    // 狀態
    let settings: SfShippingSettings = loadSfSettings();
    const orders: OrderRow[] = [];

    const tbody = panel.querySelector<HTMLTableSectionElement>('[data-role="orders-tbody"]')!;
    const orderCountEl = panel.querySelector<HTMLElement>('[data-role="order-count"]')!;
    const nameInput = panel.querySelector<HTMLInputElement>('#sf-rcv-name')!;
    const phoneInput = panel.querySelector<HTMLInputElement>('#sf-rcv-phone')!;
    const addressInput = panel.querySelector<HTMLInputElement>('#sf-rcv-address')!;
    const dateInput = panel.querySelector<HTMLInputElement>('#sf-order-date')!;
    const seqInput = panel.querySelector<HTMLInputElement>('#sf-start-seq')!;
    const filenameInput = panel.querySelector<HTMLInputElement>('#sf-filename')!;
    const downloadBtn = panel.querySelector<HTMLButtonElement>('[data-role="download"]')!;
    const openFavBtn = panel.querySelector<HTMLButtonElement>('[data-role="open-favorites"]')!;
    const toggleFavBtn = panel.querySelector<HTMLButtonElement>('[data-role="toggle-favorite"]')!;
    const favStarEl = panel.querySelector<HTMLElement>('[data-role="fav-star"]')!;
    const favLabelEl = panel.querySelector<HTMLElement>('[data-role="fav-label"]')!;

    // 預設今日
    dateInput.value = toDateInputValue(new Date());

    // 收件人最愛：載入清單 / 加入最愛
    openFavBtn.addEventListener('click', () => {
        openFavoriteReceiversDialog({
            onApply: (r) => {
                nameInput.value = r.name;
                phoneInput.value = r.phone;
                addressInput.value = r.address;
                refreshFavoriteButton();
            },
        });
    });

    toggleFavBtn.addEventListener('click', () => {
        const name = nameInput.value.trim();
        const phone = phoneInput.value.trim();
        const address = addressInput.value.trim();
        if (!name || !phone || !address) {
            showToast({
                variant: 'warning',
                title: '欄位不可空白',
                message: '姓名、手機、地址都要填好才能加入最愛',
            });
            return;
        }
        const {isNew} = addFavoriteReceiver({name, phone, address});
        showToast({
            variant: 'success',
            title: isNew ? '已加入最愛' : '已更新最愛',
            message: `${name}　${phone}`,
        });
        refreshFavoriteButton();
    });

    [nameInput, phoneInput, addressInput].forEach((el) => {
        el.addEventListener('input', refreshFavoriteButton);
    });

    function refreshFavoriteButton(): void {
        const name = nameInput.value.trim();
        const phone = phoneInput.value.trim();
        const address = addressInput.value.trim();
        const canSave = name.length > 0 && phone.length > 0 && address.length > 0;
        toggleFavBtn.disabled = !canSave;
        const existing = canSave ? findFavoriteByKey(name, phone) : null;
        if (existing) {
            favStarEl.textContent = '★';
            const addressDiffers = existing.address !== address;
            favLabelEl.textContent = addressDiffers ? '更新最愛' : '已加入';
            toggleFavBtn.disabled = !addressDiffers && existing.address === address;
            // 已加入且地址沒變 → 維持 disabled 視覺
            toggleFavBtn.classList.toggle('sf-fav-saved', !addressDiffers);
        } else {
            favStarEl.textContent = '☆';
            favLabelEl.textContent = '加入最愛';
            toggleFavBtn.classList.remove('sf-fav-saved');
        }
    }

    refreshFavoriteButton();

    // 開啟設定面板
    panel
        .querySelector<HTMLButtonElement>('[data-role="settings-btn"]')!
        .addEventListener('click', () => {
            openSfShippingSettingsDialog({
                onSaved: (next) => {
                    settings = next;
                },
            });
        });

    // 新增訂單
    panel
        .querySelector<HTMLButtonElement>('[data-role="add-order"]')!
        .addEventListener('click', () => {
            orders.push(buildOrderFromSettings(settings));
            rerenderTbody();
        });

    dateInput.addEventListener('input', rerenderTbody);
    seqInput.addEventListener('input', rerenderTbody);

    downloadBtn.addEventListener('click', () => {
        void handleDownload();
    });

    function buildOrderFromSettings(s: SfShippingSettings): OrderRow {
        return {
            productName: s.productName,
            productQty: s.productQty,
            productPrice: s.productPrice,
            parcelCount: s.parcelCount,
            totalWeight: s.totalWeight,
        };
    }

    function rerenderTbody(): void {
        orderCountEl.textContent = `(${orders.length})`;
        const orderNumbers = computeOrderNumbers();
        tbody.innerHTML = '';
        orders.forEach((order, idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
        <td class="sf-orders-index">${idx + 1}</td>
        <td class="sf-orders-no">${orderNumbers[idx] ?? '-'}</td>
        <td><input class="app-form-input sf-cell sf-cell-product" type="text" data-field="productName" value="${escapeAttr(order.productName)}"></td>
        <td><input class="app-form-input sf-cell sf-cell-num" type="number" step="any" data-field="productQty" value="${order.productQty}"></td>
        <td><input class="app-form-input sf-cell sf-cell-num" type="number" step="any" data-field="productPrice" value="${order.productPrice}"></td>
        <td><input class="app-form-input sf-cell sf-cell-num" type="number" step="any" data-field="parcelCount" value="${order.parcelCount}"></td>
        <td><input class="app-form-input sf-cell sf-cell-num" type="number" step="any" data-field="totalWeight" value="${order.totalWeight}"></td>
        <td><button type="button" class="icon-btn sf-delete-btn" data-role="delete" title="刪除">${icon('trash', 14)}</button></td>
      `;
            tr.querySelectorAll<HTMLInputElement>('input[data-field]').forEach((el) => {
                el.addEventListener('input', () => {
                    const field = el.dataset.field as keyof OrderRow;
                    if (field === 'productName') {
                        orders[idx].productName = el.value;
                    } else {
                        const n = Number(el.value);
                        (orders[idx][field] as number) = Number.isFinite(n) ? n : 0;
                    }
                });
            });
            tr.querySelector<HTMLButtonElement>('[data-role="delete"]')!.addEventListener(
                'click',
                () => {
                    orders.splice(idx, 1);
                    rerenderTbody();
                }
            );
            tbody.appendChild(tr);
        });
    }

    function computeOrderNumbers(): string[] {
        const date = parseDateInputValue(dateInput.value) ?? new Date();
        const startSeq = Math.max(1, Number(seqInput.value) || 1);
        return buildOrderNumbers(date, orders.length, startSeq);
    }

    async function handleDownload(): Promise<void> {
        const receiverName = nameInput.value.trim();
        const receiverPhone = phoneInput.value.trim();
        const receiverAddress = addressInput.value.trim();
        const filename = filenameInput.value.trim();

        if (orders.length === 0) {
            showToast({variant: 'warning', title: '尚無訂單', message: '請至少新增一筆訂單'});
            return;
        }
        if (!receiverName || !receiverPhone || !receiverAddress) {
            showToast({
                variant: 'warning',
                title: '收件人資訊不完整',
                message: '姓名、手機、地址三項皆為必填',
            });
            return;
        }
        if (!filename) {
            showToast({variant: 'warning', title: '請填寫檔名', message: '輸出檔名不可空白'});
            filenameInput.focus();
            return;
        }

        const orderNumbers = computeOrderNumbers();
        const orderPayload: SfShippingOrder[] = orders.map((o, i) => ({
            orderNo: orderNumbers[i],
            productName: o.productName,
            productQty: o.productQty,
            productPrice: o.productPrice,
            parcelCount: o.parcelCount,
            totalWeight: o.totalWeight,
        }));

        downloadBtn.disabled = true;
        const original = downloadBtn.innerHTML;
        downloadBtn.textContent = '產生中...';
        try {
            const blob = await buildSfShippingWorkbook({
                sheetName: receiverName,
                receiver: {name: receiverName, phone: receiverPhone, address: receiverAddress},
                orders: orderPayload,
                settings,
            });
            saveAs(blob, `${filename}.xlsx`);
            showToast({
                variant: 'success',
                title: '已產生託運單',
                message: `${filename}.xlsx　共 ${orders.length} 筆`,
            });
        } catch (err) {
            console.error('[sf-shipping] download failed', err);
            showToast({
                variant: 'error',
                title: '產生失敗',
                message: err instanceof Error ? err.message : String(err),
            });
        } finally {
            downloadBtn.innerHTML = original;
            downloadBtn.disabled = false;
        }
    }

    // 初始展示空表格
    rerenderTbody();

    return panel;
}

function toDateInputValue(d: Date): string {
    const y = d.getFullYear();
    const m = (d.getMonth() + 1).toString().padStart(2, '0');
    const dd = d.getDate().toString().padStart(2, '0');
    return `${y}-${m}-${dd}`;
}

function parseDateInputValue(s: string): Date | null {
    if (!s) return null;
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (!m) return null;
    return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function escapeAttr(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
