import {
    DEFAULT_SF_SETTINGS,
    loadSfSettings,
    resetSfSettings,
    saveSfSettings,
    type SfShippingSettings,
} from '@/infra/sf-shipping-settings';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';

export interface SfShippingSettingsDialogOptions {
    onSaved?: (settings: SfShippingSettings) => void;
}

interface FieldDef {
    key: keyof SfShippingSettings;
    label: string;
    type: 'text' | 'number';
    group: string;
}

const FIELDS: ReadonlyArray<FieldDef> = [
    {key: 'shipperName', label: '寄件方姓名', type: 'text', group: '寄件方'},
    {key: 'shipperPhone', label: '寄件方手機', type: 'text', group: '寄件方'},
    {key: 'shipperAddress', label: '寄件方詳細地址', type: 'text', group: '寄件方'},
    {key: 'shipperCity', label: '寄件方城市 (G)', type: 'text', group: '寄件方'},
    {key: 'shipperState', label: '寄件方州/省 (H)', type: 'text', group: '寄件方'},
    {key: 'shipperCountry', label: '寄件方國家 (I)', type: 'text', group: '寄件方'},
    {key: 'shipperZip', label: '寄件方郵編 (J)', type: 'text', group: '寄件方'},
    {key: 'shipperType', label: '寄件類型 (L)', type: 'text', group: '寄件方'},
    {key: 'shipperCompany', label: '寄件方公司 (M)', type: 'text', group: '寄件方'},
    {key: 'receiverCity', label: '收件方城市 (U)', type: 'text', group: '收件方預設'},
    {key: 'receiverState', label: '收件方州/省 (V)', type: 'text', group: '收件方預設'},
    {key: 'receiverCountry', label: '收件方國家 (W)', type: 'text', group: '收件方預設'},
    {key: 'receiverZip', label: '收件方郵編 (X)', type: 'text', group: '收件方預設'},
    {key: 'productName', label: '商品名稱預設', type: 'text', group: '商品預設'},
    {key: 'productQty', label: '商品數量預設', type: 'number', group: '商品預設'},
    {key: 'productUnit', label: '單位 (AD)', type: 'text', group: '商品預設'},
    {key: 'productPrice', label: '商品單價預設', type: 'number', group: '商品預設'},
    {key: 'parcelCount', label: '包裹件數預設', type: 'number', group: '商品預設'},
    {key: 'totalWeight', label: '總重量預設', type: 'number', group: '商品預設'},
    {key: 'currency', label: '商品貨幣 (AL)', type: 'text', group: '運送資訊'},
    {key: 'expressType', label: '快件類型 (AM)', type: 'text', group: '運送資訊'},
    {key: 'pickupMethod', label: '寄件方式 (BC)', type: 'text', group: '運送資訊'},
    {key: 'paymentMethod', label: '付款方式 (BF)', type: 'text', group: '運送資訊'},
    {key: 'monthlyCardNo', label: '月結卡號 (BG)', type: 'text', group: '運送資訊'},
];

export function openSfShippingSettingsDialog(
    options: SfShippingSettingsDialogOptions = {}
): void {
    const current = loadSfSettings();

    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');

    const dialog = document.createElement('div');
    dialog.className = 'app-modal sf-settings-modal';
    dialog.setAttribute('role', 'dialog');
    dialog.setAttribute('aria-modal', 'true');
    dialog.setAttribute('aria-labelledby', 'sf-settings-dialog-title');

    const groups = groupFields();

    dialog.innerHTML = `
    <header class="app-modal-header">
      <h2 id="sf-settings-dialog-title" class="app-modal-title">順豐託運單 — 預設值設定</h2>
      <button type="button" class="app-modal-close" aria-label="關閉" data-role="close">${icon('close', 16)}</button>
    </header>
    <div class="app-modal-body sf-settings-body">
      ${groups
        .map(
            (g) => `
        <section class="sf-settings-group">
          <h3 class="sf-settings-group-title">${g.title}</h3>
          <div class="sf-settings-grid">
            ${g.fields
                .map((f) => {
                    const v = current[f.key];
                    return `
                <div class="app-form-row">
                  <label class="app-form-label" for="sfset-${f.key}">${f.label}</label>
                  <input id="sfset-${f.key}" data-key="${f.key}" data-type="${f.type}"
                         class="app-form-input"
                         type="${f.type === 'number' ? 'number' : 'text'}"
                         ${f.type === 'number' ? 'step="any"' : ''}
                         value="${escapeAttr(String(v))}"
                         autocomplete="off" />
                </div>`;
                })
                .join('')}
          </div>
        </section>`
        )
        .join('')}
    </div>
    <footer class="app-modal-footer">
      <button type="button" class="btn btn-secondary" data-role="reset">重設為預設</button>
      <span class="sf-settings-footer-spacer"></span>
      <button type="button" class="btn btn-secondary" data-role="cancel">取消</button>
      <button type="button" class="btn btn-primary" data-role="save">儲存</button>
    </footer>
  `;

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

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

    dialog
        .querySelector<HTMLButtonElement>('[data-role="reset"]')!
        .addEventListener('click', () => {
            if (!confirm('將所有欄位重設為預設值？')) return;
            resetSfSettings();
            const defaults = {...DEFAULT_SF_SETTINGS};
            dialog.querySelectorAll<HTMLInputElement>('input[data-key]').forEach((el) => {
                const key = el.dataset.key as keyof SfShippingSettings;
                el.value = String(defaults[key]);
            });
            showToast({
                variant: 'success',
                title: '已重設為預設值',
                message: '尚未儲存，可繼續修改或點儲存。',
            });
        });

    dialog
        .querySelector<HTMLButtonElement>('[data-role="save"]')!
        .addEventListener('click', () => {
            const next: SfShippingSettings = {...current};
            dialog.querySelectorAll<HTMLInputElement>('input[data-key]').forEach((el) => {
                const key = el.dataset.key as keyof SfShippingSettings;
                const type = el.dataset.type;
                if (type === 'number') {
                    const n = Number(el.value);
                    (next[key] as number) = Number.isFinite(n) ? n : 0;
                } else {
                    (next[key] as string) = el.value;
                }
            });
            saveSfSettings(next);
            showToast({
                variant: 'success',
                title: '已儲存設定',
                message: '下次產生託運單時會自動套用',
            });
            options.onSaved?.(next);
            close();
        });
}

interface FieldGroup {
    title: string;
    fields: ReadonlyArray<FieldDef>;
}

function groupFields(): FieldGroup[] {
    const order = ['寄件方', '收件方預設', '商品預設', '運送資訊'];
    return order.map((title) => ({
        title,
        fields: FIELDS.filter((f) => f.group === title),
    }));
}

function escapeAttr(value: string): string {
    return value
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
}
