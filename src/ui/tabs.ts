import {icon, type IconName} from './icons';

export interface TabDefinition {
  id: string;
  label: string;
  icon: IconName;
  hash: string;
  hidden?: boolean;
}

export const TABS: readonly TabDefinition[] = [
  { id: 'bill', label: '帳單', icon: 'receipt', hash: '#bill' },
  { id: 'overview', label: '明細', icon: 'list', hash: '#overview' },
  { id: 'delivery', label: '代送費', icon: 'truck', hash: '#delivery' },
  { id: 'bank', label: '對帳', icon: 'bank', hash: '#bank' },
    {id: 'bank-v2', label: '對帳 2.0', icon: 'bank', hash: '#bank-v2'},
    {id: 'daily', label: '單日數量', icon: 'clipboard-list', hash: '#daily'},
    {id: 'handfill', label: '生成手填本', icon: 'document', hash: '#handfill'},
    {id: 'sf-shipping', label: '順豐託運單', icon: 'truck', hash: '#sf-shipping'},
  {id: 'analytics', label: '數據分析', icon: 'chart', hash: '#analytics', hidden: true},
    {
        id: 'cross-month-analytics',
        label: '跨月數據分析',
        icon: 'calendar',
        hash: '#cross-month-analytics',
        hidden: true
    },
    {id: 'settings', label: '設定', icon: 'settings', hash: '#settings'},
] as const;

const DEFAULT_TAB_ID = 'bill';

const REVEALED_STORAGE_KEY = 'revealed-hidden-tabs';

const revealedHiddenTabs = loadRevealedFromStorage();

function loadRevealedFromStorage(): Set<string> {
  try {
    const raw = localStorage.getItem(REVEALED_STORAGE_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    const validIds = new Set(TABS.filter((t) => t.hidden).map((t) => t.id));
    return new Set(
      Array.isArray(arr) ? arr.filter((x) => typeof x === 'string' && validIds.has(x)) : []
    );
  } catch {
    return new Set();
  }
}

function persistRevealed(): void {
  try {
    localStorage.setItem(REVEALED_STORAGE_KEY, JSON.stringify([...revealedHiddenTabs]));
  } catch (err) {
    console.warn('[tabs] persist revealed hidden tabs failed', err);
  }
}

function isTabVisible(tab: TabDefinition): boolean {
  return !tab.hidden || revealedHiddenTabs.has(tab.id);
}

export function revealHiddenTab(tabId: string): boolean {
  const tab = TABS.find((t) => t.id === tabId);
  if (!tab || !tab.hidden || revealedHiddenTabs.has(tabId)) return false;
  revealedHiddenTabs.add(tabId);
  persistRevealed();
  return true;
}

export function hideHiddenTab(tabId: string): boolean {
  const tab = TABS.find((t) => t.id === tabId);
  if (!tab || !tab.hidden || !revealedHiddenTabs.has(tabId)) return false;
  revealedHiddenTabs.delete(tabId);
  persistRevealed();
  return true;
}

export function isHiddenTabRevealed(tabId: string): boolean {
  return revealedHiddenTabs.has(tabId);
}

export function renderSideNav(host: HTMLElement, onSelect: (tabId: string) => void): void {
  host.innerHTML = '';
  for (const tab of TABS) {
    if (!isTabVisible(tab)) continue;
    const item = document.createElement('button');
    item.type = 'button';
    item.className = 'sidenav-item';
    item.setAttribute('role', 'tab');
    item.setAttribute('aria-selected', 'false');
    item.dataset.tabId = tab.id;
    item.innerHTML = `
      <span class="sidenav-item-icon">${icon(tab.icon, 18)}</span>
      <span>${tab.label}</span>
    `;
    item.addEventListener('click', () => onSelect(tab.id));
    host.appendChild(item);
  }
}

export function setActiveTab(navHost: HTMLElement, panelHost: HTMLElement, tabId: string): void {
  navHost.querySelectorAll<HTMLElement>('.sidenav-item').forEach((el) => {
    el.setAttribute('aria-selected', el.dataset.tabId === tabId ? 'true' : 'false');
  });
  panelHost.querySelectorAll<HTMLElement>('.tab-panel').forEach((el) => {
    el.classList.toggle('is-active', el.dataset.tabId === tabId);
  });
}

export function resolveInitialTab(): string {
  const hash = window.location.hash;
  const found = TABS.find((t) => t.hash === hash);
  if (found && isTabVisible(found)) return found.id;
  return DEFAULT_TAB_ID;
}

export function tabHash(tabId: string): string {
  return TABS.find((t) => t.id === tabId)?.hash ?? `#${DEFAULT_TAB_ID}`;
}

/**
 * 切到「設定」分頁並開啟指定子分頁（如 'cargo' / 'daily'）。
 * 設定面板於啟動時已渲染並常駐 DOM，故設好 hash 後直接點擊對應子分頁按鈕即可。
 */
export function jumpToSettingsSubtab(subtab: string): void {
  window.location.hash = '#settings';
  const btn = document.querySelector<HTMLButtonElement>(
    `.settings-subnav-item[data-subtab="${subtab}"]`,
  );
  btn?.click();
}

/** 產生「前往設定」跳轉按鈕 HTML；點擊由 main.ts 的全域委派統一處理。 */
export function settingsJumpButtonHtml(subtab: string, label: string): string {
  return `<button type="button" class="btn btn-secondary btn-sm settings-jump-btn" data-settings-jump="${subtab}">
      <span class="btn-icon">${icon('settings', 14)}</span>前往設定：${label}
    </button>`;
}
