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
  { id: 'daily', label: '單日數量', icon: 'calendar', hash: '#daily' },
  {id: 'analytics', label: '數據分析', icon: 'chart', hash: '#analytics', hidden: true},
    {id: 'settings', label: '設定', icon: 'settings', hash: '#settings'},
] as const;

const DEFAULT_TAB_ID = 'bill';

const revealedHiddenTabs = new Set<string>();

function isTabVisible(tab: TabDefinition): boolean {
  return !tab.hidden || revealedHiddenTabs.has(tab.id);
}

export function revealHiddenTab(tabId: string): boolean {
  const tab = TABS.find((t) => t.id === tabId);
  if (!tab || !tab.hidden || revealedHiddenTabs.has(tabId)) return false;
  revealedHiddenTabs.add(tabId);
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
