import './ui/theme.css';

import {icon} from '@/ui/icons';
import {bindThemeToggle} from '@/ui/theme';
import {showToast} from '@/ui/toast';
import {renderSideNav, resolveInitialTab, setActiveTab, tabHash, TABS,} from '@/ui/tabs';
import {renderPlaceholderPanel} from '@/tabs/placeholder';
import {renderBillReformatPanel} from '@/tabs/bill-reformat';
import {renderBillOverviewPanel} from '@/tabs/bill-overview';
import {renderDeliveryFeePanel} from '@/tabs/delivery-fee';
import {renderBankNameFormatPanel} from '@/tabs/bank-name-format';
import {renderDailyCountPanel} from '@/tabs/daily-count';
import {renderSettingsPanel} from '@/tabs/settings';
import {loadSortingList} from '@/domain/sorting-list';

const APP_VERSION = __APP_VERSION__;

function bootstrap(): void {
  const app = document.getElementById('app');
  if (!app) throw new Error('#app root not found');

  app.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div class="topbar-brand">
          <span class="topbar-brand-icon">${icon('receipt', 18)}</span>
          <span>青坊食品行 帳單處理工具</span>
          <span class="topbar-version">v${APP_VERSION}</span>
        </div>
        <div class="topbar-actions">
          <button type="button" class="icon-btn" id="theme-toggle"></button>
        </div>
      </header>
      <nav class="sidenav" role="tablist" aria-label="功能分頁"></nav>
      <main class="main" id="main"></main>
    </div>
  `;

  const themeBtn = app.querySelector<HTMLButtonElement>('#theme-toggle');
  if (themeBtn) bindThemeToggle(themeBtn);

  const navHost = app.querySelector<HTMLElement>('.sidenav');
  const mainHost = app.querySelector<HTMLElement>('.main');
  if (!navHost || !mainHost) throw new Error('UI shell not constructed');

  for (const tab of TABS) {
    if (tab.id === 'bill') {
        mainHost.appendChild(renderBillReformatPanel(tab));
    } else if (tab.id === 'overview') {
        mainHost.appendChild(renderBillOverviewPanel(tab));
    } else if (tab.id === 'delivery') {
      mainHost.appendChild(renderDeliveryFeePanel());
    } else if (tab.id === 'bank') {
      mainHost.appendChild(renderBankNameFormatPanel(tab));
    } else if (tab.id === 'daily') {
      mainHost.appendChild(renderDailyCountPanel(tab));
    } else if (tab.id === 'settings') {
        mainHost.appendChild(renderSettingsPanel(tab));
    } else {
      mainHost.appendChild(renderPlaceholderPanel(tab));
    }
  }

  // 啟動時非同步載入 SortingList，失敗不擋 UI（只跳 toast 提醒）
  loadSortingList().catch((err) => {
    console.error('SortingList 載入失敗', err);
    showToast({
      variant: 'error',
      title: '排序資料載入失敗',
      message: err instanceof Error ? err.message : String(err),
    });
  });

  const switchTab = (tabId: string) => {
    setActiveTab(navHost, mainHost, tabId);
    const nextHash = tabHash(tabId);
    if (window.location.hash !== nextHash) {
      window.history.replaceState(null, '', nextHash);
    }
  };

  renderSideNav(navHost, switchTab);
  switchTab(resolveInitialTab());

  window.addEventListener('hashchange', () => {
    switchTab(resolveInitialTab());
  });
}

bootstrap();
