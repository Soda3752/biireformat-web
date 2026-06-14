import './ui/theme.css';

import {icon} from '@/ui/icons';
import {bindThemeToggle} from '@/ui/theme';
import {showToast} from '@/ui/toast';
import {
    hideHiddenTab,
    isHiddenTabRevealed,
    jumpToSettingsSubtab,
    renderSideNav,
    resolveInitialTab,
    revealHiddenTab,
    setActiveTab,
    tabHash,
    TABS,
} from '@/ui/tabs';
import {renderPlaceholderPanel} from '@/tabs/placeholder';
import {renderBillReformatPanel} from '@/tabs/bill-reformat';
import {renderBillOverviewPanel} from '@/tabs/bill-overview';
import {renderDeliveryFeePanel} from '@/tabs/delivery-fee';
import {renderBankNameFormatPanel} from '@/tabs/bank-name-format';
import {renderBankNameFormatV2Panel} from '@/tabs/bank-name-format-v2';
import {renderDailyCountPanel} from '@/tabs/daily-count';
import {renderHandfillPanel} from '@/tabs/handfill';
import {renderSfShippingPanel} from '@/tabs/sf-shipping';
import {renderDataAnalyticsPanel} from '@/tabs/data-analytics';
import {renderCrossMonthAnalyticsPanel} from '@/tabs/cross-month-analytics';
import {hideCostColumn, isCostColumnRevealed, renderSettingsPanel, revealCostColumn,} from '@/tabs/settings';
import {loadSortingList} from '@/domain/sorting-list';

const APP_VERSION = __APP_VERSION__;

function bootstrap(): void {
  const app = document.getElementById('app');
  if (!app) throw new Error('#app root not found');

  app.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div class="topbar-brand">
          <span class="topbar-brand-icon" id="brand-icon" role="button" tabindex="-1" aria-hidden="true">${icon('receipt', 18)}</span>
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
    } else if (tab.id === 'bank-v2') {
        mainHost.appendChild(renderBankNameFormatV2Panel(tab));
    } else if (tab.id === 'daily') {
      mainHost.appendChild(renderDailyCountPanel(tab));
    } else if (tab.id === 'handfill') {
        mainHost.appendChild(renderHandfillPanel(tab));
    } else if (tab.id === 'sf-shipping') {
        mainHost.appendChild(renderSfShippingPanel(tab));
    } else if (tab.id === 'analytics') {
        mainHost.appendChild(renderDataAnalyticsPanel(tab));
    } else if (tab.id === 'cross-month-analytics') {
        mainHost.appendChild(renderCrossMonthAnalyticsPanel(tab));
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

  // 各功能頁的「前往設定」跳轉按鈕（data-settings-jump）統一委派處理
  app.addEventListener('click', (e) => {
    const target = (e.target as HTMLElement).closest<HTMLElement>('[data-settings-jump]');
    if (target) jumpToSettingsSubtab(target.dataset.settingsJump!);
  });

    bindHiddenTabUnlock(app, navHost, switchTab);
}

function bindHiddenTabUnlock(
    app: HTMLElement,
    navHost: HTMLElement,
    switchTab: (tabId: string) => void,
): void {
    const brandIcon = app.querySelector<HTMLElement>('#brand-icon');
    if (!brandIcon) return;
    brandIcon.style.cursor = 'default';

    const REQUIRED_CLICKS = 10;
    const WINDOW_MS = 5000;
    let count = 0;
    let firstClickAt = 0;

    brandIcon.addEventListener('click', () => {
        const now = Date.now();
        if (count === 0 || now - firstClickAt > WINDOW_MS) {
            count = 1;
            firstClickAt = now;
            return;
        }
        count += 1;
        if (count < REQUIRED_CLICKS) return;
        count = 0;
        firstClickAt = 0;

        const allRevealed =
            isHiddenTabRevealed('analytics') &&
            isHiddenTabRevealed('cross-month-analytics') &&
            isCostColumnRevealed();

        if (allRevealed) {
            // 已全部解鎖 → 連點再次收起
            hideHiddenTab('analytics');
            hideHiddenTab('cross-month-analytics');
            hideCostColumn();
            renderSideNav(navHost, switchTab);
            // 若使用者目前停在被收起的分頁，切回預設分頁
            const currentHash = window.location.hash;
            if (currentHash === '#analytics' || currentHash === '#cross-month-analytics') {
                switchTab(resolveInitialTab());
            }
            showToast({
                variant: 'success',
                title: '已收起隱藏功能',
                message: '「數據分析」「跨月數據分析」與「成本」欄位已隱藏',
            });
            return;
        }

        const analyticsRevealed = revealHiddenTab('analytics');
        const crossMonthRevealed = revealHiddenTab('cross-month-analytics');
        const costRevealed = revealCostColumn();
        if (analyticsRevealed || crossMonthRevealed) {
            renderSideNav(navHost, switchTab);
            if (analyticsRevealed) {
                switchTab('analytics');
            }
        } else if (costRevealed) {
            // 數據分析已解鎖過，這次只解鎖成本欄位 → 提示使用者
            showToast({
                variant: 'success',
                title: '已解鎖隱藏設定',
                message: '「單日數量」現在會顯示「成本」欄位',
            });
        }
    });
}

bootstrap();
