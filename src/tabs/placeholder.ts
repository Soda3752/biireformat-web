import type { TabDefinition } from '@/ui/tabs';

export function renderPlaceholderPanel(tab: TabDefinition): HTMLElement {
  const panel = document.createElement('section');
  panel.className = 'tab-panel';
  panel.dataset.tabId = tab.id;
  panel.setAttribute('role', 'tabpanel');
  panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">${tab.label}</h1>
        <p class="card-subtitle">這個分頁將在後續階段（P2~P4）實作</p>
      </header>
      <div class="placeholder-state">
        <p>尚未實作</p>
        <p class="muted">P1 階段僅建立 UI 骨架</p>
      </div>
    </div>
  `;
  return panel;
}
