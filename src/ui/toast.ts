import { icon, type IconName } from './icons';

export type ToastVariant = 'success' | 'error' | 'warning' | 'info';

interface ToastOptions {
  title?: string;
  message: string;
  variant?: ToastVariant;
  duration?: number;
}

const VARIANT_ICON: Record<ToastVariant, IconName> = {
  success: 'check',
  error: 'alert',
  warning: 'alert',
  info: 'info',
};

let container: HTMLDivElement | null = null;

function ensureContainer(): HTMLDivElement {
  if (container) return container;
  container = document.createElement('div');
  container.className = 'toast-container';
  container.setAttribute('aria-live', 'polite');
  container.setAttribute('aria-atomic', 'false');
  document.body.appendChild(container);
  return container;
}

export function showToast(options: ToastOptions): void {
  const variant = options.variant ?? 'info';
  const duration = options.duration ?? 4000;

  const root = ensureContainer();
  const el = document.createElement('div');
  el.className = `toast toast-${variant}`;
  el.setAttribute('role', 'status');
  el.innerHTML = `
    <span class="toast-icon">${icon(VARIANT_ICON[variant], 18)}</span>
    <div class="toast-body">
      ${options.title ? `<div class="toast-title">${escapeHtml(options.title)}</div>` : ''}
      <div>${escapeHtml(options.message)}</div>
    </div>
    <button class="toast-close" aria-label="關閉">${icon('close', 14)}</button>
  `;

  const dismiss = () => {
    el.classList.add('is-leaving');
    el.addEventListener('animationend', () => el.remove(), { once: true });
  };

  el.querySelector<HTMLButtonElement>('.toast-close')?.addEventListener('click', dismiss);
  root.appendChild(el);

  if (duration > 0) {
    setTimeout(dismiss, duration);
  }
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
