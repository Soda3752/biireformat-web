import { icon } from './icons';

const THEME_KEY = 'biireformat-theme';

export type Theme = 'light' | 'dark';

export function getTheme(): Theme {
  const stored = localStorage.getItem(THEME_KEY);
  if (stored === 'light' || stored === 'dark') return stored;
  return 'light';
}

export function applyTheme(theme: Theme): void {
  document.documentElement.dataset.theme = theme;
  localStorage.setItem(THEME_KEY, theme);
}

export function bindThemeToggle(button: HTMLButtonElement): void {
  const update = () => {
    const current = getTheme();
    button.innerHTML = icon(current === 'light' ? 'moon' : 'sun');
    button.setAttribute('aria-label', current === 'light' ? '切換暗色主題' : '切換淺色主題');
  };

  applyTheme(getTheme());
  update();

  button.addEventListener('click', () => {
    const next: Theme = getTheme() === 'light' ? 'dark' : 'light';
    applyTheme(next);
    update();
  });
}
