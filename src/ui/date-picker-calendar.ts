/**
 * 自寫的輕量月曆選擇器。trigger 按鈕點擊後展開 popup，
 * 顯示某月的網格；指定 availableDates 可讓「有資料」的日期亮起，
 * 其餘日期 disabled。鍵盤支援 Esc 關閉。
 */

const WEEKDAY_LABEL = ['日', '一', '二', '三', '四', '五', '六'] as const;
const CHINESE_WEEKDAY = ['日', '一', '二', '三', '四', '五', '六'] as const;

export interface DatePickerOptions {
    availableDates?: ReadonlyArray<string> | ReadonlySet<string>;
    initialDate?: string | null;
    onSelect: (dateKey: string) => void;
    /** 沒有資料的日期是否仍可被選擇。預設 false（disabled）。 */
    allowUnavailable?: boolean;
    emptyHint?: string;
}

export interface DatePickerHandle {
    readonly element: HTMLElement;

    setAvailableDates(dates: ReadonlyArray<string> | ReadonlySet<string>): void;

    setSelected(dateKey: string | null): void;

    getSelected(): string | null;

    destroy(): void;
}

interface InternalState {
    selected: string | null;
    available: Set<string>;
    viewYear: number;
    viewMonth: number; // 0-based
    open: boolean;
}

const today = (): { y: number; m: number; d: number } => {
    const now = new Date();
    return {y: now.getFullYear(), m: now.getMonth(), d: now.getDate()};
};

const formatTriggerLabel = (key: string | null, hint: string): string => {
    if (!key) return hint;
    const [y, m, d] = key.split('-').map((v) => Number.parseInt(v, 10));
    const date = new Date(y, m - 1, d);
    const weekday = CHINESE_WEEKDAY[date.getDay()];
    return `${y}/${String(m).padStart(2, '0')}/${String(d).padStart(2, '0')}（${weekday}）`;
};

const toKey = (y: number, m: number, d: number): string =>
    `${y}-${String(m + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

const parseKey = (key: string): { y: number; m: number; d: number } | null => {
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(key);
    if (!m) return null;
    return {y: Number(m[1]), m: Number(m[2]) - 1, d: Number(m[3])};
};

export const createDateCalendar = (
    options: DatePickerOptions
): DatePickerHandle => {
    const allowUnavailable = options.allowUnavailable ?? false;
    const emptyHint = options.emptyHint ?? '請選擇日期';

    const state: InternalState = {
        selected: options.initialDate ?? null,
        available: new Set(options.availableDates ? [...options.availableDates] : []),
        viewYear: 0,
        viewMonth: 0,
        open: false,
    };

    // 預設展開到 selected 月份；若無 selected，使用第一個 available；最後 fallback today
    const initial = state.selected
        ? parseKey(state.selected)
        : state.available.size > 0
            ? parseKey([...state.available].sort()[0])
            : null;
    if (initial) {
        state.viewYear = initial.y;
        state.viewMonth = initial.m;
    } else {
        const t = today();
        state.viewYear = t.y;
        state.viewMonth = t.m;
    }

    const root = document.createElement('div');
    root.className = 'date-picker';

    const trigger = document.createElement('button');
    trigger.type = 'button';
    trigger.className = 'date-picker-trigger';
    trigger.setAttribute('aria-haspopup', 'dialog');
    trigger.setAttribute('aria-expanded', 'false');
    root.appendChild(trigger);

    const popup = document.createElement('div');
    popup.className = 'date-picker-popup';
    popup.setAttribute('role', 'dialog');
    popup.hidden = true;
    root.appendChild(popup);

    // popup 內部結構：header + grid
    popup.innerHTML = `
    <header class="date-picker-header">
      <button type="button" class="date-picker-nav" data-role="prev" aria-label="上個月">‹</button>
      <div class="date-picker-title" data-role="title"></div>
      <button type="button" class="date-picker-nav" data-role="next" aria-label="下個月">›</button>
    </header>
    <div class="date-picker-weekdays" data-role="weekdays"></div>
    <div class="date-picker-grid" data-role="grid"></div>
  `;

    const titleEl = popup.querySelector<HTMLElement>('[data-role="title"]')!;
    const gridEl = popup.querySelector<HTMLElement>('[data-role="grid"]')!;
    const weekdaysEl = popup.querySelector<HTMLElement>('[data-role="weekdays"]')!;
    const prevBtn = popup.querySelector<HTMLButtonElement>('[data-role="prev"]')!;
    const nextBtn = popup.querySelector<HTMLButtonElement>('[data-role="next"]')!;

    for (const w of WEEKDAY_LABEL) {
        const cell = document.createElement('span');
        cell.className = 'date-picker-weekday';
        cell.textContent = w;
        weekdaysEl.appendChild(cell);
    }

    const renderTrigger = () => {
        trigger.textContent = formatTriggerLabel(state.selected, emptyHint);
        trigger.classList.toggle('is-empty', state.selected == null);
    };

    const renderGrid = () => {
        titleEl.textContent = `${state.viewYear} 年 ${state.viewMonth + 1} 月`;
        gridEl.innerHTML = '';

        const firstDay = new Date(state.viewYear, state.viewMonth, 1);
        const lastDay = new Date(state.viewYear, state.viewMonth + 1, 0);
        const daysInMonth = lastDay.getDate();
        const leading = firstDay.getDay(); // 0=Sun

        // 前導空白：填上一個月尾巴的日期，視為 disabled out-of-month
        const prevLastDay = new Date(state.viewYear, state.viewMonth, 0).getDate();
        for (let i = leading - 1; i >= 0; i--) {
            const d = prevLastDay - i;
            const y = state.viewMonth === 0 ? state.viewYear - 1 : state.viewYear;
            const m = state.viewMonth === 0 ? 11 : state.viewMonth - 1;
            gridEl.appendChild(buildCell(y, m, d, true));
        }

        // 當月日期
        for (let d = 1; d <= daysInMonth; d++) {
            gridEl.appendChild(buildCell(state.viewYear, state.viewMonth, d, false));
        }

        // 補滿到 6 列 (42 格) — 視覺一致
        const totalCells = leading + daysInMonth;
        const trailing = (7 - (totalCells % 7)) % 7;
        for (let i = 1; i <= trailing; i++) {
            const y = state.viewMonth === 11 ? state.viewYear + 1 : state.viewYear;
            const m = state.viewMonth === 11 ? 0 : state.viewMonth + 1;
            gridEl.appendChild(buildCell(y, m, i, true));
        }
    };

    const buildCell = (y: number, m: number, d: number, outOfMonth: boolean): HTMLElement => {
        const key = toKey(y, m, d);
        const has = state.available.has(key);
        const cell = document.createElement('button');
        cell.type = 'button';
        cell.className = 'date-picker-cell';
        cell.textContent = String(d);
        cell.dataset.key = key;
        if (outOfMonth) cell.classList.add('is-out');
        if (has) cell.classList.add('has-data');
        if (state.selected === key) cell.classList.add('is-selected');

        const t = today();
        if (y === t.y && m === t.m && d === t.d) cell.classList.add('is-today');

        const clickable = has || (allowUnavailable && !outOfMonth);
        cell.disabled = !clickable;

        if (clickable) {
            cell.addEventListener('click', () => {
                state.selected = key;
                state.viewYear = y;
                state.viewMonth = m;
                renderTrigger();
                renderGrid();
                closePopup();
                options.onSelect(key);
            });
        }

        return cell;
    };

    const openPopup = () => {
        if (state.open) return;
        state.open = true;
        popup.hidden = false;
        trigger.setAttribute('aria-expanded', 'true');
        renderGrid();
        document.addEventListener('mousedown', onOutside, true);
        document.addEventListener('keydown', onKeydown, true);
    };

    const closePopup = () => {
        if (!state.open) return;
        state.open = false;
        popup.hidden = true;
        trigger.setAttribute('aria-expanded', 'false');
        document.removeEventListener('mousedown', onOutside, true);
        document.removeEventListener('keydown', onKeydown, true);
    };

    const onOutside = (e: MouseEvent) => {
        if (!root.contains(e.target as Node)) closePopup();
    };

    const onKeydown = (e: KeyboardEvent) => {
        if (e.key === 'Escape') {
            closePopup();
            trigger.focus();
        }
    };

    trigger.addEventListener('click', () => {
        if (state.open) closePopup();
        else openPopup();
    });
    prevBtn.addEventListener('click', () => {
        if (state.viewMonth === 0) {
            state.viewMonth = 11;
            state.viewYear -= 1;
        } else {
            state.viewMonth -= 1;
        }
        renderGrid();
    });
    nextBtn.addEventListener('click', () => {
        if (state.viewMonth === 11) {
            state.viewMonth = 0;
            state.viewYear += 1;
        } else {
            state.viewMonth += 1;
        }
        renderGrid();
    });

    renderTrigger();

    return {
        element: root,
        setAvailableDates(dates) {
            state.available = new Set([...dates]);
            // 如果目前選中的日期不在新的 available 集合中且不允許 unavailable，就清掉
            if (state.selected && !state.available.has(state.selected) && !allowUnavailable) {
                state.selected = null;
                renderTrigger();
            }
            // 如果有 available 但沒選任何日期，把視圖切到最早一筆
            if (!state.selected && state.available.size > 0) {
                const first = [...state.available].sort()[0];
                const parsed = parseKey(first);
                if (parsed) {
                    state.viewYear = parsed.y;
                    state.viewMonth = parsed.m;
                }
            }
            if (state.open) renderGrid();
        },
        setSelected(dateKey) {
            state.selected = dateKey;
            const parsed = dateKey ? parseKey(dateKey) : null;
            if (parsed) {
                state.viewYear = parsed.y;
                state.viewMonth = parsed.m;
            }
            renderTrigger();
            if (state.open) renderGrid();
        },
        getSelected() {
            return state.selected;
        },
        destroy() {
            closePopup();
            root.remove();
        },
    };
};
