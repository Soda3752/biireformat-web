import {icon} from './icons';

export interface DropZoneOptions {
  title: string;
  hint?: string;
  accept?: string;
    /**
     * 是否允許一次拖入多個檔案。預設 false。
     * 開啟後：input 加上 multiple、drag/drop 與 picker 會依序對每個檔呼叫 onFile，
     * 並且不會把 dropzone 自身切到 loaded/error 狀態（由外層自管清單）。
     */
    multiple?: boolean;
  onFile: (file: File) => Promise<void> | void;
}

export interface DropZoneController {
  element: HTMLElement;
  setStatus(status: 'idle' | 'loaded' | 'error', meta?: string): void;
  reset(): void;
}

export function createDropZone(options: DropZoneOptions): DropZoneController {
  const root = document.createElement('div');
  root.className = 'dropzone';
  root.setAttribute('role', 'button');
  root.setAttribute('tabindex', '0');

  const accept = options.accept ?? '.xlsx,.xls';

  const render = (status: 'idle' | 'loaded' | 'error', meta?: string) => {
    const baseClass = 'dropzone';
    root.className = baseClass + (status === 'idle' ? '' : ` is-${status}`);
    const iconName = status === 'loaded' ? 'check' : status === 'error' ? 'alert' : 'upload';
    const title =
      status === 'loaded'
        ? options.title
        : status === 'error'
          ? '讀取失敗'
          : options.title;
    const metaLine = meta
      ? `<div class="dropzone-meta">${escapeHtml(meta)}</div>`
      : '';
    const hintLine =
      status === 'idle' && options.hint
        ? `<div class="dropzone-hint">${escapeHtml(options.hint)}</div>`
        : '';
    const clearBtn =
      status === 'loaded' || status === 'error'
        ? `<button type="button" class="dropzone-clear">清除</button>`
        : '';

    root.innerHTML = `
      <span class="dropzone-icon">${icon(iconName, 28)}</span>
      <div class="dropzone-title">${escapeHtml(title)}</div>
      ${metaLine}
      ${hintLine}
      ${clearBtn}
    `;

    root.querySelector<HTMLButtonElement>('.dropzone-clear')?.addEventListener('click', (e) => {
      e.stopPropagation();
      reset();
    });
  };

    const multiple = options.multiple === true;

  const fileInput = document.createElement('input');
  fileInput.type = 'file';
  fileInput.accept = accept;
    if (multiple) fileInput.multiple = true;
  root.appendChild(fileInput);

  const triggerPick = () => fileInput.click();

  root.addEventListener('click', triggerPick);
  root.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      triggerPick();
    }
  });

  fileInput.addEventListener('change', () => {
      const files = Array.from(fileInput.files ?? []);
    fileInput.value = '';
      if (files.length > 0) void handleFiles(files);
  });

  root.addEventListener('dragover', (e) => {
    e.preventDefault();
    root.classList.add('is-dragover');
  });
  root.addEventListener('dragleave', () => root.classList.remove('is-dragover'));
  root.addEventListener('drop', (e) => {
    e.preventDefault();
    root.classList.remove('is-dragover');
      const files = Array.from(e.dataTransfer?.files ?? []);
      if (files.length > 0) void handleFiles(files);
  });

    const handleFiles = async (files: File[]) => {
        if (multiple) {
            for (const file of files) {
                try {
                    await options.onFile(file);
                } catch (err) {
                    console.error(err);
                }
            }
            return;
        }
        const file = files[0];
        if (file) await handleSingle(file);
    };

    const handleSingle = async (file: File) => {
    try {
      await options.onFile(file);
    } catch (err) {
      console.error(err);
      const message = err instanceof Error ? err.message : '處理檔案時發生錯誤';
      render('error', message);
    }
  };

  const setStatus = (status: 'idle' | 'loaded' | 'error', meta?: string) => render(status, meta);
  const reset = () => render('idle');

  render('idle');

  return { element: root, setStatus, reset };
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
