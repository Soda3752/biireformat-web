/**
 * 自動催款分頁的「收款登記」視窗：搜尋客人 → 勾選收到的期別 → 一次寫回各月的「X月_已自動對帳.xlsx」。
 *
 * - Excel 所在資料夾跟對帳 2.0「全部存檔」用同一個（選一次就記住）；覆蓋前把舊檔備份到「_備份」。
 * - 瀏覽器不能直接存資料夾時，改成請使用者選那幾個月的檔案，登記好後下載修改過的檔案。
 * - 登記完視窗不關，游標回到搜尋框，可以馬上接著登記下一位。
 */

import ExcelJS from 'exceljs';
import {saveAs} from 'file-saver';

import {showToast} from '@/ui/toast';
import {type History, localIso, type Marks, money} from '@/domain/auto-collect/ledger';
import {
    buildEntries,
    candidateFilenames,
    type RegisterCustomer,
    type RegisterLine,
    registerCustomers,
    searchRegister,
    todayIso,
} from '@/domain/auto-collect/payment-register';
import {detectMonth, parseStatement} from '@/domain/auto-collect/statement-parser';
import {type PaymentEntry, registerPayments} from '@/domain/auto-reconcile-engine';
import {restoreExternalLinks} from '@/domain/xlsx-external-links';
import {
    BACKUP_DIR, backupFiles, describeWriteError, ensureWritable, loadFolderHandle, pickFolder, saveFolderHandle, supportsFolderSave, writeFiles,
} from '@/infra/folder-save';

/** 跟對帳 2.0 存檔同一個資料夾。 */
const FOLDER_KEY = 'recon-v2-save-dir';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const esc = (v: unknown): string =>
    String(v ?? '').replace(/[&<>"']/g, (c) => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]!));

export interface RegisterDeps {
    getHistory: () => History;
    getMarks: () => Marks;
    /** 存進瀏覽器；失敗回 false（呼叫端已提示）。 */
    persist: (h: History) => boolean;
    /** 登記完重畫分頁。 */
    onChanged: () => void;
    /** 直接從某位客人那列打開時帶入（客人鍵）。 */
    preselect?: string;
}

interface MonthJob {
    key: string;
    month: number;
    filename: string;
    src: ArrayBuffer;
    entries: PaymentEntry[];
}

export function openRegisterDialog(deps: RegisterDeps): void {
    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');
    backdrop.innerHTML = `
      <div class="app-modal ac-reg-modal" role="dialog" aria-modal="true" aria-label="收款登記">
        <header class="app-modal-header"><h2 class="app-modal-title">收款登記</h2></header>
        <div class="app-modal-body">
          <div class="ac-reg-top">
            <label class="ac-field ac-reg-date"><span>收款日</span>
              <input class="app-form-input" type="date" data-role="date"></label>
            <label class="ac-field ac-reg-search"><span>客人（打代號或店名）</span>
              <input class="app-form-input" type="search" placeholder="例如：賴、1001" autocomplete="off" data-role="q"></label>
          </div>
          <div class="ac-reg-hits" data-role="hits"></div>
          <div class="ac-reg-cust" data-role="cust"></div>
          <div class="ac-reg-folder" data-role="folder"></div>
          <div class="ac-reg-log" data-role="log"></div>
        </div>
        <footer class="app-modal-footer">
          <button type="button" class="btn btn-secondary" data-act="close">關閉</button>
          <button type="button" class="btn btn-primary" data-act="save" disabled>登記</button>
        </footer>
      </div>`;
    document.body.appendChild(backdrop);
    const dialog = backdrop.querySelector<HTMLElement>('.app-modal')!;
    const $ = <T extends HTMLElement = HTMLElement>(role: string): T => dialog.querySelector<T>(`[data-role="${role}"]`)!;
    const dateEl = $<HTMLInputElement>('date');
    const q = $<HTMLInputElement>('q');
    const saveBtn = dialog.querySelector<HTMLButtonElement>('[data-act="save"]')!;
    dateEl.value = todayIso();

    let list: RegisterCustomer[] = registerCustomers(deps.getHistory(), deps.getMarks());
    let current: RegisterCustomer | null = null;
    let hits: RegisterCustomer[] = [];
    let busy = false;
    let dir: FileSystemDirectoryHandle | null = null;

    const close = () => {
        backdrop.remove();
        document.removeEventListener('keydown', onKey, true);
    };

    // ---------- 資料夾

    const renderFolder = () => {
        $('folder').innerHTML = supportsFolderSave()
            ? (dir
                ? `Excel 會直接改「<b>${esc(dir.name)}</b>」資料夾裡的檔（舊檔先備份到「${BACKUP_DIR}」）　<a href="#" data-act="folder">換資料夾</a>`
                : '第一次按「登記」會請你選放對帳 Excel（X月_已自動對帳.xlsx）的資料夾，之後就記住。')
            : '這個瀏覽器不能直接改資料夾裡的檔：按「登記」後請選那幾個月的 Excel，改好會下載新的檔，再自己蓋掉原本的檔。';
    };
    void loadFolderHandle(FOLDER_KEY).then((h) => {
        dir = h;
        renderFolder();
    });
    renderFolder();

    // ---------- 搜尋與選客人

    const renderHits = () => {
        const kw = q.value.trim();
        hits = current && kw === current.name ? [] : searchRegister(list, kw);
        $('hits').innerHTML = kw && !hits.length && !(current && kw === current.name)
            ? '<div class="ac-reg-none">找不到這位客人（只找已記住月份裡的客人）</div>'
            : hits.map((c, i) => `<button type="button" class="ac-reg-hit${i === 0 ? ' is-first' : ''}" data-pick="${esc(c.ckey)}">
                <span><b>${esc(c.code)}</b> ${esc(c.name)}${c.monthly ? '<span class="ac-tag">月結</span>' : ''}</span>
                <span class="${c.lines.length ? 'ac-reg-owe' : 'ac-reg-clear'}">${c.lines.length ? `未付 ${c.lines.length} 筆，${money(c.total)} 元` : '都付清了'}</span>
              </button>`).join('');
    };

    const lineRow = (l: RegisterLine, i: number): string => `<label class="ac-reg-line">
        <input type="checkbox" data-line="${i}" checked>
        <span class="ac-reg-label">${esc(l.label)}</span>
        <span class="ac-reg-bill">帳單 ${money(l.amt)}${l.paid ? `，已收 ${money(l.paid)}` : ''}</span>
        <span class="ac-reg-amt">實收 <input class="app-form-input" type="number" inputmode="numeric" min="0" step="1" value="${l.owed}" data-amt="${i}"></span>
      </label>`;

    const pick = (c: RegisterCustomer) => {
        current = c;
        q.value = c.name;
        renderHits();
        renderCust();
        if (c.lines.length) saveBtn.focus();
    };

    const renderCust = () => {
        const c = current;
        if (!c) {
            $('cust').innerHTML = '';
            saveBtn.disabled = true;
            return;
        }
        $('cust').innerHTML = `<div class="ac-reg-head"><b>${esc(c.code)} ${esc(c.name)}</b>${c.monthly ? '<span class="ac-tag">月結</span>' : ''}
            <span class="muted">目前未付 ${money(c.total)} 元</span></div>`
            + (c.lines.length
                ? `<div class="ac-reg-lines">${c.lines.map(lineRow).join('')}</div>
                   <div class="ac-reg-sum" data-role="sum"></div>`
                : '<div class="ac-reg-none">這位客人在已記住的月份裡都付清了。</div>');
        updateSum();
    };

    const picked = (): Array<{line: RegisterLine; amount: number}> => {
        if (!current) return [];
        return current.lines.flatMap((line, i) => {
            const box = dialog.querySelector<HTMLInputElement>(`[data-line="${i}"]`);
            const amt = dialog.querySelector<HTMLInputElement>(`[data-amt="${i}"]`);
            if (!box?.checked || !amt) return [];
            return [{line, amount: Math.round(Number(amt.value) || 0)}];
        });
    };

    const updateSum = () => {
        const p = picked();
        const sum = p.reduce((t, x) => t + x.amount, 0);
        const sumEl = dialog.querySelector('[data-role="sum"]');
        if (sumEl) sumEl.innerHTML = p.length ? `這次登記 <b>${p.length}</b> 筆，合計 <b>${money(sum)}</b> 元，收款日 ${esc(dateEl.value.replace(/-/g, '/'))}` : '請勾選收到的期別';
        saveBtn.disabled = busy || !p.some((x) => x.amount > 0) || !dateEl.value;
    };

    q.addEventListener('input', () => {
        if (current && q.value.trim() !== current.name) {
            current = null;
            renderCust();
        }
        renderHits();
    });
    q.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && hits.length) {
            e.preventDefault();
            pick(hits[0]);
        }
    });
    dateEl.addEventListener('change', updateSum);
    dialog.addEventListener('input', (e) => {
        if ((e.target as HTMLElement).matches('[data-amt]')) updateSum();
    });
    dialog.addEventListener('change', (e) => {
        if ((e.target as HTMLElement).matches('[data-line]')) updateSum();
    });
    // 改完金額直接按 Enter 就登記
    dialog.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && (e.target as HTMLElement).matches('[data-amt]')) {
            e.preventDefault();
            save();
        }
    });

    // ---------- 寫回

    const log = (html: string) => $('log').insertAdjacentHTML('afterbegin', `<div>${html}</div>`);

    /** 改好各月的 workbook；有問題回傳說明（一份都不寫）。 */
    const applyAll = async (jobs: MonthJob[]): Promise<{out: Array<{job: MonthJob; data: ArrayBuffer}>; problems: string[]}> => {
        const problems: string[] = [];
        const out: Array<{job: MonthJob; data: ArrayBuffer}> = [];
        for (const job of jobs) {
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(job.src.slice(0));
            const probs = registerPayments(wb, job.month, job.entries);
            if (probs.length) {
                problems.push(...probs.map((p) => `${job.month}月${p.half}期：${p.reason}`));
                continue;
            }
            // ExcelJS 會丟掉外部活頁簿連結，從原檔搬回去，否則 Excel 打不開
            out.push({job, data: await restoreExternalLinks(await wb.xlsx.writeBuffer() as ArrayBuffer, job.src)});
        }
        return {out, problems};
    };

    /** 寫好的檔重新讀進網站記住的月份，欠款立刻更新。 */
    const remember = async (written: Array<{job: MonthJob; data: ArrayBuffer; mtime: number}>): Promise<boolean> => {
        const h = deps.getHistory();
        const months = {...h.months};
        const now = localIso(new Date());
        for (const {job, data, mtime} of written) {
            const st = await parseStatement(data, job.filename);
            months[job.key] = {...months[job.key], source: job.filename, file_mtime: localIso(new Date(mtime)), imported_at: now, periods: st.periods};
        }
        return deps.persist({version: 1, months});
    };

    const finish = (c: RegisterCustomer, picks: Array<{line: RegisterLine; amount: number}>, how: string) => {
        const sum = picks.reduce((t, x) => t + x.amount, 0);
        log(`✓ ${esc(c.code)} ${esc(c.name)}：${picks.map((x) => esc(x.line.label)).join('、')}，共 ${money(sum)} 元（${esc(how)}）`);
        showToast({variant: 'success', message: `已登記 ${c.name} ${picks.length} 筆，共 ${money(sum)} 元`, duration: 2500});
        deps.onChanged();
        list = registerCustomers(deps.getHistory(), deps.getMarks());
        current = null;
        q.value = '';
        renderHits();
        renderCust();
        q.focus();
    };

    const jobsFor = (byMonth: Map<string, PaymentEntry[]>): Array<Omit<MonthJob, 'filename' | 'src'>> =>
        [...byMonth].sort(([a], [b]) => (a < b ? -1 : 1)).map(([key, entries]) => ({key, month: +key.slice(5), entries}));

    const saveToFolder = async (c: RegisterCustomer, picks: Array<{line: RegisterLine; amount: number}>) => {
        // 權限要在使用者按下後馬上要，不能等讀完檔
        let d = dir;
        if (d && !(await ensureWritable(d))) d = null;
        if (!d) {
            d = await pickFolder();
            if (!d) return;
            if (!(await ensureWritable(d))) {
                showToast({variant: 'error', title: '沒有寫入權限', message: '請在選資料夾時允許「編輯檔案」'});
                return;
            }
            dir = d;
            await saveFolderHandle(FOLDER_KEY, d);
            renderFolder();
        }
        const h = deps.getHistory();
        const jobs: MonthJob[] = [];
        const missing: string[] = [];
        for (const j of jobsFor(buildEntries(picks, dateEl.value))) {
            let found: {filename: string; src: ArrayBuffer} | null = null;
            for (const name of candidateFilenames(h, j.key)) {
                try {
                    found = {filename: name, src: await (await (await d.getFileHandle(name)).getFile()).arrayBuffer()};
                    break;
                } catch (err) {
                    if (!(err instanceof DOMException && (err.name === 'NotFoundError' || err.name === 'TypeMismatchError'))) throw err;
                }
            }
            if (found) jobs.push({...j, ...found});
            else missing.push(`${j.month}月（${candidateFilenames(h, j.key).join(' 或 ')}）`);
        }
        if (missing.length) {
            showToast({variant: 'error', duration: 0, title: '沒有登記',
                message: `「${d.name}」資料夾裡找不到 ${missing.join('、')}。請按「換資料夾」選放對帳 Excel 的資料夾。`});
            return;
        }
        const {out, problems} = await applyAll(jobs);
        if (problems.length) {
            showToast({variant: 'error', duration: 0, title: '沒有登記（Excel 跟網站記住的不一樣）', message: problems.join('；')});
            return;
        }
        try {
            await backupFiles(d, out.map((x) => x.job.filename));
        } catch (err) {
            showToast({variant: 'error', title: '備份舊檔失敗，沒有登記', message: describeWriteError(err)});
            return;
        }
        const written: Array<{job: MonthJob; data: ArrayBuffer; mtime: number}> = [];
        const failed = await writeFiles(d, out.map((x) => ({filename: x.job.filename, data: x.data})), (filename, file) => {
            const x = out.find((o) => o.job.filename === filename)!;
            written.push({...x, mtime: file.lastModified});
        });
        if (written.length) await remember(written);
        if (failed.length) {
            deps.onChanged();
            list = registerCustomers(deps.getHistory(), deps.getMarks());
            showToast({variant: 'error', duration: 0, title: `有 ${failed.length} 個月沒寫進去`,
                message: failed.map((f) => `${f.filename}：${describeWriteError(f.error)}`).join('；')
                    + (written.length ? `（${written.map((w) => `${w.job.month}月`).join('、')}已寫好）` : '')});
            if (current) current = list.find((x) => x.ckey === current!.ckey) ?? null;
            renderCust();
            return;
        }
        finish(c, picks, `已寫回 ${written.map((w) => `${w.job.month}月`).join('、')} 的 Excel`);
    };

    /** 不支援資料夾存取：請使用者選那幾個月的檔案，改好後下載。 */
    const saveByDownload = (c: RegisterCustomer, picks: Array<{line: RegisterLine; amount: number}>) => {
        const h = deps.getHistory();
        const need = jobsFor(buildEntries(picks, dateEl.value));
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.xlsx';
        input.multiple = true;
        input.addEventListener('change', () => void run(async () => {
            const files = [...(input.files ?? [])];
            const jobs: MonthJob[] = [];
            const missing: number[] = [];
            for (const j of need) {
                const f = files.find((x) => x.name === h.months[j.key]?.source) ?? files.find((x) => detectMonth(x.name) === j.month);
                if (f) jobs.push({...j, filename: f.name, src: await f.arrayBuffer()});
                else missing.push(j.month);
            }
            if (missing.length) {
                showToast({variant: 'error', duration: 0, title: '沒有登記', message: `請一起選 ${need.map((j) => `${j.month}月`).join('、')} 的 Excel（少了 ${missing.map((m) => `${m}月`).join('、')}）`});
                return;
            }
            const {out, problems} = await applyAll(jobs);
            if (problems.length) {
                showToast({variant: 'error', duration: 0, title: '沒有登記（Excel 跟網站記住的不一樣）', message: problems.join('；')});
                return;
            }
            for (const x of out) saveAs(new Blob([x.data], {type: XLSX_MIME}), x.job.filename);
            await remember(out.map((x) => ({...x, mtime: Date.now()})));
            finish(c, picks, `已下載改好的 ${out.map((x) => `${x.job.month}月`).join('、')} Excel，請蓋掉原本的檔`);
        }));
        input.click();
    };

    const run = async (job: () => Promise<void>) => {
        busy = true;
        saveBtn.disabled = true;
        saveBtn.textContent = '登記中…';
        try {
            await job();
        } catch (err) {
            console.error('[auto-collect-register]', err);
            showToast({variant: 'error', title: '登記失敗', message: describeWriteError(err)});
        } finally {
            busy = false;
            saveBtn.textContent = '登記';
            updateSum();
        }
    };

    const save = () => {
        const c = current;
        const picks = picked().filter((x) => x.amount > 0);
        if (busy || !c || !picks.length || !dateEl.value) return;
        if (supportsFolderSave()) void run(() => saveToFolder(c, picks));
        else saveByDownload(c, picks);
    };

    // ---------- 按鈕與鍵盤

    dialog.addEventListener('click', (e) => {
        const t = e.target as HTMLElement;
        const hit = t.closest<HTMLElement>('[data-pick]');
        if (hit) {
            const c = list.find((x) => x.ckey === hit.dataset.pick);
            if (c) pick(c);
            return;
        }
        const act = t.closest<HTMLElement>('[data-act]')?.dataset.act;
        if (act === 'close') close();
        else if (act === 'save') save();
        else if (act === 'folder') {
            e.preventDefault();
            void pickFolder().then(async (h) => {
                if (!h) return;
                dir = h;
                await saveFolderHandle(FOLDER_KEY, h);
                renderFolder();
            }).catch((err) => showToast({variant: 'error', title: '選資料夾失敗', message: describeWriteError(err)}));
        }
    });

    const onKey = (e: KeyboardEvent) => {
        if (!backdrop.isConnected) return;
        if (e.key === 'Escape' && !busy) {
            e.preventDefault();
            close();
        }
    };
    document.addEventListener('keydown', onKey, true);

    const pre = deps.preselect ? list.find((x) => x.ckey === deps.preselect) : undefined;
    if (pre) pick(pre);
    else q.focus();
}
