/**
 * 「二聯發票」分頁：兩個月對帳檔＋結帳原始檔＋三聯金額＋總額 → 算出要開哪些二聯發票。
 *
 * 功能與流程照 invoice-helper 網頁版（web/index.html、web/app.ts）：
 * - ① 選兩個月對帳檔（月份從檔名「N月」判斷，判斷不到可選）＋結帳原始檔（可多份，用來排除有統編的客戶）
 * - ② 三聯：前月／後月各一個大框框連續打金額（空白或換行分隔），即時張數／合計、看不懂的紅字列出
 * - ③ 總額（可打逗號）＋進階設定（每月模式、門檻、超出範圍、張數偏好、每期最多幾張、民國年、固定排除名單）
 * - 結果：摘要卡、四期卡、需人工確認、二聯清單、未選上／排除名單、下載客戶清單 Excel
 *
 * 計算核心在 src/invoice/（移植自 invoice-helper/src），在 Web Worker 裡跑，檔案不離開瀏覽器。
 * 三聯、總額、進階設定記在這台電腦的瀏覽器（localStorage）；檔案要重選。
 */

import './invoice.css';
import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';
import {monthFromFilename} from '@/invoice/dates';
import {parseTriText, type TriParse} from '@/invoice/triText';
import {parseExcludeList} from '@/invoice/select';
import {FIXED_EXCLUDE_TEXT} from '@/invoice/exclude';
import type {WebInput, WebOutput} from '@/invoice/webCalc';

const fmt = (n: number) => n.toLocaleString('zh-TW');

type Kid = Node | string | null | undefined | false;
function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Record<string, string> = {}, ...kids: Kid[]): HTMLElementTagNameMap[K] {
    const el = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
    for (const c of kids) if (c !== null && c !== undefined && c !== false) el.append(c);
    return el;
}

/** 金額字串 → 數字；空白回 null；看不懂回 NaN。可含逗號、空白、全形數字、$、元 */
function parseAmount(s: string): number | null {
    const t = s
        .replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
        .replace(/[，,\s　$＄元]/g, '')
        .replace(/^[(（](.*)[)）]$/, '-$1');
    if (t === '') return null;
    return /^-?\d+(\.\d+)?$/.test(t) ? Number(t) : NaN;
}

// ---------- 本機記憶（只存在這台電腦的瀏覽器，失敗就算了） ----------
const STORE = 'bii.invoice.v1';
type Adv = { mode: 'period' | 'monthly'; threshold: string; overMin: string; overMax: string; prefer: 'many' | 'few'; maxPerHalf: string; year: string };
type Saved = { triText: [string, string]; total: string; monthTotals: [string, string]; adv: Partial<Adv> };
const DEFAULT_ADV: Adv = { mode: 'period', threshold: '7000', overMin: '300', overMax: '5000', prefer: 'many', maxPerHalf: '31', year: '' };

function loadSaved(): Partial<Saved> {
    try {
        return JSON.parse(localStorage.getItem(STORE) ?? '{}') as Partial<Saved>;
    } catch {
        return {};
    }
}

type BillFile = { file: File; month: number | null };
type WorkerReply = { ok: true; out: WebOutput } | { ok: false; error: string };

export function renderInvoicePanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card inv-card">
      <header class="card-header">
        <h1 class="card-title">二聯發票</h1>
        <p class="card-subtitle">
          選好兩個月的檔案、打上三聯金額和總額，按「開始計算」，算出這期要開哪些客戶的二聯發票。
          檔案只在這台電腦的瀏覽器裡計算，不會上傳。
        </p>
      </header>

      <section class="inv-step">
        <h2 class="inv-step-title"><span class="inv-step-no">1</span>選擇檔案</h2>
        <p class="inv-hint">可以點框框選檔，也可以直接把檔案拖進框框裡。</p>
        <div class="inv-files">
          <div class="inv-uploader">
            <div class="analytics-dropzone" data-role="bill-drop" tabindex="0" role="button">
              <span class="analytics-dropzone-icon">${icon('upload', 28)}</span>
              <div class="analytics-dropzone-title">對帳檔（兩個月）</div>
              <div class="analytics-dropzone-hint">例如「9月_已自動對帳.xlsx」「10月_已自動對帳.xlsx」</div>
              <input type="file" accept=".xlsx" multiple hidden data-role="bill-input">
            </div>
            <ul class="inv-filelist" data-role="bill-list"></ul>
          </div>
          <div class="inv-uploader">
            <div class="analytics-dropzone" data-role="source-drop" tabindex="0" role="button">
              <span class="analytics-dropzone-icon">${icon('upload', 28)}</span>
              <div class="analytics-dropzone-title">結帳原始檔（可多份）</div>
              <div class="analytics-dropzone-hint">「應收帳款對帳單--明細表」，用來找出有統編的客戶（不開二聯）</div>
              <input type="file" accept=".xlsx" multiple hidden data-role="source-input">
            </div>
            <ul class="inv-filelist" data-role="source-list"></ul>
          </div>
        </div>
      </section>

      <section class="inv-step">
        <h2 class="inv-step-title"><span class="inv-step-no">2</span>三聯金額</h2>
        <p class="inv-hint">
          在框框裡連續打每張三聯的金額（含稅），<b>數字之間按空白或換行就好</b>，網頁會自動拆成一張一張。
          也可以用 Tab、全形空白、頓號「、」隔開，或從 Excel 整欄／整塊複製貼上（空白格沒關係）。<br>
          逗號：<b>1,290</b> 這種千分位寫法（逗號後面剛好三位數字）當成同一個數字；要分成兩張請用空白，例如「500 300」。
          全形數字可以直接打；0 和「-」不算一張。
        </p>
        <div class="inv-tri" data-role="tri"></div>
        <label class="inv-ack" data-role="tri-ack" hidden>
          <input type="checkbox" data-role="tri-ack-box">
          紅字標出的內容看不懂，<b>略過不算</b>，照樣計算
        </label>
        <div class="inv-tri-total">兩個月三聯合計：<b data-role="tri-all">0</b> 元</div>
      </section>

      <section class="inv-step">
        <h2 class="inv-step-title"><span class="inv-step-no">3</span>總額</h2>
        <p class="inv-hint">這兩個月要開發票的總額（含稅），可以打逗號，例如 1,550,000。</p>
        <div class="inv-total-row" data-role="total-period">
          <label for="inv-total">整期總額</label>
          <input class="inv-total-input" id="inv-total" data-role="total" inputmode="numeric" autocomplete="off" placeholder="例如 1,550,000">
          <span>元</span>
        </div>
        <div data-role="total-monthly" hidden></div>

        <details class="inv-adv" data-role="adv">
          <summary>進階設定（一般不用改）</summary>
          <div class="inv-adv-grid">
            <div class="inv-adv-label">總額模式</div>
            <div class="inv-adv-field">
              <label><input type="radio" name="inv-mode" value="period" checked> 整期一個總額（預設）</label>
              <label><input type="radio" name="inv-mode" value="monthly"> 每月各一個總額</label>
            </div>
            <div class="inv-adv-label">候選門檻</div>
            <div class="inv-adv-field">單期帳單金額大於 <input type="text" class="inv-num" data-adv="threshold" inputmode="numeric"> 元才列入候選</div>
            <div class="inv-adv-label">二聯合計要超出差額</div>
            <div class="inv-adv-field">
              <input type="text" class="inv-num" data-adv="overMin" inputmode="numeric"> ～
              <input type="text" class="inv-num" data-adv="overMax" inputmode="numeric"> 元
              <span class="inv-note">（二聯＋三聯合計尾數也不會是 00）</span>
            </div>
            <div class="inv-adv-label">張數偏好</div>
            <div class="inv-adv-field">
              <label><input type="radio" name="inv-prefer" value="many" checked> 張數多一點（偏小額客戶，預設）</label>
              <label><input type="radio" name="inv-prefer" value="few"> 張數少一點（偏大客戶）</label>
            </div>
            <div class="inv-adv-label">每期最多幾張</div>
            <div class="inv-adv-field">
              <input type="text" class="inv-num" data-adv="maxPerHalf" inputmode="numeric"> 張
              <span class="inv-note">（二聯發票工作簿每欄能填的格數，115.9.10 是 31 格；空白＝不限）</span>
            </div>
            <div class="inv-adv-label">民國年</div>
            <div class="inv-adv-field">
              <input type="text" class="inv-num" data-adv="year" inputmode="numeric" placeholder="自動">
              <span class="inv-note">（空白＝自動；跨年時例如 1 月算去年 11～12 月請填）</span>
            </div>
            <div class="inv-adv-label">固定排除名單</div>
            <div class="inv-adv-field">
              <div class="inv-exclude" data-role="exclude"></div>
              <div class="inv-note">這些客戶本來就有開三聯，不開二聯（內建名單）。</div>
            </div>
            <div></div>
            <div class="inv-adv-field"><button type="button" class="btn btn-secondary btn-sm" data-role="reset-adv">恢復預設值</button></div>
          </div>
        </details>
      </section>

      <div class="inv-go">
        <button type="button" class="btn btn-primary btn-lg" data-role="run">開始計算</button>
        <div class="inv-run-block" data-role="run-block" hidden></div>
        <div class="inv-status" data-role="status"></div>
        <div class="inv-error" data-role="error" hidden></div>
      </div>

      <section class="inv-result" data-role="result" hidden></section>
    </div>
  `;

    const $ = <T extends HTMLElement = HTMLElement>(role: string) => panel.querySelector<T>(`[data-role="${role}"]`)!;
    const advInput = (k: string) => panel.querySelector<HTMLInputElement>(`[data-adv="${k}"]`)!;
    const radio = (name: string) => panel.querySelector<HTMLInputElement>(`input[name="inv-${name}"]:checked`)?.value;

    // ---------- 狀態 ----------
    const saved = loadSaved();
    const state = {
        bills: [] as BillFile[],
        sources: [] as File[],
        triText: (Array.isArray(saved.triText) && saved.triText.length === 2 ? [...saved.triText] : ['', '']) as [string, string],
        monthTotals: (Array.isArray(saved.monthTotals) && saved.monthTotals.length === 2 ? [...saved.monthTotals] : ['', '']) as [string, string],
    };
    const totalInp = $<HTMLInputElement>('total');

    function save(): void {
        try {
            const s: Saved = { triText: state.triText, total: totalInp.value, monthTotals: state.monthTotals, adv: readAdv() };
            localStorage.setItem(STORE, JSON.stringify(s));
        } catch {
            /* 無痕視窗等情況：不記也沒關係 */
        }
    }

    /** 對帳檔月份（由小到大）；不足兩個時回 null */
    function billMonths(): [number, number] | null {
        const ms = state.bills.map((b) => b.month).filter((m): m is number => m !== null).sort((a, b) => a - b);
        return ms.length === 2 && ms[0] !== ms[1] ? [ms[0]!, ms[1]!] : null;
    }
    const colLabel = (i: 0 | 1) => {
        const ms = billMonths();
        return ms ? `${ms[i]}月` : i === 0 ? '前月' : '後月';
    };

    // ---------- 1. 檔案 ----------
    const isXlsx = (f: File) => {
        if (/\.xlsx$/i.test(f.name)) return true;
        showToast({ variant: 'warning', message: `「${f.name}」不是 .xlsx 檔，略過` });
        return false;
    };
    bindDropzone($('bill-drop'), $<HTMLInputElement>('bill-input'), (fs) => {
        for (const f of fs.filter(isXlsx)) {
            const i = state.bills.findIndex((b) => b.file.name === f.name);
            const item = { file: f, month: monthFromFilename(f.name) };
            if (i >= 0) state.bills[i] = item;
            else state.bills.push(item);
        }
        if (state.bills.length > 2) showToast({ variant: 'warning', message: '對帳檔只要兩個月，請刪掉多的' });
        renderBills();
    });
    bindDropzone($('source-drop'), $<HTMLInputElement>('source-input'), (fs) => {
        for (const f of fs.filter(isXlsx)) if (!state.sources.some((s) => s.name === f.name)) state.sources.push(f);
        renderSources();
    });

    const removeBtn = (onClick: () => void) => {
        const b = h('button', { type: 'button', class: 'analytics-file-chip-remove', 'aria-label': '移除', title: '移除' });
        b.innerHTML = icon('close', 14);
        b.addEventListener('click', onClick);
        return b;
    };

    function renderBills(): void {
        $('bill-list').replaceChildren(
            ...state.bills.map((b, i) => {
                const sel = h('select', { class: 'inv-month-select', 'aria-label': '月份' }, h('option', { value: '' }, '選月份'), ...Array.from({ length: 12 }, (_, k) => h('option', { value: String(k + 1) }, `${k + 1}月`)));
                sel.value = b.month ? String(b.month) : '';
                sel.addEventListener('change', () => {
                    b.month = sel.value ? Number(sel.value) : null;
                    renderBills();
                });
                const del = removeBtn(() => {
                    state.bills.splice(i, 1);
                    renderBills();
                });
                return h('li', {}, h('span', { class: 'inv-fname' }, b.file.name), sel, b.month ? null : h('span', { class: 'inv-need' }, '← 請選月份'), del);
            }),
        );
        updateTriLabels(); // 欄名跟著月份變
        if (readAdv().mode === 'monthly') renderMonthTotals();
    }
    function renderSources(): void {
        $('source-list').replaceChildren(
            ...state.sources.map((f, i) =>
                h('li', {}, h('span', { class: 'inv-fname' }, f.name), removeBtn(() => {
                    state.sources.splice(i, 1);
                    renderSources();
                })),
            ),
        );
    }

    // ---------- 2. 三聯大框框連續輸入 ----------
    const triParsed: [TriParse, TriParse] = [parseTriText(state.triText[0]), parseTriText(state.triText[1])];
    let ackSig = ''; // 使用者勾「略過」時的錯誤內容；錯誤內容變了就要重新確認
    const errSig = () => triParsed.map((p) => p.errors.map((e) => `${e.start}:${e.text}`).join('|')).join('||');
    const hasTriErrors = () => triParsed.some((p) => p.errors.length > 0);
    const triBlocked = () => hasTriErrors() && ackSig !== errSig();
    const triEls: { label: HTMLElement; sum: HTMLElement; box: HTMLTextAreaElement; err: HTMLElement; listSum: HTMLElement; list: HTMLElement }[] = [];

    function buildTri(): void {
        $('tri').replaceChildren(
            ...([0, 1] as const).map((col) => {
                const ta = h('textarea', { class: 'inv-tri-box', spellcheck: 'false', autocomplete: 'off', placeholder: '例如：\n1290 3000 12,600\n4500 8820 …' });
                ta.value = state.triText[col];
                const setText = (v: string) => {
                    ta.value = v;
                    state.triText[col] = v;
                    triParsed[col] = parseTriText(v);
                    updateTri();
                };
                ta.addEventListener('input', () => setText(ta.value));
                const bar = h('div', { class: 'inv-tri-bar' });
                const clear = h('button', { type: 'button', class: 'btn btn-secondary btn-sm' }, '清空');
                clear.addEventListener('click', () => {
                    if (!ta.value.trim()) return;
                    const backup = ta.value;
                    setText('');
                    const undo = h('button', { type: 'button', class: 'btn btn-secondary btn-sm' }, '復原');
                    undo.addEventListener('click', () => {
                        setText(backup);
                        undo.remove();
                    });
                    bar.append(undo);
                });
                bar.append(clear);
                const label = h('span', {});
                const sum = h('span', { class: 'inv-tri-sum' });
                const err = h('div', { class: 'inv-tri-err' });
                err.hidden = true;
                const listSum = h('summary', {}, '拆出來的金額');
                const list = h('div', { class: 'inv-chips' });
                triEls[col] = { label, sum, box: ta, err, listSum, list };
                return h('div', { class: 'inv-tri-col' }, h('h3', { class: 'inv-tri-head' }, label, sum), ta, err, h('details', { class: 'inv-tri-list' }, listSum, list), bar);
            }),
        );
        $<HTMLInputElement>('tri-ack-box').addEventListener('change', () => {
            ackSig = $<HTMLInputElement>('tri-ack-box').checked ? errSig() : '';
            updateTri();
        });
        updateTri();
    }

    /** 欄名跟著對帳檔月份變 */
    function updateTriLabels(): void {
        for (const col of [0, 1] as const) {
            const el = triEls[col];
            if (!el) continue;
            el.label.textContent = `${colLabel(col)}三聯`;
            el.box.setAttribute('aria-label', `${colLabel(col)}三聯金額`);
        }
    }

    function updateTri(): void {
        updateTriLabels();
        for (const col of [0, 1] as const) {
            const p = triParsed[col];
            const text = state.triText[col];
            const el = triEls[col]!;
            el.sum.textContent = `${p.amounts.length} 張　${fmt(p.sum)}`;
            el.box.classList.toggle('is-invalid', p.errors.length > 0);
            el.err.hidden = !p.errors.length;
            if (p.errors.length) {
                el.err.replaceChildren(
                    h('div', {}, `⚠ 有 ${p.errors.length} 段看不懂，沒有算進去（點一下可跳到那裡修改）：`),
                    h(
                        'ul',
                        {},
                        ...p.errors.slice(0, 20).map((e) => {
                            const before = text.slice(Math.max(0, e.start - 12), e.start).replace(/\s+/g, ' ');
                            const after = text.slice(e.end, e.end + 12).replace(/\s+/g, ' ');
                            const li = h('li', { title: '點一下跳到這裡' }, h('span', { class: 'inv-ctx' }, `${e.start > 12 ? '…' : ''}${before}`), h('b', {}, e.text), h('span', { class: 'inv-ctx' }, `${after}${e.end + 12 < text.length ? '…' : ''}`));
                            li.addEventListener('click', () => {
                                el.box.focus();
                                el.box.setSelectionRange(e.start, e.end);
                            });
                            return li;
                        }),
                        p.errors.length > 20 ? h('li', {}, `…還有 ${p.errors.length - 20} 段`) : null,
                    ),
                );
            }
            el.listSum.textContent = `拆出來的金額（${p.amounts.length} 張，點開核對）`;
            el.list.replaceChildren(...p.amounts.map((n, i) => h('span', {}, h('i', {}, String(i + 1)), fmt(n))));
        }
        $('tri-all').textContent = fmt(triParsed[0].sum + triParsed[1].sum);
        // 有看不懂的內容：要修正或勾「略過」才能計算
        const errs = hasTriErrors();
        if (!errs) ackSig = '';
        $('tri-ack').hidden = !errs;
        $<HTMLInputElement>('tri-ack-box').checked = errs && ackSig === errSig();
        const blocked = triBlocked();
        const btn = $<HTMLButtonElement>('run');
        if (!running) btn.disabled = blocked;
        $('run-block').hidden = !blocked;
        $('run-block').textContent = blocked ? '三聯金額有紅字標出看不懂的內容，請修正，或勾選「略過不算」後再計算。' : '';
        save();
    }

    // ---------- 3. 總額與進階設定 ----------
    function bindMoney(inp: HTMLInputElement, onChange: () => void): void {
        inp.addEventListener('input', () => {
            const n = parseAmount(inp.value);
            inp.classList.toggle('is-invalid', n !== null && !Number.isFinite(n));
            onChange();
        });
        inp.addEventListener('blur', () => {
            const n = parseAmount(inp.value);
            if (n !== null && Number.isFinite(n)) inp.value = fmt(n);
            onChange();
        });
    }
    totalInp.value = saved.total ?? '';
    bindMoney(totalInp, save);

    function renderMonthTotals(): void {
        $('total-monthly').replaceChildren(
            ...([0, 1] as const).map((i) => {
                const inp = h('input', { class: 'inv-total-input', inputmode: 'numeric', autocomplete: 'off', 'aria-label': `${colLabel(i)}總額` });
                inp.value = state.monthTotals[i];
                bindMoney(inp, () => {
                    state.monthTotals[i] = inp.value;
                    save();
                });
                return h('div', { class: 'inv-total-row' }, h('label', {}, `${colLabel(i)}總額`), inp, h('span', {}, '元'));
            }),
        );
    }

    function readAdv(): Adv {
        return {
            mode: radio('mode') === 'monthly' ? 'monthly' : 'period',
            threshold: advInput('threshold').value,
            overMin: advInput('overMin').value,
            overMax: advInput('overMax').value,
            prefer: radio('prefer') === 'few' ? 'few' : 'many',
            maxPerHalf: advInput('maxPerHalf').value,
            year: advInput('year').value,
        };
    }
    function writeAdv(a: Adv): void {
        for (const k of ['threshold', 'overMin', 'overMax', 'maxPerHalf', 'year'] as const) advInput(k).value = a[k];
        panel.querySelector<HTMLInputElement>(`input[name="inv-mode"][value="${a.mode}"]`)!.checked = true;
        panel.querySelector<HTMLInputElement>(`input[name="inv-prefer"][value="${a.prefer}"]`)!.checked = true;
        applyMode();
    }
    function applyMode(): void {
        const monthly = readAdv().mode === 'monthly';
        $('total-period').hidden = monthly;
        $('total-monthly').hidden = !monthly;
        if (monthly) renderMonthTotals();
    }
    const advSaved: Adv = { ...DEFAULT_ADV, ...saved.adv };
    writeAdv(advSaved);
    if (JSON.stringify(advSaved) !== JSON.stringify(DEFAULT_ADV)) $<HTMLDetailsElement>('adv').open = true; // 有改過就展開提醒
    for (const el of panel.querySelectorAll<HTMLInputElement>('[data-role="adv"] input'))
        el.addEventListener('change', () => {
            applyMode();
            save();
        });
    $('reset-adv').addEventListener('click', () => {
        writeAdv(DEFAULT_ADV);
        save();
        showToast({ variant: 'success', message: '進階設定已恢復預設值' });
    });

    const excludeNames = parseExcludeList(FIXED_EXCLUDE_TEXT);
    $('exclude').textContent = excludeNames.length ? excludeNames.join('、') : '（名單是空的）';

    // ---------- 計算 ----------
    let running = false;
    let worker: Worker | null = null;
    const getWorker = () => (worker ??= new Worker(new URL('../invoice/worker.ts', import.meta.url), { type: 'module' }));

    function showError(msg: string): void {
        const e = $('error');
        e.textContent = msg;
        e.hidden = false;
    }

    function intField(v: string, what: string, opt = false): number | undefined {
        const n = parseAmount(v);
        if (n === null) {
            if (opt) return undefined;
            throw new Error(`進階設定「${what}」沒有填`);
        }
        if (!Number.isFinite(n)) throw new Error(`進階設定「${what}」不是數字：${v}`);
        return Math.round(n);
    }

    async function run(): Promise<void> {
        $('error').hidden = true;
        const btn = $<HTMLButtonElement>('run');
        try {
            if (state.bills.length !== 2) throw new Error(`請選兩個月的對帳檔（目前選了 ${state.bills.length} 個）`);
            const noMonth = state.bills.find((b) => !b.month);
            if (noMonth) throw new Error(`對帳檔「${noMonth.file.name}」看不出是幾月，請在檔名旁邊選月份`);
            const ms = billMonths();
            if (!ms) throw new Error('兩個對帳檔的月份一樣，請檢查月份');
            if (triBlocked()) {
                const list = ([0, 1] as const).filter((i) => triParsed[i].errors.length).map((i) => `${colLabel(i)}：${triParsed[i].errors.slice(0, 5).map((e) => `「${e.text}」`).join('')}`);
                throw new Error(`三聯金額有看不懂的內容（${list.join('；')}），請修正，或勾選「略過不算」後再計算`);
            }
            const adv = readAdv();
            const money = (s: string, what: string) => {
                const n = parseAmount(s);
                if (n === null) throw new Error(`請輸入${what}`);
                if (!Number.isFinite(n) || n <= 0) throw new Error(`${what}不是正確的金額：${s}`);
                return Math.round(n);
            };
            const total = adv.mode === 'period' ? money(totalInp.value, '總額') : undefined;
            const monthTotals: [number, number][] | undefined =
                adv.mode === 'monthly' ? ms.map((m, i) => [m, money(state.monthTotals[i as 0 | 1], `${m}月總額`)] as [number, number]) : undefined;
            const input: WebInput = {
                bills: await Promise.all(state.bills.map(async (b) => ({ name: b.file.name, month: b.month!, data: await b.file.arrayBuffer() }))),
                sources: await Promise.all(state.sources.map(async (f) => ({ name: f.name, data: await f.arrayBuffer() }))),
                triplicate: ms.map((m, i) => ({ month: m, amounts: triParsed[i as 0 | 1].amounts })),
                mode: adv.mode,
                total,
                monthTotals,
                threshold: intField(adv.threshold, '候選門檻')!,
                overMin: intField(adv.overMin, '超出下限')!,
                overMax: intField(adv.overMax, '超出上限')!,
                prefer: adv.prefer,
                maxPerHalf: intField(adv.maxPerHalf, '每期最多幾張', true),
                year: intField(adv.year, '民國年', true),
                excludeText: FIXED_EXCLUDE_TEXT,
            };
            if (input.overMin < 1 || input.overMax < input.overMin) throw new Error(`進階設定「超出差額」範圍不合理：${input.overMin}～${input.overMax} 元`);

            running = true;
            btn.disabled = true;
            btn.textContent = '計算中…';
            $('status').textContent = '正在讀取 Excel 並計算，結帳原始檔比較大，約需 5～20 秒，請稍候。';
            const t0 = performance.now();
            const res = await new Promise<WorkerReply>((resolve) => {
                const w = getWorker();
                w.onmessage = (e: MessageEvent<WorkerReply>) => resolve(e.data);
                w.onerror = (e) => {
                    worker?.terminate();
                    worker = null;
                    resolve({ ok: false, error: `計算程式出錯：${e.message || '未知錯誤'}` });
                };
                w.postMessage(input, [...input.bills.map((b) => b.data), ...input.sources.map((s) => s.data)]);
            });
            running = false;
            btn.textContent = '重新計算';
            btn.disabled = triBlocked();
            $('status').textContent = res.ok ? `計算完成（${((performance.now() - t0) / 1000).toFixed(1)} 秒）` : '';
            if (!res.ok) throw new Error(res.error);
            renderResult(res.out);
        } catch (e) {
            if (running) {
                running = false;
                btn.textContent = '開始計算';
            }
            btn.disabled = triBlocked();
            $('status').textContent = '';
            showError(e instanceof Error ? e.message : String(e));
        }
    }
    $('run').addEventListener('click', () => void run());

    // ---------- 結果 ----------
    function kpi(k: string, v: string, s?: string, cls = ''): HTMLElement {
        return h('div', { class: `inv-kpi ${cls}` }, h('div', { class: 'inv-kpi-label' }, k), h('div', { class: 'inv-kpi-value' }, v), s ? h('div', { class: 'inv-kpi-hint' }, s) : null);
    }

    type Row = { month: number; half: string; code: string; name: string; amount: number };
    function entryTable(rows: Row[], extra?: { head: string; get: (i: number) => string }[], subtotals = false): HTMLElement {
        const head = h('tr', {}, h('th', {}, '月份'), h('th', {}, '期別'), h('th', {}, '客戶代號'), h('th', {}, '客戶名稱'), h('th', { class: 'is-right' }, '金額'), ...(extra ?? []).map((x) => h('th', {}, x.head)));
        const body: HTMLElement[] = [];
        let groupKey = '', groupSum = 0, groupN = 0;
        const flush = () => {
            if (subtotals && groupN) body.push(h('tr', { class: 'inv-row-sub' }, h('td', { colspan: '4' }, `${groupKey} 小計 ${groupN} 張`), h('td', { class: 'is-right' }, fmt(groupSum))));
        };
        rows.forEach((e, i) => {
            const key = `${e.month}月${e.half}`;
            if (key !== groupKey) {
                flush();
                groupKey = key;
                groupSum = 0;
                groupN = 0;
            }
            groupSum += e.amount;
            groupN++;
            body.push(h('tr', {}, h('td', {}, `${e.month}月`), h('td', {}, e.half), h('td', {}, e.code || ''), h('td', {}, e.name), h('td', { class: 'is-right' }, fmt(e.amount)), ...(extra ?? []).map((x) => h('td', {}, x.get(i)))));
        });
        flush();
        if (subtotals) body.push(h('tr', { class: 'inv-row-total' }, h('td', { colspan: '4' }, `合計 ${rows.length} 張`), h('td', { class: 'is-right' }, fmt(rows.reduce((a, e) => a + e.amount, 0)))));
        return h('div', { class: 'inv-table-wrap' }, h('table', { class: 'inv-table' }, h('thead', {}, head), h('tbody', {}, ...body)));
    }

    const more = (title: string, ...kids: Kid[]) => h('details', { class: 'inv-more' }, h('summary', {}, title), ...kids);

    let downloadUrl: string | null = null;
    function renderResult(o: WebOutput): void {
        const R = o.result;
        const box = $('result');
        box.hidden = false;
        if (downloadUrl) URL.revokeObjectURL(downloadUrl);
        downloadUrl = URL.createObjectURL(new Blob([o.report], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
        const dl = h('a', { href: downloadUrl, download: o.reportFileName, class: 'btn btn-primary inv-download' });
        dl.innerHTML = `${icon('download', 16)}<span>下載客戶清單 Excel</span>`;

        const calcBlocks = R.calcs.map((c) => {
            const ok = !c.noSolution && c.over >= (c.overMin ?? 1) && c.over <= (c.overMax ?? Infinity) && c.grandTotal % 100 !== 0;
            const triNote = c.months
                .map((m) => {
                    const t = o.triplicateSource.match(new RegExp(`${m}月 (\\d+) 張 ([\\d,]+)`));
                    return t ? `${m}月 ${t[2]}（${t[1]} 張）` : null;
                })
                .filter(Boolean)
                .join('／');
            const n = c.halfStats.reduce((a, x) => a + x.count, 0);
            return h(
                'div',
                { class: 'inv-calc' },
                R.calcs.length > 1 ? h('h3', { class: 'inv-sub-title' }, c.label) : null,
                h(
                    'div',
                    { class: 'inv-kpis' },
                    kpi('總額', fmt(c.total), '含稅'),
                    kpi('三聯合計', fmt(c.triplicateSum), triNote),
                    kpi('差額（要用二聯補）', fmt(c.gap), '總額 − 三聯'),
                    kpi('二聯', `${n} 張`, `合計 ${fmt(c.selectedSum)} 元`),
                    kpi('超出差額', `${fmt(c.over)} 元`, `規則：超出 ${fmt(c.overMin ?? 1)}～${fmt(c.overMax ?? 0)} 元`, ok ? 'is-good' : 'is-bad'),
                    kpi('二聯＋三聯合計', fmt(c.grandTotal), c.grandTotal % 100 === 0 ? '⚠ 尾數是 00' : '尾數不是 00', ok ? 'is-good' : 'is-bad'),
                ),
                h('div', { class: 'inv-halves' }, ...c.halfStats.map((x) => h('div', { class: 'inv-half' }, h('div', { class: 'inv-half-label' }, `${x.month}月${x.half}`), h('div', { class: 'inv-half-count' }, `${x.count} 張`), h('div', { class: 'inv-half-sum' }, fmt(x.sum))))),
                c.noSolution
                    ? h(
                          'div',
                          { class: 'inv-nosol' },
                          h('b', {}, `⚠ ${c.label}：在規定範圍內找不到合適的組合，這次沒有開二聯。`),
                          h('div', {}, `低於範圍最接近：${c.noSolution.below ? `${c.noSolution.below.count} 張 ${fmt(c.noSolution.below.sum)}（不足 ${fmt(c.gap - c.noSolution.below.sum)}）` : '無'}`),
                          h('div', {}, `高於範圍最接近：${c.noSolution.above ? `${c.noSolution.above.count} 張 ${fmt(c.noSolution.above.sum)}（超出 ${fmt(c.noSolution.above.sum - c.gap)}）` : '無'}`),
                          h('div', {}, '可以到「進階設定」調整超出範圍，或確認總額後再算一次。'),
                      )
                    : null,
            );
        });

        const pend = o.pending.length
            ? h('div', { class: 'inv-pending' }, h('h3', {}, `⚠ 需人工確認（${o.pending.length} 項）`), h('ol', {}, ...o.pending.map((p) => h('li', {}, p))))
            : h('div', { class: 'inv-allgood' }, '✓ 沒有需要人工確認的項目');

        const parts: Kid[] = [
            h('h2', { class: 'inv-result-title' }, `計算結果　${o.periodLabel}`),
            h('p', { class: 'inv-hint' }, `${o.modeLabel}；三聯來源：${o.triplicateSource}`),
            ...calcBlocks,
            pend,
            o.uniformLine ? h('p', { class: 'inv-uniform' }, o.uniformLine) : null,
            h('div', { class: 'inv-result-bar' }, dl, h('span', { class: 'inv-note' }, '（內容與命令列版的「客戶清單」Excel 相同）')),
            h('h3', { class: 'inv-sub-title' }, `二聯清單（${R.selected.length} 張）`),
            entryTable(R.selected, undefined, true),
            more(`未選上的候選（${R.unselected.length} 筆）`, h('p', { class: 'inv-hint' }, '單期金額大於門檻、沒被排除，但這次沒挑到的客戶。'), entryTable(R.unselected)),
            more(
                `排除名單（${R.excluded.length} 筆）`,
                h('p', { class: 'inv-hint' }, '有統編的客戶整期不開二聯；固定排除名單與扣代送客戶也不列入。'),
                entryTable(
                    R.excluded.map((x) => ({ ...x.entry })),
                    [
                        { head: '排除原因', get: (i) => R.excluded[i]!.reason },
                        { head: '統一編號', get: (i) => R.excluded[i]!.uniform?.uniform ?? '' },
                    ],
                ),
            ),
            R.uniformStatus.length
                ? more(
                      `有統編客戶比對（${R.uniformStatus.length} 位）`,
                      h(
                          'div',
                          { class: 'inv-table-wrap' },
                          h(
                              'table',
                              { class: 'inv-table' },
                              h('thead', {}, h('tr', {}, h('th', {}, '客戶編號'), h('th', {}, '結帳檔名稱'), h('th', {}, '統一編號'), h('th', {}, '對帳檔'), h('th', {}, '說明'))),
                              h('tbody', {}, ...R.uniformStatus.map((u) => h('tr', {}, h('td', {}, u.customer.code), h('td', {}, u.customer.name), h('td', {}, u.customer.uniform), h('td', {}, u.matchedBy ? `${u.billNames.join('、')}（用${u.matchedBy}對到）` : '找不到'), h('td', {}, u.hint)))),
                          ),
                      ),
                  )
                : null,
            more(
                '對帳檔讀取結果',
                h('ul', { class: 'inv-warn-list' }, ...o.halvesRead.map((x) => h('li', {}, `${x.label}：客戶 ${x.customers} 家，客戶金額合計 ${fmt(x.customerSum)}；含零售與扣代送淨額共 ${fmt(x.parsedTotal)}，${x.compare}`))),
                ...o.boundNotes.map((b) => h('p', { class: 'inv-hint' }, `張數上限：${b}`)),
            ),
            more(`其他警告（${o.warnings.length} 則）`, o.warnings.length ? h('ul', { class: 'inv-warn-list' }, ...o.warnings.map((w) => h('li', {}, w))) : h('p', {}, '（無）')),
        ];
        box.replaceChildren(...parts.filter((x): x is Node | string => !!x));
        box.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    // ---------- 開始 ----------
    buildTri();
    renderBills();
    renderSources();
    return panel;
}

function bindDropzone(zone: HTMLElement, input: HTMLInputElement, onFiles: (files: File[]) => void): void {
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            input.click();
        }
    });
    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        zone.classList.add('is-dragover');
    });
    zone.addEventListener('dragleave', () => zone.classList.remove('is-dragover'));
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('is-dragover');
        if (e.dataTransfer?.files) onFiles(Array.from(e.dataTransfer.files));
    });
    input.addEventListener('change', () => {
        if (input.files) onFiles(Array.from(input.files));
        input.value = '';
    });
}
