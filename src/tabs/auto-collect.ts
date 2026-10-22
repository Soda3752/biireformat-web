/**
 * 自動催款分頁（對應獨立工具 auto-collect）。
 *
 * 拖入對帳 2.0 的「X月_已自動對帳.xlsx」，跨好幾期累計每位客人的欠款（含部分付款差額），
 * 產出欠款總表（可下載 Excel）與每位客人一段 LINE 催款訊息。只產生訊息，不會傳給客人。
 * 各月明細記在這台電腦的瀏覽器裡，之後每月只要丟新的一份就會自動接上。
 */

import './auto-collect.css';

import {saveAs} from 'file-saver';

import {icon} from '@/ui/icons';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';
import {
    type Choices,
    type CollectResult,
    type Debtor,
    addManualDebt,
    clearMessageEdit,
    compute,
    type ConflictOption,
    filterCustomers,
    type DebtGroup,
    type DebtItem,
    GROUPS,
    type History,
    type ImportFile,
    importFiles,
    type ImportReportRow,
    type Marks,
    localIso,
    type ManualDebt,
    type ManualDebtInput,
    type MonthConflict,
    money,
    sumMessageLines,
    monthLabel,
    removeManualDebt,
    removeMonth,
    setMessageEdit,
    setMonthly,
    settleDebtor,
    unsettle,
    updateManualDebt,
    validateManualDebt,
} from '@/domain/auto-collect/ledger';
import {defaultUiPrefs, expandAll, type SectionKey, type Side, type UiPrefs} from '@/domain/auto-collect/ui-prefs';
import {addReasonChip, defaultReasonChips, removeReasonChip} from '@/domain/auto-collect/reason-chips';
import {defaultTemplates, DEFAULT_TEMPLATES, type MessageTemplates, type TemplateKey} from '@/domain/auto-collect/templates';
import {clearHistory, clearMarks, loadHistory, loadMarks, loadReasonChips, loadTemplates, resetTemplates, saveHistory, saveMarks, saveReasonChips, saveTemplates} from '@/infra/auto-collect-store';
import {buildMessagesText, buildSummaryXlsx} from '@/writers/auto-collect-writer';

const esc = (v: unknown): string =>
    String(v ?? '').replace(/[&<>"']/g, (c) => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[c]!));
const when = (iso: string | undefined): string => esc(String(iso ?? '').replace('T', ' ').slice(0, 16));
const custId = (c: {code: string; name: string}): string => `${c.code}|${c.name}`;
const GROUP_LABEL = Object.fromEntries(GROUPS.map((g) => [g.key, g.label])) as Record<DebtGroup, string>;

export function renderAutoCollectPanel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel auto-collect';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">自動催款</h1>
        <p class="card-subtitle">把「對帳 2.0」做好的 <b>X月_已自動對帳.xlsx</b> 拖進來，自動算出每位客人跨好幾期累計的欠款（包含只付一部分的差額），產生可以貼到 LINE 的催款訊息。</p>
      </header>
      <div class="notice-banner">這裡<b>不會傳任何訊息給客人</b>，請自己按「複製」後貼到 LINE。</div>

      <div class="card-section">
        <div class="card-section-label">放入檔案</div>
        <div class="dropzone" role="button" tabindex="0" data-role="drop">
          <span class="dropzone-icon">${icon('upload', 28)}</span>
          <div class="dropzone-title">把 X月_已自動對帳.xlsx 拖到這裡，或點一下選擇檔案</div>
          <div class="dropzone-hint">第一次可以一次放好幾個月；之後每個月只要放新的那一份。同一個月放了好幾份、或這個月之前已經放過，會先請你選要用哪一份。</div>
          <input type="file" accept=".xlsx" multiple hidden data-role="file">
        </div>
        <div class="ac-report" data-role="report"></div>
      </div>

      <div class="card-section">
        <div class="ac-row-between">
          <div class="card-section-label">已記住的月份（存在這台電腦的瀏覽器裡）</div>
          <button type="button" class="btn btn-danger-outline btn-sm" data-role="clear-all" hidden>${icon('trash', 14)}<span>清除全部記錄</span></button>
        </div>
        <div data-role="months"></div>
      </div>
    </div>

    <div class="card" data-role="result-card">
      <header class="card-header ac-row-between">
        <h2 class="card-title">欠款總表</h2>
        <div class="ac-actions">
          <button type="button" class="btn btn-primary" data-role="dl-xlsx">${icon('download', 14)}<span>下載總表 Excel</span></button>
          <button type="button" class="btn btn-secondary" data-role="dl-txt">${icon('download', 14)}<span>下載全部訊息（文字檔）</span></button>
        </div>
      </header>
      <div class="metric-grid" data-role="stats"></div>
      <div class="ac-filter-hint" data-role="filter-hint" hidden></div>
      <input class="app-form-input ac-search" type="search" placeholder="搜尋代號或店名…" autocomplete="off" data-role="search">
      <div class="ac-table-wrap">
        <table class="ac-table">
          <thead><tr><th>代號</th><th>客人</th><th>欠哪幾期</th><th class="ac-num">合計</th><th>備註</th><th></th></tr></thead>
          <tbody data-role="tbody"></tbody>
        </table>
      </div>
    </div>

    <div class="card" data-role="settled-card">
      <header class="card-header">
        <button type="button" class="ac-fold-btn ac-fold-card" data-role="settled-fold" data-fold="table:settled"></button>
        <p class="card-subtitle">這裡的欠款不算進上面的總數、也不會出現在下載的檔案和訊息裡。按「取消」就會恢復。</p>
      </header>
      <div data-role="settled"></div>
    </div>

    <div class="card">
      <details class="ac-tpl" data-role="tpl">
        <summary class="ac-tpl-summary">訊息範本（改催款訊息最後那一句的語氣）</summary>
        <p class="card-subtitle">每則催款訊息最後一句，會依客人欠幾期自動換語氣。改了之後下面所有自動產生的訊息會馬上更新；已經手動編輯過的訊息不受影響。這些設定只存在這台電腦的瀏覽器，「清除全部記錄」不會清掉它。</p>
        <p class="card-subtitle">小技巧：句子裡寫 <b>{期數}</b>（連大括號一起打），傳出去時會自動換成這位客人欠的期數，例如「目前已累積 <b>4</b> 期帳款未付」。三句都可以用。</p>
        <label class="ac-tpl-field"><span>欠 4 期以上</span><textarea class="app-form-input" rows="2" data-tpl="four_plus"></textarea></label>
        <label class="ac-tpl-field"><span>欠 2～3 期</span><textarea class="app-form-input" rows="2" data-tpl="two_three"></textarea></label>
        <label class="ac-tpl-field"><span>月結：本月該付</span><textarea class="app-form-input" rows="2" data-tpl="monthly_due"></textarea></label>
        <div class="ac-tpl-bar"><button type="button" class="btn btn-secondary btn-sm" data-role="tpl-reset">恢復預設</button></div>
      </details>
    </div>

    <div class="card">
      <header class="card-header">
        <h2 class="card-title">每位客人的催款訊息</h2>
        <p class="card-subtitle">欠 2 期以上、或月結客人「本月該付」上個月帳款的才有訊息（其餘欠 1 期的只在上面總表顯示、不催）。按「複製」後到 LINE 貼上。複製過的會變灰（重新整理頁面就會清掉標記）。有<span class="ac-warn-tag">注意</span>的請先確認再傳。</p>
      </header>
      <div class="ac-messages" data-role="messages"></div>
    </div>
  `;

    const $ = <T extends HTMLElement = HTMLElement>(role: string): T => panel.querySelector<T>(`[data-role="${role}"]`)!;
    const drop = $('drop');
    const fileInput = $<HTMLInputElement>('file');
    const reportHost = $('report');
    const search = $<HTMLInputElement>('search');

    let history: History = loadHistory();
    let marks: Marks = loadMarks(history);
    let templates: MessageTemplates = loadTemplates();
    let result: CollectResult = compute(history, marks, {}, new Date(), templates);
    const copied = new Set<string>();
    const editing = new Map<string, string>(); // 正在編輯訊息的客人 → 文字框目前內容（重畫清單時保留）
    let ui: UiPrefs = defaultUiPrefs(); // 各組收合狀態：只在記憶體，每次載入一律全部收起（不記住）
    let onlyNotes = false; // 點「需要先確認」格子 → 只顯示有注意事項的客人
    let uiBeforeFilter: UiPrefs | null = null; // 篩選期間暫時全展開，取消時還原成這份

    const setOnlyNotes = (on: boolean) => {
        if (on === onlyNotes) return;
        onlyNotes = on;
        if (on) {
            uiBeforeFilter = ui;
            ui = expandAll(ui);
        } else if (uiBeforeFilter) {
            ui = uiBeforeFilter;
            uiBeforeFilter = null;
        }
    };

    /** 重新匯入後重畫：開合回到全部收起（篩選中則取消篩選時才回到收起）。 */
    const resetFolds = () => {
        if (onlyNotes) uiBeforeFilter = defaultUiPrefs();
        else ui = defaultUiPrefs();
    };

    const persist = (h: History): boolean => {
        try {
            saveHistory(h);
            history = h;
            return true;
        } catch (err) {
            console.error('[auto-collect] save failed', err);
            showToast({variant: 'error', title: '存不進瀏覽器', message: '瀏覽器的儲存空間可能滿了，請先移除用不到的月份再試一次。'});
            return false;
        }
    };

    const persistMarks = (m: Marks): boolean => {
        try {
            saveMarks(m);
            marks = m;
            return true;
        } catch (err) {
            console.error('[auto-collect] save marks failed', err);
            showToast({variant: 'error', title: '存不進瀏覽器', message: '瀏覽器的儲存空間可能滿了，已結清標記或手動欠款沒有存下來。'});
            return false;
        }
    };

    // ---------- 復原提示（固定在畫面上方，約 9 秒後消失）

    const undoHost = document.createElement('div');
    undoHost.className = 'ac-undo-host';
    undoHost.setAttribute('role', 'status');
    document.body.appendChild(undoHost);
    // 離開頁面（分頁面板被拿掉）後清掉提示
    new MutationObserver(() => {
        if (!panel.isConnected) undoHost.remove();
    }).observe(document.body, {childList: true, subtree: true});

    const showUndoBar = (text: string, onUndo: () => void) => {
        const bar = document.createElement('div');
        bar.className = 'ac-undo';
        bar.innerHTML = `<span>${esc(text)}</span><button type="button" class="ac-undo-btn">復原</button>`;
        const timer = setTimeout(() => bar.remove(), 9000);
        bar.querySelector('button')!.addEventListener('click', () => {
            clearTimeout(timer);
            bar.remove();
            onUndo();
        });
        undoHost.appendChild(bar);
    };

    // ---------- 匯入

    const setReport = (html: string) => (reportHost.innerHTML = html);

    const handleFiles = async (list: File[]) => {
        const picked = list.filter((f) => /\.xlsx$/i.test(f.name) && !f.name.startsWith('~$'));
        if (!picked.length) {
            showToast({variant: 'warning', message: '請選對帳 2.0 做好的 .xlsx 檔'});
            return;
        }
        setReport('<div class="muted">讀取中…</div>');
        try {
            const files: ImportFile[] = await Promise.all(picked.map(async (f) => ({
                name: f.name, mtime: f.lastModified, data: await f.arrayBuffer(),
            })));
            let choices: Choices = {};
            for (;;) {
                const res = await importFiles(files, history, {choices});
                if (res.conflicts.length) {
                    setReport('<div class="muted">請在跳出的視窗選擇要用哪一份檔案…</div>');
                    const answer = await askPick(res.conflicts);
                    if (!answer) {
                        setReport('<ul class="ac-report-list"><li class="is-skipped">已取消，這次放的檔案都沒有匯入。</li></ul>');
                        return;
                    }
                    choices = {...choices, ...answer};
                    continue;
                }
                if (res.history !== history && !persist(res.history)) return;
                renderReport(res.report);
                resetFolds();
                render();
                return;
            }
        } catch (err) {
            console.error(err);
            setReport(`<ul class="ac-report-list"><li class="is-error">匯入失敗：${esc(err instanceof Error ? err.message : err)}</li></ul>`);
        }
    };

    const renderReport = (rows: ImportReportRow[]) => {
        const label = {used: '✓ 採用', skipped: '沒採用', error: '✗ 無法讀取'};
        setReport(`<ul class="ac-report-list">${rows.map((r) =>
            `<li class="is-${r.status}">${label[r.status]}：${esc(r.name)} — ${esc(r.reason)}</li>`).join('')}</ul>`);
    };

    drop.addEventListener('click', () => fileInput.click());
    drop.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            fileInput.click();
        }
    });
    fileInput.addEventListener('change', () => {
        const files = Array.from(fileInput.files ?? []);
        fileInput.value = '';
        if (files.length) void handleFiles(files);
    });
    drop.addEventListener('dragover', (e) => {
        e.preventDefault();
        drop.classList.add('is-dragover');
    });
    drop.addEventListener('dragleave', () => drop.classList.remove('is-dragover'));
    drop.addEventListener('drop', (e) => {
        e.preventDefault();
        drop.classList.remove('is-dragover');
        const files = Array.from(e.dataTransfer?.files ?? []);
        if (files.length) void handleFiles(files);
    });

    // ---------- 已記住的月份

    $('months').addEventListener('click', async (e) => {
        const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-rm]');
        if (!btn) return;
        const ok = await confirmDialog('移除月份', `確定要把「${btn.dataset.label}」從記錄中移除嗎？之後可以再把檔案放進來。`, '移除');
        if (ok && persist(removeMonth(history, btn.dataset.rm!))) render();
    });

    $('clear-all').addEventListener('click', async () => {
        const hasMarks = marks.settled.length > 0 || marks.manual_debts.length > 0 || Object.keys(marks.monthly).length > 0
            || Object.keys(marks.edits).length > 0;
        const ans = await clearDialog(hasMarks);
        if (!ans) return;
        clearHistory();
        history = loadHistory();
        if (ans.clearMarks) {
            clearMarks();
            marks = loadMarks(history);
        }
        copied.clear();
        editing.clear();
        setReport('');
        render();
        showToast({variant: 'success', message: hasMarks && !ans.clearMarks ? '已清除全部月份記錄（已結清標記、手動加的欠款、月結設定、手動修改過的訊息保留）' : '已清除全部記錄'});
    });

    // ---------- 下載

    const stamp = (): string => localIso(new Date()).slice(0, 16).replace('T', '_').replace(/[-:]/g, '');

    $('dl-xlsx').addEventListener('click', async () => {
        if (!result.months.length) return showToast({variant: 'warning', message: '還沒有資料，請先放入檔案'});
        saveAs(await buildSummaryXlsx(result), `催款總表_${stamp()}.xlsx`);
    });
    $('dl-txt').addEventListener('click', () => {
        if (!result.months.length) return showToast({variant: 'warning', message: '還沒有資料，請先放入檔案'});
        const text = buildMessagesText(result, localIso(new Date()).slice(0, 16).replace('T', ' '));
        saveAs(new Blob([text], {type: 'text/plain;charset=utf-8'}), `催款訊息_${stamp()}.txt`);
    });

    // ---------- 複製

    const copyText = async (text: string) => {
        try {
            // 視窗沒取得焦點時 clipboard API 可能一直不回應，超過 1.5 秒改用舊方法
            await Promise.race([navigator.clipboard.writeText(text),
                new Promise((_, fail) => setTimeout(() => fail(new Error('timeout')), 1500))]);
        } catch {
            const ta = document.createElement('textarea');
            ta.value = text;
            document.body.appendChild(ta);
            ta.select();
            document.execCommand('copy');
            ta.remove();
        }
    };

    panel.addEventListener('click', async (e) => {
        const t = e.target as HTMLElement;
        const fold = t.closest<HTMLElement>('[data-fold]');
        if (fold) {
            const [side, key] = fold.dataset.fold!.split(':') as [Side, SectionKey];
            setFolded(side, key, !ui[side][key]);
            return;
        }
        const jump = t.closest<HTMLElement>('[data-jump]');
        if (jump) {
            jumpTo('messages', jump.dataset.jump!);
            return;
        }
        const back = t.closest<HTMLElement>('[data-back]');
        if (back) {
            jumpTo('table', back.dataset.back!);
            return;
        }
        const monthlyBtn = t.closest<HTMLButtonElement>('[data-monthly]');
        if (monthlyBtn) {
            const c = result.customers.find((x) => custId(x) === monthlyBtn.dataset.monthly);
            if (c && persistMarks(setMonthly(marks, c, !c.monthly))) {
                render();
                showToast({variant: 'success', message: c.monthly ? `已取消 ${c.name} 的月結` : `已將 ${c.name} 設為月結（同月上下期算 1 期）`, duration: 2500});
            }
            return;
        }
        const settleBtn = t.closest<HTMLButtonElement>('[data-settle]');
        if (settleBtn) {
            const c = result.customers.find((x) => custId(x) === settleBtn.dataset.settle);
            const ans = c && await settleDialog(c);
            if (c && ans) {
                const next = settleDebtor(marks, c, ans.reason);
                const id = next.settled[next.settled.length - 1]?.id;
                if (next !== marks && persistMarks(next)) {
                    render();
                    showUndoBar(`已將 ${c.name} 標為已結清`, () => {
                        if (persistMarks(unsettle(marks, id))) render();
                    });
                }
            }
            return;
        }
        const undo = t.closest<HTMLButtonElement>('[data-unsettle]');
        if (undo) {
            if (persistMarks(unsettle(marks, undo.dataset.unsettle!))) render();
            return;
        }
        const addBtn = t.closest<HTMLButtonElement>('[data-adddebt]');
        if (addBtn) {
            const c = result.customers.find((x) => custId(x) === addBtn.dataset.adddebt);
            const ans = c && await debtDialog(c, history);
            if (c && ans && persistMarks(addManualDebt(marks, c, ans))) {
                render();
                showToast({variant: 'success', message: `已替 ${c.name} 加上 ${periodText(ans.month_key, ans.half)} ${money(ans.amount)} 元`});
            }
            return;
        }
        const editBtn = t.closest<HTMLButtonElement>('[data-editdebt]');
        if (editBtn) {
            const d = marks.manual_debts.find((x) => x.id === editBtn.dataset.editdebt);
            const c = d && result.customers.find((x) => custId(x) === editBtn.dataset.cust);
            const ans = d && c && await debtDialog(c, history, d);
            if (d && ans && persistMarks(updateManualDebt(marks, d.id, ans))) render();
            return;
        }
        const delBtn = t.closest<HTMLButtonElement>('[data-deldebt]');
        if (delBtn) {
            const d = marks.manual_debts.find((x) => x.id === delBtn.dataset.deldebt);
            if (d && await confirmDialog('移除手動欠款', `確定要移除 ${d.name} 的「${periodText(d.month_key, d.half)} ${money(d.amount)} 元${d.reason ? `（${d.reason}）` : ''}」嗎？`, '移除')
                && persistMarks(removeManualDebt(marks, d.id))) render();
            return;
        }
        const msgBtn = t.closest<HTMLButtonElement>('[data-msg]');
        if (msgBtn) {
            const who = msgBtn.dataset.cust!;
            const c = result.customers.find((x) => custId(x) === who);
            if (!c) return;
            const act = msgBtn.dataset.msg;
            if (act === 'edit') {
                editing.set(who, c.message);
            } else if (act === 'cancel') {
                editing.delete(who);
            } else if (act === 'save') {
                const text = editing.get(who) ?? c.message;
                if (!text.trim()) return showToast({variant: 'warning', message: '訊息不能是空白；想回到自動產生的請按「還原成自動產生」'});
                if (!persistMarks(setMessageEdit(marks, c, text))) return;
                editing.delete(who);
                render();
                showToast({variant: 'success', message: `已儲存 ${c.name} 的訊息`, duration: 2000});
                return;
            } else if (act === 'revert') {
                const prev = marks;
                if (persistMarks(clearMessageEdit(marks, c))) {
                    render();
                    showUndoBar(`已把 ${c.name} 的訊息還原成自動產生`, () => {
                        if (persistMarks(prev)) render();
                    });
                }
                return;
            }
            renderList();
            if (act === 'edit') panel.querySelector<HTMLTextAreaElement>(`textarea[data-msg-input="${CSS.escape(who)}"]`)?.focus();
            return;
        }
        const btn = t.closest<HTMLButtonElement>('[data-copy]');
        if (!btn) return;
        const c = result.chase.find((x) => custId(x) === btn.dataset.copy);
        if (!c) return;
        await copyText(c.message);
        copied.add(custId(c));
        showToast({variant: 'success', message: `已複製 ${c.name} 的訊息，請到 LINE 貼上`, duration: 2000});
        renderList();
    });

    search.addEventListener('input', () => renderList());

    $('messages').addEventListener('input', (e) => {
        const ta = e.target as HTMLTextAreaElement;
        if (ta.dataset?.msgInput) editing.set(ta.dataset.msgInput, ta.value);
    });

    $('stats').addEventListener('click', (e) => {
        const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-role="only-notes"]');
        if (!btn || btn.disabled) return;
        setOnlyNotes(!onlyNotes);
        render();
    });

    // ---------- 訊息範本（設定，存 localStorage，改了即時更新所有自動訊息）

    const tplInputs = Array.from(panel.querySelectorAll<HTMLTextAreaElement>('[data-tpl]'));
    const fillTemplates = () => {
        for (const ta of tplInputs) ta.value = templates[ta.dataset.tpl as TemplateKey];
    };
    fillTemplates();
    for (const ta of tplInputs) {
        ta.addEventListener('input', () => {
            const key = ta.dataset.tpl as TemplateKey;
            // 清空時暫用預設句（框裡先留空讓她繼續打字）
            templates = {...templates, [key]: ta.value.trim() ? ta.value : DEFAULT_TEMPLATES[key]};
            try {
                saveTemplates(templates);
            } catch (err) {
                console.error('[auto-collect] save templates failed', err);
                showToast({variant: 'error', title: '存不進瀏覽器', message: '訊息範本沒有存下來，重新整理後會回到上一次的版本。'});
            }
            render();
        });
    }
    $('tpl-reset').addEventListener('click', () => {
        templates = defaultTemplates();
        resetTemplates();
        fillTemplates();
        render();
        showToast({variant: 'success', message: '訊息範本已恢復預設', duration: 2500});
    });

    // ---------- 收合與跳轉

    const setFolded = (side: Side, key: SectionKey, folded: boolean) => {
        ui = {...ui, [side]: {...ui[side], [key]: folded}};
        if (key === 'settled') renderSettled();
        else renderList();
    };

    /** 捲到另一邊該客人那一格並閃一下；那一組是收起的就先展開。 */
    const jumpTo = (side: Side, who: string) => {
        const c = result.customers.find((x) => custId(x) === who);
        if (!c) return;
        const key: SectionKey = c.group;
        if (ui[side][key]) setFolded(side, key, false);
        const attr = side === 'messages' ? 'data-card' : 'data-row';
        const el = Array.from(panel.querySelectorAll<HTMLElement>(`[${attr}]`)).find((x) => x.getAttribute(attr) === who);
        if (!el) {
            showToast({variant: 'info', message: side === 'messages' ? '這位沒有催款訊息' : '總表裡找不到這位（可能被搜尋或篩選隱藏了）', duration: 2500});
            return;
        }
        el.scrollIntoView({behavior: 'smooth', block: 'center'});
        el.classList.remove('is-flash');
        void el.offsetWidth; // 重新觸發動畫
        el.classList.add('is-flash');
        setTimeout(() => el.classList.remove('is-flash'), 1800);
    };

    // ---------- 畫面

    const render = () => {
        result = compute(history, marks, {}, new Date(), templates);
        const {months} = result;
        $('clear-all').hidden = !months.length;
        $('months').innerHTML = months.length
            ? `<div class="ac-table-wrap"><table class="ac-table">
                <thead><tr><th>月份</th><th>採用的檔案</th><th>檔案修改時間</th><th>匯入時間</th><th class="ac-num">筆數</th><th></th></tr></thead>
                <tbody>${months.map((m) => `<tr><td><b>${esc(m.label)}</b>${m.skipped.map((s) =>
                    `<div class="ac-skip" title="${s.reason === 'same_as_prev' ? `${esc(s.ref ? monthLabel(s.ref) : '')}下期幾乎一樣，判定還沒結帳` : '整期沒有任何人付款，判定還沒結帳'}">${s.half}期未結帳，未計入</div>`).join('')}</td><td>${esc(m.source)}</td>
                  <td>${when(m.file_mtime)}</td><td>${when(m.imported_at)}</td><td class="ac-num">${m.count}</td>
                  <td><button type="button" class="btn btn-secondary btn-xs" data-rm="${esc(m.key)}" data-label="${esc(m.label)}">移除</button></td></tr>`).join('')}
                </tbody></table></div>`
            : '<div class="ac-empty">還沒有資料。請把對帳 2.0 的「X月_已自動對帳.xlsx」放到上面。</div>';

        const warn = result.chase.filter((c) => c.notes.length).length;
        if (!warn) setOnlyNotes(false);
        const metric = (label: string, value: string) =>
            `<div class="metric-label">${label}</div><div class="metric-value">${value}</div>`;
        $('stats').innerHTML = [
            `<div class="metric">${metric('位客人要催款（欠 2 期以上＋月結本月該付）', `${result.chase.length}`)}</div>`,
            `<div class="metric">${metric('欠款合計', `${money(result.chase_total)} 元`)}</div>`,
            `<button type="button" class="metric ac-metric-toggle${onlyNotes ? ' is-on' : ''}" data-role="only-notes"
               aria-pressed="${onlyNotes}" ${warn ? `title="${onlyNotes ? '再點一次顯示全部' : '點一下只顯示這幾位'}"` : 'disabled'}>
               ${metric('位需要先確認（有注意事項）', `${warn}`)}${warn && !onlyNotes ? '<div class="ac-metric-cta">點一下只看這幾位</div>' : ''}</button>`,
        ].join('');
        const hint = $('filter-hint');
        hint.hidden = !onlyNotes;
        hint.innerHTML = onlyNotes ? `目前只顯示<b>需確認的 ${warn} 位</b>，再點一次「需要先確認」的格子顯示全部。` : '';
        renderSettled();
        renderList();
    };

    const renderSettled = () => {
        const groups = result.settled;
        $('settled-card').hidden = !groups.length;
        const sf = $('settled-fold');
        const folded = ui.table.settled;
        sf.setAttribute('aria-expanded', String(!folded));
        sf.innerHTML = `<span class="ac-fold-icon" aria-hidden="true">${folded ? '▸' : '▾'}</span><span class="ac-fold-title">已結清（不催）：${groups.length} 位，合計 ${money(result.settled_total)} 元</span>`;
        $('settled').hidden = folded;
        $('settled').innerHTML = groups.length && !folded
            ? `<div class="ac-table-wrap"><table class="ac-table">
                <thead><tr><th>代號</th><th>客人</th><th>哪幾期</th><th class="ac-num">金額</th><th>原因</th><th>標記日期</th><th></th></tr></thead>
                <tbody>${groups.map((g) => `<tr><td>${esc(g.code)}</td><td><b>${esc(g.name)}</b></td>
                  <td class="ac-periods">${g.lines.length
                    ? g.lines.map((i) => `${esc(i.period)}${i.manual?.reason ? `（${esc(i.manual.reason)}）` : ''} ${money(i.owed)}`).join('<br>')
                    : '<span class="muted">目前沒有符合的期別</span>'}</td>
                  <td class="ac-num">${money(g.total)}</td>
                  <td>${esc(g.reason || '（未填）')}</td>
                  <td>${when(g.marked_at).slice(0, 10)}</td>
                  <td><button type="button" class="btn btn-secondary btn-xs" data-unsettle="${esc(g.id)}">取消</button></td></tr>`).join('')}
                </tbody></table></div>
              <div class="ac-settled-total">不催的金額合計 ${money(result.settled_total)} 元</div>`
            : '';
    };

    const itemHtml = (c: Debtor, i: DebtItem): string => {
        if (i.manual) {
            const label = `${esc(i.period)} ${money(i.owed)}${i.manual.reason ? `（${esc(i.manual.reason)}）` : ''}`;
            return `<span class="ac-manual"><span class="ac-tag">手動</span>${label}
              <button type="button" class="ac-link" data-editdebt="${esc(i.manual.id)}" data-cust="${esc(custId(c))}">修改</button>
              <button type="button" class="ac-link" data-deldebt="${esc(i.manual.id)}">移除</button></span>`;
        }
        return i.partial
            ? `<span class="ac-partial" title="帳單 ${money(i.amt)}，已收 ${money(i.paid)}">${esc(i.period)} ${money(i.owed)}（只付一部分）</span>`
            : `${esc(i.period)} ${money(i.owed)}`;
    };

    const monthlyTag = (c: Debtor): string => (c.monthly
        ? `<span class="ac-tag-monthly" title="${c.monthly_auto ? '備註有「月」，自動認成月結' : '手動設為月結'}；同月上下期算 1 期">月結</span>` : '');
    const pendingInfo = (c: Debtor): string => {
        const nowY = new Date().getFullYear();
        const byMonth = new Map<string, number>();
        for (const i of c.pending) byMonth.set(i.month_key, (byMonth.get(i.month_key) ?? 0) + i.owed);
        return [...byMonth].map(([k, amt]) => {
            const m = `${+k.slice(0, 4) === nowY ? '' : `${+k.slice(0, 4)}年`}${+k.slice(5)}月`;
            return `<div class="ac-pending" title="月結客人這個月還沒結完，不算欠款、不催款">${m}未到期 ${money(amt)}</div>`;
        }).join('');
    };
    const periodsInfo = (c: Debtor): string => `<div class="ac-periods-count">${c.periods ? `欠 ${c.periods} 期` : '只有手動欠款'}</div>`;
    const monthlyBtn = (c: Debtor): string =>
        `<button type="button" class="btn btn-secondary btn-xs" data-monthly="${esc(custId(c))}">${c.monthly ? '取消月結' : '設為月結'}</button>`;

    /** 分組標題列：整條可點來收合／展開，收起時仍看得到人數與金額。 */
    const groupHead = (side: Side, key: DebtGroup, list: Debtor[]): string => {
        const total = list.reduce((t, c) => t + c.total, 0);
        const note = key === 'one' ? '　不催款，也不算進上面的欠款總額' : '';
        const folded = ui[side][key];
        return `<button type="button" class="ac-fold-btn" data-fold="${side}:${key}" aria-expanded="${!folded}"
            title="${folded ? '點一下展開' : '點一下收起'}"><span class="ac-fold-icon" aria-hidden="true">${folded ? '▸' : '▾'}</span>
            <span class="ac-fold-title">${esc(GROUP_LABEL[key])}：${list.length} 位，合計 ${money(total)} 元${note}</span></button>`;
    };

    const msgBody = (c: Debtor): string => {
        const who = esc(custId(c));
        if (editing.has(custId(c))) {
            return `<textarea class="app-form-input ac-msg-edit" rows="${Math.min(16, c.message.split('\n').length + 2)}" data-msg-input="${who}">${esc(editing.get(custId(c)))}</textarea>
              <div class="ac-msg-edit-bar"><span class="muted">只改訊息文字，不影響上面的金額與合計；儲存時會自動重算訊息裡的「合計」。</span>
                <span class="ac-btns"><button type="button" class="btn btn-secondary btn-xs" data-msg="cancel" data-cust="${who}">取消</button>
                <button type="button" class="btn btn-primary btn-xs" data-msg="save" data-cust="${who}">儲存</button></span></div>`;
        }
        const stale = c.edit?.stale
            ? `<div class="ac-stale">⚠ 欠款已變動，請檢查
                <details><summary>看目前自動產生的版本</summary><pre class="ac-msg-text ac-msg-auto">${esc(c.auto_message)}</pre></details></div>` : '';
        const sum = c.edit ? sumMessageLines(c.message) : null;
        const mismatch = sum !== null && sum !== c.total
            ? `<div class="ac-stale">⚠ 訊息合計 ${money(sum)} 元 與總表 ${money(c.total)} 元 不同，請檢查</div>` : '';
        return `${stale}${mismatch}<pre class="ac-msg-text">${esc(c.message)}</pre>`;
    };
    const msgTools = (c: Debtor): string => `<div class="ac-msg-tools">${c.edit
        ? `<span class="ac-tag-edited">已手動修改</span>
           <button type="button" class="ac-link" data-msg="revert" data-cust="${esc(custId(c))}">還原成自動產生</button>` : ''}
        ${editing.has(custId(c)) ? '' : `<button type="button" class="ac-link" data-msg="edit" data-cust="${esc(custId(c))}">編輯</button>`}</div>`;

    const renderList = () => {
        const list = filterCustomers(result.customers, search.value, onlyNotes);
        const byGroup = (key: DebtGroup): Debtor[] => list.filter((c) => c.group === key);
        const nameCell = (c: Debtor): string => (c.group === 'one'
            ? `<b>${esc(c.name)}</b>`
            : `<button type="button" class="ac-name-link" data-jump="${esc(custId(c))}" title="跳到下面這位的催款訊息"><b>${esc(c.name)}</b></button>`);
        const row = (c: Debtor): string => `<tr data-row="${esc(custId(c))}" class="${copied.has(custId(c)) ? 'is-done' : ''}">
                <td>${esc(c.code)}</td><td>${nameCell(c)}${monthlyTag(c)}${periodsInfo(c)}</td>
                <td class="ac-periods">${c.lines.map((i) => itemHtml(c, i)).join('<br>')}${pendingInfo(c)}</td>
                <td class="ac-num"><b>${money(c.total)}</b></td>
                <td>${c.notes.map((n) => `<div class="ac-note">${esc(n)}</div>`).join('')}</td>
                <td class="ac-btns">${c.group !== 'one' ? `<button type="button" class="btn btn-secondary btn-xs" data-copy="${esc(custId(c))}">複製訊息</button>` : ''}
                  <button type="button" class="btn btn-secondary btn-xs" data-adddebt="${esc(custId(c))}">加一筆欠款</button>
                  ${monthlyBtn(c)}
                  <button type="button" class="btn btn-secondary btn-xs" data-settle="${esc(custId(c))}">已結清</button></td></tr>`;
        const rows = GROUPS.flatMap((g) => {
            const part = byGroup(g.key);
            if (!part.length) return [];
            const head = `<tr class="ac-group-row${g.key === 'one' ? ' is-light' : ''}"><td colspan="6">${groupHead('table', g.key, part)}</td></tr>`;
            return ui.table[g.key] ? [head] : [head, ...part.map(row)];
        });
        $('tbody').innerHTML = rows.length
            ? rows.join('')
            : `<tr><td colspan="6" class="ac-empty">${result.months.length
                ? (onlyNotes ? '需確認的客人裡沒有符合搜尋的' : '沒有符合的欠款客人') : '還沒有資料'}</td></tr>`;
        const msgs = GROUPS.filter((g) => g.key !== 'one').flatMap((g) => {
            const part = byGroup(g.key);
            if (!part.length) return [];
            const head = `<div class="ac-msg-group">${groupHead('messages', g.key, part)}</div>`;
            if (ui.messages[g.key]) return [head];
            return [head, ...part.map((c) => `<div class="ac-msg ${copied.has(custId(c)) ? 'is-done' : ''}" data-card="${esc(custId(c))}">
            <div class="ac-msg-head"><b>${esc(c.code)} ${esc(c.name)}${monthlyTag(c)}${c.notes.length ? '<span class="ac-warn-tag">注意</span>' : ''}</b>
              <span class="ac-btns"><button type="button" class="btn btn-secondary btn-xs" data-adddebt="${esc(custId(c))}">加一筆欠款</button>
              ${monthlyBtn(c)}
              <button type="button" class="btn btn-secondary btn-xs" data-settle="${esc(custId(c))}">已結清</button>
              <button type="button" class="btn btn-secondary btn-xs" data-back="${esc(custId(c))}" title="回到上面總表這位的那一列">↑ 回到總表</button>
              <button type="button" class="btn btn-primary btn-xs" data-copy="${esc(custId(c))}">複製</button></span></div>
            ${c.notes.map((n) => `<div class="ac-note">⚠ ${esc(n)}</div>`).join('')}
            ${msgTools(c)}${msgBody(c)}</div>`)];
        });
        $('messages').innerHTML = msgs.join('') || '<div class="ac-empty">沒有訊息</div>';
    };

    render();
    return panel;
}

// ---------- 對話框

function openModal(title: string, bodyHtml: string, footerHtml: string, extraClass = ''): {
    dialog: HTMLElement;
    close: () => void;
} {
    const backdrop = document.createElement('div');
    backdrop.className = 'app-modal-backdrop';
    backdrop.setAttribute('role', 'presentation');
    backdrop.innerHTML = `
      <div class="app-modal ${extraClass}" role="dialog" aria-modal="true">
        <header class="app-modal-header"><h2 class="app-modal-title">${esc(title)}</h2></header>
        <div class="app-modal-body">${bodyHtml}</div>
        <footer class="app-modal-footer">${footerHtml}</footer>
      </div>`;
    document.body.appendChild(backdrop);
    return {dialog: backdrop.querySelector<HTMLElement>('.app-modal')!, close: () => backdrop.remove()};
}

function confirmDialog(title: string, message: string, okLabel: string): Promise<boolean> {
    const {dialog, close} = openModal(title, `<p>${esc(message)}</p>`,
        `<button type="button" class="btn btn-secondary" data-act="cancel">取消</button>
         <button type="button" class="btn btn-primary" data-act="ok">${esc(okLabel)}</button>`);
    return new Promise((done) => {
        dialog.addEventListener('click', (e) => {
            const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
            if (!act) return;
            close();
            done(act === 'ok');
        });
    });
}

/** 清除全部記錄：有已結清標記時多問一個「要不要一起清」（預設保留）。 */
function clearDialog(hasMarks: boolean): Promise<{clearMarks: boolean} | null> {
    const body = `<p>確定要清除這台電腦記住的所有月份嗎？清除後要重新放入各月的檔案才能再算欠款。</p>`
        + (hasMarks ? `<label class="ac-check"><input type="checkbox" data-role="also-marks">
            <span>同時清除「已結清」標記（沒勾的話標記會保留，之後放回檔案仍然有效）</span></label>` : '');
    const {dialog, close} = openModal('清除全部記錄', body,
        `<button type="button" class="btn btn-secondary" data-act="cancel">取消</button>
         <button type="button" class="btn btn-primary" data-act="ok">全部清除</button>`);
    return new Promise((done) => {
        dialog.addEventListener('click', (e) => {
            const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
            if (!act) return;
            const also = dialog.querySelector<HTMLInputElement>('[data-role="also-marks"]')?.checked ?? false;
            close();
            done(act === 'ok' ? {clearMarks: also} : null);
        });
    });
}

/** 「已結清」視窗：選填原因。 */
function settleDialog(c: Debtor): Promise<{reason: string} | null> {
    const periods = c.lines.map((i) => `${esc(i.period)}${i.manual?.reason ? `（${esc(i.manual.reason)}）` : ''} ${money(i.owed)} 元`).join('<br>');
    const body = `<p><b>${esc(c.code)} ${esc(c.name)}</b>　目前欠款 ${money(c.total)} 元</p>
      <div class="ac-settle-periods">${periods}</div>
      <label class="ac-field"><span>原因（選填）</span>
        <input class="app-form-input" type="text" data-role="reason" placeholder="例如：已補款、回饋金自扣" autocomplete="off"></label>
      <div class="ac-chips" data-role="chips"></div>
      <div class="ac-settle-note">確定後，上面這 ${c.lines.length} 筆不再算欠款。之後新的月份沒付仍會出現；按錯可以在畫面上方按「復原」。</div>`;
    const {dialog, close} = openModal('標記已結清', body,
        `<button type="button" class="btn btn-secondary" data-act="cancel">取消</button>
         <button type="button" class="btn btn-primary" data-act="ok">確定</button>`);
    const reason = dialog.querySelector<HTMLInputElement>('[data-role="reason"]')!;
    const chipBox = dialog.querySelector<HTMLElement>('[data-role="chips"]')!;
    let chips = loadReasonChips();
    let adding = false;
    const renderChips = (): void => {
        chipBox.innerHTML = chips.map((t) => `<span class="ac-chip">
            <button type="button" class="btn btn-secondary btn-xs ac-chip-main" data-chip="${esc(t)}">${esc(t)}</button><button type="button" class="ac-chip-x" data-chip-del="${esc(t)}" title="刪除這個按鈕" aria-label="刪除按鈕：${esc(t)}">×</button></span>`).join('')
            + (adding
                ? `<span class="ac-chip-add"><input class="app-form-input" type="text" data-role="chip-new" maxlength="20" placeholder="輸入新的原因" autocomplete="off">
                    <button type="button" class="btn btn-primary btn-xs" data-chip-ok>加入</button>
                    <button type="button" class="btn btn-secondary btn-xs" data-chip-cancel>取消</button></span>`
                : `<button type="button" class="btn btn-secondary btn-xs" data-chip-new>＋ 新增</button>`)
            + `<a href="#" class="ac-chip-reset" data-chip-reset>恢復預設</a>`;
        if (adding) chipBox.querySelector<HTMLInputElement>('[data-role="chip-new"]')!.focus();
    };
    const updateChips = (next: string[]): void => {
        chips = next;
        saveReasonChips(chips);
    };
    const commitNew = (): void => {
        const input = chipBox.querySelector<HTMLInputElement>('[data-role="chip-new"]');
        if (input) updateChips(addReasonChip(chips, input.value));
        adding = false;
        renderChips();
    };
    chipBox.addEventListener('keydown', (e) => {
        if (!(e.target as HTMLElement).matches('[data-role="chip-new"]')) return;
        if (e.key === 'Enter') {
            e.preventDefault();
            e.stopPropagation();
            commitNew();
        } else if (e.key === 'Escape') {
            e.preventDefault();
            e.stopPropagation();
            adding = false;
            renderChips();
        }
    });
    renderChips();
    return new Promise((done) => {
        dialog.addEventListener('click', (e) => {
            const t = e.target as HTMLElement;
            const del = t.closest<HTMLElement>('[data-chip-del]');
            if (del) {
                updateChips(removeReasonChip(chips, del.dataset.chipDel!));
                renderChips();
                return;
            }
            if (t.closest('[data-chip-new]')) {
                adding = true;
                renderChips();
                return;
            }
            if (t.closest('[data-chip-ok]')) {
                commitNew();
                return;
            }
            if (t.closest('[data-chip-cancel]')) {
                adding = false;
                renderChips();
                return;
            }
            if (t.closest('[data-chip-reset]')) {
                e.preventDefault();
                updateChips(defaultReasonChips());
                renderChips();
                return;
            }
            const chip = t.closest<HTMLElement>('[data-chip]');
            if (chip) {
                reason.value = chip.dataset.chip!;
                return;
            }
            const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
            if (!act) return;
            const ans = {reason: reason.value.trim()};
            close();
            done(act === 'ok' ? ans : null);
        });
    });
}

const periodText = (key: string, half: '上' | '下'): string => `${monthLabel(key)}${half}期`;

/**
 * 「加一筆欠款」／修改視窗：期別（記住的月份＋上下期，或輸入更早的年月）、金額（正整數）、原因（選填）。
 * 月結客人只選月份（記成下期；修改時沿用原本的上下期）。
 */
function debtDialog(c: Debtor, history: History, editing?: ManualDebt): Promise<ManualDebtInput | null> {
    const OTHER = '__other';
    const monthKeys = [...new Set([...Object.keys(history.months), ...(editing ? [editing.month_key] : [])])].sort().reverse();
    const halfOf = (k: string): '上' | '下' => (editing?.month_key === k ? editing.half : '下');
    const opts = c.monthly
        ? monthKeys.map((k) => `<option value="${k}|${halfOf(k)}">${esc(monthLabel(k))}</option>`)
        : monthKeys.flatMap((k) => (['上', '下'] as const).map((h) =>
            `<option value="${k}|${h}">${esc(periodText(k, h))}</option>`));
    const cur = editing ? `${editing.month_key}|${editing.half}` : '';
    const body = `<p><b>${esc(c.code)} ${esc(c.name)}</b>　目前欠款 ${money(c.total)} 元</p>
      <label class="ac-field"><span>期別</span>
        <select class="app-form-input" data-role="period">
          <option value="">請選擇期別…</option>${opts}<option value="${OTHER}">其他月份（自己輸入年月）…</option></select></label>
      <div class="ac-field-row" data-role="other" hidden>
        <label class="ac-field"><span>年月</span><input class="app-form-input" type="month" data-role="month"></label>
        <label class="ac-field"${c.monthly ? ' hidden' : ''}><span>上／下期</span>
          <select class="app-form-input" data-role="half"><option value="上">上期</option><option value="下">下期</option></select></label>
      </div>
      <label class="ac-field"><span>金額（元，正整數）</span>
        <input class="app-form-input" type="number" min="1" step="1" inputmode="numeric" data-role="amount" placeholder="例如：64" autocomplete="off"></label>
      <label class="ac-field"><span>原因（選填）</span>
        <input class="app-form-input" type="text" data-role="reason" placeholder="例如：少付" autocomplete="off"></label>
      <div class="ac-settle-note">加入後會算進這位客人的合計、總表與下載的檔案，訊息也會多一行。</div>
      <div class="ac-error" data-role="error" hidden></div>`;
    const {dialog, close} = openModal(editing ? '修改手動欠款' : '加一筆欠款', body,
        `<button type="button" class="btn btn-secondary" data-act="cancel">取消</button>
         <button type="button" class="btn btn-primary" data-act="ok">${editing ? '儲存' : '加入'}</button>`);
    const q = <T extends HTMLElement>(r: string): T => dialog.querySelector<T>(`[data-role="${r}"]`)!;
    const period = q<HTMLSelectElement>('period');
    const other = q('other');
    const month = q<HTMLInputElement>('month');
    const half = q<HTMLSelectElement>('half');
    const amount = q<HTMLInputElement>('amount');
    const reason = q<HTMLInputElement>('reason');
    const error = q('error');
    if (editing) {
        if (cur && [...period.options].some((o) => o.value === cur)) period.value = cur;
        amount.value = String(editing.amount);
        reason.value = editing.reason;
    }
    period.addEventListener('change', () => {
        other.hidden = period.value !== OTHER;
    });
    return new Promise((done) => {
        dialog.addEventListener('click', (e) => {
            const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
            if (!act) return;
            if (act !== 'ok') {
                close();
                done(null);
                return;
            }
            const [mk, h] = period.value === OTHER ? [month.value, half.value] : period.value.split('|');
            const amt = amount.value.trim() === '' ? NaN : Number(amount.value);
            const err = validateManualDebt({month_key: mk ?? '', half: h ?? '', amount: amt});
            if (err) {
                error.textContent = err;
                error.hidden = false;
                return;
            }
            close();
            done({month_key: mk, half: h as '上' | '下', amount: amt, reason: reason.value.trim()});
        });
    });
}

/** 同月份有多份 → 讓使用者選；回傳 {月份: 選項} 或 null（取消）。 */
function askPick(conflicts: MonthConflict[]): Promise<Choices | null> {
    const optionHtml = (c: MonthConflict, ci: number, o: ConflictOption, oi: number, single: boolean): string => {
        const title = o.kind === 'old'
            ? `${single ? '保留原本的：' : ''}<span class="ac-fn">${esc(o.name)}</span><span class="ac-tag is-old">之前記住的</span>`
            : `${single ? '用新的取代：' : ''}<span class="ac-fn">${esc(o.name)}</span><span class="ac-tag">這次放的</span>`
            + (o.same_as_old ? '<span class="ac-tag is-old">內容和之前記住的一樣</span>' : '');
        const meta = `檔案修改時間 ${when(o.file_mtime)}${o.imported_at ? `　・　匯入時間 ${when(o.imported_at)}` : ''}`;
        return `<label class="ac-pick-opt"><input type="radio" name="ac-pick-${ci}" data-key="${esc(c.month_key)}" data-i="${oi}">
          <div><div>${title}</div><div class="ac-pick-meta">${meta}</div>
          <div class="ac-pick-sum">上期 ${o.upper} 位、下期 ${o.lower} 位客人，其中 ${o.paid} 筆有記到付款<br>
          只看這個月：<b>${o.debtors} 位</b>客人欠款，合計 <b>${money(o.owed_total)} 元</b></div></div></label>`;
    };
    const body = `<p class="ac-pick-intro">同一個月的不同份檔案，可能是用不同批的銀行資料對帳出來的，記到的付款不一樣。用錯份可能會催到已經付款的客人，所以請你自己選。<b>選好之前，這次放的檔案都還不會算進去。</b></p>`
        + conflicts.map((c, ci) => {
            const news = c.options.filter((o) => o.kind === 'new');
            const old = c.options.find((o) => o.kind === 'old');
            const single = news.length === 1;
            const intro = single
                ? `${esc(c.label)} 之前已經記住「${esc(old?.name)}」，這次又放了新的一份。要用新的取代，還是保留原本的？`
                : `這次放了 ${news.length} 份 ${esc(c.label)} 的檔案${old ? '，而且之前已經記住一份' : ''}，只能用其中一份。`;
            return `<section class="ac-pick-month"><h3>${esc(c.label)}</h3><p>${intro}</p>
              ${c.options.map((o, oi) => optionHtml(c, ci, o, oi, single)).join('')}</section>`;
        }).join('');
    const title = conflicts.some((c) => c.options.filter((o) => o.kind === 'new').length > 1)
        ? '同一個月有好幾份檔案，請選一份'
        : '這個月之前已經放過，要用哪一份？';
    const {dialog, close} = openModal(title, body,
        `<button type="button" class="btn btn-secondary" data-act="cancel">取消（這次都不匯入）</button>
         <button type="button" class="btn btn-primary" data-act="ok" disabled>用我選的開始計算</button>`, 'ac-pick-modal');
    const picked: Choices = {};
    const ok = dialog.querySelector<HTMLButtonElement>('[data-act="ok"]')!;
    dialog.addEventListener('change', (e) => {
        const r = e.target as HTMLInputElement;
        if (r.type !== 'radio') return;
        const c = conflicts.find((x) => x.month_key === r.dataset.key)!;
        picked[c.month_key] = c.options[+r.dataset.i!].value;
        dialog.querySelectorAll<HTMLInputElement>(`input[name="${r.name}"]`)
            .forEach((x) => x.closest('.ac-pick-opt')!.classList.toggle('is-on', x.checked));
        ok.disabled = Object.keys(picked).length < conflicts.length;
    });
    return new Promise((done) => {
        dialog.addEventListener('click', (e) => {
            const act = (e.target as HTMLElement).closest<HTMLElement>('[data-act]')?.dataset.act;
            if (!act) return;
            close();
            done(act === 'ok' ? picked : null);
        });
    });
}
