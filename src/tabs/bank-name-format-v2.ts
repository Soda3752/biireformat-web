/**
 * 對帳 2.0 分頁（自動對帳）
 *
 * 上傳：
 *  - 銀行對帳單 .csv / .xlsx（可多檔；沿用 trans-record-reader，Big5 / UTF-8 / 郵局格式）
 *  - 月結明細 .xlsx（含「上期」「下期」工作表，可放多個月；檔名含「N月」自動帶月份，可手動改）
 *  - 末五碼資料 .xlsx（選用；含「歷史 / 已確認 / 現付」標記的對照表）
 * 末五碼另外自動套用設定頁的對照表，以「店家編號 = 明細代號」直接對。
 *
 * 比對演算法見 domain/auto-reconcile-engine.ts。
 * 輸出：回填「實收金額 / 收款日」後的各月明細檔（最新月份加上「銀行對帳結果」工作表），以及含已確認的末五碼資料。
 */

import ExcelJS from 'exceljs';
import {saveAs} from 'file-saver';

import type {AcctRow, AutoReconcileResult, AutoReconcileStatus, BankEntry, MemoryEntry, ReconcileEntry} from '@/domain/auto-reconcile-engine';
import {
    cellVal,
    isStatementWorkbook,
    readAcct,
    readBankRows,
    readDone,
    readStmt,
    reconcile,
    writeBack,
} from '@/domain/auto-reconcile-engine';
import {getBankInfos, loadBankInfos} from '@/domain/bank-info-loader';
import {localSettings} from '@/infra/local-settings-store';
import {parseTransRecord} from '@/readers/trans-record-reader';
import {createDropZone} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';

const MEMORY_KEY = 'bii.reconcileV2.memory';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

type FilterKey = 'all' | AutoReconcileStatus;

const STATUS_ORDER: ReadonlyArray<AutoReconcileStatus> = ['需確認', '金額不符', '找不到', '不在明細', 'OK'];
const STATUS_LABEL: Record<FilterKey, string> = {
    all: '全部',
    OK: '已對上',
    需確認: '需確認',
    金額不符: '金額不符',
    找不到: '找不到',
    不在明細: '不在明細',
};
/** 沿用 reconcile-status-pill 的配色 */
const STATUS_PILL: Record<AutoReconcileStatus, string> = {
    OK: 'is-matched',
    需確認: 'is-partial',
    金額不符: 'is-partial',
    找不到: 'is-unpaid',
    不在明細: 'is-overpaid',
};

interface BankFileEntry {
    name: string;
    rows: BankEntry[];
}

interface StmtFileEntry {
    name: string;
    buf: ArrayBuffer;
    month: number | null;
}

const monthOf = (name: string): number | null => {
    const m = /(\d{1,2})\s*月/.exec(name);
    if (!m) return null;
    const n = +m[1];
    return n >= 1 && n <= 12 ? n : null;
};

const loadWorkbook = async (buf: ArrayBuffer): Promise<ExcelJS.Workbook> => {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf.slice(0));
    return wb;
};

const escapeHtml = (t: unknown): string => String(t ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

/** 只寫「青坊」「麵包」這類共用備註的，確認只套用這一筆，不記到這台電腦。 */
const isGenericKey = (k: string): boolean => !/\*\*\d{4}\*/.test(k)
    && (/^(青坊|漢堡|麵包|麵包錢|貨款|青坊.*|.*月.*貨款|.*月初貨款|.*月底貨款)$/.test(k) || k.length <= 2);

export function renderBankNameFormatV2Panel(tab: TabDefinition): HTMLElement {
    const panel = document.createElement('section');
    panel.className = 'tab-panel';
    panel.dataset.tabId = tab.id;
    panel.setAttribute('role', 'tabpanel');

    panel.innerHTML = `
    <div class="card">
      <header class="card-header">
        <h1 class="card-title">對帳匯總 2.0</h1>
        <p class="card-subtitle">
          上傳銀行對帳單與月結明細（可放多個月），系統自動依帳號末碼、對照紀錄與金額找出客戶和期別，回填明細的「實收金額」與「收款日」。
          末五碼自動套用設定頁的對照表（店家編號 = 明細代號）。
        </p>
      </header>

      <div class="notice-banner" data-role="notice-banner" hidden></div>

      <div class="auto-recon-dropzones">
        <div class="reconcile-trans-slot">
          <div data-role="bank-slot"></div>
          <div class="reconcile-trans-list" data-role="bank-list" hidden></div>
        </div>
        <div class="reconcile-trans-slot">
          <div data-role="stmt-slot"></div>
          <div class="reconcile-trans-list" data-role="stmt-list" hidden></div>
        </div>
        <div class="reconcile-trans-slot">
          <div data-role="acct-slot"></div>
        </div>
      </div>

      <div class="auto-recon-memory">
        <span data-role="memory-text"></span>
        <button type="button" class="btn btn-secondary btn-xs" data-role="memory-clear">清除記憶</button>
      </div>

      <section class="reconcile-preview" data-role="preview" hidden>
        <div class="notice-banner auto-recon-info" data-role="cont-banner" hidden></div>
        <div class="reconcile-tabs auto-recon-tabs" role="tablist" data-role="tabs"></div>
        <div class="reconcile-tab-panel">
          <div class="reconcile-search-bar">
            <input type="search" class="reconcile-search-input" data-role="search"
                   placeholder="搜尋客戶、帳號、金額..." autocomplete="off">
            <button type="button" class="reconcile-search-clear" data-role="search-clear" aria-label="清除搜尋" hidden>✕</button>
          </div>
          <div class="reconcile-scroll">
            <table class="reconcile-table auto-recon-table">
              <thead>
                <tr>
                  <th>日期</th>
                  <th class="col-money">金額</th>
                  <th>帳號 / 備註</th>
                  <th>對應客戶與期別</th>
                  <th>狀態</th>
                  <th>說明</th>
                  <th>確認</th>
                </tr>
              </thead>
              <tbody data-role="tbody"></tbody>
            </table>
          </div>
          <p class="auto-recon-help">
            「需確認」的請核對：對的按「✓ 正確」；不對就在框裡輸入正確客戶（可從清單選），或選「現付店家」。
            用帳號確認過的會記在這台電腦，下次自動對上；只寫「青坊」「麵包」這類共用備註的，只套用這一筆。
            明細表回填後：綠色是確定的、黃色是推測或金額不符；收款日標記 (-15) 扣手續費、(月) 一次付多期或多店合併、(2筆) 多筆、? 待確認。
            下次有新入帳，把這次下載的各月檔和新的銀行對帳單一起放進來，會自動略過已處理的入帳並接著填。
          </p>
          <datalist id="auto-recon-names" data-role="names"></datalist>
        </div>
      </section>

      <footer class="action-bar">
        <div class="action-bar-status" data-role="overall-status">請上傳銀行對帳單與月結明細</div>
        <div class="action-bar-actions" data-role="actions">
          <button type="button" class="btn btn-secondary" data-role="reset">重設</button>
          <button type="button" class="btn btn-secondary" data-role="download-acct" disabled>下載末五碼資料（含已確認）</button>
          <span class="auto-recon-downloads" data-role="downloads"></span>
        </div>
      </footer>
    </div>
  `;

    const $ = <T extends HTMLElement = HTMLElement>(role: string): T =>
        panel.querySelector<T>(`[data-role="${role}"]`)!;

    const banner = $('notice-banner');
    const bankListEl = $('bank-list');
    const stmtListEl = $('stmt-list');
    const memoryText = $('memory-text');
    const previewEl = $('preview');
    const contBanner = $('cont-banner');
    const tabsEl = $('tabs');
    const searchInput = $<HTMLInputElement>('search');
    const searchClear = $<HTMLButtonElement>('search-clear');
    const tbody = $('tbody');
    const namesEl = $('names');
    const overallStatus = $('overall-status');
    const downloadsEl = $('downloads');
    const downloadAcctBtn = $<HTMLButtonElement>('download-acct');

    const bankFiles: BankFileEntry[] = [];
    let stmtFiles: StmtFileEntry[] = [];
    let acctFile: { name: string; buf: ArrayBuffer; rows: AcctRow[] } | null = null;
    let rowMemory: Record<number, MemoryEntry> = {};
    let filter: FilterKey = 'all';
    let query = '';
    let lastResult: AutoReconcileResult | null = null;
    let lastError: string | null = null;
    /** 最近一次對帳用到的明細（由舊到新），下載時依此回填 */
    let lastStmts: StmtFileEntry[] = [];
    let runSeq = 0;

    let memory: Record<string, MemoryEntry> = {};
    try {
        memory = JSON.parse(localStorage.getItem(MEMORY_KEY) || '{}') || {};
    } catch {
        memory = {};
    }
    const saveMemory = () => {
        try {
            localStorage.setItem(MEMORY_KEY, JSON.stringify(memory));
        } catch {
            // 無法寫入時只在本次有效
        }
        refreshMemoryText();
    };
    const refreshMemoryText = () => {
        memoryText.textContent = `這台電腦已記住 ${Object.keys(memory).length} 筆確認`;
    };

    // ---------- 上傳

    const bankZone = createDropZone({
        title: '銀行對帳單',
        hint: '拖曳或點擊上傳（可多檔）.csv（Big5 / UTF-8）或 .xlsx',
        accept: `.csv,text/csv,.xlsx,${XLSX_MIME}`,
        multiple: true,
        onFile: async (file) => {
            if (bankFiles.some((f) => f.name === file.name)) {
                showToast({variant: 'info', title: '已存在相同檔名', message: `${file.name} 已在清單中，跳過。`});
                return;
            }
            try {
                const rows = readBankRows(await parseTransRecord(file));
                bankFiles.push({name: file.name, rows});
                renderBankList();
                void rebuild();
            } catch (err) {
                showToast({variant: 'error', title: `${file.name} 讀取失敗`, message: err instanceof Error ? err.message : '讀取失敗'});
                throw err;
            }
        },
    });

    const stmtZone = createDropZone({
        title: '月結明細 .xlsx',
        hint: '可放多個月（例：9月、8月、7月），放越多個月越能對出付舊帳的客人；接續對帳時放上次下載的檔',
        accept: `.xlsx,${XLSX_MIME}`,
        multiple: true,
        onFile: async (file) => {
            try {
                const buf = await file.arrayBuffer();
                const wb = await loadWorkbook(buf);
                if (!isStatementWorkbook(wb)) throw new Error('找不到「上期」「下期」工作表，不像月結明細');
                stmtFiles = stmtFiles.filter((x) => x.name !== file.name);
                stmtFiles.push({name: file.name, buf, month: monthOf(file.name)});
                renderStmtList();
                void rebuild();
            } catch (err) {
                showToast({variant: 'error', title: `${file.name} 讀取失敗`, message: err instanceof Error ? err.message : '讀取失敗'});
                throw err;
            }
        },
    });

    const acctZone = createDropZone({
        title: '末五碼資料 .xlsx（選用）',
        hint: '含「歷史 / 已確認 / 現付」標記的對照表；欄位：店名、區域、末五碼/帳號/備註、標記',
        accept: `.xlsx,${XLSX_MIME}`,
        onFile: async (file) => {
            const buf = await file.arrayBuffer();
            const wb = await loadWorkbook(buf);
            if (isStatementWorkbook(wb)) throw new Error('這是月結明細，請放到月結明細欄');
            const rows = readAcct(wb);
            acctFile = {name: file.name, buf, rows};
            const hist = rows.filter((r) => r.flag === '歷史').length;
            acctZone.setStatus('loaded', `${file.name}　共 ${rows.length} 筆（歷史對照 ${hist} 筆）`);
            void rebuild();
        },
    });
    // 清除按鈕只重繪 dropzone，這裡同步清掉資料
    acctZone.element.addEventListener('click', (e) => {
        if ((e.target as HTMLElement).closest('.dropzone-clear') && acctFile) {
            acctFile = null;
            void rebuild();
        }
    }, true);

    $('bank-slot').appendChild(bankZone.element);
    $('stmt-slot').appendChild(stmtZone.element);
    $('acct-slot').appendChild(acctZone.element);

    const listItem = (name: string, meta: string | HTMLElement, onRemove: () => void): HTMLElement => {
        const item = document.createElement('div');
        item.className = 'reconcile-trans-list-item';
        const nameEl = document.createElement('span');
        nameEl.className = 'reconcile-trans-list-name';
        nameEl.textContent = name;
        nameEl.title = name;
        const metaEl = document.createElement('span');
        metaEl.className = 'reconcile-trans-list-meta';
        if (typeof meta === 'string') metaEl.textContent = meta;
        else metaEl.appendChild(meta);
        const removeBtn = document.createElement('button');
        removeBtn.type = 'button';
        removeBtn.className = 'reconcile-trans-list-remove';
        removeBtn.setAttribute('aria-label', `移除 ${name}`);
        removeBtn.textContent = '✕';
        removeBtn.addEventListener('click', onRemove);
        item.append(nameEl, metaEl, removeBtn);
        return item;
    };

    const listTitle = (text: string): HTMLElement => {
        const el = document.createElement('div');
        el.className = 'reconcile-trans-list-title';
        el.textContent = text;
        return el;
    };

    const renderBankList = () => {
        bankListEl.hidden = bankFiles.length === 0;
        bankListEl.innerHTML = '';
        if (bankFiles.length === 0) return;
        bankListEl.appendChild(listTitle(`已上傳 ${bankFiles.length} 個`));
        for (const entry of bankFiles) {
            bankListEl.appendChild(listItem(entry.name, `入帳 ${entry.rows.length} 筆`, () => {
                bankFiles.splice(bankFiles.indexOf(entry), 1);
                renderBankList();
                void rebuild();
            }));
        }
    };

    /** 依月份由舊到新排序（處理 12 月 → 1 月跨年） */
    const sortedStmts = (): StmtFileEntry[] => {
        const ms = stmtFiles.map((x) => x.month).filter((m): m is number => m !== null);
        const wrap = ms.some((m) => m >= 10) && ms.some((m) => m <= 3);
        const rank = (m: number | null) => (m ?? 0) + (wrap && m !== null && m <= 6 ? 12 : 0);
        return [...stmtFiles].sort((a, b) => rank(a.month) - rank(b.month));
    };

    const renderStmtList = () => {
        stmtListEl.hidden = stmtFiles.length === 0;
        stmtListEl.innerHTML = '';
        if (stmtFiles.length === 0) return;
        stmtListEl.appendChild(listTitle(`已上傳 ${stmtFiles.length} 個（月份可修改）`));
        const ss = sortedStmts();
        ss.forEach((entry, i) => {
            const meta = document.createElement('span');
            meta.className = 'auto-recon-month';
            const input = document.createElement('input');
            input.type = 'number';
            input.min = '1';
            input.max = '12';
            input.className = 'reconcile-options-input auto-recon-month-input';
            input.value = entry.month ? String(entry.month) : '';
            input.setAttribute('aria-label', `${entry.name} 的月份`);
            input.addEventListener('change', () => {
                const n = Number(input.value);
                entry.month = Number.isInteger(n) && n >= 1 && n <= 12 ? n : null;
                renderStmtList();
                void rebuild();
            });
            meta.append(input, document.createTextNode(i === ss.length - 1 ? '月（本月）' : '月'));
            stmtListEl.appendChild(listItem(entry.name, meta, () => {
                stmtFiles = stmtFiles.filter((x) => x !== entry);
                renderStmtList();
                void rebuild();
            }));
        });
    };

    // ---------- 對帳

    const rebuild = async () => {
        const seq = ++runSeq;
        lastError = null;
        const ss = sortedStmts();
        const bank = mergeBankFiles(bankFiles);
        if (bank.length === 0 || ss.length === 0) {
            lastResult = null;
            render();
            return;
        }
        if (ss.some((x) => !x.month)) {
            lastResult = null;
            lastError = '請在每個月結明細旁填上是幾月';
            render();
            return;
        }
        if (new Set(ss.map((x) => x.month)).size !== ss.length) {
            lastResult = null;
            lastError = '有兩個月結明細是同一個月份，請移除一個';
            render();
            return;
        }
        try {
            const stmts: Array<{ month: number; data: ReturnType<typeof readStmt> }> = [];
            let latest: ExcelJS.Workbook | null = null;
            for (const x of ss) {
                const wb = await loadWorkbook(x.buf);
                stmts.push({month: x.month!, data: readStmt(wb, x.month!)});
                latest = wb;
            }
            if (seq !== runSeq) return;
            lastResult = reconcile({
                bank,
                stmts,
                acct: acctFile?.rows ?? [],
                bankInfos: getBankInfos(),
                memory,
                rowMemory,
                done: latest ? readDone(latest) : new Set(),
            });
            lastStmts = ss;
        } catch (err) {
            if (seq !== runSeq) return;
            console.error(err);
            lastResult = null;
            lastError = '對帳失敗：' + (err instanceof Error ? err.message : String(err));
        }
        render();
    };

    /** 重算（不重讀檔）：確認紀錄變動時使用 */
    const rerun = () => void rebuild();

    // ---------- 畫面

    const refreshBanner = () => {
        const count = localSettings.hasLastFiveDigit() ? getBankInfos().length : 0;
        const withCode = getBankInfos().filter((i) => i.storeCode.trim() !== '').length;
        if (count > 0 && withCode > 0) {
            banner.hidden = true;
            banner.innerHTML = '';
            return;
        }
        banner.hidden = false;
        banner.innerHTML = count === 0
            ? '設定頁尚未建立末五碼對照表，只能靠上傳的末五碼資料、備註文字與金額推測。可到 <a href="#settings" class="notice-banner-link">設定 → 末五碼</a> 建立。'
            : '設定頁的末五碼都沒有填店家編號，只能用店名比對明細。可到 <a href="#settings" class="notice-banner-link">設定 → 末五碼</a> 補上店家編號（= 明細代號）。';
    };

    const render = () => {
        refreshBanner();
        refreshMemoryText();
        const r = lastResult;
        downloadAcctBtn.disabled = Object.keys(memory).length === 0 && !acctFile;
        renderDownloads();

        if (!r) {
            previewEl.hidden = true;
            const missing: string[] = [];
            if (bankFiles.length === 0) missing.push('銀行對帳單');
            if (stmtFiles.length === 0) missing.push('月結明細');
            overallStatus.textContent = lastError
                ?? (missing.length ? `尚缺${missing.join('與')}` : '對帳中...');
            return;
        }

        previewEl.hidden = false;
        if (r.skipped) {
            contBanner.hidden = false;
            contBanner.textContent = `接續模式：本月檔裡已有 ${r.skipped} 筆處理過的入帳，已自動略過，只處理新的 ${r.results.length} 筆。`;
        } else {
            contBanner.hidden = true;
        }

        const cnt: Record<string, number> = {all: r.results.length};
        for (const x of r.results) cnt[x.status] = (cnt[x.status] ?? 0) + 1;
        const ok = cnt.OK ?? 0;
        const pending = r.results.length - ok;
        overallStatus.textContent = `入帳 ${r.results.length} 筆　已對上 ${ok}　待處理 ${pending}`
            + (lastStmts.length > 1 ? `（明細 ${lastStmts.map((x) => x.month).join('、')} 月）` : '');

        tabsEl.innerHTML = (['all', ...STATUS_ORDER] as FilterKey[]).map((k) => {
            const n = cnt[k] ?? 0;
            const tone = k === '找不到' ? 'is-danger' : (k === '需確認' || k === '金額不符') ? 'is-warn' : '';
            return `<button type="button" class="reconcile-tab ${filter === k ? 'is-active' : ''}" role="tab"
                      aria-selected="${filter === k}" data-k="${k}">
                      <span class="reconcile-tab-label">${STATUS_LABEL[k]}</span>
                      <span class="reconcile-tab-badge ${n === 0 ? 'is-zero' : (n > 0 ? tone : '')}">${n}</span>
                    </button>`;
        }).join('');

        namesEl.innerHTML = r.names.map((n) => `<option value="${escapeHtml(n)}">`).join('') + '<option value="現付店家">';
        renderRows();
    };

    const renderRows = () => {
        const r = lastResult;
        if (!r) return;
        const q = query.trim();
        const shown = r.results
            .filter((x) => (filter === 'all' || x.status === filter)
                && (!q || [x.key, x.who, x.amt, x.note, x.date, x.kind, x.bank].join(' ').includes(q)))
            .sort((a, b) => STATUS_ORDER.indexOf(a.status) - STATUS_ORDER.indexOf(b.status) || a.idx - b.idx);
        searchClear.hidden = query === '';
        if (shown.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" class="reconcile-empty">${q ? '沒有符合搜尋的資料' : '沒有資料'}</td></tr>`;
            return;
        }
        tbody.innerHTML = shown.map((x) => {
            const canOk = x.custNames.length === 1 && !x.confirmed;
            return `<tr>
              <td class="auto-recon-date">${escapeHtml(x.date.slice(5))}</td>
              <td class="col-money">${x.amt.toLocaleString()}</td>
              <td class="auto-recon-key">${escapeHtml(x.key) || '—'}<div class="auto-recon-sub">${escapeHtml(x.kind)} ${escapeHtml(x.bank)}</div></td>
              <td>${escapeHtml(x.who) || '<span class="auto-recon-sub">—</span>'}</td>
              <td><span class="reconcile-status-pill ${STATUS_PILL[x.status]}">${STATUS_LABEL[x.status]}</span></td>
              <td class="auto-recon-note">${escapeHtml(x.note)}${x.src ? `<div class="auto-recon-sub">依據：${escapeHtml(x.src)}</div>` : ''}</td>
              <td><div class="auto-recon-act">
                ${x.confirmed ? '<span class="reconcile-status-pill is-matched">✓ 已確認</span>' : ''}
                ${canOk ? `<button type="button" class="btn btn-secondary btn-xs" data-ok="${x.idx}">✓ 正確</button>` : ''}
                <input class="reconcile-search-input auto-recon-fix" list="auto-recon-names" placeholder="改成…" data-fix="${x.idx}">
              </div></td>
            </tr>`;
        }).join('');
    };

    const byIdx = (i: string | undefined): ReconcileEntry | undefined =>
        lastResult?.results.find((x) => x.idx === Number(i));

    const remember = (entry: ReconcileEntry, name: string, flag: string) => {
        if (isGenericKey(entry.key)) {
            rowMemory[entry.idx] = {name, flag};
        } else {
            memory[entry.key] = {name, flag, at: new Date().toISOString().slice(0, 10)};
            saveMemory();
        }
        rerun();
    };

    tbody.addEventListener('click', (e) => {
        const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-ok]');
        if (!btn) return;
        const entry = byIdx(btn.dataset.ok);
        if (entry) remember(entry, entry.custNames[0], '已確認');
    });
    tbody.addEventListener('change', (e) => {
        const input = (e.target as HTMLElement).closest<HTMLInputElement>('[data-fix]');
        if (!input) return;
        const v = input.value.trim();
        const entry = byIdx(input.dataset.fix);
        if (!v || !entry) return;
        if (/現付/.test(v)) {
            const nm = v.replace(/現付(店家)?/, '').trim()
                || window.prompt('這家現付店家的名稱？', entry.key)
                || entry.key;
            remember(entry, nm, '現付（叫貨直接轉帳）');
        } else {
            remember(entry, v, '已確認');
        }
    });

    tabsEl.addEventListener('click', (e) => {
        const btn = (e.target as HTMLElement).closest<HTMLButtonElement>('[data-k]');
        if (!btn) return;
        filter = btn.dataset.k as FilterKey;
        render();
    });
    searchInput.addEventListener('input', () => {
        query = searchInput.value;
        renderRows();
    });
    searchClear.addEventListener('click', () => {
        query = '';
        searchInput.value = '';
        searchInput.focus();
        renderRows();
    });

    $('memory-clear').addEventListener('click', () => {
        if (!window.confirm('確定要清除這台電腦記住的確認紀錄嗎？')) return;
        memory = {};
        saveMemory();
        rerun();
    });

    // ---------- 下載

    const renderDownloads = () => {
        downloadsEl.innerHTML = '';
        if (!lastResult) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-primary btn-lg';
            btn.textContent = '下載對帳檔';
            btn.disabled = true;
            downloadsEl.appendChild(btn);
            return;
        }
        const ss = lastStmts;
        [...ss].reverse().forEach((x) => {
            const isLatest = x === ss[ss.length - 1];
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = isLatest ? 'btn btn-primary btn-lg' : 'btn btn-secondary';
            btn.textContent = `下載 ${x.month}月${isLatest ? '對帳檔' : '補收款檔'}`;
            btn.addEventListener('click', () => void downloadStmt(x, btn));
            downloadsEl.appendChild(btn);
        });
    };

    const saveWorkbook = async (wb: ExcelJS.Workbook, filename: string) => {
        const buf = await wb.xlsx.writeBuffer();
        saveAs(new Blob([buf], {type: XLSX_MIME}), filename);
        showToast({variant: 'success', title: '已輸出', message: filename});
    };

    const withBusy = async (btn: HTMLButtonElement, job: () => Promise<void>) => {
        const original = btn.textContent;
        btn.disabled = true;
        btn.textContent = '處理中...';
        try {
            await job();
        } catch (err) {
            console.error(err);
            showToast({variant: 'error', title: '輸出失敗', message: err instanceof Error ? err.message : String(err)});
        } finally {
            btn.textContent = original;
            btn.disabled = false;
        }
    };

    const downloadStmt = (target: StmtFileEntry, btn: HTMLButtonElement) => withBusy(btn, async () => {
        const rec = lastResult;
        if (!rec) return;
        const ss = lastStmts;
        const wbs: Record<number, ExcelJS.Workbook> = {};
        for (const x of ss) wbs[x.month!] = await loadWorkbook(x.buf);
        const latestMonth = ss[ss.length - 1].month!;
        writeBack(wbs, latestMonth, rec);
        const isLatest = target.month === latestMonth;
        await saveWorkbook(wbs[target.month!], `${target.month}月_${isLatest ? '已自動對帳' : '補收款'}.xlsx`);
    });

    downloadAcctBtn.addEventListener('click', () => void withBusy(downloadAcctBtn, async () => {
        let wb: ExcelJS.Workbook;
        if (acctFile) {
            wb = await loadWorkbook(acctFile.buf);
        } else {
            wb = new ExcelJS.Workbook();
            wb.addWorksheet('末五碼').addRow(['店名', '區域', '末五碼/帳號/備註', '標記']);
        }
        const ws = wb.worksheets[0];
        const rowOf: Record<string, number> = {};
        ws.eachRow((row, i) => {
            const v = cellVal(row.getCell(3));
            if (v !== null && v !== undefined) rowOf[String(v).trim()] = i;
        });
        for (const [k, v] of Object.entries(memory)) {
            if (rowOf[k]) {
                const row = ws.getRow(rowOf[k]);
                row.getCell(1).value = v.name;
                row.getCell(4).value = v.flag;
            } else {
                const row = ws.addRow([v.name, null, k, v.flag]);
                row.eachCell({includeEmpty: true}, (c) => {
                    c.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFE2EFDA'}};
                });
            }
        }
        await saveWorkbook(wb, '末五碼資料.xlsx');
    }));

    const reset = () => {
        bankFiles.length = 0;
        stmtFiles = [];
        acctFile = null;
        rowMemory = {};
        filter = 'all';
        query = '';
        searchInput.value = '';
        lastResult = null;
        lastError = null;
        lastStmts = [];
        runSeq++;
        bankZone.reset();
        stmtZone.reset();
        acctZone.reset();
        renderBankList();
        renderStmtList();
        render();
    };
    $('reset').addEventListener('click', reset);

    window.addEventListener('hashchange', () => {
        if (window.location.hash === tab.hash) void rebuild();
    });

    loadBankInfos();
    render();

    return panel;
}

/**
 * 多份對帳單期間重疊時，同一筆入帳只算一次：
 * 每個 (日期/金額/帳號/摘要/備註) 取「單一檔案內出現次數」的最大值，保留同日同額的真實多筆。
 */
function mergeBankFiles(files: ReadonlyArray<{ rows: ReadonlyArray<BankEntry> }>): BankEntry[] {
    const keyOf = (r: BankEntry) => `${r.date}|${r.amt}|${r.key}|${r.kind}|${r.bank}`;
    const maxCount = new Map<string, number>();
    for (const f of files) {
        const count = new Map<string, number>();
        for (const r of f.rows) count.set(keyOf(r), (count.get(keyOf(r)) ?? 0) + 1);
        for (const [k, n] of count) maxCount.set(k, Math.max(maxCount.get(k) ?? 0, n));
    }
    const emitted = new Map<string, number>();
    const out: BankEntry[] = [];
    for (const f of files) {
        for (const r of f.rows) {
            const k = keyOf(r);
            const n = emitted.get(k) ?? 0;
            if (n >= (maxCount.get(k) ?? 0)) continue;
            emitted.set(k, n + 1);
            out.push(r);
        }
    }
    return out;
}
