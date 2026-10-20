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
 * 輸出：回填「實收金額 / 收款日 / 備註」後的各月明細檔（最新月份加上「銀行對帳結果」工作表），以及含已確認的末五碼資料。
 */

import ExcelJS from 'exceljs';
import {saveAs} from 'file-saver';
import {
    downloadSequentially, ensureWritable, isHandleGone, loadFolderHandle, pickFolder, saveFolderHandle, supportsFolderSave, writeFiles,
} from '../infra/folder-save';

import type {
    AcctRow,
    AutoReconcileResult,
    AutoReconcileStatus,
    BankEntry,
    GuessPick,
    MemoryEntry,
    MultiOption,
    ReconcileEntry,
    SpecialKind,
} from '@/domain/auto-reconcile-engine';
import {
    cellVal,
    isStatementWorkbook,
    CASH_FLAG,
    CHEQUE_FLAG,
    DEPOSIT_FLAG,
    isChequeKind,
    isLowerUnsettled,
    missedLines,
    MULTI_SEP,
    NEW_CUSTOMER_FLAG,
    ONE_OFF_FLAG,
    readAcct,
    readBankRows,
    readDone,
    readStmt,
    reconcile,
    suggestNames,
    writeBack,
} from '@/domain/auto-reconcile-engine';
import {getBankInfos, loadBankInfos} from '@/domain/bank-info-loader';
import {restoreExternalLinks} from '@/domain/xlsx-external-links';
import {localSettings} from '@/infra/local-settings-store';
import {parseTransRecord} from '@/readers/trans-record-reader';
import {createDropZone} from '@/ui/drop-zone';
import {showToast} from '@/ui/toast';
import type {TabDefinition} from '@/ui/tabs';

const MEMORY_KEY = 'bii.reconcileV2.memory';
const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const CHEQUE_TAB = '支票';
const MISSED_TAB = '漏付';
type FilterKey = 'all' | AutoReconcileStatus | SpecialKind | typeof CHEQUE_TAB | typeof MISSED_TAB;

/** 銀行摘要為代收票據（含票據、支票字樣；去空白、全半形）＝存入的支票，另列「支票」分頁 */
const isCheque = isChequeKind;

const STATUS_ORDER: ReadonlyArray<AutoReconcileStatus> = ['需確認', '金額不符', '找不到', '不在明細', 'OK'];
/** 狀態分頁：「已對上」後面另列單次付款、現金存款（這些列仍算在已對上與全部裡） */
const FILTER_ORDER: ReadonlyArray<FilterKey> = ['all', ...STATUS_ORDER, ONE_OFF_FLAG, DEPOSIT_FLAG, CHEQUE_TAB, MISSED_TAB];
const STATUS_LABEL: Record<FilterKey, string> = {
    all: '全部',
    OK: '已對上',
    需確認: '需確認',
    金額不符: '金額不符',
    找不到: '找不到',
    不在明細: '不在明細',
    單次付款: '單次付款',
    現金存款: '現金存款',
    支票: '支票',
    漏付: '漏付',
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
    /** 自動判斷下期還沒結帳時，比對的是哪個月的下期（null＝判斷為已結帳）。 */
    autoSkipRef: number | null;
    /** 使用者手動勾/取消「下期未結帳」（undefined＝照自動判斷）。 */
    skipManual?: boolean;
}

const skipLowerOf = (x: StmtFileEntry): boolean => x.skipManual ?? x.autoSkipRef !== null;

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
          上傳銀行對帳單與月結明細（可放多個月），系統自動依帳號末碼、對照紀錄與金額找出客戶和期別，回填明細的「實收金額」「收款日」與「備註」。
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
            叫貨時直接付款、不屬於任何一期的，按「現付（不歸期別）」：這筆算已對上、不回填明細，之後這個帳號的入帳都自動當現付；改回月結按「取消」。
            金額只對得上明細「已收過款」的那期時，會列為「需確認」並寫出那期的實收金額與收款日：通常是付別期，勾要回填的期別後按「✓ 回填」（可填原因備註）；真的是補付同一期就按「✓ 正確（疊加）」。
            一個帳號對到多家店時，會試各店連續 1～4 期未收的期別合起來（可扣手續費 15/30）：只有一組對得上就自動對上；好幾組都對得上就從「是哪一組？」挑；都對不上（或想改）就按「手動選各店期別」，勾每家店付哪幾期，差額是 0、-15、-30 可直接確認，其他差額要填原因備註。確認後會記住這個帳號對應這幾家店。
            帳號/備註沒有紀錄的，會用金額在還沒收款的期別裡推測客戶（可扣手續費 15/30）：只推到一個就按「✓ 正確」；推到好幾個就從「是哪一個？」挑。
            「支票」分頁（摘要含代收票據）的入帳，備註是支票號碼、不是店家：還不知道是誰就按「支票」（算已對上、不回填明細，只算這一筆，可填備註寫是誰的）；之後知道了，在「改成」填店家或從候選確認，就會照那家回填明細。備註的號碼是郵局代收流水號（同一天不同店家的支票會連號），不會拿來記憶或對應店家。支票只認金額完全相等，不扣手續費。
            「不在明細」「找不到」的入帳：自己存的現金按「現金存款」（只算這一筆）；不是月結客戶、單次付款的按「單次付款」（記住帳號，下次自動確認）；兩者都算已對上、不回填明細，可在框裡填備註。
            用帳號確認過的會記在這台電腦，下次自動對上；只寫「青坊」「麵包」這類共用備註的，只套用這一筆。
            明細表回填後：綠色是確定的、黃色是推測或金額不符；收款日只寫日期，淺藍底表示一次付多期或多店合併。收款日右邊會加一欄「備註」：寫「月」「扣手續費15」「2筆」「待確認」，你填的原因備註以紅字寫在同一格。
            少付／要補：你填的原因備註寫了「少付」「要補」「補N元」「下次補」，或回填後某期實收比明細少超過手續費 15，系統會在同一客戶的下一期（上→下、下→下個月上）備註欄寫紅字「9上少付64，這期要補」，催款時看得到；下一期不在已上傳的明細，就只在這列說明寫出「請記得補收」。下一期若剛好付「明細＋少付金額」，會自動對上並補上期差額。
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
    /** 金額不符原因備註：只記本次對帳，不寫進這台電腦的記憶 */
    let rowNotes: Record<number, string> = {};
    /** 金額不符要回填的期別（使用者勾的，可多期）：只記本次對帳 */
    let rowNotePeriods: Record<number, string[]> = {};
    /** 記為現付、但金額等於某期明細的入帳，使用者選的處理方式：只記本次對帳 */
    let rowCashMode: Record<number, 'cash' | 'period'> = {};
    /** 依金額推測、使用者確認（或從候選挑）的客戶期別：只記本次對帳 */
    let rowPick: Record<number, GuessPick> = {};
    /** 金額只對得上已收過款的期別：使用者選照勾的期別回填（pick）或照原期別疊加（stack）：只記本次對帳 */
    let rowPaidChoice: Record<number, 'pick' | 'stack'> = {};
    /** 多家店合併：使用者確認的各店期別：只記本次對帳 */
    let rowMulti: Record<number, Array<[string, string[]]>> = {};
    /** 手動選各店期別：勾到一半的選擇（店代號 → 期別）、打開選單的列、差額原因備註草稿 */
    let multiDraft: Record<number, Record<string, string[]>> = {};
    let multiOpen = new Set<number>();
    let multiMemo: Record<number, string> = {};
    /** 備註框打到一半還沒送出的字（重繪時保留） */
    let memoDrafts: Record<number, string> = {};
    /** 重繪中：瀏覽器移除有焦點的備註框時會補發 change，這時不能當成送出 */
    let repainting = false;
    /** 本次取消的末五碼資料 .xlsx「已確認 / 現付」紀錄（帳號/備註）；不改原檔，下載時移除 */
    const acctCancelled = new Set<string>();
    /** 「改成」打的名字在明細找不到一模一樣的客戶，等使用者挑選 */
    let pending: { idx: number; name: string } | null = null;
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
        hint: '拖曳或點擊上傳（可多檔）.csv（Big5 / UTF-8）、.xlsx 或郵局 .xls',
        accept: `.csv,text/csv,.xlsx,.xls,${XLSX_MIME}`,
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
                stmtFiles.push({name: file.name, buf, month: monthOf(file.name), autoSkipRef: null});
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
            const skipLabel = document.createElement('label');
            skipLabel.className = 'auto-recon-skip';
            const skip = document.createElement('input');
            skip.type = 'checkbox';
            skip.checked = skipLowerOf(entry);
            skip.addEventListener('change', () => {
                entry.skipManual = skip.checked === (entry.autoSkipRef !== null) ? undefined : skip.checked;
                void rebuild();
            });
            skipLabel.append(skip, document.createTextNode('下期未結帳（略過）'));
            meta.appendChild(skipLabel);
            const item = listItem(entry.name, meta, () => {
                stmtFiles = stmtFiles.filter((x) => x !== entry);
                renderStmtList();
                void rebuild();
            });
            if (entry.autoSkipRef !== null) {
                const hint = document.createElement('div');
                hint.className = 'auto-recon-skip-hint';
                hint.textContent = `下期與 ${entry.autoSkipRef}月下期相同，判定未結帳`;
                item.classList.add('auto-recon-has-hint');
                item.appendChild(hint);
            }
            stmtListEl.appendChild(item);
        });
    };

    // ---------- 對帳

    const rebuild = async () => {
        const seq = ++runSeq;
        lastError = null;
        const ss = sortedStmts();
        // 先判斷各明細的下期是否還沒結帳（與上一個月份檔的下期幾乎一樣），上傳後不用等銀行檔就能看到
        const datas: Array<ReturnType<typeof readStmt> | null> = [];
        try {
            for (const x of ss) datas.push(x.month ? readStmt(await loadWorkbook(x.buf), x.month) : null);
        } catch {
            datas.length = 0;
        }
        if (seq !== runSeq) return;
        ss.forEach((x, i) => {
            const prev = ss[i - 1];
            const ok = i > 0 && !!datas[i] && !!datas[i - 1] && prev.month !== x.month
                && isLowerUnsettled(datas[i - 1]!, prev.month!, datas[i]!, x.month!);
            x.autoSkipRef = ok ? prev.month : null;
        });
        renderStmtList();
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
            const stmts: Array<{ month: number; data: ReturnType<typeof readStmt>; skipLower: boolean }> = [];
            let latest: ExcelJS.Workbook | null = null;
            for (const [i, x] of ss.entries()) {
                const wb = await loadWorkbook(x.buf);
                stmts.push({month: x.month!, data: datas[i] ?? readStmt(wb, x.month!), skipLower: skipLowerOf(x)});
                latest = wb;
            }
            if (seq !== runSeq) return;
            lastResult = reconcile({
                bank,
                stmts,
                acct: (acctFile?.rows ?? []).filter((x) => !isCancelledAcct(x)),
                bankInfos: getBankInfos(),
                memory,
                rowMemory,
                rowNotes,
                rowNotePeriods,
                rowCashMode,
                rowPick,
                rowPaidChoice,
                rowMulti,
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

    const isConfirmFlag = (flag: string): boolean => /已確認|現付|單次付款/.test(flag);
    const isCancelledAcct = (x: AcctRow): boolean => isConfirmFlag(x.flag) && acctCancelled.has(x.val);

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
        for (const x of r.results) {
            cnt[x.status] = (cnt[x.status] ?? 0) + 1;
            if (x.special && x.special !== CHEQUE_FLAG) cnt[x.special] = (cnt[x.special] ?? 0) + 1;
            if (isCheque(x.kind)) cnt[CHEQUE_TAB] = (cnt[CHEQUE_TAB] ?? 0) + 1;
            if (x.missed?.length) cnt[MISSED_TAB] = (cnt[MISSED_TAB] ?? 0) + 1;
        }
        const ok = cnt.OK ?? 0;
        const pending = r.results.length - ok;
        overallStatus.textContent = `入帳 ${r.results.length} 筆　已對上 ${ok}　待處理 ${pending}`
            + (lastStmts.length > 1 ? `（明細 ${lastStmts.map((x) => x.month).join('、')} 月）` : '');

        tabsEl.innerHTML = FILTER_ORDER.map((k) => {
            const n = cnt[k] ?? 0;
            const tone = k === '找不到' || k === MISSED_TAB ? 'is-danger' : (k === '需確認' || k === '金額不符') ? 'is-warn' : '';
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
            .filter((x) => (filter === 'all' || x.status === filter || x.special === filter
                || (filter === CHEQUE_TAB && isCheque(x.kind)) || (filter === MISSED_TAB && !!x.missed?.length))
                && (!q || [x.key, x.who, x.amt, x.note, x.date, x.kind, x.bank].join(' ').includes(q)))
            .sort((a, b) => STATUS_ORDER.indexOf(a.status) - STATUS_ORDER.indexOf(b.status) || a.idx - b.idx);
        searchClear.hidden = query === '';
        if (shown.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" class="reconcile-empty">${q ? '沒有符合搜尋的資料' : '沒有資料'}</td></tr>`;
            return;
        }
        const act = document.activeElement as HTMLInputElement | null;
        const focusSel = !act || !tbody.contains(act) ? null
            : act.dataset.memo !== undefined ? `[data-memo="${act.dataset.memo}"]`
                : act.dataset.multiMemo !== undefined ? `[data-multi-memo="${act.dataset.multiMemo}"]` : null;
        const focusMemo = focusSel ? {sel: focusSel, pos: act!.selectionStart} : null;
        repainting = true;
        tbody.innerHTML = shown.map((x) => {
            // 已收過、還沒選怎麼處理：改用下面的「✓ 回填」「✓ 正確（疊加）」
            const paidOpen = !!x.paidDup && !x.paidChoice;
            const canOk = x.custNames.length === 1 && !x.confirmed && !paidOpen;
            // 多家店合併自動對上的（只有一組解）：按「✓ 正確」記住這個帳號對應這幾家店
            const canOkMulti = x.alloc.length > 1 && x.alloc.every(([, run]) => run.length) && !x.confirmed && !paidOpen && !x.multiAlts;
            const multiShown = isMultiOpen(x);
            const multiToggle = canMulti(x) && !multiShown
                ? `<button type="button" class="btn btn-secondary btn-xs" data-multi-open="${x.idx}"
                          title="自己勾每家店各付哪幾期">${x.manualMulti ? '改選各店期別' : '手動選各店期別'}</button>`
                : '';
            const altPicker = x.multiAlts && !x.manualMulti
                ? `<div class="auto-recon-period auto-recon-multi-alts" role="group" aria-label="從對得上的組合挑一組">
                    <span class="auto-recon-sub">是哪一組？</span>
                    ${x.multiAlts.map((a, ai) => `<button type="button" class="auto-recon-period-chip"
                                    data-multi-alt="${x.idx}" data-alt-i="${ai}">${escapeHtml(a.label)}</button>`).join('')}
                   </div>`
                : '';
            const canCash = (x.status === '金額不符' || x.status === '需確認') && x.custNames.length === 1 && !x.cash && !x.guesses;
            // 依金額推測的原本是「找不到」，推錯時一樣可以當現金存款 / 單次付款
            const canSpecial = (x.status === '不在明細' || x.status === '找不到' || !!x.guesses) && !x.special;
            // 支票列：任何還沒對上的狀態都可按「支票」（不指定店家）
            const canCheque = isCheque(x.kind) && x.status !== 'OK' && !x.special;
            const guessPicker = x.guesses && x.guesses.length > 1
                ? `<div class="auto-recon-period" role="group" aria-label="從推測的客戶挑一個">
                    <span class="auto-recon-sub">是哪一個？</span>
                    ${x.guesses.map((g, gi) => `<button type="button" class="auto-recon-period-chip"
                                    data-guess="${x.idx}" data-guess-i="${gi}">${escapeHtml(g.label ?? '')}</button>`).join('')}
                    ${x.guessMore ? `<span class="auto-recon-sub">還有 ${x.guessMore} 個沒列出，請用「改成」自己輸入</span>` : ''}
                   </div>`
                : '';
            const memoInput = (x.status === '金額不符' || x.cash === 'ok' || x.special || canSpecial || canCheque || (paidOpen && x.alloc.length === 1)) && !x.userNote
                ? `<input class="reconcile-search-input auto-recon-fix auto-recon-memo"
                          placeholder="${x.status === '金額不符' ? '原因備註（例：扣30手續費）' : paidOpen ? '原因備註（選填，例：付成上期金額）' : '備註（選填）'}"
                          aria-label="金額不符原因備註" data-memo="${x.idx}" value="${escapeHtml(memoDrafts[x.idx] ?? '')}">`
                : '';
            const periodSelect = x.periodOptions?.length && x.notePeriods
                ? `<div class="auto-recon-period" role="group" aria-label="回填到哪幾期"
                        title="${x.paidDup ? '按「✓ 回填」或填了備註後' : '填了備註後'}，實收金額與收款日回填到勾選的期別；勾多期時差額落在最後一期">
                    <span class="auto-recon-sub">回填</span>
                    ${x.periodOptions.map((k) => {
                        const on = x.notePeriods!.includes(k);
                        return `<button type="button" class="auto-recon-period-chip${on ? ' is-on' : ''}" aria-pressed="${on}"
                                        data-memo-period="${x.idx}" data-period="${escapeHtml(k)}">${escapeHtml(k)}</button>`;
                    }).join('')}
                   </div>${x.noteError ? `<div class="auto-recon-period-error" role="alert">⚠ ${escapeHtml(x.noteError)}</div>` : ''}`
                : '';
            const paidAct = paidOpen
                ? `${x.alloc.length === 1 && x.notePeriods ? `<button type="button" class="btn btn-secondary btn-xs" data-paid-choice="pick" data-paid-idx="${x.idx}"
                          title="照下面勾的期別回填實收金額與收款日">✓ 回填 ${escapeHtml(x.notePeriods.join('+'))}</button>` : ''}
                   <button type="button" class="btn btn-secondary btn-xs" data-paid-choice="stack" data-paid-idx="${x.idx}"
                          title="真的是補付同一期：實收金額疊加到已收過的那期">✓ 正確（照 ${escapeHtml(x.paidDup!.map((d) => d.per).join('+'))}，疊加）</button>`
                : x.paidChoice
                    ? `<span class="reconcile-status-pill is-matched">${x.paidChoice === 'stack' ? '疊加' : '已選期別'}</span>
                       <button type="button" class="btn btn-secondary btn-xs" data-paid-choice="" data-paid-idx="${x.idx}"
                               title="取消這次的選擇，恢復成需確認">取消</button>`
                    : '';
            const memoTag = x.userNote
                ? `<span class="reconcile-status-pill is-partial" title="${escapeHtml(x.userNote)}">已填備註</span>
                   <button type="button" class="btn btn-secondary btn-xs" data-memo-clear="${x.idx}">清除備註</button>`
                : '';
            const note = x.userNote
                ? `<div class="auto-recon-user-note">備註：${escapeHtml(x.userNote)}</div>${escapeHtml(x.note.replace(`備註：${x.userNote}；`, ''))}`
                : escapeHtml(x.note);
            const missedNote = missedLines(x, r.cust).map((m) => `<div class="auto-recon-missed">⚠ ${escapeHtml(m)}</div>`).join('');
            return `<tr>
              <td class="auto-recon-date">${escapeHtml(x.date.slice(5))}</td>
              <td class="col-money">${x.amt.toLocaleString()}</td>
              <td class="auto-recon-key">${escapeHtml(x.key) || '—'}<div class="auto-recon-sub">${escapeHtml(x.kind)} ${escapeHtml(x.bank)}</div></td>
              <td>${escapeHtml(x.who) || '<span class="auto-recon-sub">—</span>'}</td>
              <td><span class="reconcile-status-pill ${STATUS_PILL[x.status]}">${STATUS_LABEL[x.status]}</span></td>
              <td class="auto-recon-note">${note}${missedNote}${x.src ? `<div class="auto-recon-sub">依據：${escapeHtml(x.src)}</div>` : ''}</td>
              <td><div class="auto-recon-act">
                ${x.confirmed ? `<span class="reconcile-status-pill is-matched">${x.special ?? (x.cash === 'ok' || x.cash === 'warn' ? '現付' : '✓ 已確認')}</span>
                  <button type="button" class="btn btn-secondary btn-xs" data-unconfirm="${x.idx}"
                          title="${x.cash ? '忘掉這個帳號的現付紀錄，恢復照期別比對'
                              : x.special === ONE_OFF_FLAG ? '忘掉這個帳號的單次付款紀錄，恢復原狀'
                              : x.special ? '恢復原狀' : '取消確認，重新比對'}">取消</button>` : ''}
                ${canOk ? `<button type="button" class="btn btn-secondary btn-xs" data-ok="${x.idx}">✓ 正確</button>` : ''}
                ${canOkMulti ? `<button type="button" class="btn btn-secondary btn-xs" data-multi-ok="${x.idx}">✓ 正確</button>` : ''}
                ${canCash ? `<button type="button" class="btn btn-secondary btn-xs auto-recon-cash" data-cash="${x.idx}"
                          title="叫貨時直接付款，不屬於任何一期：算已對上、不回填明細，並記住這個帳號">現付（不歸期別）</button>` : ''}
                ${x.cash === 'warn' ? `<button type="button" class="btn btn-secondary btn-xs" data-cash-mode="cash" data-cash-idx="${x.idx}">✓ 正確（當現付）</button>
                  <button type="button" class="btn btn-secondary btn-xs" data-cash-mode="period" data-cash-idx="${x.idx}">回填 ${escapeHtml(x.alloc[0]?.[1].join('+') ?? '')}（這筆照月結）</button>` : ''}
                ${canSpecial ? `<button type="button" class="btn btn-secondary btn-xs" data-special="${DEPOSIT_FLAG}" data-special-idx="${x.idx}"
                          title="自己存的現金：算已對上、不回填明細，只算這一筆">現金存款</button>
                  <button type="button" class="btn btn-secondary btn-xs" data-special="${ONE_OFF_FLAG}" data-special-idx="${x.idx}"
                          title="不是月結客戶的單次付款：算已對上、不回填明細，並記住這個帳號">單次付款</button>` : ''}
                ${canCheque ? `<button type="button" class="btn btn-secondary btn-xs" data-special="${CHEQUE_FLAG}" data-special-idx="${x.idx}"
                          title="這是支票、先不指定店家：算已對上、不回填明細，只算這一筆（知道是誰再用「改成」）">支票</button>` : ''}
                ${x.cash === 'period' ? `<button type="button" class="btn btn-secondary btn-xs" data-cash-mode="" data-cash-idx="${x.idx}">改回現付</button>` : ''}
                ${paidAct}${guessPicker}${altPicker}${multiToggle}${memoTag}${memoInput}${periodSelect}
                <input class="reconcile-search-input auto-recon-fix" list="auto-recon-names" placeholder="改成…" data-fix="${x.idx}">
              </div></td>
            </tr>${multiShown ? multiRow(x) : ''}${pending?.idx === x.idx ? pickRow(pending.name) : ''}`;
        }).join('');
        repainting = false;
        if (focusMemo) {
            const el = tbody.querySelector<HTMLInputElement>(focusMemo.sel);
            el?.focus();
            if (el && focusMemo.pos !== null) el.setSelectionRange(focusMemo.pos, focusMemo.pos);
        }
    };

    /** 手動選各店期別：差額（入帳 − 已選合計）是這些值時當手續費，可直接確認 */
    const MULTI_OK_DIFFS = [0, -15, -30];
    const canMulti = (x: ReconcileEntry): boolean => (x.multiOptions?.length ?? 0) > 1 && !x.special && !x.cash;
    /** 多店候選都對不上時直接展開；其他（自動對上、已手動確認）要按「手動選各店期別」才展開 */
    const isMultiOpen = (x: ReconcileEntry): boolean => canMulti(x)
        && (multiOpen.has(x.idx) || (x.status === '需確認' && !x.alloc.some(([, run]) => run.length) && !x.paidDup));
    /** 目前勾的期別（還沒勾過就帶入對到的期別） */
    const multiSel = (x: ReconcileEntry): Record<string, string[]> => multiDraft[x.idx]
        ?? Object.fromEntries((x.multiOptions ?? []).map((o) => [o.code, x.alloc.find(([c]) => c === o.code)?.[1] ?? []]));
    const multiCalc = (x: ReconcileEntry) => {
        const opts: ReadonlyArray<MultiOption> = x.multiOptions ?? [];
        const sel = multiSel(x);
        const alloc = opts
            .map((o): [string, string[]] => [o.code, o.periods.filter((p) => sel[o.code]?.includes(p.per)).map((p) => p.per)])
            .filter(([, run]) => run.length);
        const amtOf = (code: string, per: string): number =>
            opts.find((o) => o.code === code)?.periods.find((p) => p.per === per)?.amt ?? 0;
        const total = alloc.reduce((t, [c, run]) => t + run.reduce((u, per) => u + amtOf(c, per), 0), 0);
        const diff = x.amt - total;
        const needMemo = !MULTI_OK_DIFFS.includes(diff);
        const last = alloc[alloc.length - 1];
        const lastPer = last?.[1][last[1].length - 1];
        const error = !alloc.length ? '請至少勾一期'
            : last && amtOf(last[0], lastPer!) + diff <= 0
                ? `已選合計比入帳多 ${(-diff).toLocaleString()}，超過最後一期（${lastPer}）的金額，請少勾幾期` : '';
        return {alloc, total, diff, needMemo, error};
    };
    /** 差額原因備註（改選時帶入上次填的） */
    const memoOf = (x: ReconcileEntry): string => multiMemo[x.idx] ?? (x.manualMulti ? x.userNote ?? '' : '');
    const multiReady = (x: ReconcileEntry): boolean => {
        const m = multiCalc(x);
        return !m.error && (!m.needMemo || !!memoOf(x).trim());
    };

    /** 手動選各店期別：列在該列下方，勾每家店付哪幾期，即時顯示合計與差額 */
    const multiRow = (x: ReconcileEntry): string => {
        const sel = multiSel(x);
        const m = multiCalc(x);
        const sign = (n: number): string => (n > 0 ? '+' : n < 0 ? '−' : '') + Math.abs(n).toLocaleString();
        return `<tr class="auto-recon-pick-row"><td colspan="7">
          <div class="auto-recon-pick auto-recon-multi" role="group" aria-label="手動選各店付哪幾期">
            <div class="auto-recon-pick-title">多家店合併匯款：勾每家店各付哪幾期（只列還沒收款的期別）</div>
            ${(x.multiOptions ?? []).map((o) => `<div class="auto-recon-pick-body auto-recon-multi-store">
                <span class="auto-recon-multi-name">${escapeHtml(o.name)}</span>
                ${o.periods.length ? o.periods.map((p) => {
                    const on = !!sel[o.code]?.includes(p.per);
                    return `<button type="button" class="auto-recon-period-chip${on ? ' is-on' : ''}" aria-pressed="${on}"
                                    data-multi-chip="${x.idx}" data-code="${escapeHtml(o.code)}" data-period="${escapeHtml(p.per)}">${escapeHtml(p.per)} ${p.amt.toLocaleString()}</button>`;
                }).join('') : '<span class="auto-recon-sub">沒有未收的期別</span>'}
              </div>`).join('')}
            <div class="auto-recon-multi-sum ${m.needMemo ? 'is-warn' : 'is-ok'}" data-multi-sum="${x.idx}">
              已選合計 <b>${m.total.toLocaleString()}</b>　/　入帳 <b>${x.amt.toLocaleString()}</b>　/　差額 <b>${sign(m.diff)}</b>
              ${!m.needMemo && m.diff ? `（扣手續費${-m.diff}，落在最後一期）` : m.needMemo && m.total ? '（差額落在最後一期，需填原因備註）' : ''}
            </div>
            ${m.error ? `<div class="auto-recon-period-error" role="alert">⚠ ${escapeHtml(m.error)}</div>` : ''}
            <div class="auto-recon-pick-body">
              ${m.needMemo && !m.error ? `<input class="reconcile-search-input auto-recon-fix auto-recon-memo" data-multi-memo="${x.idx}"
                       placeholder="差額原因備註（必填）" aria-label="差額原因備註" value="${escapeHtml(memoOf(x))}">` : ''}
              <button type="button" class="btn btn-primary btn-xs" data-multi-confirm="${x.idx}" ${multiReady(x) ? '' : 'disabled'}>✓ 確認各店期別</button>
              ${multiOpen.has(x.idx) ? `<button type="button" class="btn btn-secondary btn-xs" data-multi-close="${x.idx}">收起</button>` : ''}
            </div>
          </div>
        </td></tr>`;
    };

    /** 「改成」找不到一模一樣的客戶時，在該列下方列出候選讓使用者挑 */
    const pickRow = (name: string): string => {
        const names = lastResult?.names ?? [];
        const sug = suggestNames(name, names);
        return `<tr class="auto-recon-pick-row"><td colspan="7">
          <div class="auto-recon-pick" role="group" aria-label="選擇正確客戶">
            <div class="auto-recon-pick-title">明細裡沒有叫「${escapeHtml(name)}」的客戶，請選正確的客戶：</div>
            <div class="auto-recon-pick-body">
              ${sug.map((n) => `<button type="button" class="btn btn-secondary btn-xs" data-pick-name="${escapeHtml(n)}">${escapeHtml(n)}</button>`).join('')}
              <select class="reconcile-options-input auto-recon-pick-select" data-pick-select aria-label="從全部客戶選">
                <option value="">從全部客戶選…</option>
                ${names.map((n) => `<option value="${escapeHtml(n)}">${escapeHtml(n)}</option>`).join('')}
              </select>
            </div>
            <div class="auto-recon-pick-body">
              <button type="button" class="btn btn-primary btn-xs" data-pick-new>這是新客戶（清單沒有），照「${escapeHtml(name)}」記下</button>
              <button type="button" class="btn btn-secondary btn-xs" data-pick-cancel>取消</button>
            </div>
          </div>
        </td></tr>`;
    };

    const byIdx = (i: string | undefined): ReconcileEntry | undefined =>
        lastResult?.results.find((x) => x.idx === Number(i));

    const remember = (entry: ReconcileEntry, name: string, flag: string, pick?: GuessPick) => {
        pending = null;
        delete rowPick[entry.idx];
        delete rowMulti[entry.idx];
        if (pick) rowPick[entry.idx] = pick;
        saveConfirm(entry, {name, flag});
    };

    const saveConfirm = (entry: ReconcileEntry, m: MemoryEntry) => {
        if (isCheque(entry.kind) || isGenericKey(entry.key)) {
            // 支票的備註是郵局代收流水號（每張不同），只記本次這一列，不寫記憶
            rowMemory[entry.idx] = m;
        } else {
            memory[entry.key] = {...m, at: new Date().toISOString().slice(0, 10)};
            saveMemory();
        }
        rerun();
    };

    /** 多家店合併確認：照選的各店期別回填，並記住此帳號對應這幾家店 */
    const confirmMulti = (entry: ReconcileEntry, alloc: Array<[string, string[]]>) => {
        if (!lastResult || !alloc.length) return;
        const names = [...new Set(alloc.map(([c]) => lastResult!.cust[c].name))];
        pending = null;
        delete rowPick[entry.idx];
        delete multiDraft[entry.idx];
        delete multiMemo[entry.idx];
        multiOpen.delete(entry.idx);
        rowMulti[entry.idx] = alloc;
        saveConfirm(entry, {name: names.join(MULTI_SEP), flag: '已確認', ...(names.length > 1 ? {names} : {})});
    };

    /** 取消確認：本次逐筆紀錄、這台電腦的記憶、上傳的末五碼資料（本次）一併取消，重新比對 */
    const unconfirm = (entry: ReconcileEntry) => {
        delete rowMemory[entry.idx];
        delete rowCashMode[entry.idx];
        delete rowPick[entry.idx];
        delete rowMulti[entry.idx];
        if (entry.special || entry.manualMulti) delete rowNotes[entry.idx];
        if (!isCheque(entry.kind) && entry.key in memory) {
            delete memory[entry.key];
            saveMemory();
        }
        if (acctFile?.rows.some((x) => x.val === entry.key && isConfirmFlag(x.flag))) acctCancelled.add(entry.key);
        rerun();
    };

    tbody.addEventListener('click', (e) => {
        const el = e.target as HTMLElement;
        const okBtn = el.closest<HTMLButtonElement>('[data-ok]');
        if (okBtn) {
            const entry = byIdx(okBtn.dataset.ok);
            if (entry) remember(entry, entry.custNames[0], '已確認', entry.guesses?.length === 1 ? entry.guesses[0] : undefined);
            return;
        }
        const multiChip = el.closest<HTMLButtonElement>('[data-multi-chip]');
        if (multiChip) {
            const entry = byIdx(multiChip.dataset.multiChip);
            if (!entry) return;
            const code = multiChip.dataset.code!;
            const per = multiChip.dataset.period!;
            const sel = {...multiSel(entry)};
            const cur = sel[code] ?? [];
            sel[code] = cur.includes(per) ? cur.filter((p) => p !== per) : [...cur, per];
            multiDraft[entry.idx] = sel;
            renderRows();
            return;
        }
        const multiOpenBtn = el.closest<HTMLButtonElement>('[data-multi-open]');
        if (multiOpenBtn) {
            multiOpen.add(Number(multiOpenBtn.dataset.multiOpen));
            renderRows();
            return;
        }
        const multiCloseBtn = el.closest<HTMLButtonElement>('[data-multi-close]');
        if (multiCloseBtn) {
            const i = Number(multiCloseBtn.dataset.multiClose);
            multiOpen.delete(i);
            delete multiDraft[i];
            renderRows();
            return;
        }
        const multiConfirmBtn = el.closest<HTMLButtonElement>('[data-multi-confirm]');
        if (multiConfirmBtn) {
            const entry = byIdx(multiConfirmBtn.dataset.multiConfirm);
            if (!entry || !multiReady(entry)) return;
            const m = multiCalc(entry);
            const memo = memoOf(entry).trim();
            if (m.needMemo && memo) rowNotes[entry.idx] = memo;
            else delete rowNotes[entry.idx];
            confirmMulti(entry, m.alloc);
            return;
        }
        const multiAltBtn = el.closest<HTMLButtonElement>('[data-multi-alt]');
        if (multiAltBtn) {
            const entry = byIdx(multiAltBtn.dataset.multiAlt);
            const alt = entry?.multiAlts?.[Number(multiAltBtn.dataset.altI)];
            if (entry && alt) confirmMulti(entry, alt.alloc);
            return;
        }
        const multiOkBtn = el.closest<HTMLButtonElement>('[data-multi-ok]');
        if (multiOkBtn) {
            const entry = byIdx(multiOkBtn.dataset.multiOk);
            if (entry) confirmMulti(entry, entry.alloc);
            return;
        }
        const guessBtn = el.closest<HTMLButtonElement>('[data-guess]');
        if (guessBtn) {
            const entry = byIdx(guessBtn.dataset.guess);
            const g = entry?.guesses?.[Number(guessBtn.dataset.guessI)];
            if (entry && g && lastResult) remember(entry, lastResult.cust[g.code].name, '已確認', g);
            return;
        }
        const cashBtn = el.closest<HTMLButtonElement>('[data-cash]');
        if (cashBtn) {
            const entry = byIdx(cashBtn.dataset.cash);
            if (!entry) return;
            // 備註框打了還沒送出的字一併當成現付備註
            const draft = memoDrafts[entry.idx]?.trim();
            if (draft) rowNotes[entry.idx] = draft;
            delete memoDrafts[entry.idx];
            delete rowNotePeriods[entry.idx];
            delete rowCashMode[entry.idx];
            remember(entry, entry.custNames[0], CASH_FLAG);
            return;
        }
        const spBtn = el.closest<HTMLButtonElement>('[data-special]');
        if (spBtn) {
            const entry = byIdx(spBtn.dataset.specialIdx);
            if (!entry) return;
            const draft = memoDrafts[entry.idx]?.trim();
            if (draft) rowNotes[entry.idx] = draft;
            delete memoDrafts[entry.idx];
            if (spBtn.dataset.special === DEPOSIT_FLAG || spBtn.dataset.special === CHEQUE_FLAG) {
                // 現金存款、支票：只算這一筆，不記到這台電腦
                const kind = spBtn.dataset.special;
                pending = null;
                delete rowPick[entry.idx];
                delete rowMulti[entry.idx];
                rowMemory[entry.idx] = {name: kind, flag: kind};
                rerun();
            } else {
                // 單次付款：客戶名稱優先用對到的名稱，沒有就用帳號/備註文字
                const nm = (entry.guesses ? '' : entry.custNames[0]) || entry.nameHint || entry.key || `${entry.kind}${entry.bank}`;
                remember(entry, nm, ONE_OFF_FLAG);
            }
            return;
        }
        const modeBtn = el.closest<HTMLButtonElement>('[data-cash-mode]');
        if (modeBtn) {
            const i = Number(modeBtn.dataset.cashIdx);
            const mode = modeBtn.dataset.cashMode;
            if (mode === 'cash' || mode === 'period') rowCashMode[i] = mode;
            else delete rowCashMode[i];
            rerun();
            return;
        }
        const paidBtn = el.closest<HTMLButtonElement>('[data-paid-choice]');
        if (paidBtn) {
            const i = Number(paidBtn.dataset.paidIdx);
            const mode = paidBtn.dataset.paidChoice;
            if (mode === 'pick' || mode === 'stack') {
                // 備註框打了還沒送出的字一併帶入
                const draft = memoDrafts[i]?.trim();
                if (draft) rowNotes[i] = draft;
                delete memoDrafts[i];
                rowPaidChoice[i] = mode;
            } else {
                delete rowPaidChoice[i];
                delete rowNotes[i];
            }
            rerun();
            return;
        }
        const unBtn = el.closest<HTMLButtonElement>('[data-unconfirm]');
        if (unBtn) {
            const entry = byIdx(unBtn.dataset.unconfirm);
            if (entry) unconfirm(entry);
            return;
        }
        const clearBtn = el.closest<HTMLButtonElement>('[data-memo-clear]');
        if (clearBtn) {
            const cleared = byIdx(clearBtn.dataset.memoClear);
            if (cleared?.manualMulti) {
                multiDraft[cleared.idx] = Object.fromEntries(cleared.alloc);
                multiOpen.add(cleared.idx);
                delete rowMulti[cleared.idx];
            }
            delete rowNotes[Number(clearBtn.dataset.memoClear)];
            delete rowNotePeriods[Number(clearBtn.dataset.memoClear)];
            delete rowPaidChoice[Number(clearBtn.dataset.memoClear)];
            rerun();
            return;
        }
        const chip = el.closest<HTMLButtonElement>('[data-memo-period]');
        const chipEntry = chip ? byIdx(chip.dataset.memoPeriod) : undefined;
        if (chip && chipEntry?.notePeriods && chipEntry.periodOptions) {
            const k = chip.dataset.period!;
            const cur = chipEntry.notePeriods;
            const next = cur.includes(k) ? cur.filter((p) => p !== k) : [...cur, k];
            if (!next.length) return; // 至少留一期
            rowNotePeriods[chipEntry.idx] = chipEntry.periodOptions.filter((p) => next.includes(p));
            rerun();
            return;
        }
        const pickBtn = el.closest<HTMLButtonElement>('[data-pick-name]');
        const pendEntry = pending ? byIdx(String(pending.idx)) : undefined;
        if (pickBtn && pendEntry) {
            remember(pendEntry, pickBtn.dataset.pickName!, '已確認');
            return;
        }
        if (el.closest('[data-pick-new]') && pending && pendEntry) {
            remember(pendEntry, pending.name, NEW_CUSTOMER_FLAG);
            return;
        }
        if (el.closest('[data-pick-cancel]')) {
            pending = null;
            renderRows();
        }
    });
    // 備註框打到一半就記下，重繪（例如切換期別）時不會被清掉
    tbody.addEventListener('input', (e) => {
        const memoInput = (e.target as HTMLElement).closest<HTMLInputElement>('[data-memo]');
        if (memoInput) memoDrafts[Number(memoInput.dataset.memo)] = memoInput.value;
        const multiMemoInput = (e.target as HTMLElement).closest<HTMLInputElement>('[data-multi-memo]');
        if (multiMemoInput) {
            const entry = byIdx(multiMemoInput.dataset.multiMemo);
            if (!entry) return;
            multiMemo[entry.idx] = multiMemoInput.value;
            const btn = tbody.querySelector<HTMLButtonElement>(`[data-multi-confirm="${entry.idx}"]`);
            if (btn) btn.disabled = !multiReady(entry);
        }
    });
    // 點期別按鈕不讓備註框失焦：避免打到一半就被當成送出
    tbody.addEventListener('mousedown', (e) => {
        if ((e.target as HTMLElement).closest('[data-memo-period], [data-multi-chip]')) e.preventDefault();
    });
    tbody.addEventListener('change', (e) => {
        const sel = (e.target as HTMLElement).closest<HTMLSelectElement>('[data-pick-select]');
        if (sel) {
            const entry = pending ? byIdx(String(pending.idx)) : undefined;
            if (sel.value && entry) remember(entry, sel.value, '已確認');
            return;
        }
        const memoInput = (e.target as HTMLElement).closest<HTMLInputElement>('[data-memo]');
        if (memoInput) {
            if (repainting) return;
            const v = memoInput.value.trim();
            if (!v) return;
            const i = Number(memoInput.dataset.memo);
            const entry = byIdx(memoInput.dataset.memo);
            // 不在明細 / 找不到：先留著，按「現金存款」「單次付款」時一併帶入
            if (entry && (entry.status === '不在明細' || entry.status === '找不到' || entry.guesses
                || (entry.status === '需確認' && isCheque(entry.kind))) && !entry.special) return;
            rowNotes[i] = v;
            delete memoDrafts[i];
            rerun();
            return;
        }
        const input = (e.target as HTMLElement).closest<HTMLInputElement>('[data-fix]');
        if (!input) return;
        const v = input.value.trim();
        const entry = byIdx(input.dataset.fix);
        if (!v || !entry) return;
        if (/現付/.test(v)) {
            const nm = v.replace(/現付(店家)?/, '').trim()
                || window.prompt('這家現付店家的名稱？', entry.key)
                || entry.key;
            remember(entry, nm, CASH_FLAG);
        } else if (lastResult && (lastResult.names.includes(v) || lastResult.cust[v])) {
            remember(entry, v, '已確認');
        } else {
            // 名字不完全一致：不自動猜，讓使用者挑或標成新客戶
            pending = {idx: entry.idx, name: v};
            renderRows();
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

    const FOLDER_KEY = 'recon-v2-save-dir';
    let saveDir: FileSystemDirectoryHandle | null = null;
    void loadFolderHandle(FOLDER_KEY).then((h) => {
        if (!h) return;
        saveDir = h;
        renderDownloads();
    });

    const stmtFilename = (month: number) => `${month}月_已自動對帳.xlsx`;

    const renderDownloads = () => {
        downloadsEl.innerHTML = '';
        if (!lastResult) {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-primary btn-lg';
            btn.textContent = '全部存檔';
            btn.disabled = true;
            downloadsEl.appendChild(btn);
            return;
        }
        const ss = lastStmts;
        const allBtn = document.createElement('button');
        allBtn.type = 'button';
        allBtn.className = 'btn btn-primary btn-lg';
        allBtn.dataset.role = 'save-all';
        allBtn.textContent = `全部存檔（${ss.length} 份）`;
        allBtn.addEventListener('click', () => void saveAllStmts(allBtn));
        downloadsEl.appendChild(allBtn);

        if (supportsFolderSave()) {
            const folder = document.createElement('span');
            folder.className = 'auto-recon-folder';
            folder.dataset.role = 'save-folder';
            folder.append(saveDir ? `存到「${saveDir.name}」` : '尚未選資料夾（按下後會請你選）');
            if (saveDir) {
                const change = document.createElement('a');
                change.href = '#';
                change.textContent = '換資料夾';
                change.dataset.role = 'change-folder';
                change.addEventListener('click', (e) => {
                    e.preventDefault();
                    void changeSaveFolder();
                });
                folder.append(' ', change);
            }
            downloadsEl.appendChild(folder);
        }

        [...ss].reverse().forEach((x) => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-secondary';
            btn.textContent = `下載 ${x.month}月檔`;
            btn.addEventListener('click', () => void downloadStmt(x, btn));
            downloadsEl.appendChild(btn);
        });
    };

    const changeSaveFolder = async () => {
        try {
            const h = await pickFolder();
            if (!h) return;
            saveDir = h;
            await saveFolderHandle(FOLDER_KEY, h);
            renderDownloads();
        } catch (err) {
            showToast({variant: 'error', title: '選資料夾失敗', message: err instanceof Error ? err.message : String(err)});
        }
    };

    /** 產生每個月的檔（回填一次，各月共用同一批活頁簿） */
    const buildStmtFiles = async (): Promise<{month: number; filename: string; data: ArrayBuffer}[]> => {
        const rec = lastResult;
        if (!rec) return [];
        const ss = lastStmts;
        const wbs: Record<number, ExcelJS.Workbook> = {};
        for (const x of ss) wbs[x.month!] = await loadWorkbook(x.buf);
        const latestMonth = ss[ss.length - 1].month!;
        writeBack(wbs, latestMonth, rec);
        const out: {month: number; filename: string; data: ArrayBuffer}[] = [];
        for (const x of [...ss].reverse()) {
            // ExcelJS 會丟掉外部活頁簿連結，從原檔搬回去，否則 Excel 打不開
            const data = await restoreExternalLinks(await wbs[x.month!].xlsx.writeBuffer() as ArrayBuffer, x.buf);
            out.push({month: x.month!, filename: stmtFilename(x.month!), data});
        }
        return out;
    };

    const saveAllStmts = (btn: HTMLButtonElement) => withBusy(btn, async () => {
        const files = await buildStmtFiles();
        if (!files.length) return;

        if (!supportsFolderSave()) {
            await downloadSequentially(files, XLSX_MIME, saveAs);
            showToast({
                variant: 'success',
                title: `已觸發 ${files.length} 份下載`,
                message: '瀏覽器可能會問是否允許下載多個檔案，請按允許；舊檔需自行刪除',
            });
            return;
        }

        // 權限仍在就直接寫；被拒絕或 handle 失效就重新選資料夾
        let dir = saveDir;
        if (dir && !(await ensureWritable(dir))) dir = null;
        if (!dir) {
            dir = await pickFolder();
            if (!dir) return;
            if (!(await ensureWritable(dir))) {
                showToast({variant: 'error', title: '沒有寫入權限', message: '請在選資料夾時允許「編輯檔案」'});
                return;
            }
            saveDir = dir;
            await saveFolderHandle(FOLDER_KEY, dir);
        }

        let failed = await writeFiles(dir, files);
        // 資料夾被搬走/刪除等 handle 失效：重選一次再寫
        if (failed.length === files.length && failed.every((f) => isHandleGone(f.error))) {
            const again = await pickFolder();
            if (!again || !(await ensureWritable(again))) return;
            dir = again;
            saveDir = again;
            await saveFolderHandle(FOLDER_KEY, again);
            failed = await writeFiles(dir, files);
        }
        renderDownloads();

        if (failed.length) {
            const names = failed.map((f) => f.filename).join('、');
            const months = failed.map((f) => `${f.filename.split('月')[0]}月`).join('、');
            showToast({
                variant: 'error',
                title: `有 ${failed.length} 份沒存成功：${names}`,
                message: `請先關閉 Excel 裡的 ${months} 檔再按一次`,
            });
            return;
        }
        showToast({variant: 'success', title: '全部存檔完成', message: `已存 ${files.length} 份到『${dir.name}』，覆蓋舊檔`});
    });

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
        const f = (await buildStmtFiles()).find((x) => x.month === target.month);
        if (!f) return;
        saveAs(new Blob([f.data], {type: XLSX_MIME}), f.filename);
        showToast({variant: 'success', title: '已輸出', message: f.filename});
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
        // 本次取消的「已確認 / 現付」列：從下載的檔案移除（原檔不動）
        if (acctFile && acctCancelled.size) {
            const drop: number[] = [];
            ws.eachRow((row, i) => {
                const v = String(cellVal(row.getCell(3)) ?? '').trim();
                const flag = String(cellVal(row.getCell(4)) ?? '').trim();
                if (acctCancelled.has(v) && isConfirmFlag(flag)) drop.push(i);
            });
            for (const i of drop.reverse()) ws.spliceRows(i, 1);
        }
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
        rowNotes = {};
        rowNotePeriods = {};
        rowCashMode = {};
        rowPick = {};
        rowPaidChoice = {};
        rowMulti = {};
        multiDraft = {};
        multiOpen = new Set();
        multiMemo = {};
        memoDrafts = {};
        acctCancelled.clear();
        pending = null;
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
