/**
 * 銀行對帳預覽表（精簡 5 欄：日期 / 摘要 / 存入 / 帳號 / 配對）。
 *
 * 用於上傳對帳單後、輸出前的視覺檢查。已配對列僅顯示客戶/線別 chip（唯讀）；
 * 未配對列在配對欄顯示警示文字 + 「補建」按鈕，呼叫端可在 onAddClick 開啟 dialog。
 */

import type {BankMatchResult, BankRowMatch} from '@/domain/bank-match-service';
import {icon} from '@/ui/icons';

export interface BankPreviewTableOptions {
    onAddClick: (row: BankRowMatch) => void;
}

export interface BankPreviewTableHandle {
    readonly element: HTMLElement;

    setData(result: BankMatchResult): void;

    clear(): void;
}

export function createBankPreviewTable(options: BankPreviewTableOptions): BankPreviewTableHandle {
    const element = document.createElement('section');
    element.className = 'bank-preview';
    element.hidden = true;

    element.innerHTML = `
      <div class="bank-preview-summary" data-role="summary"></div>
      <div class="bank-preview-scroll">
        <table class="bank-preview-table">
          <thead>
            <tr>
              <th class="col-date">日期</th>
              <th class="col-summary">摘要</th>
              <th class="col-deposit">存入</th>
              <th class="col-account">帳號</th>
              <th class="col-match">配對</th>
            </tr>
          </thead>
          <tbody data-role="tbody"></tbody>
        </table>
      </div>
    `;

    const summaryEl = element.querySelector<HTMLElement>('[data-role="summary"]')!;
    const tbody = element.querySelector<HTMLElement>('[data-role="tbody"]')!;

    const setData = (result: BankMatchResult) => {
        renderSummary(summaryEl, result);
        renderRows(tbody, result.rows, options.onAddClick);
        element.hidden = false;
    };

    const clear = () => {
        summaryEl.innerHTML = '';
        tbody.innerHTML = '';
        element.hidden = true;
    };

    return {element, setData, clear};
}

function renderSummary(host: HTMLElement, result: BankMatchResult): void {
    const total = result.rows.length;
    const matched = result.matchedCount;
    const unmatched = result.unmatchedCount;
    const unmatchedClass = unmatched > 0 ? 'bank-summary-chip is-warn' : 'bank-summary-chip';
    host.innerHTML = `
      <span class="bank-summary-chip">總筆數 <strong>${total}</strong></span>
      <span class="bank-summary-chip is-ok">已配對 <strong>${matched}</strong></span>
      <span class="${unmatchedClass}">未配對 <strong>${unmatched}</strong></span>
    `;
}

function renderRows(
    tbody: HTMLElement,
    rows: ReadonlyArray<BankRowMatch>,
    onAddClick: (row: BankRowMatch) => void
): void {
    tbody.innerHTML = '';
    if (rows.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td colspan="5" class="bank-preview-empty">無資料列</td>`;
        tbody.appendChild(tr);
        return;
    }

    const fragment = document.createDocumentFragment();
    for (const row of rows) {
        const tr = document.createElement('tr');
        if (row.matches.length === 0) tr.classList.add('is-unmatched');

        const dateTd = document.createElement('td');
        dateTd.className = 'col-date';
        dateTd.textContent = row.date;

        const summaryTd = document.createElement('td');
        summaryTd.className = 'col-summary';
        summaryTd.textContent = row.summary;

        const depositTd = document.createElement('td');
        depositTd.className = 'col-deposit';
        depositTd.textContent = row.deposit;

        const accountTd = document.createElement('td');
        accountTd.className = 'col-account';
        accountTd.textContent = row.account;
        accountTd.title = row.account;

        const matchTd = document.createElement('td');
        matchTd.className = 'col-match';
        if (row.matches.length === 0) {
            const warn = document.createElement('span');
            warn.className = 'bank-match-warn';
            warn.innerHTML = `${icon('alert', 14)}<span>未配對</span>`;
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-secondary btn-xs bank-match-add';
            btn.textContent = '補建';
            btn.addEventListener('click', () => onAddClick(row));
            matchTd.appendChild(warn);
            matchTd.appendChild(btn);
        } else {
            for (const info of row.matches) {
                const chip = document.createElement('span');
                chip.className = 'bank-match-chip';
                const name = document.createElement('strong');
                name.textContent = info.customerName;
                chip.appendChild(name);
                if (info.customerLine) {
                    const sep = document.createTextNode(' / ');
                    const line = document.createElement('span');
                    line.textContent = info.customerLine;
                    chip.appendChild(sep);
                    chip.appendChild(line);
                }
                matchTd.appendChild(chip);
            }
        }

        tr.appendChild(dateTd);
        tr.appendChild(summaryTd);
        tr.appendChild(depositTd);
        tr.appendChild(accountTd);
        tr.appendChild(matchTd);
        fragment.appendChild(tr);
    }
    tbody.appendChild(fragment);
}
