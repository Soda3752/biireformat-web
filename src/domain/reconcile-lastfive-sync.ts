/**
 * 青坊對帳 2.0 按確認時，把帳號/備註同步寫進設定頁「末五碼」對照表。
 *
 * - 設定表欄位不變（客戶名稱 / 店家編號 / 線別 / 末五碼），CSV 與匯出格式照舊
 * - 自動加入的列另記在 AUTO_LASTFIVE_KEY：設定頁據此顯示「對帳自動加入 M/D」標籤，
 *   使用者在對帳頁按「取消」時也只移除這些列，原本就有的列不動
 */

import type {BankInfo} from '@/domain/models/bank-info';
import {toHalfWidth} from '@/domain/models/bank-info';
import {invalidateBankInfos, loadBankInfos, serializeBankInfoCsv} from '@/domain/bank-info-loader';
import {localSettings} from '@/infra/local-settings-store';

export const AUTO_LASTFIVE_KEY = 'bii.reconcileV2.autoLastFive';
/** 對帳頁改了設定表時發出；設定頁（開站就建好、不會重建）收到後重新載入，免得之後編輯時把新列蓋掉 */
export const LASTFIVE_CHANGED_EVENT = 'bii:lastfive-changed';

export interface AutoLastFiveRow {
    customerName: string;
    storeCode: string;
    customerLine: string;
    lastFiveDigit: string;
    /** 寫入這列的入帳帳號/備註（記憶的 key）；全部取消後才移除這列。 */
    keys: string[];
    /** 加入日期 M/D。 */
    at: string;
}

/**
 * 入帳帳號/備註 → 設定表的末五碼寫法：
 * 銀行隱碼「000011350**6460*」取 ** 後的數字（設定表多數寫 4 碼）；
 * 郵局帳號（含隱碼）存完整字串；其他文字（戶名、備註）存原字。
 */
export function lastFiveValueOf(key: string): string | null {
    const k = key.trim();
    if (!k) return null;
    const half = toHalfWidth(k).replace(/\s/g, '');
    const masked = /\*\*(\d+)\*/.exec(half);
    if (masked) return masked[1];
    if (/^[\d*]+$/.test(half)) return half;
    return k;
}

/** 設定表裡「同客戶＋同末五碼值」視為同一列。 */
const sameRow = (a: Pick<BankInfo, 'customerName' | 'lastFiveDigit'>, b: Pick<BankInfo, 'customerName' | 'lastFiveDigit'>): boolean =>
    a.customerName === b.customerName && a.lastFiveDigit === b.lastFiveDigit;

export function loadAutoLastFive(): AutoLastFiveRow[] {
    try {
        const v = JSON.parse(localStorage.getItem(AUTO_LASTFIVE_KEY) || '[]');
        return Array.isArray(v) ? v : [];
    } catch {
        return [];
    }
}

export function saveAutoLastFive(rows: ReadonlyArray<AutoLastFiveRow>): void {
    try {
        localStorage.setItem(AUTO_LASTFIVE_KEY, JSON.stringify(rows));
    } catch {
        // 存不了就算了：設定表本身已寫入，只是少了標籤
    }
}

/** 設定頁用：這列是不是對帳自動加入的（回傳紀錄，否則 undefined）。 */
export function autoLastFiveOf(row: Pick<BankInfo, 'customerName' | 'storeCode' | 'lastFiveDigit'>, autos: ReadonlyArray<AutoLastFiveRow>): AutoLastFiveRow | undefined {
    return autos.find((a) => sameRow(a, row) && a.storeCode === row.storeCode);
}

const writeInfos = (infos: BankInfo[]): void => {
    localSettings.setLastFiveDigit(serializeBankInfoCsv(infos));
    invalidateBankInfos();
    window.dispatchEvent(new Event(LASTFIVE_CHANGED_EVENT));
};

/**
 * 確認後把對到的各客戶（明細代號、名稱）加進設定表；已有同客戶＋同末五碼值的列就不重複加。
 * 回傳新加入的列數。
 */
export function addConfirmedLastFive(key: string, custs: ReadonlyArray<{ code: string; name: string }>, now = new Date()): number {
    const val = lastFiveValueOf(key);
    if (!val || !custs.length) return 0;
    const infos = [...loadBankInfos()];
    const autos = loadAutoLastFive();
    const at = `${now.getMonth() + 1}/${now.getDate()}`;
    let added = 0;
    for (const {code, name} of custs) {
        const row: BankInfo = {customerName: name, storeCode: code, customerLine: code.charAt(0), lastFiveDigit: val};
        const auto = autos.find((a) => sameRow(a, row));
        if (infos.some((x) => sameRow(x, row))) {
            // 已是自動加入的列：記下這個帳號也用到它，取消時要全部取消才移除
            if (auto && !auto.keys.includes(key)) auto.keys.push(key);
            continue;
        }
        infos.push(row);
        if (auto) Object.assign(auto, row, {keys: [key], at});
        else autos.push({...row, keys: [key], at});
        added++;
    }
    // 先存自動加入清單再寫設定表：寫設定表會通知設定頁重新載入，要讀得到新標籤
    saveAutoLastFive(autos);
    if (added) writeInfos(infos);
    return added;
}

/** 取消確認：移除這個帳號自動加入的列（其他帳號也用到的、使用者原本就有的不動）。回傳移除列數。 */
export function removeConfirmedLastFive(key: string): number {
    const autos = loadAutoLastFive();
    const gone: AutoLastFiveRow[] = [];
    const keep: AutoLastFiveRow[] = [];
    for (const a of autos) {
        if (!a.keys.includes(key)) {
            keep.push(a);
            continue;
        }
        a.keys = a.keys.filter((k) => k !== key);
        (a.keys.length ? keep : gone).push(a);
    }
    saveAutoLastFive(keep);
    if (!gone.length) return 0;
    const infos = loadBankInfos();
    const next = infos.filter((x) => !gone.some((a) => sameRow(a, x) && a.storeCode === x.storeCode));
    if (next.length !== infos.length) writeInfos(next);
    return infos.length - next.length;
}
