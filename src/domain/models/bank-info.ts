/**
 * 對應桌面版 src/main/kotlin/bankNameFormat/core/models/BankInfo.kt
 *
 * 用 lastFiveDigit 與隱碼帳號 (account) 比對。隱碼字 '*' 視為 wildcard。
 * - 純數字（含 *）走「反向比對」：accountInHalf.reversed() vs lastFiveDigit.reversed()
 * - 中文則走 prefix 比對（account 取 lastFiveDigit.length 前綴）
 */
export interface BankInfo {
  readonly customerName: string;
  /** 設定頁紀錄的店家編號；舊資料或未填寫時為空字串。不參與比對 / 去重邏輯。 */
  readonly storeCode: string;
  readonly customerLine: string;
  readonly lastFiveDigit: string;
}

const isAllDigit = (s: string): boolean => /^[0-9*]+$/.test(s);

export const toHalfWidth = (input: string): string => {
  let out = '';
  for (const ch of input) {
    const code = ch.codePointAt(0)!;
    if (code >= 0xff01 && code <= 0xff5e) {
      out += String.fromCodePoint(code - 0xfee0);
    } else if (code === 0x3000) {
      out += ' ';
    } else {
      out += ch;
    }
  }
  return out;
};

const reverseString = (s: string): string => [...s].reverse().join('');

export const isSameAccount = (info: BankInfo, account: string): boolean => {
  if (isAllDigit(info.lastFiveDigit) && isAllDigit(account)) {
    const accountInHalf = reverseString(toHalfWidth(account));
    const lastFiveSource = info.lastFiveDigit.length === 4 ? `${info.lastFiveDigit}*` : info.lastFiveDigit;
    const lastFiveReversed = reverseString(lastFiveSource);

    for (let i = 0; i < lastFiveReversed.length; i++) {
      const ch = lastFiveReversed[i];
      const accountCh = accountInHalf[i];
      if (accountCh === undefined) return false;
      if (ch !== accountCh && accountCh !== '*' && ch !== '*') return false;
    }
    return true;
  }
  // 中文字 → prefix 比對
  return info.lastFiveDigit === account.slice(0, info.lastFiveDigit.length);
};

export const equalsBankInfo = (a: BankInfo, b: BankInfo): boolean =>
  a.customerName === b.customerName &&
  a.customerLine === b.customerLine &&
  a.lastFiveDigit === b.lastFiveDigit;
