import type { BankInfo } from '@/domain/models/bank-info';
import { isSameAccount } from '@/domain/models/bank-info';
import { createWorkbook, workbookToBlob } from '@/infra/excel-service';

/**
 * 對應桌面版 BankNameInfoWriter.kt
 * - 將 originData（CSV）逐列輸出，並把每列尾端的隱碼帳號比對 bankInfos，
 *   把比對到的 customerName / customerLine 緊接在該列尾端附加上去。
 * - 每筆 BankInfo 占兩欄（name、line），多筆則持續展開往右。
 */
export const writeBankNameMerged = async (
  bankInfos: ReadonlyArray<BankInfo>,
  originData: ReadonlyArray<ReadonlyArray<string>>
): Promise<Blob> => {
  const workbook = createWorkbook();
  const sheet = workbook.addWorksheet('對帳結果');

  for (const rowInfo of originData) {
    const trimmed = dropTrailingBlanks(rowInfo);
    const account = trimmed[trimmed.length - 1];
    const matches = findMatches(account, bankInfos);

    const cells: Array<string | number | null> = [...trimmed];
    for (const info of matches) {
      cells.push(info.customerName);
      cells.push(info.customerLine);
    }
    sheet.addRow(cells);
  }

  return workbookToBlob(workbook);
};

const dropTrailingBlanks = (row: ReadonlyArray<string>): string[] => {
  const out = [...row];
  while (out.length > 0 && (out[out.length - 1] ?? '').trim() === '') {
    out.pop();
  }
  return out;
};

const findMatches = (
  account: string | undefined,
  bankInfos: ReadonlyArray<BankInfo>
): BankInfo[] => {
  if (!account || account.trim().length === 0) return [];
  const trimmed = account.trim();
  return bankInfos.filter((info) => isSameAccount(info, trimmed));
};

/** 對帳結果_${YYYYMMDD}.xlsx */
export const buildBankResultFilename = (date = new Date()): string => {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `對帳結果_${yyyy}${mm}${dd}.xlsx`;
};
