import Papa from 'papaparse';

/**
 * 對應桌面版 TransRecordReader.kt
 * 讀取銀行對帳單 .csv（桌面版預設 Big5 編碼）。
 * 透過 TextDecoder 偵測 / fallback：先試 Big5，若解碼後出現大量替換字元 (U+FFFD) 改用 UTF-8。
 * 解析後輸出 string[][]（與桌面版 List<List<String>> 等價）。
 */
export const parseTransRecord = async (file: File): Promise<string[][]> => {
  const buffer = await file.arrayBuffer();
  const text = decodeCsv(buffer);

  const result = Papa.parse<string[]>(text, {
    header: false,
    skipEmptyLines: 'greedy',
    transform: (value) => (typeof value === 'string' ? value : String(value ?? '')),
  });

  if (result.errors.length > 0) {
    const firstFatal = result.errors.find((e) => e.type !== 'FieldMismatch');
    if (firstFatal) {
      throw new Error(`CSV 解析失敗：${firstFatal.message}`);
    }
  }

  return (result.data as string[][]).filter((row) => Array.isArray(row));
};

/**
 * 偵測編碼：UTF-8 BOM 直接走 UTF-8；
 * 否則先嘗試 Big5，若 replacement char 比例過高就退回 UTF-8。
 */
const decodeCsv = (buffer: ArrayBuffer): string => {
  const bytes = new Uint8Array(buffer);
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }

  try {
    const big5 = new TextDecoder('big5', { fatal: false }).decode(bytes);
    if (replacementRatio(big5) < 0.02) return big5;
  } catch {
    // ignore，走 utf-8
  }

  return new TextDecoder('utf-8', { fatal: false }).decode(bytes);
};

const replacementRatio = (s: string): number => {
  if (s.length === 0) return 0;
  let count = 0;
  for (const ch of s) if (ch === '�') count++;
  return count / s.length;
};
