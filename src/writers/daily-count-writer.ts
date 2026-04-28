import {
  getTotalCount,
  type DailyProductMap,
} from '@/domain/models/daily-product';
import { createWorkbook, workbookToBlob } from '@/infra/excel-service';

/**
 * 對應桌面版 DailyCountViewModel.exportToExcelFile()
 * 輸出格式：
 * - 第 1 列：日期 MM/dd(週X)
 * - 第 2 列：類型 / 品名 / 數量 / 總量
 * - 後續每組：第一列填類型 + 該組總量；同組其餘列只填品名與數量
 */
export const writeDailyCount = async (
  map: DailyProductMap,
  date: Date = new Date()
): Promise<Blob> => {
  const workbook = createWorkbook();
  const sheet = workbook.addWorksheet('單日統計');

  sheet.addRow([formatDateDisplay(date)]);
  sheet.addRow(['類型', '品名', '數量', '總量']);

  for (const [groupName, products] of map) {
    if (products.length === 0) continue;
    const groupTotal = products.reduce((sum, p) => sum + getTotalCount(p), 0);
    products.forEach((product, index) => {
      sheet.addRow([
        index === 0 ? groupName : '',
        product.name,
        getTotalCount(product),
        index === 0 ? groupTotal : '',
      ]);
    });
  }

  return workbookToBlob(workbook);
};

const CHINESE_WEEKDAY = ['日', '一', '二', '三', '四', '五', '六'] as const;

/** MM/dd(X) — 對應桌面版 DailyCountViewModel：dateStr + "($dayOfWeek)" */
export const formatDateDisplay = (date: Date): string => {
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  const weekday = CHINESE_WEEKDAY[date.getDay()];
  return `${mm}/${dd}(${weekday})`;
};

/** 單日數量_${YYYYMMDD}.xlsx */
export const buildDailyCountFilename = (date = new Date()): string => {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const dd = String(date.getDate()).padStart(2, '0');
  return `單日數量_${yyyy}${mm}${dd}.xlsx`;
};
