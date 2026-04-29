import {loadDailyReportTemplate} from '@/domain/daily-report-loader';
import {ExcelRowType} from '@/domain/excel-row-data';
import {cloneProductMap, type DailyProduct, type DailyProductMap,} from '@/domain/models/daily-product';
import {DEFAULT_PRODUCT_SETTING, parseProductSetting, type ProductSetting,} from '@/domain/models/product-setting';
import {parseBillFile} from '@/readers/bill-reader';

/**
 * 多日 xlsx 解析結果。每個日期 key 對應該日商品名 → 加總數量。
 * dates 為遞增排序的 YYYY-MM-DD 列表，供月曆顯示「有資料」的日期。
 */
export interface DailyCountParseResult {
    readonly byDate: ReadonlyMap<string, ReadonlyMap<string, number>>;
    readonly availableDates: ReadonlyArray<string>;
}

/** 對應 UI 端單日的 metric 資料；維持與舊介面相容。 */
export interface DailyCountResult {
  readonly map: DailyProductMap;
  readonly matched: number;
  readonly otherCount: number;
}

const DATE_RE = /^(\d{4})\/(\d{1,2})\/(\d{1,2})/;

/**
 * 讀取「應收帳款對帳單明細表」格式 xlsx，依日期分桶累積商品數量。
 * 每位客戶各有自己的欄位設定（ProductRowSetting）；遇到新客戶（CustomerData）時重置為 default。
 */
export const parseDailyCount = async (
    file: File | Blob | ArrayBuffer
): Promise<DailyCountParseResult> => {
    const byDate = new Map<string, Map<string, number>>();
    let setting: ProductSetting = DEFAULT_PRODUCT_SETTING;

    await parseBillFile(file, ({type, values}) => {
        switch (type) {
            case ExcelRowType.CustomerData:
                setting = DEFAULT_PRODUCT_SETTING;
                return;
            case ExcelRowType.ProductRowSetting:
                setting = parseProductSetting(values);
                return;
            case ExcelRowType.ProductSellData: {
                const dateRaw = values[setting.orderDateIndex] ?? '';
                const dateStr = dateRaw.replace('銷', '').trim();
                const m = DATE_RE.exec(dateStr);
                if (!m) return;
                const key = `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;

                const productName = (values[setting.productNameIndex] ?? '').trim();
                if (!productName) return;

                const count = Number(values[setting.productCountIndex]);
                if (!Number.isFinite(count) || count === 0) return;

                let day = byDate.get(key);
                if (!day) {
                    day = new Map();
                    byDate.set(key, day);
                }
                day.set(productName, (day.get(productName) ?? 0) + count);
                return;
            }
            default:
                return;
        }
    });

    const availableDates = [...byDate.keys()].sort();
    return {byDate, availableDates};
};

/**
 * 依所選日期，從 byDate 取出該日商品數量，再對應 daily_report_list.csv 模板分組；
 * 對不到的商品歸入「其他」。回傳的 map 已是該日獨立資料（不會汙染 cache）。
 */
export const buildDailyResultForDate = async (
    byDate: ReadonlyMap<string, ReadonlyMap<string, number>>,
    dateKey: string
): Promise<DailyCountResult> => {
  const template = await loadDailyReportTemplate();
  const productMap = cloneProductMap(template);

    const nameIndex = new Map<string, DailyProduct>();
  for (const products of productMap.values()) {
      for (const p of products) nameIndex.set(p.name, p);
  }

    const dayMap = byDate.get(dateKey);
  let matched = 0;
  let otherCount = 0;

    if (dayMap) {
        for (const [productName, count] of dayMap) {
            const existing = nameIndex.get(productName);
            if (existing) {
                existing.count = count;
                matched++;
            } else {
                const fallback: DailyProduct = {
                    code: '',
                    name: productName,
                    groupName: '其他',
                    costRaw: '',
                    count,
                };
                let other = productMap.get('其他');
                if (!other) {
                    other = [];
                    productMap.set('其他', other);
                }
                other.push(fallback);
                nameIndex.set(productName, fallback);
                otherCount++;
            }
    }
  }

  return { map: productMap, matched, otherCount };
};
