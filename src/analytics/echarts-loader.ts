/**
 * ECharts dynamic loader：
 * 切換到「數據分析」分頁第一次需要繪圖時才載入 ~1MB 的 echarts，
 * 其他 5 個分頁的首屏載入完全不受影響。
 */

type EChartsModule = typeof import('echarts');

let cached: Promise<EChartsModule> | null = null;

export const loadECharts = (): Promise<EChartsModule> => {
    if (!cached) {
        cached = import('echarts');
    }
    return cached;
};
