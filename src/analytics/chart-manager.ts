/**
 * Chart Manager：管理一張 ECharts 圖的生命週期。
 *
 * 提供：
 * - mount(option)：第一次或切換主題時 init
 * - update(option)：篩選或資料變動時 setOption
 * - resize()：視窗尺寸變化時呼叫
 * - dispose()：分頁卸載時釋放
 *
 * 故意不直接吐 ECharts instance，而是用 facade，
 * 避免各 adapter 直接耦合 echarts 模組。
 */

import {loadECharts} from './echarts-loader';

type EChartsType = import('echarts').ECharts;
type EChartsOption = import('echarts').EChartsOption;

export interface ChartHandle {
    setOption(option: EChartsOption): void;

    resize(): void;

    dispose(): void;

    on(event: string, handler: (params: unknown) => void): void;

    off(event: string): void;

    getInstance(): EChartsType;
}

export async function createChart(container: HTMLElement): Promise<ChartHandle> {
    const echarts = await loadECharts();
    const instance = echarts.init(container, undefined, {renderer: 'canvas'});

    return {
        setOption(option) {
            instance.setOption(option, true);
        },
        resize() {
            instance.resize();
        },
        dispose() {
            instance.dispose();
        },
        on(event, handler) {
            instance.on(event, handler);
        },
        off(event) {
            instance.off(event);
        },
        getInstance() {
            return instance;
        },
    };
}

/**
 * 建立一個 ResizeObserver，當任一個 chart container 尺寸變化時自動 resize。
 * 回傳清理函式。
 */
export function observeChartsResize(charts: ReadonlyArray<ChartHandle>): () => void {
    const ro = new ResizeObserver(() => {
        for (const c of charts) c.resize();
    });
    for (const c of charts) {
        const dom = c.getInstance().getDom() as HTMLElement;
        if (dom) ro.observe(dom);
    }
    // 視窗 resize 也觸發
    const onWindowResize = () => {
        for (const c of charts) c.resize();
    };
    window.addEventListener('resize', onWindowResize);
    return () => {
        ro.disconnect();
        window.removeEventListener('resize', onWindowResize);
    };
}
