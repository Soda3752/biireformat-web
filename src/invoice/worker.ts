// 二聯發票：在背景執行緒做 Excel 解析與計算，畫面不會卡住；檔案不離開瀏覽器。
import {computeFromBuffers, type WebInput} from './webCalc.ts';

const ctx = self as unknown as {
    onmessage: ((e: MessageEvent<WebInput>) => void) | null;
    postMessage: (msg: unknown, transfer?: Transferable[]) => void;
};
ctx.onmessage = async (e) => {
    try {
        const out = await computeFromBuffers(e.data);
        ctx.postMessage({ok: true, out}, [out.report]);
    } catch (err) {
        ctx.postMessage({ok: false, error: err instanceof Error ? err.message : String(err)});
    }
};
