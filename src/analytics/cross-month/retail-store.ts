/**
 * 零售客戶標記 store：以「客戶代碼」為 key 的全域 Set，localStorage 持久化。
 *
 * 用途：跨月分析中，使用者可把「新增/流失」列表裡的某客戶標為零售（一次性零買、
 * 非真正的新客或流失客）。被標記的客戶會從 新增/流失 tab 移到 零售新增/零售流失 tab。
 *
 * 標記是全域的（不限於某對月份），同一客戶代碼在所有相鄰月對中皆視為零售。
 */

const STORAGE_KEY = 'cross-month-retail-codes';

let cache: Set<string> | null = null;
const listeners = new Set<() => void>();

function load(): Set<string> {
    if (cache) return cache;
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        const arr = raw ? JSON.parse(raw) : [];
        cache = new Set(Array.isArray(arr) ? arr.filter((x) => typeof x === 'string') : []);
    } catch {
        cache = new Set();
    }
    return cache;
}

function persist(): void {
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify([...load()]));
    } catch (err) {
        console.warn('[retail-store] persist failed', err);
    }
}

export function isRetail(customerCode: string): boolean {
    return load().has(customerCode);
}

export function setRetail(customerCode: string, retail: boolean): void {
    const set = load();
    const before = set.has(customerCode);
    if (retail) set.add(customerCode);
    else set.delete(customerCode);
    if (set.has(customerCode) === before) return;
    persist();
    listeners.forEach((fn) => fn());
}

export function subscribeRetail(fn: () => void): () => void {
    listeners.add(fn);
    return () => {
        listeners.delete(fn);
    };
}
