/**
 * 順豐託運單分頁的「收件人最愛」本地儲存。
 *
 * 以「姓名 + 手機」組合做唯一鍵；資料整包以 JSON 存 localStorage。
 */

export const SF_FAVORITE_RECEIVERS_KEY = 'bii.sfShipping.favoriteReceivers.json';
const STORAGE_KEY = SF_FAVORITE_RECEIVERS_KEY;

export interface FavoriteReceiver {
    id: string;
    name: string;
    phone: string;
    address: string;
    updatedAt: string;
}

function read(): FavoriteReceiver[] {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return [];
        const arr = JSON.parse(raw);
        if (!Array.isArray(arr)) return [];
        return arr.filter(isValidReceiver);
    } catch {
        return [];
    }
}

function write(list: FavoriteReceiver[]): void {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(list));
}

function isValidReceiver(x: unknown): x is FavoriteReceiver {
    if (!x || typeof x !== 'object') return false;
    const o = x as Record<string, unknown>;
    return (
        typeof o.id === 'string' &&
        typeof o.name === 'string' &&
        typeof o.phone === 'string' &&
        typeof o.address === 'string' &&
        typeof o.updatedAt === 'string'
    );
}

export function listFavoriteReceivers(): FavoriteReceiver[] {
    return read().sort(byUpdatedAtDesc);
}

function byUpdatedAtDesc(a: FavoriteReceiver, b: FavoriteReceiver): number {
    return b.updatedAt.localeCompare(a.updatedAt);
}

export function findFavoriteByKey(name: string, phone: string): FavoriteReceiver | null {
    const n = name.trim();
    const p = phone.trim();
    if (n.length === 0 || p.length === 0) return null;
    const list = read();
    return list.find((r) => r.name === n && r.phone === p) ?? null;
}

export interface AddFavoriteInput {
    name: string;
    phone: string;
    address: string;
}

/** 新增；若「姓名+手機」已存在則回傳既有項目並更新地址。 */
export function addFavoriteReceiver(input: AddFavoriteInput): {
    receiver: FavoriteReceiver;
    isNew: boolean;
} {
    const list = read();
    const name = input.name.trim();
    const phone = input.phone.trim();
    const address = input.address.trim();
    const now = new Date().toISOString();

    const idx = list.findIndex((r) => r.name === name && r.phone === phone);
    if (idx >= 0) {
        const updated: FavoriteReceiver = {
            ...list[idx],
            address,
            updatedAt: now,
        };
        list[idx] = updated;
        write(list);
        return {receiver: updated, isNew: false};
    }

    const receiver: FavoriteReceiver = {
        id: generateId(),
        name,
        phone,
        address,
        updatedAt: now,
    };
    list.push(receiver);
    write(list);
    return {receiver, isNew: true};
}

export function updateFavoriteReceiver(
    id: string,
    patch: Partial<Pick<FavoriteReceiver, 'name' | 'phone' | 'address'>>
): FavoriteReceiver | null {
    const list = read();
    const idx = list.findIndex((r) => r.id === id);
    if (idx < 0) return null;
    const next: FavoriteReceiver = {
        ...list[idx],
        ...(patch.name !== undefined ? {name: patch.name.trim()} : {}),
        ...(patch.phone !== undefined ? {phone: patch.phone.trim()} : {}),
        ...(patch.address !== undefined ? {address: patch.address.trim()} : {}),
        updatedAt: new Date().toISOString(),
    };
    list[idx] = next;
    write(list);
    return next;
}

export function deleteFavoriteReceiver(id: string): boolean {
    const list = read();
    const next = list.filter((r) => r.id !== id);
    if (next.length === list.length) return false;
    write(next);
    return true;
}

export function getFavoritesJson(): string | null {
    try {
        return localStorage.getItem(STORAGE_KEY);
    } catch {
        return null;
    }
}

export function setFavoritesJson(value: string | null): void {
    if (value === null) {
        try {
            localStorage.removeItem(STORAGE_KEY);
        } catch {
            // ignore
        }
        return;
    }
    // 驗證後寫入
    try {
        const arr = JSON.parse(value);
        if (!Array.isArray(arr)) throw new Error('not an array');
        const valid = arr.filter(isValidReceiver);
        write(valid);
    } catch {
        // ignore invalid JSON; 不寫入避免清除既有資料
    }
}

export function countFavoriteReceivers(): number {
    return read().length;
}

function generateId(): string {
    if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
        return crypto.randomUUID();
    }
    return `r_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
}
