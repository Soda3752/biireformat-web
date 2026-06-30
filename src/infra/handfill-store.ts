/**
 * 「生成手填本」localStorage 儲存層。
 *
 * Schema：
 *  - 每份手填本獨立一個 key：`bii.handfill.book.<id>`，value = `HandfillBook` JSON
 *  - 「當前工作中 id」存於 `bii.handfill.activeId`（重整後可恢復上次編輯）
 */

import type {HandfillBook} from '@/domain/models/handfill-book';
import {lineFullName} from '@/domain/models/handfill-book';

const KEY_PREFIX = 'bii.handfill.book.';
const ACTIVE_KEY = 'bii.handfill.activeId';

/** 「匯出全部歷史紀錄」備份檔信封標記，供寫出與匯入驗證共用。 */
export const HANDFILL_BACKUP_TYPE = 'bii.handfill.backup';
export const HANDFILL_BACKUP_VERSION = 1;

function bookKey(id: string): string {
    return KEY_PREFIX + id;
}

export interface HandfillBookSummary {
    id: string;
    lineNo: number;
    lineName: string;
    year: number;
    month: number;
    customerCount: number;
    updatedAt: number;
    fullName: string;
}

export function saveBook(book: HandfillBook): void {
    book.updatedAt = Date.now();
    try {
        localStorage.setItem(bookKey(book.id), JSON.stringify(book));
    } catch (err) {
        console.error('[handfill-store] save failed', err);
        throw err;
    }
}

export function loadBook(id: string): HandfillBook | null {
    try {
        const raw = localStorage.getItem(bookKey(id));
        if (!raw) return null;
        return JSON.parse(raw) as HandfillBook;
    } catch (err) {
        console.warn('[handfill-store] load failed', err);
        return null;
    }
}

export function deleteBook(id: string): void {
    try {
        localStorage.removeItem(bookKey(id));
        if (getActiveId() === id) {
            setActiveId(null);
        }
    } catch (err) {
        console.warn('[handfill-store] delete failed', err);
    }
}

export function listBooks(): HandfillBookSummary[] {
    const summaries: HandfillBookSummary[] = [];
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key || !key.startsWith(KEY_PREFIX)) continue;
        try {
            const raw = localStorage.getItem(key);
            if (!raw) continue;
            const book = JSON.parse(raw) as HandfillBook;
            summaries.push({
                id: book.id,
                lineNo: book.lineNo,
                lineName: book.lineName,
                year: book.year,
                month: book.month,
                customerCount: book.customers?.length ?? 0,
                updatedAt: book.updatedAt ?? 0,
                fullName: lineFullName(book),
            });
        } catch {
            // skip corrupted entry
        }
    }
    summaries.sort((a, b) => b.updatedAt - a.updatedAt);
    return summaries;
}

/** 讀出 localStorage 中所有手填本完整內容（用於「匯出全部歷史紀錄」備份）。 */
export function exportAllBooks(): HandfillBook[] {
    const books: HandfillBook[] = [];
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key || !key.startsWith(KEY_PREFIX)) continue;
        try {
            const raw = localStorage.getItem(key);
            if (!raw) continue;
            books.push(JSON.parse(raw) as HandfillBook);
        } catch {
            // skip corrupted entry
        }
    }
    books.sort((a, b) => (b.updatedAt ?? 0) - (a.updatedAt ?? 0));
    return books;
}

/**
 * 合併＋覆蓋匯入：依 id 寫入。
 *  - 同 id 已存在 → 以匯入內容覆蓋（updated）
 *  - 新 id → 新增（added）
 * 保留 book 自身的 createdAt / updatedAt（不重新蓋上現在時間），確保備份還原後時間正確。
 */
export function importBooks(books: HandfillBook[]): {added: number; updated: number} {
    let added = 0;
    let updated = 0;
    for (const book of books) {
        const exists = localStorage.getItem(bookKey(book.id)) !== null;
        localStorage.setItem(bookKey(book.id), JSON.stringify(book));
        if (exists) updated++;
        else added++;
    }
    return {added, updated};
}

export function getActiveId(): string | null {
    try {
        return localStorage.getItem(ACTIVE_KEY);
    } catch {
        return null;
    }
}

export function setActiveId(id: string | null): void {
    try {
        if (id === null) localStorage.removeItem(ACTIVE_KEY);
        else localStorage.setItem(ACTIVE_KEY, id);
    } catch {
        // ignore
    }
}
