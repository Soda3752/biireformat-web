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
