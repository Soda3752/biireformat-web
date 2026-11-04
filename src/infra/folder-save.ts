/**
 * 一鍵存多個檔到使用者選的資料夾（File System Access API）。
 * - 資料夾 handle 存在 IndexedDB（per-origin）；IndexedDB 失敗不影響功能，只是每次都要重選。
 * - 同名檔直接覆蓋（createWritable 預設行為），不會產生 (1)。
 */

const DB_NAME = 'biireformat-folder-save';
const STORE = 'handles';

export interface SaveFile {
    filename: string;
    data: ArrayBuffer;
}

export type FolderSaveError = {filename: string; error: unknown};

export const supportsFolderSave = (): boolean =>
    typeof (window as unknown as {showDirectoryPicker?: unknown}).showDirectoryPicker === 'function';

const openDb = (): Promise<IDBDatabase> =>
    new Promise((resolve, reject) => {
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = () => req.result.createObjectStore(STORE);
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => reject(req.error);
    });

export const loadFolderHandle = async (key: string): Promise<FileSystemDirectoryHandle | null> => {
    try {
        const db = await openDb();
        return await new Promise((resolve) => {
            const req = db.transaction(STORE).objectStore(STORE).get(key);
            req.onsuccess = () => resolve((req.result as FileSystemDirectoryHandle | undefined) ?? null);
            req.onerror = () => resolve(null);
        });
    } catch {
        return null;
    }
};

export const saveFolderHandle = async (key: string, handle: FileSystemDirectoryHandle): Promise<void> => {
    try {
        const db = await openDb();
        await new Promise<void>((resolve) => {
            const tx = db.transaction(STORE, 'readwrite');
            tx.objectStore(STORE).put(handle, key);
            tx.oncomplete = () => resolve();
            tx.onerror = () => resolve();
            tx.onabort = () => resolve();
        });
    } catch {
        // 存不起來就算了，下次再選
    }
};

type PermHandle = FileSystemDirectoryHandle & {
    queryPermission(d: {mode: 'readwrite'}): Promise<PermissionState>;
    requestPermission(d: {mode: 'readwrite'}): Promise<PermissionState>;
};

/** 確認有 readwrite 權限；沒有就嘗試要（需在使用者點擊手勢內呼叫） */
export const ensureWritable = async (handle: FileSystemDirectoryHandle): Promise<boolean> => {
    const h = handle as PermHandle;
    try {
        if (await h.queryPermission({mode: 'readwrite'}) === 'granted') return true;
        return await h.requestPermission({mode: 'readwrite'}) === 'granted';
    } catch {
        return false;
    }
};

/** 請使用者選資料夾；使用者取消回 null */
export const pickFolder = async (): Promise<FileSystemDirectoryHandle | null> => {
    try {
        return await (window as unknown as {
            showDirectoryPicker(o: {mode: 'readwrite'; id?: string}): Promise<FileSystemDirectoryHandle>;
        }).showDirectoryPicker({mode: 'readwrite', id: 'biireformat-recon'});
    } catch (err) {
        if (err instanceof DOMException && err.name === 'AbortError') return null;
        throw err;
    }
};

/** 檔案的大小與修改時間（用來判斷資料夾裡的檔在上傳後有沒有被改過）。 */
export interface FileStamp {
    size: number;
    lastModified: number;
}

export const stampOf = (f: Pick<File, 'size' | 'lastModified'>): FileStamp => ({size: f.size, lastModified: f.lastModified});

export const sameStamp = (a: FileStamp, b: FileStamp): boolean => a.size === b.size && a.lastModified === b.lastModified;

/** 資料夾裡某個檔的大小與修改時間；沒有這個檔回 null。 */
export const statFile = async (dir: FileSystemDirectoryHandle, filename: string): Promise<FileStamp | null> => {
    try {
        return stampOf(await (await dir.getFileHandle(filename)).getFile());
    } catch (err) {
        if (err instanceof DOMException && (err.name === 'NotFoundError' || err.name === 'TypeMismatchError')) return null;
        throw err;
    }
};

export const BACKUP_DIR = '_備份';
export const BACKUP_KEEP = 10;

const pad2 = (n: number): string => String(n).padStart(2, '0');
/** 備份檔名用的時間戳：20261104-161306 */
export const backupStamp = (d: Date): string =>
    `${d.getFullYear()}${pad2(d.getMonth() + 1)}${pad2(d.getDate())}-${pad2(d.getHours())}${pad2(d.getMinutes())}${pad2(d.getSeconds())}`;

const splitExt = (filename: string): [string, string] => {
    const i = filename.lastIndexOf('.');
    return i > 0 ? [filename.slice(0, i), filename.slice(i)] : [filename, ''];
};
const escapeRe = (s: string): string => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * 覆蓋前備份：資料夾裡已有的同名檔複製到「_備份」子資料夾，檔名加時間戳（例：9月_已自動對帳_20261104-161306.xlsx）。
 * 每個檔只留最近 BACKUP_KEEP 份，只刪符合「原檔名_時間戳」格式、由這裡產生的舊備份。
 * 回傳備份了幾個檔；備份失敗直接丟錯（呼叫端應中止覆蓋）。
 */
export const backupFiles = async (
    dir: FileSystemDirectoryHandle,
    filenames: ReadonlyArray<string>,
    now: Date = new Date(),
    keep: number = BACKUP_KEEP,
): Promise<number> => {
    const existing: Array<{name: string; file: File}> = [];
    for (const name of filenames) {
        try {
            existing.push({name, file: await (await dir.getFileHandle(name)).getFile()});
        } catch (err) {
            if (err instanceof DOMException && err.name === 'NotFoundError') continue;
            throw err;
        }
    }
    if (!existing.length) return 0;
    const bdir = await dir.getDirectoryHandle(BACKUP_DIR, {create: true});
    const stamp = backupStamp(now);
    for (const {name, file} of existing) {
        const [base, ext] = splitExt(name);
        const fh = await bdir.getFileHandle(`${base}_${stamp}${ext}`, {create: true});
        const w = await fh.createWritable();
        try {
            await w.write(await file.arrayBuffer());
            await w.close();
        } catch (err) {
            await w.abort().catch(() => undefined);
            throw err;
        }
    }
    // 清掉多的舊備份（時間戳格式固定，字串排序＝時間排序）
    const names: string[] = [];
    for await (const [n, h] of (bdir as unknown as {entries(): AsyncIterable<[string, FileSystemHandle]>}).entries()) {
        if (h.kind === 'file') names.push(n);
    }
    for (const {name} of existing) {
        const [base, ext] = splitExt(name);
        const re = new RegExp(`^${escapeRe(base)}_\\d{8}-\\d{6}${escapeRe(ext)}$`);
        const mine = names.filter((n) => re.test(n)).sort();
        for (const old of mine.slice(0, Math.max(0, mine.length - keep))) {
            await bdir.removeEntry(old).catch(() => undefined);
        }
    }
    return existing.length;
};

/** 逐檔寫入（同名覆蓋）；close() 成功後再讀回大小比對，一致才算成功。回傳失敗清單，不因單檔失敗中斷其他檔。
 *  onWritten：每個檔寫入並確認成功後呼叫，帶回寫好的檔（可取大小與修改時間）。 */
export const writeFiles = async (
    dir: FileSystemDirectoryHandle,
    files: SaveFile[],
    onWritten?: (filename: string, written: File) => void,
): Promise<FolderSaveError[]> => {
    const failed: FolderSaveError[] = [];
    for (const f of files) {
        try {
            const fh = await dir.getFileHandle(f.filename, {create: true});
            const w = await fh.createWritable();
            try {
                await w.write(f.data);
                await w.close();
            } catch (err) {
                await w.abort().catch(() => undefined);
                throw err;
            }
            const written = await fh.getFile();
            if (written.size !== f.data.byteLength) {
                throw new Error(`寫入後檔案大小不符（預期 ${f.data.byteLength}，實際 ${written.size}）`);
            }
            onWritten?.(f.filename, written);
        } catch (error) {
            failed.push({filename: f.filename, error});
        }
    }
    return failed;
};

/** 把寫入錯誤翻成白話原因 */
export const describeWriteError = (err: unknown): string => {
    const name = err instanceof DOMException ? err.name : '';
    if (name === 'NoModificationAllowedError' || name === 'InvalidStateError' || name === 'NotReadableError') {
        return '檔案可能被 Excel 開著，請關閉後再按';
    }
    if (name === 'NotAllowedError' || name === 'SecurityError') return '沒有寫入權限，請在 Chrome 上方允許後再按';
    if (name === 'NotFoundError') return '找不到資料夾，可能被搬走或刪除';
    return err instanceof Error ? err.message : String(err);
};

/** 依序觸發瀏覽器下載（不支援資料夾存檔時用），間隔避免被擋 */
export const downloadSequentially = async (
    files: SaveFile[],
    mime: string,
    save: (blob: Blob, filename: string) => void,
    gapMs = 500,
): Promise<void> => {
    for (let i = 0; i < files.length; i++) {
        save(new Blob([files[i].data], {type: mime}), files[i].filename);
        if (i < files.length - 1) await new Promise((r) => setTimeout(r, gapMs));
    }
};

/** 資料夾被搬走/刪除/權限失效這類「handle 已不能用」的錯誤 */
export const isHandleGone = (err: unknown): boolean =>
    err instanceof DOMException && (err.name === 'NotFoundError' || err.name === 'NotAllowedError' || err.name === 'InvalidStateError');
