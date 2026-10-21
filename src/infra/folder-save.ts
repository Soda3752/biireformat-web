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

/** 逐檔寫入（同名覆蓋）；close() 成功後再讀回大小比對，一致才算成功。回傳失敗清單，不因單檔失敗中斷其他檔 */
export const writeFiles = async (
    dir: FileSystemDirectoryHandle,
    files: SaveFile[],
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
