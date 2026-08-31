/**
 * IndexedDB cache for heavy API payloads.
 * Pattern: show cached data immediately, then refresh from network.
 */

const DB_NAME = "army-grid-data-cache";
const DB_VERSION = 1;
const STORE_NAME = "entries";

type CacheEntry<T = unknown> = {
  key: string;
  value: T;
  savedAt: number;
};

const memoryCache = new Map<string, CacheEntry>();

const openDb = () =>
  new Promise<IDBDatabase>((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "key" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("IndexedDB open failed"));
  });

const withStore = async <T>(
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest<T>,
) => {
  const db = await openDb();
  try {
    return await new Promise<T>((resolve, reject) => {
      const tx = db.transaction(STORE_NAME, mode);
      const request = run(tx.objectStore(STORE_NAME));
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error("IndexedDB request failed"));
    });
  } finally {
    db.close();
  }
};

export const CacheKeys = {
  ejournalImports: "ejournal:imports",
  sheetRows: (sheetId: string, stamp: string) =>
    `ejournal:sheet-rows:${sheetId}:${stamp}`,
  rosterLatest: "personnel:roster:latest",
  staffSheetImport: "anketa:staff-sheet-import",
  overview: "personnel:overview",
  documentsAll: "personnel:documents:all",
  questionnairesMeta: "personnel:questionnaires:meta",
} as const;

export const readDataCache = async <T>(key: string): Promise<T | null> => {
  const memory = memoryCache.get(key);
  if (memory) return memory.value as T;

  try {
    const entry = await withStore<CacheEntry<T> | undefined>("readonly", (store) =>
      store.get(key),
    );
    if (!entry || entry.value === undefined) return null;
    memoryCache.set(key, entry);
    return entry.value;
  } catch {
    return null;
  }
};

export const writeDataCache = async <T>(key: string, value: T): Promise<void> => {
  const entry: CacheEntry<T> = { key, value, savedAt: Date.now() };
  memoryCache.set(key, entry);
  try {
    await withStore("readwrite", (store) => store.put(entry));
  } catch {
    /* ignore quota / private mode */
  }
};

export const deleteDataCache = async (key: string): Promise<void> => {
  memoryCache.delete(key);
  try {
    await withStore("readwrite", (store) => store.delete(key));
  } catch {
    /* ignore */
  }
};

/** Delete exact keys and any key that starts with a given prefix. */
export const invalidateDataCache = async (
  ...keysOrPrefixes: string[]
): Promise<void> => {
  if (keysOrPrefixes.length === 0) return;

  for (const token of keysOrPrefixes) {
    memoryCache.delete(token);
    for (const key of [...memoryCache.keys()]) {
      if (key.startsWith(token)) memoryCache.delete(key);
    }
  }

  try {
    const db = await openDb();
    try {
      const allKeys = await new Promise<IDBValidKey[]>((resolve, reject) => {
        const request = db
          .transaction(STORE_NAME, "readonly")
          .objectStore(STORE_NAME)
          .getAllKeys();
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });

      const toDelete = allKeys.filter((raw) => {
        const key = String(raw);
        return keysOrPrefixes.some(
          (token) => key === token || key.startsWith(token),
        );
      });

      if (toDelete.length === 0) return;

      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        const store = tx.objectStore(STORE_NAME);
        for (const key of toDelete) store.delete(key);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
    } finally {
      db.close();
    }
  } catch {
    /* ignore */
  }
};

export const invalidatePersonnelCaches = () =>
  invalidateDataCache(
    CacheKeys.ejournalImports,
    "ejournal:sheet-rows:",
    CacheKeys.rosterLatest,
    CacheKeys.overview,
    CacheKeys.documentsAll,
    CacheKeys.questionnairesMeta,
  );

/**
 * Cache-first fetch: emit cached payload ASAP, then refresh from network.
 * Returns the freshest available data (network win, cache fallback on error).
 */
export const fetchWithCache = async <T>(options: {
  key: string;
  fetcher: () => Promise<T>;
  onCached?: (data: T) => void | Promise<void>;
  isChanged?: (cached: T, fresh: T) => boolean;
}): Promise<T> => {
  const cached = await readDataCache<T>(options.key);
  if (cached != null) {
    await options.onCached?.(cached);
  }

  try {
    const fresh = await options.fetcher();
    const changed =
      cached == null ||
      (options.isChanged ? options.isChanged(cached, fresh) : true);
    if (changed) {
      await writeDataCache(options.key, fresh);
    }
    return fresh;
  } catch (error) {
    if (cached != null) return cached;
    throw error;
  }
};

/** Stable JSON compare for deciding whether to rewrite UI / disk. */
export const jsonChanged = (a: unknown, b: unknown) => {
  try {
    return JSON.stringify(a) !== JSON.stringify(b);
  } catch {
    return true;
  }
};
