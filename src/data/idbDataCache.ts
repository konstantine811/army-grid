/**
 * Shared memory + IndexedDB cache for heavy API payloads.
 * Pages read the same in-memory copy; network refresh is TTL-gated
 * and de-duplicated across concurrent callers.
 */

const DB_NAME = "army-grid-data-cache";
const DB_VERSION = 1;
const STORE_NAME = "entries";

/** Skip a network refresh if the shared copy is newer than this. */
export const SHARED_DATA_TTL_MS = 5 * 60 * 1000;

type CacheEntry<T = unknown> = {
  key: string;
  value: T;
  savedAt: number;
};

const memoryCache = new Map<string, CacheEntry>();
const inFlight = new Map<string, Promise<unknown>>();
const cacheGenerations = new Map<string, number>();
const listeners = new Map<string, Set<() => void>>();

const getCacheGeneration = (key: string) => cacheGenerations.get(key) ?? 0;

const invalidateInFlightKey = (key: string) => {
  cacheGenerations.set(key, getCacheGeneration(key) + 1);
  inFlight.delete(key);
};

const notifyDataCache = (key: string) => {
  listeners.get(key)?.forEach((listener) => listener());
};

export const subscribeDataCache = (key: string, listener: () => void) => {
  const set = listeners.get(key) ?? new Set<() => void>();
  set.add(listener);
  listeners.set(key, set);
  return () => {
    set.delete(listener);
    if (set.size === 0) listeners.delete(key);
  };
};

export const peekDataCache = <T>(key: string): T | null => {
  const entry = memoryCache.get(key);
  return entry ? (entry.value as T) : null;
};

export const getDataCacheAgeMs = (key: string): number | null => {
  const entry = memoryCache.get(key);
  return entry ? Date.now() - entry.savedAt : null;
};

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
  anketaCreatedPersonnel: "personnel:anketa-created:v1",
  staffSheetImport: "anketa:staff-sheet-import",
  staffSheetVkIndex: "anketa:staff-sheet-vk-index",
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
    notifyDataCache(key);
    return entry.value;
  } catch {
    return null;
  }
};

export const writeDataCache = async <T>(key: string, value: T): Promise<void> => {
  const entry: CacheEntry<T> = { key, value, savedAt: Date.now() };
  memoryCache.set(key, entry);
  notifyDataCache(key);
  try {
    await withStore("readwrite", (store) => store.put(entry));
  } catch {
    /* ignore quota / private mode */
  }
};

export const touchDataCache = (key: string) => {
  const entry = memoryCache.get(key);
  if (!entry) return;
  entry.savedAt = Date.now();
};

export const deleteDataCache = async (key: string): Promise<void> => {
  memoryCache.delete(key);
  notifyDataCache(key);
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

  const notified = new Set<string>();
  const activeKeys = new Set([...memoryCache.keys(), ...inFlight.keys()]);
  for (const token of keysOrPrefixes) {
    memoryCache.delete(token);
    notified.add(token);
    invalidateInFlightKey(token);
    for (const key of activeKeys) {
      if (key !== token && key.startsWith(token)) {
        memoryCache.delete(key);
        notified.add(key);
        invalidateInFlightKey(key);
      }
    }
  }
  notified.forEach((key) => notifyDataCache(key));

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
  /** Fresh shared copy younger than this skips the network. */
  ttlMs?: number;
  /** Ignore TTL and refetch (кнопка «Оновити»). */
  force?: boolean;
}): Promise<T> => {
  const pending = inFlight.get(options.key);
  if (pending) return pending as Promise<T>;

  const run = async () => {
    const generation = getCacheGeneration(options.key);
    const cached = await readDataCache<T>(options.key);
    if (cached != null) {
      await options.onCached?.(cached);
    }

    const ttl = options.ttlMs ?? SHARED_DATA_TTL_MS;
    const age = getDataCacheAgeMs(options.key);
    if (
      cached != null &&
      !options.force &&
      age != null &&
      age < ttl
    ) {
      return cached;
    }

    try {
      const fresh = await options.fetcher();
      if (generation !== getCacheGeneration(options.key)) {
        return fetchWithCache({ ...options, force: true });
      }
      if (fresh == null && cached != null) {
        return cached;
      }
      const changed =
        cached == null ||
        (options.isChanged ? options.isChanged(cached, fresh) : true);
      if (changed) {
        await writeDataCache(options.key, fresh);
      } else {
        touchDataCache(options.key);
      }
      return fresh;
    } catch (error) {
      if (cached != null) return cached;
      throw error;
    }
  };

  const promise = run().finally(() => {
    if (inFlight.get(options.key) === promise) inFlight.delete(options.key);
  });
  inFlight.set(options.key, promise);
  return promise as Promise<T>;
};

/** Test helper: drop in-memory copies between cases. */
export const resetDataCacheMemory = () => {
  memoryCache.clear();
  inFlight.clear();
  cacheGenerations.clear();
};

export { payloadChanged as jsonChanged } from "./payloadFingerprint";
