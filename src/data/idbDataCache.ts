/**
 * Shared memory + IndexedDB cache for heavy API payloads.
 * Pages read the same in-memory copy; network refresh is TTL-gated
 * and de-duplicated across concurrent callers.
 */
import {
  clearSharedDataStore,
  deleteSharedDataEntry,
  getSharedDataEntry,
  getSharedDataKeys,
  setSharedDataEntry,
} from "./sharedDataStore";

const DB_NAME = "army-grid-data-cache";
const DB_VERSION = 1;
const STORE_NAME = "entries";
export const DATA_CACHE_FORMAT_VERSION = 1;
export const DATA_CACHE_MAX_AGE_MS = 7 * 24 * 60 * 60 * 1000;
export const DATA_CACHE_MAX_SHEET_SNAPSHOTS = 3;

/** Skip a network refresh if the shared copy is newer than this. */
export const SHARED_DATA_TTL_MS = 5 * 60 * 1000;

type CacheEntry<T = unknown> = {
  key: string;
  value: T;
  savedAt: number;
  formatVersion?: number;
};

type InFlightCacheRequest = {
  promise: Promise<unknown>;
  signal?: AbortSignal;
};

const inFlight = new Map<string, InFlightCacheRequest>();
const cacheGenerations = new Map<string, number>();
const listeners = new Map<string, Set<() => void>>();

const isCacheEntryExpired = (entry: CacheEntry, now = Date.now()) =>
  now - entry.savedAt > DATA_CACHE_MAX_AGE_MS ||
  (entry.formatVersion != null &&
    entry.formatVersion !== DATA_CACHE_FORMAT_VERSION);

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
  const entry = getSharedDataEntry(key) as CacheEntry | undefined;
  if (entry && isCacheEntryExpired(entry)) {
    deleteSharedDataEntry(key);
    return null;
  }
  return entry ? (entry.value as T) : null;
};

export const getDataCacheAgeMs = (key: string): number | null => {
  const entry = getSharedDataEntry(key);
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
  personnelDataset: "personnel:dataset:memory:v3",
  overview: "personnel:overview",
  documentsAll: "personnel:documents:meta:v4",
  questionnairesMeta: "personnel:questionnaires:meta",
} as const;

const isKnownCacheKey = (key: string) =>
  key === CacheKeys.ejournalImports ||
  key.startsWith("ejournal:sheet-rows:") ||
  key === CacheKeys.rosterLatest ||
  key === CacheKeys.anketaCreatedPersonnel ||
  key === CacheKeys.staffSheetImport ||
  key === CacheKeys.staffSheetVkIndex ||
  key === CacheKeys.personnelDataset ||
  key === CacheKeys.overview ||
  key === CacheKeys.documentsAll ||
  key === CacheKeys.questionnairesMeta;

export const planDataCacheCleanup = (
  entries: Array<Pick<CacheEntry, "key" | "savedAt" | "formatVersion">>,
  now = Date.now(),
): string[] => {
  const toDelete = new Set<string>();
  const sheetSnapshots: Array<Pick<CacheEntry, "key" | "savedAt">> = [];

  for (const entry of entries) {
    if (
      !isKnownCacheKey(entry.key) ||
      now - entry.savedAt > DATA_CACHE_MAX_AGE_MS ||
      (entry.formatVersion != null &&
        entry.formatVersion !== DATA_CACHE_FORMAT_VERSION)
    ) {
      toDelete.add(entry.key);
      continue;
    }
    if (entry.key.startsWith("ejournal:sheet-rows:")) {
      sheetSnapshots.push(entry);
    }
  }

  sheetSnapshots
    .filter((entry) => !toDelete.has(entry.key))
    .sort((left, right) => right.savedAt - left.savedAt)
    .slice(DATA_CACHE_MAX_SHEET_SNAPSHOTS)
    .forEach((entry) => toDelete.add(entry.key));

  return [...toDelete];
};

export const readDataCache = async <T>(key: string): Promise<T | null> => {
  const memory = getSharedDataEntry(key) as CacheEntry<T> | undefined;
  if (memory && !isCacheEntryExpired(memory)) return memory.value as T;
  if (memory) deleteSharedDataEntry(key);

  try {
    const entry = await withStore<CacheEntry<T> | undefined>("readonly", (store) =>
      store.get(key),
    );
    if (!entry || entry.value === undefined) return null;
    if (isCacheEntryExpired(entry)) {
      void deleteDataCache(key);
      return null;
    }
    setSharedDataEntry(entry);
    notifyDataCache(key);
    return entry.value;
  } catch {
    return null;
  }
};

export const writeDataCache = async <T>(key: string, value: T): Promise<void> => {
  const entry: CacheEntry<T> = {
    key,
    value,
    savedAt: Date.now(),
    formatVersion: DATA_CACHE_FORMAT_VERSION,
  };
  setSharedDataEntry(entry);
  notifyDataCache(key);
  try {
    await withStore("readwrite", (store) => store.put(entry));
  } catch {
    // A stale workbook snapshot can exhaust the browser quota. Clean first,
    // then make one best-effort retry for the current value.
    const removed = await cleanupDataCache();
    if (removed > 0) {
      try {
        await withStore("readwrite", (store) => store.put(entry));
      } catch {
        /* ignore private mode / payload larger than available quota */
      }
    }
  }
};

/** Store a derived payload only in Zustand, avoiding an IndexedDB structured clone. */
export const writeMemoryDataCache = <T>(key: string, value: T): void => {
  setSharedDataEntry({
    key,
    value,
    savedAt: Date.now(),
    formatVersion: DATA_CACHE_FORMAT_VERSION,
  });
  notifyDataCache(key);
};

/** Release heavy L1 values while preserving their persistent IndexedDB copies. */
export const evictDataCacheMemory = (...keysOrPrefixes: string[]): void => {
  for (const key of [...getSharedDataKeys()]) {
    if (
      keysOrPrefixes.some(
        (token) => key === token || key.startsWith(token),
      )
    ) {
      deleteSharedDataEntry(key);
    }
  }
};

export const cleanupDataCache = async (): Promise<number> => {
  try {
    const db = await openDb();
    try {
      const entries = await new Promise<CacheEntry[]>((resolve, reject) => {
        const request = db
          .transaction(STORE_NAME, "readonly")
          .objectStore(STORE_NAME)
          .getAll();
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
      });
      const keys = planDataCacheCleanup(entries);
      if (!keys.length) return 0;
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction(STORE_NAME, "readwrite");
        const store = tx.objectStore(STORE_NAME);
        keys.forEach((key) => store.delete(key));
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
      });
      keys.forEach((key) => {
        deleteSharedDataEntry(key);
        invalidateInFlightKey(key);
        notifyDataCache(key);
      });
      return keys.length;
    } finally {
      db.close();
    }
  } catch {
    return 0;
  }
};

let cleanupScheduled = false;

export const scheduleDataCacheCleanup = () => {
  if (cleanupScheduled || typeof window === "undefined") return;
  cleanupScheduled = true;
  const run = () => void cleanupDataCache();
  if ("requestIdleCallback" in window) {
    window.requestIdleCallback(run, { timeout: 5_000 });
  } else {
    globalThis.setTimeout(run, 1_000);
  }
};

export const touchDataCache = (key: string) => {
  const entry = getSharedDataEntry(key);
  if (!entry) return;
  setSharedDataEntry({ ...entry, savedAt: Date.now() });
};

export const deleteDataCache = async (key: string): Promise<void> => {
  deleteSharedDataEntry(key);
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
  const activeKeys = new Set([...getSharedDataKeys(), ...inFlight.keys()]);
  for (const token of keysOrPrefixes) {
    deleteSharedDataEntry(token);
    notified.add(token);
    invalidateInFlightKey(token);
    for (const key of activeKeys) {
      if (key !== token && key.startsWith(token)) {
        deleteSharedDataEntry(key);
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
    CacheKeys.personnelDataset,
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
  signal?: AbortSignal;
  onCached?: (data: T) => void | Promise<void>;
  isChanged?: (cached: T, fresh: T) => boolean;
  /** Fresh shared copy younger than this skips the network. */
  ttlMs?: number;
  /** Ignore TTL and refetch (кнопка «Оновити»). */
  force?: boolean;
}): Promise<T> => {
  const pending = inFlight.get(options.key);
  if (pending && !pending.signal?.aborted) {
    return pending.promise as Promise<T>;
  }
  if (pending) inFlight.delete(options.key);

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
    if (inFlight.get(options.key)?.promise === promise) {
      inFlight.delete(options.key);
    }
  });
  inFlight.set(options.key, { promise, signal: options.signal });
  return promise as Promise<T>;
};

/** Test helper: drop in-memory copies between cases. */
export const resetDataCacheMemory = () => {
  clearSharedDataStore();
  inFlight.clear();
  cacheGenerations.clear();
  cleanupScheduled = false;
};

export { payloadChanged as jsonChanged } from "./payloadFingerprint";

scheduleDataCacheCleanup();
