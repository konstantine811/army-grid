import { useStore } from "zustand";
import { api } from "../api";
import type {
  BackendEjournalImport,
  BackendPersonDocument,
  BackendPersonnelOverview,
  BackendPersonnelRosterLatest,
  BackendPersonnelRosterVersion,
} from "../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
} from "./idbDataCache";
import { sharedDataStore } from "./sharedDataStore";
import type { PersonnelDataset } from "./personnelDataset";

export { SHARED_DATA_TTL_MS } from "./idbDataCache";

type SharedLoadOptions = {
  force?: boolean;
  signal?: AbortSignal;
};

export const ROSTER_VERSION_CHECK_TTL_MS = 30_000;
let rosterVersionCheckedAt = 0;

export const shouldCheckRosterVersion = (
  checkedAt: number,
  now = Date.now(),
) => now - checkedAt >= ROSTER_VERSION_CHECK_TTL_MS;

const fetchSharedRosterLatest = (options: SharedLoadOptions) =>
  fetchWithCache({
    key: CacheKeys.rosterLatest,
    signal: options.signal,
    fetcher: () => api.getLatestPersonnelRoster({ signal: options.signal }),
    isChanged: jsonChanged,
    force: options.force,
  });

export const isRosterVersionCurrent = (
  cached: BackendPersonnelRosterLatest,
  version: BackendPersonnelRosterVersion | null,
) =>
  version?.importId === cached.importId &&
  version.rowCount === (cached.sheet?.rowCount ?? cached.rows.length) &&
  String(version.sheetUpdatedAt ?? "") ===
    String(cached.sheet?.updatedAt ?? "");

export const loadSharedRosterLatest = async (
  options: SharedLoadOptions = {},
) => {
  if (options.force) return fetchSharedRosterLatest(options);
  const cached = await readDataCache<BackendPersonnelRosterLatest | null>(
    CacheKeys.rosterLatest,
  );
  if (!cached) return fetchSharedRosterLatest(options);
  if (!shouldCheckRosterVersion(rosterVersionCheckedAt)) return cached;

  try {
    const version = await api.getLatestPersonnelRosterVersion({
      signal: options.signal,
    });
    rosterVersionCheckedAt = Date.now();
    const sameVersion = isRosterVersionCurrent(cached, version);
    if (sameVersion) return cached;
    return fetchSharedRosterLatest({ ...options, force: true });
  } catch (error) {
    if (options.signal?.aborted) throw error;
    return cached;
  }
};

export const loadSharedDocumentsAll = (options: SharedLoadOptions = {}) =>
  fetchWithCache({
    key: CacheKeys.documentsAll,
    signal: options.signal,
    fetcher: () => api.listAllPersonDocuments({ signal: options.signal }),
    isChanged: jsonChanged,
    force: options.force,
  });

export const loadSharedEjournalImports = (options: SharedLoadOptions = {}) =>
  fetchWithCache({
    key: CacheKeys.ejournalImports,
    signal: options.signal,
    fetcher: () => api.listEjournalImports({ signal: options.signal }),
    isChanged: jsonChanged,
    force: options.force,
  });

export const useSharedCacheValue = <T>(key: string) => {
  const entry = useStore(sharedDataStore, (state) => state.entries.get(key));
  return (entry?.value as T | undefined) ?? null;
};

export const useSharedRosterLatest = () =>
  useSharedCacheValue<BackendPersonnelRosterLatest | null>(CacheKeys.rosterLatest);

export const useSharedOverview = () =>
  useSharedCacheValue<BackendPersonnelOverview>(CacheKeys.overview);

export const useSharedPersonnelDataset = () =>
  useSharedCacheValue<PersonnelDataset>(CacheKeys.personnelDataset);

export const useSharedDocumentsAll = () =>
  useSharedCacheValue<BackendPersonDocument[]>(CacheKeys.documentsAll);

export const useSharedEjournalImports = () =>
  useSharedCacheValue<BackendEjournalImport[]>(CacheKeys.ejournalImports);
