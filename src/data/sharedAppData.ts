import { useEffect, useState } from "react";
import { api } from "../api";
import type {
  BackendEjournalImport,
  BackendPersonDocument,
  BackendPersonnelOverview,
  BackendPersonnelRosterLatest,
} from "../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  peekDataCache,
  subscribeDataCache,
} from "./idbDataCache";

export { SHARED_DATA_TTL_MS } from "./idbDataCache";

type SharedLoadOptions = {
  force?: boolean;
};

export const loadSharedRosterLatest = (
  options: SharedLoadOptions = {},
) =>
  fetchWithCache({
    key: CacheKeys.rosterLatest,
    fetcher: () => api.getLatestPersonnelRoster(),
    isChanged: jsonChanged,
    force: options.force,
  });

export const loadSharedDocumentsAll = (options: SharedLoadOptions = {}) =>
  fetchWithCache({
    key: CacheKeys.documentsAll,
    fetcher: () => api.listAllPersonDocuments(),
    isChanged: jsonChanged,
    force: options.force,
  });

export const loadSharedEjournalImports = (options: SharedLoadOptions = {}) =>
  fetchWithCache({
    key: CacheKeys.ejournalImports,
    fetcher: () => api.listEjournalImports(),
    isChanged: jsonChanged,
    force: options.force,
  });

export const useSharedCacheValue = <T>(key: string) => {
  const [value, setValue] = useState<T | null>(() => peekDataCache<T>(key));

  useEffect(() => {
    setValue(peekDataCache<T>(key));
    return subscribeDataCache(key, () => {
      setValue(peekDataCache<T>(key));
    });
  }, [key]);

  return value;
};

export const useSharedRosterLatest = () =>
  useSharedCacheValue<BackendPersonnelRosterLatest | null>(CacheKeys.rosterLatest);

export const useSharedOverview = () =>
  useSharedCacheValue<BackendPersonnelOverview>(CacheKeys.overview);

export const useSharedDocumentsAll = () =>
  useSharedCacheValue<BackendPersonDocument[]>(CacheKeys.documentsAll);

export const useSharedEjournalImports = () =>
  useSharedCacheValue<BackendEjournalImport[]>(CacheKeys.ejournalImports);
