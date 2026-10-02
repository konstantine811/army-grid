import { api } from "../../api";
import {
  CacheKeys,
  peekDataCache,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import {
  rosterRowsFromPersonnelLatest,
  type PersonnelDataset,
} from "../../data/personnelDataset";
import { loadSharedRosterLatest } from "../../data/sharedAppData";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { loadAvailablePersonPhotoIds } from "./personAttachments";
import {
  buildStaffScopePreview,
  datasetCarriesSpreadsheetPersonIds,
  isPersonnelFromArchive,
  isPersonnelInStaffRoster,
} from "./personnelRosterMerge";

const readCachedPersonnelDataset = async () => {
  const peeked = peekDataCache<PersonnelDataset>(CacheKeys.personnelDataset);
  if (peeked?.rows?.length) return peeked;
  return readDataCache<PersonnelDataset>(CacheKeys.personnelDataset);
};

const datasetWithSpreadsheetIds = async (
  cached: PersonnelDataset | null | undefined,
  signal?: AbortSignal,
) => {
  if (datasetCarriesSpreadsheetPersonIds(cached?.rows)) return cached ?? null;
  const serverDataset = await api.getPersonnelDataset({ signal }).catch(() => null);
  if (!serverDataset || !datasetCarriesSpreadsheetPersonIds(serverDataset.rows)) {
    return cached ?? null;
  }
  void writeDataCache(CacheKeys.personnelDataset, serverDataset);
  return serverDataset;
};

/** Штатка + картки ООС з числовими id, у порядку рядків штатки. */
export const loadStaffScopePreview = async (options?: {
  force?: boolean;
  signal?: AbortSignal;
}): Promise<
  | { aborted: true; preview: null }
  | { aborted: false; preview: DbPreviewState | null }
> => {
  const [roster, cachedDataset] = await Promise.all([
    loadSharedRosterLatest({
      force: options?.force,
      signal: options?.signal,
    }),
    readCachedPersonnelDataset(),
    loadAvailablePersonPhotoIds(),
  ]);
  if (options?.signal?.aborted) return { aborted: true, preview: null };

  const dataset = await datasetWithSpreadsheetIds(cachedDataset, options?.signal);
  if (options?.signal?.aborted) return { aborted: true, preview: null };

  const cachedPreview = dataset?.rows?.length
    ? {
        rows: dataset.rows,
        columns: dataset.columns ?? [],
        sheet: dataset.sheet,
      }
    : null;
  const preview = roster
    ? buildStaffScopePreview(
        rosterRowsFromPersonnelLatest(roster),
        roster.sheet,
        cachedPreview,
      )
    : null;
  return { aborted: false, preview };
};

export const staffQuestionnaireRows = (rows: EjournalPreviewRow[]) =>
  rows.filter(
    (row) => isPersonnelInStaffRoster(row) && !isPersonnelFromArchive(row),
  );
