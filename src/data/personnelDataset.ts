import type {
  BackendPersonnelRosterLatest,
} from "../api";
import type {
  DbPreviewState,
} from "../pages/ejournal/ejournalTypes";
import {
  applyPersonnelMergeDelta,
  buildRosterOnlyPreviewState,
} from "../pages/personnel/personnelRosterMerge";
import {
  buildPersonnelDatasetSync,
  buildPersonnelDatasetVersion,
  buildPersonnelDatasetVersionFromRosterVersion,
  dedupePersonnelDatasetRows,
  personnelDatasetFingerprint,
  rosterRowsFromPersonnelLatest,
  sortPersonnelRowsByRosterOrder,
  type PersonnelDataset,
  type PersonnelDatasetVersion,
} from "./personnelDatasetCore";
import { runHeavyJob } from "../workers/runHeavyJob";
import { loadSharedEjournalImports, loadSharedRosterLatest } from "./sharedAppData";
import {
  CacheKeys,
  deleteDataCache,
  evictDataCacheMemory,
  readDataCache,
  writeDataCache,
} from "./idbDataCache";
import {
  findEjournalPersonnelSheet,
  loadAllEjournalSheetRows,
} from "../pages/personnel/personnelUtils";
import { api } from "../api";

export type { PersonnelDataset, PersonnelDatasetVersion } from "./personnelDatasetCore";
export {
  buildPersonnelDatasetSync,
  buildPersonnelDatasetVersion,
  buildPersonnelDatasetVersionFromRosterVersion,
  dedupePersonnelDatasetRows,
  personnelDatasetFingerprint,
  rosterRowsFromPersonnelLatest,
  sortPersonnelRowsByRosterOrder,
} from "./personnelDatasetCore";

type LoadPersonnelDatasetOptions = {
  force?: boolean;
  signal?: AbortSignal;
  onCached?: (dataset: PersonnelDataset) => void | Promise<void>;
};

let inFlight: Promise<PersonnelDataset> | null = null;
let inFlightSignal: AbortSignal | undefined;
let removedLegacyDatasetCaches = false;

const buildDataset = async (
  preview: DbPreviewState | null,
  roster: BackendPersonnelRosterLatest | null,
  version: PersonnelDatasetVersion,
): Promise<PersonnelDataset> => {
  const synced = buildPersonnelDatasetSync(preview, roster, version);
  if (!preview?.rows.length || !roster?.rows?.length) return synced;

  const rosterRows = rosterRowsFromPersonnelLatest(roster);
  const fallback = buildRosterOnlyPreviewState(rosterRows, roster?.sheet);
  const base = preview ?? fallback;
  const mergedRows = applyPersonnelMergeDelta(
    base.rows,
    await runHeavyJob({
      type: "mergePersonnel",
      preview: base,
      rosterRows,
    }),
  );
  const rows = dedupePersonnelDatasetRows(
    sortPersonnelRowsByRosterOrder(mergedRows),
  );
  return { ...synced, rows, total: rows.length, mergedAt: Date.now() };
};

const buildRosterBootstrap = (
  roster: BackendPersonnelRosterLatest,
  version: PersonnelDatasetVersion,
): PersonnelDataset | null => {
  const dataset = buildPersonnelDatasetSync(null, roster, version);
  if (!dataset.rows.length) return null;
  return { ...dataset, complete: false };
};

export const personnelDatasetToPreview = (
  dataset: PersonnelDataset,
): DbPreviewState | null =>
  dataset.sheet
    ? {
        sheet: dataset.sheet,
        columns: dataset.columns,
        rows: dataset.rows,
        total: dataset.total,
        offset: 0,
        limit: dataset.rows.length,
      }
    : null;

export const loadPersonnelDataset = async (
  options: LoadPersonnelDatasetOptions = {},
): Promise<PersonnelDataset> => {
  if (!removedLegacyDatasetCaches) {
    removedLegacyDatasetCaches = true;
    void deleteDataCache("personnel:dataset:v1");
    void deleteDataCache("personnel:dataset:memory:v3");
  }
  if (inFlight && !options.force && !inFlightSignal?.aborted) return inFlight;
  if (inFlightSignal?.aborted) {
    inFlight = null;
    inFlightSignal = undefined;
  }

  const tryServerDataset = async (fingerprint: string) => {
    if (options.force) return null;
    try {
      const serverDataset = await api.getPersonnelDataset({
        fingerprint,
        signal: options.signal,
      });
      if (
        serverDataset?.fingerprint === fingerprint &&
        Array.isArray(serverDataset.rows) &&
        serverDataset.rows.length
      ) {
        await options.onCached?.(serverDataset);
        await writeDataCache(CacheKeys.personnelDataset, serverDataset);
        evictDataCacheMemory("ejournal:sheet-rows:");
        return serverDataset;
      }
    } catch {
      /* server snapshot missing or stale — fall back to client merge */
    }
    return null;
  };

  const run = async () => {
    const cached = await readDataCache<PersonnelDataset>(
      CacheKeys.personnelDataset,
    );
    if (cached) await options.onCached?.(cached);

    const [imports, rosterVersion] = await Promise.all([
      loadSharedEjournalImports({
        force: options.force,
        signal: options.signal,
      }),
      api
        .getLatestPersonnelRosterVersion({ signal: options.signal })
        .catch(() => null),
    ]);
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    const sheet = findEjournalPersonnelSheet(imports);
    const versionFromMeta = buildPersonnelDatasetVersionFromRosterVersion(
      sheet,
      rosterVersion,
    );
    const fingerprintFromMeta = personnelDatasetFingerprint(versionFromMeta);
    if (!options.force && cached?.fingerprint === fingerprintFromMeta) {
      return cached;
    }
    const serverHit = await tryServerDataset(fingerprintFromMeta);
    if (serverHit) return serverHit;

    const roster = await loadSharedRosterLatest({
      force: options.force,
      signal: options.signal,
    });
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    const version = buildPersonnelDatasetVersion(sheet, roster);
    const fingerprint = personnelDatasetFingerprint(version);
    if (!options.force && cached?.fingerprint === fingerprint) return cached;
    if (fingerprint !== fingerprintFromMeta) {
      const retryServerHit = await tryServerDataset(fingerprint);
      if (retryServerHit) return retryServerHit;
    }
    if (!cached && roster) {
      const bootstrap = buildRosterBootstrap(roster, version);
      if (bootstrap) await options.onCached?.(bootstrap);
    }

    const preview = sheet
      ? await loadAllEjournalSheetRows(sheet, {
          force: options.force,
          signal: options.signal,
        })
      : null;
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    const dataset = await buildDataset(preview, roster, version);
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");
    await writeDataCache(CacheKeys.personnelDataset, dataset);
    // The merged dataset supersedes the wide raw OOS snapshot in RAM. Its
    // IndexedDB copy remains available for a cold reload.
    evictDataCacheMemory("ejournal:sheet-rows:");
    return dataset;
  };

  const promise = run();
  if (!options.force) {
    inFlight = promise;
    inFlightSignal = options.signal;
  }
  try {
    return await promise;
  } finally {
    if (inFlight === promise) {
      inFlight = null;
      inFlightSignal = undefined;
    }
  }
};
