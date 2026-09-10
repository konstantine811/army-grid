import type {
  BackendPersonnelBootstrap,
  BackendPersonnelRosterLatest,
  BackendPersonnelVersionProbe,
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
import { loadPersonnelVersionProbe, loadPersonnelBootstrapMeta } from "./personnelVersion";

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
  /** Last-resort client worker merge when server rebuild is unavailable. */
  clientMergeFallback?: boolean;
  /** Skip redundant `/personnel/version` when caller already fetched it. */
  versionProbe?: BackendPersonnelVersionProbe | null;
  /** Skip redundant `/personnel/bootstrap` when caller already fetched it. */
  bootstrapMeta?: BackendPersonnelBootstrap | null;
};

let inFlight: Promise<PersonnelDataset> | null = null;
let inFlightSignal: AbortSignal | undefined;
let removedLegacyDatasetCaches = false;

const acceptServerDataset = (
  dataset: PersonnelDataset | null | undefined,
  fingerprint: string,
) =>
  dataset?.fingerprint === fingerprint &&
  Array.isArray(dataset.rows) &&
  dataset.rows.length
    ? dataset
    : null;

const buildClientMergedDataset = async (
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

  const persistDataset = async (dataset: PersonnelDataset) => {
    await options.onCached?.(dataset);
    await writeDataCache(CacheKeys.personnelDataset, dataset);
    evictDataCacheMemory("ejournal:sheet-rows:");
    return dataset;
  };

  const tryServerDataset = async (fingerprint: string) => {
    if (options.force) return null;
    try {
      const serverDataset = await api.getPersonnelDataset({
        fingerprint,
        signal: options.signal,
      });
      const accepted = acceptServerDataset(serverDataset, fingerprint);
      if (accepted) return persistDataset(accepted);
    } catch {
      /* server snapshot missing or stale */
    }
    return null;
  };

  const tryRebuildServerDataset = async (fingerprint: string) => {
    if (options.force) return null;
    try {
      const rebuilt = await api.rebuildPersonnelDataset({
        signal: options.signal,
      });
      const accepted = acceptServerDataset(rebuilt, fingerprint);
      if (accepted) return persistDataset(accepted);
    } catch {
      /* backend rebuild unavailable */
    }
    return null;
  };

  const run = async () => {
    const cached = await readDataCache<PersonnelDataset>(
      CacheKeys.personnelDataset,
    );
    if (cached) await options.onCached?.(cached);

    const versionProbe =
      options.versionProbe !== undefined
        ? options.versionProbe
        : await loadPersonnelVersionProbe({
            force: options.force,
            signal: options.signal,
          });
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    let fingerprintFromMeta = versionProbe?.fingerprint?.trim() ?? "";

    if (!fingerprintFromMeta) {
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
      fingerprintFromMeta = personnelDatasetFingerprint(versionFromMeta);
    }

    if (!options.force && cached?.fingerprint === fingerprintFromMeta) {
      return cached;
    }

    const bootstrapMeta =
      options.bootstrapMeta !== undefined
        ? options.bootstrapMeta
        : await loadPersonnelBootstrapMeta({
            force: options.force,
            signal: options.signal,
          }).catch(() => null);

    if (
      !options.force &&
      cached &&
      versionProbe?.snapshot?.matchesLive &&
      cached.fingerprint === versionProbe.fingerprint
    ) {
      return cached;
    }

    const serverHit =
      bootstrapMeta?.dataset.available === false
        ? null
        : (await tryServerDataset(fingerprintFromMeta)) ??
          (options.force || !cached
            ? await tryRebuildServerDataset(fingerprintFromMeta)
            : null);
    if (serverHit) return serverHit;
    if (!options.force && cached) return cached;

    const imports = await loadSharedEjournalImports({
      force: options.force,
      signal: options.signal,
    });
    const roster = await loadSharedRosterLatest({
      force: options.force,
      signal: options.signal,
    });
    if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

    const sheet = findEjournalPersonnelSheet(imports);
    const version = buildPersonnelDatasetVersion(sheet, roster);
    const fingerprint = personnelDatasetFingerprint(version);
    if (!options.force && cached?.fingerprint === fingerprint) return cached;

    if (fingerprint !== fingerprintFromMeta) {
      const retryHit =
        (await tryServerDataset(fingerprint)) ??
        (await tryRebuildServerDataset(fingerprint));
      if (retryHit) return retryHit;
    }

    let bootstrap: PersonnelDataset | null = null;
    if (roster) {
      bootstrap = buildRosterBootstrap(roster, version);
      if (bootstrap && !cached) await options.onCached?.(bootstrap);
    }

    if (options.clientMergeFallback !== false) {
      const preview = sheet
        ? await loadAllEjournalSheetRows(sheet, {
            force: options.force,
            signal: options.signal,
          })
        : null;
      if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");

      const dataset = await buildClientMergedDataset(preview, roster, version);
      if (options.signal?.aborted) throw new DOMException("Aborted", "AbortError");
      return persistDataset(dataset);
    }

    if (bootstrap) return bootstrap;
    if (cached) return cached;
    throw new Error(
      "Не вдалося завантажити особовий склад з сервера. Запустіть backend або увімкніть client merge.",
    );
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

/** Prefer merged roster rows; fall back to full personnel rows. */
export const rosterRowsFromDataset = (dataset: PersonnelDataset) =>
  dataset.rosterRows.length ? dataset.rosterRows : dataset.rows;
