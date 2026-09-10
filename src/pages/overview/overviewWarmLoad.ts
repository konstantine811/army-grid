import type {
  BackendPersonnelBootstrap,
  BackendPersonnelOverview,
  BackendPersonnelVersionProbe,
} from "../../api";
import type { PersonnelDataset } from "../../data/personnelDatasetCore";
import {
  overviewStaffCacheKey,
  overviewStaffMetaCacheKey,
  peekDataCache,
  writeDataCache,
  type OverviewStaffCacheMeta,
} from "../../data/idbDataCache";

export type PersonnelBootstrapInclude =
  | "dataset"
  | "staff"
  | "assets"
  | "merge";

const hasOverviewRows = (overview: BackendPersonnelOverview | null | undefined) =>
  Boolean(overview?.rows?.length);

/** Warm IDB hit on Overview (staff tab) — paint instantly without blocking network. */
export const canUseOverviewWarmCacheOnly = (options: {
  force?: boolean;
  datasetOnly?: boolean;
  cachedDataset: PersonnelDataset | null | undefined;
}) =>
  !options.force &&
  Boolean(options.datasetOnly) &&
  Boolean(options.cachedDataset?.rows?.length);

/** Warm IDB dataset is still live — skip full loadPersonnelDataset on Overview. */
export const canSkipOverviewDatasetReload = (options: {
  force?: boolean;
  datasetOnly?: boolean;
  cachedDataset: PersonnelDataset | null | undefined;
  versionProbe: BackendPersonnelVersionProbe | null | undefined;
}) => {
  if (options.force || !options.datasetOnly) return false;
  const cached = options.cachedDataset;
  if (!cached?.rows?.length) return false;
  const probe = options.versionProbe;
  if (!probe?.snapshot?.matchesLive) return false;
  return cached.fingerprint === probe.fingerprint;
};

/** IDB staff snapshot + meta — enough to paint without bootstrap network. */
export const hasOverviewStaffIdbSnapshot = (fingerprint: string): boolean => {
  const staff = peekDataCache<BackendPersonnelOverview>(
    overviewStaffCacheKey(fingerprint),
  );
  if (!hasOverviewRows(staff)) return false;
  const meta = peekDataCache<OverviewStaffCacheMeta>(
    overviewStaffMetaCacheKey(fingerprint),
  );
  return Boolean(meta);
};

export const isOverviewStaffCacheFresh = (
  fingerprint: string,
  bootstrap: BackendPersonnelBootstrap | null | undefined,
): boolean => {
  if (!bootstrap || bootstrap.fingerprint !== fingerprint) return false;
  if (!bootstrap.staff.available || !bootstrap.staff.updatedAt) return false;
  const staff = peekDataCache<BackendPersonnelOverview>(
    overviewStaffCacheKey(fingerprint),
  );
  if (!hasOverviewRows(staff)) return false;
  const meta = peekDataCache<OverviewStaffCacheMeta>(
    overviewStaffMetaCacheKey(fingerprint),
  );
  if (!meta) return false;
  return (
    meta.serverUpdatedAt === bootstrap.staff.updatedAt &&
    meta.rowCount === staff!.rows.length
  );
};

export const persistOverviewStaffCache = async (
  fingerprint: string,
  staff: BackendPersonnelOverview,
  bootstrap: BackendPersonnelBootstrap | null | undefined,
) => {
  await writeDataCache(overviewStaffCacheKey(fingerprint), staff);
  await writeDataCache(overviewStaffMetaCacheKey(fingerprint), {
    serverUpdatedAt: bootstrap?.staff.updatedAt ?? null,
    rowCount: staff.rows.length,
  } satisfies OverviewStaffCacheMeta);
};

export const readBootstrapStaffPayload = (
  bootstrap: BackendPersonnelBootstrap | null | undefined,
  fingerprint: string,
) => {
  const staff = bootstrap?.payloads?.staff;
  if (!hasOverviewRows(staff)) return null;
  if (bootstrap?.fingerprint !== fingerprint) return null;
  return staff;
};

export const readBootstrapDatasetPayload = (
  bootstrap: BackendPersonnelBootstrap | null | undefined,
) => {
  const dataset = bootstrap?.payloads?.dataset;
  if (!dataset?.rows?.length) return null;
  if (dataset.fingerprint !== bootstrap?.fingerprint) return null;
  return dataset;
};
