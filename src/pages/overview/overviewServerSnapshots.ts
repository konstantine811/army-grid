import {
  api,
  type BackendPersonnelBootstrap,
  type BackendPersonnelOverview,
  type OverviewPersonnelAssetsSnapshot,
} from "../../api";
import {
  overviewStaffCacheKey,
  peekDataCache,
  readDataCache,
} from "../../data/idbDataCache";
import { loadPersonnelBootstrapMeta } from "../../data/personnelVersion";
import {
  hasOverviewStaffIdbSnapshot,
  isOverviewStaffCacheFresh,
  persistOverviewStaffCache,
} from "./overviewWarmLoad";

const hasOverviewRows = (overview: BackendPersonnelOverview | null | undefined) =>
  Boolean(overview?.rows?.length);

const peekLocalStaffSnapshot = (fingerprint: string) =>
  peekDataCache<BackendPersonnelOverview>(overviewStaffCacheKey(fingerprint));

const readLocalStaffSnapshot = async (fingerprint: string) => {
  const peeked = peekLocalStaffSnapshot(fingerprint);
  if (hasOverviewRows(peeked)) return peeked;
  return readDataCache<BackendPersonnelOverview>(
    overviewStaffCacheKey(fingerprint),
  );
};

const resolveBootstrap = async (options: {
  fingerprint: string;
  signal?: AbortSignal;
  bootstrap?: BackendPersonnelBootstrap | null;
}) => {
  if (options.bootstrap !== undefined) {
    if (
      !options.bootstrap ||
      options.bootstrap.fingerprint !== options.fingerprint
    ) {
      return null;
    }
    return options.bootstrap;
  }
  if (hasOverviewStaffIdbSnapshot(options.fingerprint)) return null;
  const bootstrap = await loadPersonnelBootstrapMeta({
    signal: options.signal,
  });
  if (!bootstrap || bootstrap.fingerprint !== options.fingerprint) return null;
  return bootstrap;
};

type StaffSnapshotRequest = {
  fingerprint: string;
  signal?: AbortSignal;
  bootstrap?: BackendPersonnelBootstrap | null;
  force?: boolean;
  allowRebuild?: boolean;
};

const staffSnapshotRequests = new Map<string, Promise<BackendPersonnelOverview | null>>();
let staffRebuildInFlight: Promise<BackendPersonnelOverview | null> | null = null;

const requestStaffSnapshotKey = (options: StaffSnapshotRequest) =>
  [
    options.fingerprint,
    options.force ? "1" : "0",
    options.allowRebuild ? "1" : "0",
  ].join(":");

const rebuildOverviewStaffOnce = async (signal?: AbortSignal) => {
  if (staffRebuildInFlight) return staffRebuildInFlight;
  const promise = api.rebuildOverviewStaff({ signal }).catch(() => null);
  staffRebuildInFlight = promise;
  try {
    return await promise;
  } finally {
    if (staffRebuildInFlight === promise) staffRebuildInFlight = null;
  }
};

const fetchOverviewStaffSnapshotInner = async (
  options: StaffSnapshotRequest,
): Promise<BackendPersonnelOverview | null> => {
  if (!options.force) {
    const local =
      peekLocalStaffSnapshot(options.fingerprint) ??
      (await readLocalStaffSnapshot(options.fingerprint));
    if (
      hasOverviewRows(local) &&
      (isOverviewStaffCacheFresh(options.fingerprint, options.bootstrap) ||
        hasOverviewStaffIdbSnapshot(options.fingerprint))
    ) {
      return local;
    }
  }

  const bootstrap = await resolveBootstrap(options);

  if (!options.force && bootstrap && bootstrap.staff.available === false) {
    const local = await readLocalStaffSnapshot(options.fingerprint);
    if (hasOverviewRows(local)) return local;
  }

  if (!options.force && bootstrap && bootstrap.staff.available) {
    const local = await readLocalStaffSnapshot(options.fingerprint);
    if (
      isOverviewStaffCacheFresh(options.fingerprint, bootstrap) &&
      hasOverviewRows(local)
    ) {
      return local;
    }
  }

  if (!bootstrap || bootstrap.staff.available) {
    const cached = await api
      .getOverviewStaff({
        fingerprint: options.fingerprint,
        signal: options.signal,
      })
      .catch(() => null);
    if (hasOverviewRows(cached)) {
      void persistOverviewStaffCache(
        options.fingerprint,
        cached!,
        bootstrap,
      );
      return cached;
    }
  }

  if (options.allowRebuild !== true) {
    const local = await readLocalStaffSnapshot(options.fingerprint);
    return hasOverviewRows(local) ? local : null;
  }

  try {
    const rebuilt = await rebuildOverviewStaffOnce(options.signal);
    if (hasOverviewRows(rebuilt)) {
      void persistOverviewStaffCache(
        options.fingerprint,
        rebuilt!,
        bootstrap,
      );
      return rebuilt;
    }
  } catch {
    /* backend rebuild unavailable */
  }

  const local = await readLocalStaffSnapshot(options.fingerprint);
  return hasOverviewRows(local) ? local : null;
};

/** GET snapshot for fingerprint; POST rebuild only when explicitly allowed. */
export const fetchOverviewStaffSnapshot = async (
  options: StaffSnapshotRequest,
): Promise<BackendPersonnelOverview | null> => {
  const key = requestStaffSnapshotKey(options);
  if (!options.force) {
    const pending = staffSnapshotRequests.get(key);
    if (pending) return pending;
  }

  const promise = fetchOverviewStaffSnapshotInner(options).finally(() => {
    if (staffSnapshotRequests.get(key) === promise) {
      staffSnapshotRequests.delete(key);
    }
  });
  if (!options.force) staffSnapshotRequests.set(key, promise);
  return promise;
};

/** GET assets snapshot → POST rebuild on miss. */
export const fetchOverviewAssetsSnapshot = async (options: {
  fingerprint: string;
  signal?: AbortSignal;
  bootstrap?: BackendPersonnelBootstrap | null;
}): Promise<OverviewPersonnelAssetsSnapshot | null> => {
  const bootstrap = await resolveBootstrap(options);
  if (!bootstrap || bootstrap.assets.available) {
    const cached = await api
      .getOverviewAssets({
        fingerprint: options.fingerprint,
        signal: options.signal,
      })
      .catch(() => null);
    if (cached) return cached;
  }

  try {
    const rebuilt = await api.rebuildOverviewAssets({ signal: options.signal });
    if (rebuilt) return rebuilt;
  } catch {
    /* backend rebuild unavailable */
  }

  return null;
};

/** GET merged overview → POST rebuild on miss. */
export const fetchMergedOverviewSnapshot = async (options: {
  mergeFingerprint: string;
  signal?: AbortSignal;
  bootstrap?: BackendPersonnelBootstrap | null;
}): Promise<BackendPersonnelOverview | null> => {
  const bootstrap = options.bootstrap ?? null;
  const shouldTryGet =
    !bootstrap ||
    (bootstrap.merge.available &&
      bootstrap.merge.fingerprint === options.mergeFingerprint);

  if (shouldTryGet) {
    const cached = await api
      .getMergedPersonnelOverview({
        fingerprint: options.mergeFingerprint,
        signal: options.signal,
      })
      .catch(() => null);
    if (hasOverviewRows(cached)) return cached;
  }

  try {
    const rebuilt = await api.rebuildMergedPersonnelOverview();
    if (hasOverviewRows(rebuilt)) return rebuilt;
  } catch {
    /* backend rebuild unavailable */
  }

  return null;
};
