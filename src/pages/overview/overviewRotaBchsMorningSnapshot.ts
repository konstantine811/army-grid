import type { BackendPersonnelOverviewRow } from "../../api";
import {
  bchsMorningSnapshotCacheKey,
  CacheKeys,
  deleteDataCache,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import { KYIV_TIME_ZONE } from "../../shared/format";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { morningUnitKey, morningUnitMatches } from "./overviewMorningUnit";
import type { MorningReportBucket } from "./overviewRotaBchsMorningStatusChanges";

export type BchsMorningSnapshotPerson = {
  key: string;
  rank: string;
  name: string;
  callsign: string;
  position: string;
  status: string;
  location: string;
  bucket: MorningReportBucket;
  stateLabel: string;
  unit?: string;
  staffUnit?: string;
  /** Raw staff sheet cols 21/31/32 for stable «Штатка» baseline comparison. */
  staffCol21?: string;
  staffCol31?: string;
  staffCol32?: string;
};

export type MorningComparisonBaselineKind = "bchs" | "staff";

export type BchsMorningDailySnapshot = {
  date: string;
  unitLabel: string;
  savedAt: string;
  people: BchsMorningSnapshotPerson[];
  source?: "daily" | "manual";
  sourceLabel?: string;
  baselineKind?: MorningComparisonBaselineKind;
};

export type BchsMorningManualBaseline = {
  fileName: string;
  uploadedAt: string;
  baselineKind: MorningComparisonBaselineKind;
  snapshot: BchsMorningDailySnapshot;
};

type BchsMorningSnapshotStore = {
  byDate: Record<string, BchsMorningDailySnapshot>;
};

const SNAPSHOT_RETENTION_DAYS = 30;

export const kyivIsoDateLabel = (date = new Date()) =>
  new Intl.DateTimeFormat("en-CA", {
    timeZone: KYIV_TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).format(date);

export const bchsMorningPersonKey = (row: BackendPersonnelOverviewRow) => {
  const externalId = row.externalId?.trim();
  if (externalId) return `id:${externalId}`;
  return bchsMorningPersonKeyFromName(row.name ?? "");
};

export const bchsMorningPersonKeyFromName = (name: string) => {
  const normalized = normalizeRosterMatchText(name.trim());
  return normalized ? `name:${normalized}` : "";
};

export const bchsMorningPersonNameKey = (name: string) =>
  normalizeRosterMatchText(name.trim());

export const snapshotPersonMatchesUnit = (
  person: BchsMorningSnapshotPerson,
  unitLabel: string,
) => {
  if (!person.unit?.trim()) return false;
  return morningUnitMatches(person.unit, unitLabel);
};

export const filterSnapshotPeopleForUnit = (
  people: BchsMorningSnapshotPerson[],
  unitLabel: string,
) =>
  people.filter((person) => snapshotPersonMatchesUnit(person, unitLabel));

const pruneSnapshotStore = (
  store: BchsMorningSnapshotStore,
  keepFromDate: string,
): BchsMorningSnapshotStore => {
  const byDate: Record<string, BchsMorningDailySnapshot> = {};
  for (const [date, snapshot] of Object.entries(store.byDate)) {
    if (date >= keepFromDate) byDate[date] = snapshot;
  }
  return { byDate };
};

export const subtractKyivDays = (isoDate: string, days: number) => {
  const [year, month, day] = isoDate.split("-").map(Number);
  const utc = Date.UTC(year, month - 1, day - days);
  return kyivIsoDateLabel(new Date(utc));
};

export const findPreviousBchsMorningSnapshot = (
  store: BchsMorningSnapshotStore,
  todayIso = kyivIsoDateLabel(),
): BchsMorningDailySnapshot | null => {
  const candidates = Object.keys(store.byDate)
    .filter((date) => date < todayIso)
    .sort((left, right) => right.localeCompare(left));
  const yesterday = subtractKyivDays(todayIso, 1);
  return store.byDate[yesterday] ?? (candidates[0] ? store.byDate[candidates[0]] : null);
};

export const readBchsMorningSnapshotStore = async (
  unitLabel: string,
): Promise<BchsMorningSnapshotStore> =>
  (await readDataCache<BchsMorningSnapshotStore>(
    bchsMorningSnapshotCacheKey(unitLabel),
  )) ?? { byDate: {} };

export const readPreviousBchsMorningSnapshot = async (
  unitLabel: string,
  reportDate = new Date(),
): Promise<BchsMorningDailySnapshot | null> => {
  const store = await readBchsMorningSnapshotStore(unitLabel);
  return findPreviousBchsMorningSnapshot(store, kyivIsoDateLabel(reportDate));
};

export const saveBchsMorningDailySnapshot = async (
  snapshot: BchsMorningDailySnapshot,
) => {
  const store = await readBchsMorningSnapshotStore(snapshot.unitLabel);
  const keepFromDate = subtractKyivDays(snapshot.date, SNAPSHOT_RETENTION_DAYS);
  const next = pruneSnapshotStore(store, keepFromDate);
  next.byDate[snapshot.date] = { ...snapshot, source: snapshot.source ?? "daily" };
  await writeDataCache(bchsMorningSnapshotCacheKey(snapshot.unitLabel), next);
};

export const extractBchsMorningRotaNumber = (unitLabel: string) => {
  const key = morningUnitKey(unitLabel);
  return key.startsWith("company:") ? key.slice("company:".length) : "";
};

export const bchsMorningManualBaselineCacheKey = (unitLabel: string) => {
  const rotaNumber = extractBchsMorningRotaNumber(unitLabel);
  if (rotaNumber) {
    return `${CacheKeys.bchsMorningSnapshotPrefix}rota-${rotaNumber}:manual`;
  }
  return `${bchsMorningSnapshotCacheKey(unitLabel)}:manual`;
};

const legacyManualBaselineCacheKey = (unitLabel: string) =>
  `${bchsMorningSnapshotCacheKey(unitLabel)}:manual`;

export const readManualBchsMorningBaseline = async (
  unitLabel: string,
): Promise<BchsMorningManualBaseline | null> => {
  const primary = await readDataCache<BchsMorningManualBaseline>(
    bchsMorningManualBaselineCacheKey(unitLabel),
  );
  if (primary) return primary;
  return (
    (await readDataCache<BchsMorningManualBaseline>(
      legacyManualBaselineCacheKey(unitLabel),
    )) ?? null
  );
};

export const saveManualBchsMorningBaseline = async (
  baseline: BchsMorningManualBaseline,
) => {
  const key = bchsMorningManualBaselineCacheKey(baseline.snapshot.unitLabel);
  await writeDataCache(key, baseline);
  const legacyKey = legacyManualBaselineCacheKey(baseline.snapshot.unitLabel);
  if (legacyKey !== key) {
    await deleteDataCache(legacyKey);
  }
};

export const clearManualBchsMorningBaseline = async (unitLabel: string) => {
  await deleteDataCache(bchsMorningManualBaselineCacheKey(unitLabel));
  await deleteDataCache(legacyManualBaselineCacheKey(unitLabel));
};

export const readBchsMorningComparisonSnapshot = async (
  unitLabel: string,
  reportDate = new Date(),
): Promise<BchsMorningDailySnapshot | null> => {
  const manual = await readManualBchsMorningBaseline(unitLabel);
  if (manual?.snapshot?.people?.length) {
    return manual.snapshot;
  }
  return readPreviousBchsMorningSnapshot(unitLabel, reportDate);
};
