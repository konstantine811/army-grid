import JSZip from "jszip";
import type { BackendEjournalLiveVersion } from "../../api";
import type { EjoosOpKind, EjoosSyncOp } from "./ejoosSyncPlan";
import { selectVersionsForHistorySheet, findLastImportVersion } from "./ejoosChangeHistorySheet";
import {
  personChangesFromOps,
  type PersonChange,
  type PersonChangeCategory,
  type PersonTimesheetPreview,
  type SheetImpactItem,
  type PersonSourceInfluence,
} from "./ejoosPersonDiff";

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Повний журнал застосувань всередині .xlsx — не залежить від аркуша «10». */
export const EJOOS_APPLIED_HISTORY_ZIP_PATH = "xl/ejoosAppliedChanges.json";

const OP_KINDS = new Set<EjoosOpKind>([
  "timesheet_day",
  "shpo_occupant",
  "absent_upsert",
  "absent_close",
  "exclude_transfer",
  "move_to_disposition",
  "data_mismatch",
  "position_change",
  "rank_change",
  "arrival",
  "other_manual",
]);

const CATEGORIES = new Set<PersonChangeCategory>([
  "status",
  "position",
  "arrival",
  "data",
  "error",
  "mixed",
]);

export type StoredAppliedPerson = {
  id: string;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  category: PersonChangeCategory;
  summaryBefore: string;
  summaryAfter: string;
  ejoosWillDo: string[];
  sheetImpacts: SheetImpactItem[];
  sourceInfluences: PersonSourceInfluence[];
  timesheetPreview: PersonTimesheetPreview | null;
  ops: Array<Record<string, unknown>>;
};

export type StoredApplyRecord = {
  version: number;
  appliedAt: string;
  pbFileName: string;
  timesheetDay: number;
  timesheetDayLabel: string;
  people: StoredAppliedPerson[];
};

export type AppliedHistoryEntry = {
  id: string;
  person: PersonChange;
  version: number;
  versionId: string;
  appliedAt: string;
  timesheetDay: number;
  timesheetDayLabel: string;
  pbFileName: string;
};

const asRecord = (value: unknown): Record<string, unknown> | null =>
  value && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;

const asString = (value: unknown) =>
  typeof value === "string"
    ? value.trim()
    : typeof value === "number" && Number.isFinite(value)
      ? String(value)
      : "";

const asStringArray = (value: unknown) =>
  Array.isArray(value)
    ? value.map((item) => asString(item)).filter(Boolean)
    : [];

const asPayload = (value: unknown): Record<string, string> => {
  const rec = asRecord(value);
  if (!rec) return {};
  const out: Record<string, string> = {};
  for (const [key, item] of Object.entries(rec)) {
    if (item == null) continue;
    out[key] = String(item);
  }
  return out;
};

const isOpKind = (value: string): value is EjoosOpKind =>
  OP_KINDS.has(value as EjoosOpKind);

export const serializeSyncOp = (op: EjoosSyncOp): Record<string, unknown> => ({
  id: op.id,
  kind: op.kind,
  sheet: op.sheet,
  personId: op.personId,
  fullName: op.fullName,
  positionIndex: op.positionIndex,
  rank: op.rank,
  before: op.before,
  after: op.after,
  sourceRef: op.sourceRef,
  why: op.why,
  movementKey: op.movementKey,
  payload: op.payload,
});

export const serializeAppliedPerson = (
  person: PersonChange,
): StoredAppliedPerson => ({
  id: person.id,
  personId: person.personId,
  fullName: person.fullName,
  rank: person.rank,
  positionIndex: person.positionIndex,
  category: person.category,
  summaryBefore: person.summaryBefore,
  summaryAfter: person.summaryAfter,
  ejoosWillDo: [...person.ejoosWillDo],
  sheetImpacts: person.sheetImpacts ?? [],
  sourceInfluences: person.sourceInfluences ?? [],
  timesheetPreview: person.timesheetPreview,
  ops: person.ops.map(serializeSyncOp),
});

export const hydrateOp = (
  raw: unknown,
  index: number,
): EjoosSyncOp | null => {
  const rec = asRecord(raw);
  if (!rec) return null;
  const kind = asString(rec.kind);
  if (!isOpKind(kind)) return null;
  const payload = asPayload(rec.payload);
  const personId = asString(rec.personId);
  const fullName = asString(rec.fullName);
  return {
    id: asString(rec.id) || `hist:${kind}:${personId || fullName || index}`,
    kind,
    class: "ready",
    sheet: asString(rec.sheet),
    personId,
    fullName,
    positionIndex: asString(rec.positionIndex),
    rank: asString(rec.rank) || payload.fromRank || payload.nextRank || "",
    before: asString(rec.before),
    after: asString(rec.after),
    sourceRef: asString(rec.sourceRef),
    why: asString(rec.why),
    confidence: "high",
    payload,
    movementKey: asString(rec.movementKey) || undefined,
    checkedDefault: true,
  };
};

const hydrateStoredPerson = (
  raw: unknown,
  timesheetDay: number,
): PersonChange | null => {
  const rec = asRecord(raw);
  if (!rec) return null;
  const ops = Array.isArray(rec.ops)
    ? rec.ops
        .map((item, index) => hydrateOp(item, index))
        .filter((op): op is EjoosSyncOp => Boolean(op))
    : [];
  const rebuilt = ops.length
    ? personChangesFromOps(ops, timesheetDay, { decision: "accepted" })[0]
    : null;
  const fullName = asString(rec.fullName) || rebuilt?.fullName || "";
  if (!fullName && !ops.length) return null;
  const categoryRaw = asString(rec.category) as PersonChangeCategory;
  const category = CATEGORIES.has(categoryRaw)
    ? categoryRaw
    : rebuilt?.category || "mixed";
  return {
    id: asString(rec.id) || rebuilt?.id || `name:${fullName}`,
    personId: asString(rec.personId) || rebuilt?.personId || "",
    fullName: fullName || "Без ПІБ",
    rank: asString(rec.rank) || rebuilt?.rank || "",
    positionIndex: asString(rec.positionIndex) || rebuilt?.positionIndex || "",
    category,
    severity: "ready",
    summaryBefore:
      asString(rec.summaryBefore) || rebuilt?.summaryBefore || "—",
    summaryAfter: asString(rec.summaryAfter) || rebuilt?.summaryAfter || "—",
    ejoosWillDo: asStringArray(rec.ejoosWillDo).length
      ? asStringArray(rec.ejoosWillDo)
      : rebuilt?.ejoosWillDo || [],
    sheetActions: rebuilt?.sheetActions || [],
    sheetImpacts: Array.isArray(rec.sheetImpacts)
      ? (rec.sheetImpacts as SheetImpactItem[])
      : rebuilt?.sheetImpacts || [],
    sourceInfluences: Array.isArray(rec.sourceInfluences)
      ? (rec.sourceInfluences as PersonSourceInfluence[])
      : rebuilt?.sourceInfluences || [],
    timesheetPreview: rec.timesheetPreview
      ? (rec.timesheetPreview as PersonTimesheetPreview)
      : rebuilt?.timesheetPreview || null,
    ops: ops.length ? ops : rebuilt?.ops || [],
    decision: "accepted",
  };
};

const protocolOpsOf = (version: BackendEjournalLiveVersion): EjoosSyncOp[] => {
  const protocol = asRecord(version.changeProtocol);
  if (!protocol) return [];
  const kind = asString(protocol.kind);
  if (kind === "rollback" || kind === "import" || kind.startsWith("anketa")) {
    return [];
  }
  if (kind && kind !== "apply") return [];
  const raw = Array.isArray(protocol.ops) ? protocol.ops : [];
  return raw
    .map((item, index) => hydrateOp(item, index))
    .filter((op): op is EjoosSyncOp => Boolean(op));
};

const protocolPeopleOf = (
  version: BackendEjournalLiveVersion,
  timesheetDay: number,
): PersonChange[] => {
  const protocol = asRecord(version.changeProtocol);
  const raw = Array.isArray(protocol?.people) ? protocol.people : [];
  const stored = raw
    .map((item) => hydrateStoredPerson(item, timesheetDay))
    .filter((person): person is PersonChange => Boolean(person));
  if (stored.length) return stored;
  const ops = protocolOpsOf(version);
  if (!ops.length) return [];
  return personChangesFromOps(ops, timesheetDay, { decision: "accepted" });
};

const readHistoryFile = async (zip: JSZip): Promise<StoredApplyRecord[]> => {
  const raw = await zip.file(EJOOS_APPLIED_HISTORY_ZIP_PATH)?.async("string");
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw) as { applies?: unknown };
    return Array.isArray(parsed.applies)
      ? parsed.applies.filter(
          (item): item is StoredApplyRecord =>
            Boolean(asRecord(item) && Number(asRecord(item)?.version)),
        )
      : [];
  } catch {
    return [];
  }
};

/** Журнал застосувань, який лежить у самому .xlsx поточного файлу. */
export const readAppliedHistoryFromWorkbook = async (
  file: Blob | File,
): Promise<StoredApplyRecord[]> => {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  return readHistoryFile(zip);
};

const entriesFromPeople = (
  people: PersonChange[],
  meta: {
    version: number;
    versionId: string;
    appliedAt: string;
    timesheetDay: number;
    timesheetDayLabel: string;
    pbFileName: string;
  },
): AppliedHistoryEntry[] =>
  people.map((person) => {
    const id = `${meta.versionId}:${person.id}`;
    return {
      id,
      person: { ...person, id, decision: "accepted" as const },
      ...meta,
    };
  });

const entriesFromFileApplies = (
  applies: StoredApplyRecord[],
): AppliedHistoryEntry[] => {
  const entries: AppliedHistoryEntry[] = [];
  for (const apply of applies) {
    const timesheetDay =
      Number.isFinite(apply.timesheetDay) && apply.timesheetDay > 0
        ? apply.timesheetDay
        : 31;
    const people = (apply.people ?? [])
      .map((item) => hydrateStoredPerson(item, timesheetDay))
      .filter((person): person is PersonChange => Boolean(person));
    entries.push(
      ...entriesFromPeople(people, {
        version: apply.version,
        versionId: `file:${apply.version}`,
        appliedAt: apply.appliedAt,
        timesheetDay,
        timesheetDayLabel: apply.timesheetDayLabel || "",
        pbFileName: apply.pbFileName || "",
      }),
    );
  }
  return entries;
};

/**
 * Історія застосувань поточного файлу ЕЖООС — не вся стрічка версій БД.
 * Спочатку журнал усередині .xlsx; якщо його ще немає — протоколи після
 * останнього імпорту цього файлу.
 */
export const collectAppliedHistoryEntries = (
  versions: BackendEjournalLiveVersion[],
  currentVersionNumber: number,
  fileApplies?: StoredApplyRecord[] | null,
): AppliedHistoryEntry[] => {
  const fromFile = fileApplies?.length ? entriesFromFileApplies(fileApplies) : [];
  const lastImport = findLastImportVersion(versions, currentVersionNumber);
  const inScope = selectVersionsForHistorySheet(
    versions,
    currentVersionNumber,
  ).filter((version) => version.version > lastImport);

  const fromProtocols: AppliedHistoryEntry[] = [];
  for (const version of inScope) {
    const protocol = asRecord(version.changeProtocol);
    const timesheetDayRaw = Number(protocol?.timesheetDay);
    const timesheetDay =
      Number.isFinite(timesheetDayRaw) && timesheetDayRaw > 0
        ? timesheetDayRaw
        : 31;
    const people = protocolPeopleOf(version, timesheetDay);
    fromProtocols.push(
      ...entriesFromPeople(people, {
        version: version.version,
        versionId: version.id,
        appliedAt: version.createdAt,
        timesheetDay,
        timesheetDayLabel: asString(protocol?.timesheetDayLabel),
        pbFileName:
          asString(protocol?.pbFileName) ||
          asString(version.sourcePbFileName) ||
          "",
      }),
    );
  }

  const jsonVersions = new Set(fromFile.map((entry) => entry.version));
  const entries = [
    ...fromFile,
    ...fromProtocols.filter((entry) => !jsonVersions.has(entry.version)),
  ];
  entries.sort((a, b) => {
    const time = Date.parse(b.appliedAt) - Date.parse(a.appliedAt);
    if (Number.isFinite(time) && time !== 0) return time;
    if (b.version !== a.version) return b.version - a.version;
    return a.person.fullName.localeCompare(b.person.fullName, "uk");
  });
  return entries;
};

/** Не пишемо custom JSON у xl/ — Excel тоді відкриває книгу через Repair. */
export const appendAppliedHistoryToWorkbook = async (
  file: Blob | File,
  _apply?: StoredApplyRecord,
): Promise<Blob> => {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  if (!zip.file(EJOOS_APPLIED_HISTORY_ZIP_PATH)) return file;
  zip.remove(EJOOS_APPLIED_HISTORY_ZIP_PATH);
  return zip.generateAsync({
    type: "blob",
    mimeType: XLSX_MIME,
    compression: "DEFLATE",
  });
};
