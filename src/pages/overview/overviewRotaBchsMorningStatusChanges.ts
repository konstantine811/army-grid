import type { BackendPersonnelOverviewRow } from "../../api";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildBchsMorningPersonRow,
  filterBchsMorningUnitRows,
  isMorningLeaveSectionPerson,
  isMorningMissionLocation,
  resolveReportStatus,
  type BchsMorningPersonRow,
  type BchsMorningSection,
} from "./overviewRotaBchsMorningExport";
import type {
  BchsMorningDailySnapshot,
  BchsMorningSnapshotPerson,
} from "./overviewRotaBchsMorningSnapshot";
import {
  bchsMorningPersonKey,
  bchsMorningPersonNameKey,
  filterSnapshotPeopleForUnit,
  kyivIsoDateLabel,
  subtractKyivDays,
} from "./overviewRotaBchsMorningSnapshot";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { looksLikePersonnelName } from "../personnel/personnelUtils";
import { resolveBchsMorningStaffUnit } from "./overviewRotaBchsMorningExport";

export type MorningReportBucket =
  | "inService"
  | "mission"
  | "hospital"
  | "medPoint"
  | "leave"
  | "detached"
  | "missing"
  | "training"
  | "dead"
  | "awol"
  | "newcomer";

const normalizeText = (value: string) =>
  value.trim().toLowerCase().replace(/\s+/g, " ");

const normalizeField = (value: string) =>
  normalizeRosterMatchText(value).replace(/\s+/g, " ");

const staffValue = (row: BackendPersonnelOverviewRow, columnNumber: number) =>
  row.staffSheetColumns?.[`staff_${columnNumber}`]?.trim() ?? "";

type MorningStaffColumnLookup = (
  row: BackendPersonnelOverviewRow,
  columnNumber: number,
) => string;

const buildMorningStaffColumnLookup = (
  rosterRows?: EjournalPreviewRow[] | null,
): MorningStaffColumnLookup => {
  const rosterById = new Map<string, EjournalPreviewRow>();
  for (const rosterRow of rosterRows ?? []) {
    const externalId = String(rosterRow.externalId ?? rosterRow.column_1 ?? "").trim();
    if (externalId) rosterById.set(externalId, rosterRow);
  }
  return (row, columnNumber) => {
    const externalId = row.externalId?.trim();
    const rosterRow = externalId ? rosterById.get(externalId) : undefined;
    if (rosterRow) return readRosterColumnValue(rosterRow, columnNumber);
    return staffValue(row, columnNumber);
  };
};

const isMorningInServiceStatus = (status: string) =>
  normalizeText(status).includes("в строю");

const isNewcomer = (row: BackendPersonnelOverviewRow) => {
  const status = normalizeText(
    `${staffValue(row, 21)} ${row.staffStatusLabel ?? ""} ${row.staffStatus ?? ""}`,
  );
  return status.includes("новоприбул");
};

const isDetachedPerson = (row: BackendPersonnelOverviewRow, readColumn: MorningStaffColumnLookup) => {
  const status = readColumn(row, 21) || resolveReportStatus(row);
  const normalized = normalizeText(status);
  return normalized.includes("відком") && !normalized.includes("приком");
};

const isAttachedPerson = (row: BackendPersonnelOverviewRow, readColumn: MorningStaffColumnLookup) => {
  const status = readColumn(row, 21) || resolveReportStatus(row);
  const normalized = normalizeText(status);
  return normalized.includes("приком") && !normalized.includes("відком");
};

const isHospitalMedical = (row: BackendPersonnelOverviewRow, readColumn: MorningStaffColumnLookup) => {
  const blob = normalizeText(`${readColumn(row, 31)} ${readColumn(row, 32)}`);
  return /шпит|госпітал|мед\.?\s*рот|цмкл|вмг|лікарн/.test(blob);
};

const isHospitalFromStaffFields = (staffLocationRaw: string, staffNoteRaw: string) => {
  const blob = normalizeText(`${staffLocationRaw} ${staffNoteRaw}`);
  return /шпит|госпітал|мед\.?\s*рот|цмкл|вмг|лікарн/.test(blob);
};

/** Classify presence using only staff columns — ignores EJOOS `row.status`. */
export const classifyMorningReportBucketFromStaffFields = (
  staffStatusRaw: string,
  staffLocationRaw: string,
  staffNoteRaw = "",
): MorningReportBucket => {
  const staffStatus = normalizeText(staffStatusRaw);
  const staffLocation = normalizeText(staffLocationRaw);
  const statusBlob = normalizeText(`${staffStatusRaw} ${staffLocationRaw} ${staffNoteRaw}`);
  if (staffStatus.includes("сзч") || staffStatus.includes("самовіл")) return "awol";
  if (staffStatus.includes("новоприбул")) return "newcomer";
  if (/загиб|^200$/.test(staffStatus)) return "dead";
  if (/навч|відряд/.test(staffStatus)) return "training";
  if (isMorningLeaveSectionPerson(staffStatusRaw, staffLocationRaw)) return "leave";
  if (/500|безв/.test(staffStatus) || staffLocation.includes("безв")) return "missing";
  if (staffStatus.includes("відком") || staffStatus.includes("приком")) return "detached";
  if (/шпит|госпітал|мед\.?\s*рот/.test(statusBlob)) return "hospital";
  if (/мед\.?\s*пункт|лік\.?пор|лік\.?відп|ліку|лік[.\s]|лікарн/.test(statusBlob)) {
    return isHospitalFromStaffFields(staffLocationRaw, staffNoteRaw)
      ? "hospital"
      : "medPoint";
  }
  if (isMorningMissionLocation(staffLocationRaw) || isMorningMissionLocation(staffStatusRaw)) {
    return "mission";
  }
  if (staffStatus.includes("в строю")) return "inService";
  return "inService";
};

type StaffMorningFields = {
  col21: string;
  col31: string;
  col32: string;
};

const staffMorningComparisonSignature = (fields: StaffMorningFields) =>
  [
    normalizeField(fields.col21),
    normalizeField(fields.col31),
    normalizeField(fields.col32),
  ].join("|");

const staffFieldsFromSnapshotPerson = (
  person: BchsMorningSnapshotPerson,
): StaffMorningFields => ({
  col21: person.staffCol21 ?? person.status ?? "",
  col31: person.staffCol31 ?? person.location ?? "",
  col32: person.staffCol32 ?? "",
});

const staffFieldsFromOverviewRow = (
  row: BackendPersonnelOverviewRow,
  readColumn: MorningStaffColumnLookup,
): StaffMorningFields => ({
  col21: readColumn(row, 21),
  col31: readColumn(row, 31),
  col32: readColumn(row, 32),
});

const buildSnapshotPersonFromStaffFields = (
  fields: StaffMorningFields,
  base: Omit<BchsMorningSnapshotPerson, "bucket" | "stateLabel">,
): BchsMorningSnapshotPerson => {
  const bucket = classifyMorningReportBucketFromStaffFields(
    fields.col21,
    fields.col31,
    fields.col32,
  );
  return {
    ...base,
    status: fields.col21,
    location: fields.col31,
    staffCol21: fields.col21,
    staffCol31: fields.col31,
    staffCol32: fields.col32,
    bucket,
    stateLabel: formatMorningPersonStateLabel(
      { status: fields.col21, location: fields.col31 },
      bucket,
    ),
  };
};

export const morningReportBucketLabel = (bucket: MorningReportBucket) => {
  switch (bucket) {
    case "inService":
      return "В строю";
    case "mission":
      return "На виконанні";
    case "hospital":
      return "Шпиталь/Мед. рота";
    case "medPoint":
      return "Мед. пункт";
    case "leave":
      return "Відпустка/Лік. відпустка";
    case "detached":
      return "Відкомандировані";
    case "missing":
      return "Зниклі безвісти";
    case "training":
      return "Навчання/відрядження";
    case "dead":
      return "Загиблі";
    case "awol":
      return "СЗЧ";
    case "newcomer":
      return "Новоприбулий";
    default:
      return "—";
  }
};

export const classifyMorningReportBucket = (
  row: BackendPersonnelOverviewRow,
  rosterRows?: EjournalPreviewRow[] | null,
): MorningReportBucket => {
  const readColumn = buildMorningStaffColumnLookup(rosterRows);
  const staffStatusEarly = normalizeText(readColumn(row, 21));
  const reportStatusEarly = normalizeText(resolveReportStatus(row));
  if (
    row.status === "AWOL" ||
    staffStatusEarly.includes("сзч") ||
    staffStatusEarly.includes("самовіл") ||
    reportStatusEarly.includes("сзч")
  ) {
    return "awol";
  }
  if (isNewcomer(row)) return "newcomer";
  const reportStatus = normalizeText(resolveReportStatus(row));
  if (/загиб|^200$/.test(reportStatus)) return "dead";
  if (/навч|відряд/.test(reportStatus)) return "training";
  const staffStatus = readColumn(row, 21);
  const staffLocation = readColumn(row, 31);
  if (isMorningLeaveSectionPerson(staffStatus, staffLocation)) return "leave";
  if (/500|безв/.test(reportStatus) || row.status === "MISSING") return "missing";
  if (isDetachedPerson(row, readColumn) || isAttachedPerson(row, readColumn)) {
    return "detached";
  }
  if (row.status === "MEDICAL") {
    return isHospitalMedical(row, readColumn) ? "hospital" : "medPoint";
  }
  const statusBlob = normalizeText(
    `${staffStatus} ${row.statusLabel ?? ""} ${staffLocation} ${readColumn(row, 32)}`,
  );
  if (/ліку|лік[.\s]|лікарн|госпіт|шпит|мед\.?\s*(?:пункт|рот)/.test(statusBlob)) {
    return isHospitalMedical(row, readColumn) ? "hospital" : "medPoint";
  }
  if (isMorningMissionLocation(staffLocation) || isMorningMissionLocation(staffStatus)) return "mission";
  if (isMorningInServiceStatus(staffStatus || reportStatus)) return "inService";
  return "inService";
};

export const formatMorningPersonStateLabel = (
  person: Pick<BchsMorningPersonRow, "status" | "location">,
  bucket: MorningReportBucket,
) => {
  const status = bucket === "mission" || bucket === "awol"
    ? morningReportBucketLabel(bucket)
    : person.status?.trim() || morningReportBucketLabel(bucket);
  const location = person.location?.trim();
  if (location && bucket !== "inService") return `${status} · ${location}`;
  return status;
};

const MEDICAL_BUCKETS = new Set<MorningReportBucket>(["hospital", "medPoint"]);
const IN_RANKS_BUCKETS = new Set<MorningReportBucket>(["inService", "mission"]);

const isMedicalLeavePerson = (
  person: Pick<BchsMorningSnapshotPerson, "status" | "location" | "bucket">,
) =>
  person.bucket === "leave" &&
  /лік/.test(normalizeField(`${person.status ?? ""} ${person.location ?? ""}`));

/** Collapse buckets that often drift between BCHS import and live classification. */
const morningStatusComparisonKey = (
  bucket: MorningReportBucket,
  person?: Pick<BchsMorningSnapshotPerson, "status" | "location" | "bucket">,
) => {
  if (MEDICAL_BUCKETS.has(bucket)) return "medical";
  if (person && isMedicalLeavePerson(person)) return "medical";
  return bucket;
};

const looksLikeMorningStatusChangePerson = (name: string) =>
  looksLikePersonnelName(name) &&
  !/^(лікуван\w*|лікувальн\w*|відпуст\w*|разом|на\s*виконан\w*|навчання)/i.test(
    name.trim(),
  );

export const describeMorningStatusChange = (
  from: MorningReportBucket,
  to: MorningReportBucket,
): string => {
  if (from === to) return "";
  if (from === "mission" && (to === "inService" || to === "mission")) {
    return to === "inService" ? "Прийшов з виконання" : "";
  }
  if (from === "mission" && IN_RANKS_BUCKETS.has(to) && to !== "mission") {
    return "Прийшов з виконання";
  }
  if (MEDICAL_BUCKETS.has(from) && (to === "inService" || to === "mission")) {
    return "Прийшов з лікування";
  }
  if (MEDICAL_BUCKETS.has(from) && to === "leave") {
    return "З лікування у відпустку";
  }
  if (from === "training" && (to === "inService" || to === "mission")) {
    return "Прийшов з навчання/відрядження";
  }
  if (from === "leave" && (to === "inService" || to === "mission")) {
    return "Прийшов з відпустки";
  }
  if (from === "detached" && (to === "inService" || to === "mission")) {
    return "Повернувся з відкомандировки";
  }
  if (from === "awol" && (to === "inService" || to === "mission")) {
    return "Повернувся після СЗЧ";
  }
  if (from === "missing" && (to === "inService" || to === "mission")) {
    return "Знайдений / повернувся";
  }
  if (IN_RANKS_BUCKETS.has(from) && to === "mission") return "Пішов на виконання";
  if (IN_RANKS_BUCKETS.has(from) && MEDICAL_BUCKETS.has(to)) {
    return "Пішов на лікування";
  }
  if (IN_RANKS_BUCKETS.has(from) && to === "training") {
    return "Пішов на навчання/відрядження";
  }
  if (IN_RANKS_BUCKETS.has(from) && to === "leave") return "Пішов у відпустку";
  if (IN_RANKS_BUCKETS.has(from) && to === "detached") {
    return "Відкомандирований";
  }
  if (IN_RANKS_BUCKETS.has(from) && to === "awol") {
    return `${morningReportBucketLabel(from)} → СЗЧ`;
  }
  if (IN_RANKS_BUCKETS.has(from) && to === "missing") {
    return "Зниклий безвісти";
  }
  if (from === "newcomer" && to === "inService") return "З новоприбулих у стрій";
  if (to === "newcomer") return "Новоприбулий";
  if (to === "dead") return "Загиблий";
  return `${morningReportBucketLabel(from)} → ${morningReportBucketLabel(to)}`;
};

const formatStaffComparisonState = (
  fields: StaffMorningFields,
  bucket: MorningReportBucket,
) =>
  formatMorningComparisonState({
    status: fields.col21,
    location: fields.col31,
    bucket,
    stateLabel: "",
  });

export const formatMorningComparisonState = (
  person: Pick<BchsMorningSnapshotPerson, "status" | "location" | "stateLabel" | "bucket">,
) => {
  const status = person.bucket === "mission" || person.bucket === "awol"
    ? morningReportBucketLabel(person.bucket)
    : person.status?.trim();
  const location = person.location?.trim();
  if (status && location && normalizeField(status) === normalizeField(location)) return status;
  if (status && location) return `${status} · ${location}`;
  return status || location || person.stateLabel?.trim() || "—";
};

export const describeMorningPersonChanges = (
  previous: BchsMorningSnapshotPerson,
  current: BchsMorningSnapshotPerson,
): string => describeMorningStatusChange(previous.bucket, current.bucket);

export type BchsMorningStatusChangeRow = BchsMorningPersonRow & {
  previousState: string;
  newState: string;
  changeNote: string;
  tableTransition: string;
};

export const formatMorningStatusTransition = (
  previous: BchsMorningSnapshotPerson,
  current: BchsMorningSnapshotPerson,
) => {
  const stateLabel = (person: BchsMorningSnapshotPerson) => {
    // In-service destinations are clearer as the actual PPD name. Mission,
    // medical and AWOL remain explicit even if an old location is still stored.
    if (person.bucket === "inService" && person.location?.trim()) {
      return person.location.trim();
    }
    return morningReportBucketLabel(person.bucket);
  };
  return `${stateLabel(previous)} → ${stateLabel(current)}`;
};

export const buildBchsMorningDailySnapshot = (
  rows: BackendPersonnelOverviewRow[],
  unitLabel: string,
  rosterRows?: EjournalPreviewRow[] | null,
  reportDate = new Date(),
): BchsMorningDailySnapshot => {
  const readColumn = buildMorningStaffColumnLookup(rosterRows);
  const staffUnit = resolveBchsMorningStaffUnit(unitLabel);
  const people: BchsMorningSnapshotPerson[] = [];
  const unitRows = filterBchsMorningUnitRows(rows, unitLabel);
  unitRows.forEach((row, index) => {
    const key = bchsMorningPersonKey(row);
    if (!key) return;
    const person = buildBchsMorningPersonRow(row, index, staffUnit, readColumn);
    if (!person.name) return;
    const staffFields = staffFieldsFromOverviewRow(row, readColumn);
    const bucket = classifyMorningReportBucket(row, rosterRows);
    people.push({
      key,
      rank: person.rank,
      name: person.name,
      callsign: person.callsign,
      position: person.position,
      status: person.status,
      location: person.location,
      bucket,
      stateLabel: formatMorningPersonStateLabel(person, bucket),
      unit: row.unit?.trim() || "",
      staffUnit: person.staffUnit,
      staffCol21: staffFields.col21,
      staffCol31: staffFields.col31,
      staffCol32: staffFields.col32,
    });
  });
  return {
    date: kyivIsoDateLabel(reportDate),
    unitLabel,
    savedAt: new Date().toISOString(),
    people,
  };
};

export const buildBchsMorningStatusChangeRows = (
  rows: BackendPersonnelOverviewRow[],
  previousSnapshot: BchsMorningDailySnapshot | null,
  unitLabel: string,
  rosterRows?: EjournalPreviewRow[] | null,
): BchsMorningStatusChangeRow[] => {
  if (!previousSnapshot) return [];
  const compareStaffFields = previousSnapshot.baselineKind === "staff";
  const unitRows = filterBchsMorningUnitRows(rows, unitLabel);
  const previousPeople = filterSnapshotPeopleForUnit(
    previousSnapshot.people,
    unitLabel,
  );
  const readColumn = buildMorningStaffColumnLookup(rosterRows);
  const staffUnit = resolveBchsMorningStaffUnit(unitLabel);
  const previousByKey = new Map(
    previousPeople.map((person) => [person.key, person]),
  );
  const previousByName = new Map<string, BchsMorningSnapshotPerson[]>();
  for (const person of previousPeople) {
    const name = bchsMorningPersonNameKey(person.name);
    previousByName.set(name, [...(previousByName.get(name) ?? []), person]);
  }
  const changes: BchsMorningStatusChangeRow[] = [];

  const resolvePreviousPerson = (
    key: string,
    name: string,
  ): BchsMorningSnapshotPerson | undefined => {
    const nameKey = bchsMorningPersonNameKey(name);
    const byKey = previousByKey.get(key);
    if (byKey && bchsMorningPersonNameKey(byKey.name) === nameKey) return byKey;
    const candidates = previousByName.get(nameKey) ?? [];
    return candidates.length === 1 ? candidates[0] : undefined;
  };

  unitRows.forEach((row, index) => {
    const key = bchsMorningPersonKey(row);
    if (!key) return;
    const person = buildBchsMorningPersonRow(row, index, staffUnit, readColumn);
    if (!person.name || !looksLikeMorningStatusChangePerson(person.name)) return;
    const bucket = classifyMorningReportBucket(row, rosterRows);
    const currentSnapshot: BchsMorningSnapshotPerson = {
      key,
      rank: person.rank,
      name: person.name,
      callsign: person.callsign,
      position: person.position,
      status: person.status,
      location: person.location,
      bucket,
      stateLabel: formatMorningPersonStateLabel(person, bucket),
      unit: row.unit?.trim() || "",
      staffUnit: person.staffUnit,
    };
    const previous = resolvePreviousPerson(key, person.name);
    if (!previous) return;

    const currentStaffFields = staffFieldsFromOverviewRow(row, readColumn);
    const previousStaffFields = staffFieldsFromSnapshotPerson(previous);

    if (compareStaffFields) {
      if (
        staffMorningComparisonSignature(previousStaffFields) ===
        staffMorningComparisonSignature(currentStaffFields)
      ) {
        return;
      }
      const previousBucket = classifyMorningReportBucketFromStaffFields(
        previousStaffFields.col21,
        previousStaffFields.col31,
        previousStaffFields.col32,
      );
      const currentBucket = classifyMorningReportBucketFromStaffFields(
        currentStaffFields.col21,
        currentStaffFields.col31,
        currentStaffFields.col32,
      );
      if (
        morningStatusComparisonKey(previousBucket) ===
        morningStatusComparisonKey(currentBucket)
      ) {
        return;
      }
      const changeNote =
        describeMorningStatusChange(previousBucket, currentBucket) ||
        `${morningReportBucketLabel(previousBucket)} → ${morningReportBucketLabel(currentBucket)}`;
      const previousView = buildSnapshotPersonFromStaffFields(previousStaffFields, {
        ...previous,
        status: previousStaffFields.col21,
        location: previousStaffFields.col31,
      });
      const currentView = buildSnapshotPersonFromStaffFields(currentStaffFields, {
        ...currentSnapshot,
        status: currentStaffFields.col21,
        location: currentStaffFields.col31,
      });
      const previousState = formatStaffComparisonState(
        previousStaffFields,
        previousBucket,
      );
      const newState = formatStaffComparisonState(currentStaffFields, currentBucket);
      const tableTransition = formatMorningStatusTransition(previousView, currentView);
      changes.push({
        ...person,
        previousState,
        newState,
        changeNote,
        tableTransition,
        status: previousState,
        location: newState,
        staffUnit: changeNote,
      });
      return;
    }

    if (
      morningStatusComparisonKey(previous.bucket, previous) ===
      morningStatusComparisonKey(currentSnapshot.bucket, currentSnapshot)
    ) {
      return;
    }
    const changeNote = describeMorningPersonChanges(previous, currentSnapshot);
    if (!changeNote.trim()) {
      return;
    }
    const previousState = formatMorningComparisonState(previous);
    const newState = formatMorningComparisonState(currentSnapshot);
    const tableTransition = formatMorningStatusTransition(previous, currentSnapshot);
    changes.push({
      ...person,
      previousState,
      newState,
      changeNote,
      tableTransition,
      status: previousState,
      location: newState,
      staffUnit: changeNote,
    });
  });

  return changes.sort((left, right) => left.name.localeCompare(right.name, "uk"));
};

export const buildBchsMorningStatusChangeSection = (
  rows: BackendPersonnelOverviewRow[],
  previousSnapshot: BchsMorningDailySnapshot | null,
  unitLabel: string,
  rosterRows?: EjournalPreviewRow[] | null,
): BchsMorningSection => {
  const people = buildBchsMorningStatusChangeRows(
    rows,
    previousSnapshot,
    unitLabel,
    rosterRows,
  ).map((row) => ({
    ...row,
    status: row.previousState,
    location: row.newState,
    staffUnit: row.changeNote,
  }));
  const previousDateLabel = previousSnapshot?.date
    ? previousSnapshot.date.split("-").reverse().join(".")
    : "";
  const comparisonTitle = (() => {
    if (!previousSnapshot) {
      return "Зміна статусу (немає бази для порівняння)";
    }
    if (previousSnapshot.source === "manual") {
      const label = previousSnapshot.sourceLabel?.trim();
      const sourceName =
        previousSnapshot.baselineKind === "staff" ? "Штатка" : "БЧС";
      const notYesterdayHint =
        previousSnapshot.baselineKind === "staff" &&
        previousSnapshot.date &&
        previousSnapshot.date < subtractKyivDays(kyivIsoDateLabel(), 1)
          ? " · не вчорашній зріз"
          : "";
      if (label && previousDateLabel) {
        return `Зміна статусу (${sourceName} «${label}», ${previousDateLabel}${notYesterdayHint})`;
      }
      if (label) return `Зміна статусу (${sourceName} «${label}»)`;
      if (previousDateLabel) {
        return `Зміна статусу (завантажена ${sourceName}, ${previousDateLabel})`;
      }
      return `Зміна статусу (завантажена ${sourceName})`;
    }
    return `Зміна статусу (відносно ${previousDateLabel})`;
  })();
  return {
    id: "statusChanges",
    title: comparisonTitle,
    side: "extra",
    people,
  };
};
