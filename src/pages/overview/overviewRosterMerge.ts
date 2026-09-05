import type {
  BackendPersonnelOverview,
  BackendPersonnelOverviewRow,
} from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  extractBchsAwayPeopleFromDbRows,
  hasBchsFullName,
  normalizeBchsText,
} from "../bchs/bchsCalc";
import {
  buildPersonIdentityFingerprint,
  classifyOverviewStatusFromRoster,
  getPersonDisplayName,
  getPersonExternalId,
  getPersonFullPositionTitle,
  isLikelyPersonnelRow,
  isUnstablePersonExternalId,
  resolvePersonIdentityKey,
  resolvePersonRankTitle,
  resolvePersonRosterStatus,
} from "../personnel/personnelUtils";
import {
  getRosterPersonName,
  getRosterUnit,
  getRosterValue,
  isPersonnelInStaffRoster,
} from "../personnel/personnelRosterMerge";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  getRosterFighterStatusOverviewFields,
  normalizeRosterMatchText,
} from "../personnel/fighterStatusImport";
import { buildStaffSheetColumnsRecord } from "./overviewStaffSheetColumns";

const normalizeRosterText = normalizeRosterMatchText;

const applyStaffRosterStatus = (
  rosterRow: EjournalPreviewRow,
  rosterLabels: Record<string, string>,
) => {
  const mapped = classifyOverviewStatusFromRoster(
    resolvePersonRosterStatus(rosterRow, rosterLabels),
  );
  return {
    staffStatus: mapped.status,
    staffStatusLabel: mapped.statusLabel,
  };
};

export type NovaStaffRosterSummary = {
  positions: number;
  people: number;
  vacant: number;
  namedRows: EjournalPreviewRow[];
  battalions: string[];
};

export const rosterBattalionLabel = (
  person: { battalion?: string } | undefined,
  row: EjournalPreviewRow,
) => {
  const fromPerson = normalizeBchsText(person?.battalion);
  if (fromPerson) return fromPerson;
  const fromRow = (
    getRosterValue(row, ["батальйон"]) ||
    readRosterColumnValue(row, 1) ||
    String(row.column_1 ?? "").trim() ||
    String(row.battalion ?? "").trim() ||
    String(row.A ?? "").trim()
  )
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase()
    .replace(/ё/g, "е");
  return fromRow;
};

const sortBattalionLabels = (labels: string[]) =>
  [...labels].sort((left, right) => {
    if (left === "нова") return -1;
    if (right === "нова") return 1;
    if (left === "стара") return -1;
    if (right === "стара") return 1;
    return left.localeCompare(right, "uk", { numeric: true, sensitivity: "base" });
  });

/** Google Sheets може зберігати роту об'єднаною коміркою — заповнюємо її для рядків посад нижче. */
export const fillDownRosterUnitRows = (
  rows: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  let currentUnit = "";
  return rows.map((row) => {
    const unit = readRosterColumnValue(row, 2).trim();
    if (unit) {
      currentUnit = unit;
      return row;
    }
    const isPositionRow = [5, 8, 13, 14].some((columnNumber) =>
      readRosterColumnValue(row, columnNumber).trim(),
    );
    if (!currentUnit || !isPositionRow) return row;
    return { ...row, column_2: currentUnit };
  });
};

/** Посади / люди зі Штатки. `battalion = ALL` — усі пункти, не лише «нова». */
export const summarizeStaffFromRoster = (
  rosterRows: EjournalPreviewRow[],
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
  battalion: string = "ALL",
): NovaStaffRosterSummary => {
  if (!rosterRows.length) {
    return { positions: 0, people: 0, vacant: 0, namedRows: [], battalions: [] };
  }

  const extracted = extractBchsAwayPeopleFromDbRows(rosterRows, columns);
  const namedRows: EjournalPreviewRow[] = [];
  const battalionSet = new Set<string>();
  let positions = 0;
  rosterRows.forEach((row, index) => {
    const person = extracted[index];
    const label = rosterBattalionLabel(person, row);
    if (label) battalionSet.add(label);
    if (!label) return;
    if (battalion !== "ALL" && label !== battalion) return;
    positions += 1;
    if (person && hasBchsFullName(person)) namedRows.push(row);
  });
  return {
    positions,
    people: namedRows.length,
    vacant: Math.max(0, positions - namedRows.length),
    namedRows,
    battalions: sortBattalionLabels([...battalionSet]),
  };
};

/** Той самий набір, що БЧС по Штатці: «нова» + є ПІБ. Посади без прізвища — вакант. */
export const summarizeNovaStaffFromRoster = (
  rosterRows: EjournalPreviewRow[],
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
): NovaStaffRosterSummary =>
  summarizeStaffFromRoster(rosterRows, columns, "нова");

const namedNovaRowsFromRoster = (
  rosterRows: EjournalPreviewRow[],
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
) => new Set(summarizeNovaStaffFromRoster(rosterRows, columns).namedRows);

export const rosterRowToOverviewRow = (
  rosterRow: EjournalPreviewRow,
  rosterLabels: Record<string, string> = {},
  options: {
    requireName?: boolean;
    inNovaStaff?: boolean;
    inStaff?: boolean;
    name?: string;
    battalion?: string;
  } = {},
): BackendPersonnelOverviewRow | null => {
  const name = options.name?.trim() || getRosterPersonName(rosterRow);
  if (options.requireName !== false && !name) return null;

  const displayName = name || "Без ПІБ";
  const identityKey = name
    ? resolvePersonIdentityKey({
        ...rosterRow,
        прізвище: name,
        ПІБ: name,
      })
    : "";
  const fallbackKey = String(
    rosterRow.__dbRowId || rosterRow.__rowNumber || "",
  ).trim();
  const rowKey =
    identityKey ||
    (name ? normalizeRosterText(name) : "") ||
    fallbackKey;
  if (!rowKey) return null;

  const staffStatus = applyStaffRosterStatus(rosterRow, rosterLabels);
  return {
    id: `roster:${rowKey}`,
    externalId: identityKey || fallbackKey,
    name: displayName,
    rank: resolvePersonRankTitle(rosterRow) || getRosterValue(rosterRow, ["звання"]),
    unit: getRosterUnit(rosterRow) || "—",
    status: staffStatus.staffStatus,
    statusLabel: staffStatus.staffStatusLabel,
    positionTitle: getPersonFullPositionTitle(rosterRow),
    validFrom: null,
    days: null,
    plannedReturn: null,
    place: "",
    updatedAt: "",
    inStaff: options.inStaff !== false,
    inNovaStaff: options.inNovaStaff === true,
    battalion: options.battalion,
    fromEjoos: false,
    ...staffStatus,
    ...getRosterFighterStatusOverviewFields(rosterRow),
    staffSheetColumns: buildStaffSheetColumnsRecord(rosterRow),
  };
};

/** Список Огляду в режимі «Штатка» — напряму з рядків Штатки, без злиття з ООС. */
export const buildStaffOverviewRowsFromRoster = (
  rosterRows: EjournalPreviewRow[],
  rosterLabels: Record<string, string> = {},
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
  battalion: string = "ALL",
): BackendPersonnelOverviewRow[] => {
  if (!rosterRows.length) return [];

  const namedNovaRows = namedNovaRowsFromRoster(rosterRows, columns);
  const extracted = extractBchsAwayPeopleFromDbRows(rosterRows, columns);
  const result: BackendPersonnelOverviewRow[] = [];

  rosterRows.forEach((row, index) => {
    const person = extracted[index];
    const battalionLabel = rosterBattalionLabel(person, row);
    if (!battalionLabel) return;
    if (battalion !== "ALL" && battalionLabel !== battalion) return;

    const name =
      getRosterPersonName(row) || String(person?.fullName ?? "").trim();
    if (!name) return;

    const overviewRow = rosterRowToOverviewRow(row, rosterLabels, {
      requireName: true,
      inNovaStaff: namedNovaRows.has(row),
      name,
      battalion: battalionLabel,
    });
    if (overviewRow) result.push(overviewRow);
  });

  return result;
};

/** Режим «Штатка» в Огляді: усі зведені картки, які видно в Особовому складі. */
export const buildStaffOverviewRowsFromPersonnel = (
  personnelRows: EjournalPreviewRow[],
  rosterLabels: Record<string, string> = {},
  battalion: string = "ALL",
): BackendPersonnelOverviewRow[] => {
  const result: BackendPersonnelOverviewRow[] = [];

  for (const row of personnelRows) {
    if (!isLikelyPersonnelRow(row)) continue;
    const inStaff = isPersonnelInStaffRoster(row);
    if (!inStaff) continue;
    const battalionLabel = rosterBattalionLabel(undefined, row);
    if (battalion !== "ALL" && battalionLabel !== battalion) continue;

    const overviewRow = rosterRowToOverviewRow(row, rosterLabels, {
      requireName: true,
      inStaff,
      inNovaStaff: battalionLabel === "нова",
      name: getPersonDisplayName(row),
      battalion: battalionLabel,
    });
    if (overviewRow) result.push(overviewRow);
  }

  return result;
};

/** Унікальні підрозділи зі Штатки (col 2) для фільтра Огляду. */
export const collectRosterUnitOptions = (
  rosterRows: EjournalPreviewRow[],
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
  battalion: string = "ALL",
): string[] => {
  if (!rosterRows.length) return [];

  const extracted = extractBchsAwayPeopleFromDbRows(rosterRows, columns);
  const units = new Set<string>();
  rosterRows.forEach((row, index) => {
    const person = extracted[index];
    const label = rosterBattalionLabel(person, row);
    if (!label) return;
    if (battalion !== "ALL" && label !== battalion) return;
    const unit = getRosterUnit(row);
    if (unit) units.add(unit);
  });
  return [...units].sort((a, b) =>
    a.localeCompare(b, "uk", { numeric: true, sensitivity: "base" }),
  );
};

export const buildOverviewMetrics = (rows: BackendPersonnelOverviewRow[]) => ({
  total: rows.length,
  onDuty: rows.filter((row) => row.status === "ON_DUTY").length,
  businessTrip: rows.filter((row) => row.status === "BUSINESS_TRIP").length,
  leave: rows.filter((row) => row.status === "LEAVE").length,
  medical: rows.filter((row) => row.status === "MEDICAL").length,
  awol: rows.filter((row) => row.status === "AWOL").length,
  other: rows.filter(
    (row) =>
      !["ON_DUTY", "BUSINESS_TRIP", "LEAVE", "MEDICAL", "AWOL"].includes(
        String(row.status),
      ),
  ).length,
});

/** Immediate Overview first paint from the same cached roster as Personnel. */
export const buildRosterOnlyOverview = (
  rosterRows: EjournalPreviewRow[],
  rosterLabels: Record<string, string> = {},
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
  meta: { importId?: string; importName?: string } = {},
): BackendPersonnelOverview => {
  const rows = buildStaffOverviewRowsFromRoster(
    rosterRows,
    rosterLabels,
    columns,
  );
  return {
    importId: meta.importId || "local-roster-cache",
    importName: meta.importName || "Штатка · локальний кеш",
    rows,
    metrics: buildOverviewMetrics(rows),
    units: collectRosterUnitOptions(rosterRows, columns),
    critical: [],
    todayChanges: {
      total: 0,
      onDuty: 0,
      businessTrip: 0,
      leave: 0,
      medical: 0,
      awol: 0,
      other: 0,
    },
    todayUpdates: 0,
  };
};

/** Ready-to-render staff Overview produced from Personnel's merged rows. */
export const buildPersonnelStaffOverview = (
  personnelRows: EjournalPreviewRow[],
  rosterLabels: Record<string, string> = {},
  meta: { importId?: string; importName?: string } = {},
): BackendPersonnelOverview => {
  const rows = buildStaffOverviewRowsFromPersonnel(
    personnelRows,
    rosterLabels,
  );
  return {
    importId: meta.importId || "personnel-zustand-preview",
    importName: meta.importName || "Особовий склад · спільний кеш",
    rows,
    metrics: buildOverviewMetrics(rows),
    units: [...new Set(rows.map((row) => row.unit).filter(Boolean))].sort(
      (left, right) =>
        left.localeCompare(right, "uk", {
          numeric: true,
          sensitivity: "base",
        }),
    ),
    critical: [],
    todayChanges: {
      total: 0,
      onDuty: 0,
      businessTrip: 0,
      leave: 0,
      medical: 0,
      awol: 0,
      other: 0,
    },
    todayUpdates: 0,
  };
};

/** Єдині назви категорій у табличному фільтрі, незалежно від тексту в штатці. */
export const overviewStatusFilterLabel = (
  row: BackendPersonnelOverviewRow,
) => {
  if (row.status === "MEDICAL") return "Лікування";
  if (row.status === "LEAVE") return "Відпустка";
  if (row.status === "BUSINESS_TRIP") return "Відрядження";
  if (row.status === "AWOL") return "СЗЧ";
  if (row.status === "MISSING") return "Безвісти";
  if (row.status === "CAPTIVITY") return "Полон";
  return row.statusLabel;
};

export const mergeRosterRowsIntoOverview = (
  overview: BackendPersonnelOverview,
  rosterRows: EjournalPreviewRow[],
  rosterLabels: Record<string, string> = {},
  columns?: Array<{ key: string; letter?: string; originalIndex?: number }>,
): BackendPersonnelOverview => {
  if (!rosterRows.length) {
    return {
      ...overview,
      rows: overview.rows.map((row) => ({
        ...row,
        inStaff: false,
        inNovaStaff: false,
        battalion: "",
        fromEjoos: true,
      })),
    };
  }

  const namedNovaRows = namedNovaRowsFromRoster(rosterRows, columns);
  const extracted = extractBchsAwayPeopleFromDbRows(rosterRows, columns);
  const extractedNameByRow = new Map<EjournalPreviewRow, string>();
  const battalionByRow = new Map<EjournalPreviewRow, string>();
  rosterRows.forEach((row, index) => {
    const person = extracted[index];
    const fullName = String(person?.fullName ?? "").trim();
    if (fullName) extractedNameByRow.set(row, fullName);
    const battalion = rosterBattalionLabel(person, row);
    if (battalion) battalionByRow.set(row, battalion);
  });
  const rosterById = new Map<string, EjournalPreviewRow>();
  const rosterByName = new Map<string, EjournalPreviewRow[]>();
  const usedRosterRows = new Set<EjournalPreviewRow>();
  rosterRows.forEach((row) => {
    const id = getPersonExternalId(row);
    const name = getRosterPersonName(row);
    if (id) rosterById.set(id, row);
    if (name) {
      const key = normalizeRosterText(name);
      const list = rosterByName.get(key);
      if (list) list.push(row);
      else rosterByName.set(key, [row]);
    }
  });

  const takeRosterByName = (name: string) => {
    const list = rosterByName.get(normalizeRosterText(name));
    if (!list?.length) return undefined;
    const unused = list.find((row) => !usedRosterRows.has(row));
    return unused ?? list[0];
  };

  const mergedRows = overview.rows.map((row) => {
    const rosterRow =
      (row.externalId &&
        !isUnstablePersonExternalId(row.externalId) &&
        rosterById.get(row.externalId)) ||
      takeRosterByName(row.name);
    const stableFromRoster = rosterRow
      ? resolvePersonIdentityKey(rosterRow)
      : "";
    const stableFromOverview =
      row.externalId && !isUnstablePersonExternalId(row.externalId)
        ? row.externalId
        : "";
    const fingerprint = buildPersonIdentityFingerprint(row.name);
    const externalId =
      stableFromRoster || stableFromOverview || fingerprint || row.externalId;

    if (!rosterRow) {
      return {
        ...row,
        externalId,
        inStaff: false,
        inNovaStaff: false,
        battalion: "",
        fromEjoos: true,
      };
    }

    usedRosterRows.add(rosterRow);
    let rosterRank = "";
    try {
      rosterRank = resolvePersonRankTitle(rosterRow);
    } catch {
      rosterRank = "";
    }
    let positionTitle = "";
    try {
      positionTitle = getPersonFullPositionTitle(rosterRow);
    } catch {
      positionTitle = "";
    }
    const rosterUnit = getRosterUnit(rosterRow);
    return {
      ...row,
      externalId,
      inStaff: true,
      inNovaStaff: namedNovaRows.has(rosterRow),
      battalion: battalionByRow.get(rosterRow) ?? "",
      fromEjoos: true,
      ...(rosterUnit ? { unit: rosterUnit } : {}),
      ...(rosterRank ? { rank: rosterRank } : {}),
      ...(positionTitle ? { positionTitle } : {}),
      ...applyStaffRosterStatus(rosterRow, rosterLabels),
      ...getRosterFighterStatusOverviewFields(rosterRow),
      staffSheetColumns: buildStaffSheetColumnsRecord(rosterRow),
    };
  });

  const rosterOnlyRows = rosterRows
    .filter(
      (row) =>
        !usedRosterRows.has(row) &&
        (namedNovaRows.has(row) || Boolean(getRosterPersonName(row))),
    )
    .map((row) =>
      rosterRowToOverviewRow(row, rosterLabels, {
        requireName: true,
        inNovaStaff: namedNovaRows.has(row),
        name: extractedNameByRow.get(row),
        battalion: battalionByRow.get(row) ?? "",
      }),
    )
    .filter((row): row is BackendPersonnelOverviewRow => Boolean(row));

  const rows = rosterOnlyRows.length
    ? [...mergedRows, ...rosterOnlyRows]
    : mergedRows;
  return {
    ...overview,
    rows,
    metrics: buildOverviewMetrics(rows),
    units: [...new Set(rows.map((row) => row.unit).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, "uk", { numeric: true, sensitivity: "base" }),
    ),
  };
};
