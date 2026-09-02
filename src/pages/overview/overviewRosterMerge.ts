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
  getPersonExternalId,
  getPersonFullPositionTitle,
  isUnstablePersonExternalId,
  resolvePersonIdentityKey,
  resolvePersonRankTitle,
  resolvePersonRosterStatus,
} from "../personnel/personnelUtils";
import {
  getRosterPersonName,
  getRosterValue,
  isNovaRosterRow,
} from "../personnel/personnelRosterMerge";
import {
  getRosterFighterStatusOverviewFields,
  normalizeRosterMatchText,
} from "../personnel/fighterStatusImport";

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

const isNovaExtractedPerson = (
  person: { battalion?: string } | undefined,
  row: EjournalPreviewRow,
) => rosterBattalionLabel(person, row) === "нова" || isNovaRosterRow(row);

const sortBattalionLabels = (labels: string[]) =>
  [...labels].sort((left, right) => {
    if (left === "нова") return -1;
    if (right === "нова") return 1;
    if (left === "стара") return -1;
    if (right === "стара") return 1;
    return left.localeCompare(right, "uk", { numeric: true, sensitivity: "base" });
  });

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
    unit:
      getRosterValue(rosterRow, ["перебування"]) ||
      getRosterValue(rosterRow, ["підрозділ"]) ||
      "—",
    status: staffStatus.staffStatus,
    statusLabel: staffStatus.staffStatusLabel,
    positionTitle: getPersonFullPositionTitle(rosterRow),
    validFrom: null,
    days: null,
    plannedReturn: null,
    place: "",
    updatedAt: "",
    inStaff: true,
    inNovaStaff: options.inNovaStaff === true,
    battalion: options.battalion,
    fromEjoos: false,
    ...staffStatus,
    ...getRosterFighterStatusOverviewFields(rosterRow),
  };
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
    return {
      ...row,
      externalId,
      inStaff: true,
      inNovaStaff: namedNovaRows.has(rosterRow),
      battalion: battalionByRow.get(rosterRow) ?? "",
      fromEjoos: true,
      ...(rosterRank ? { rank: rosterRank } : {}),
      ...(positionTitle ? { positionTitle } : {}),
      ...applyStaffRosterStatus(rosterRow, rosterLabels),
      ...getRosterFighterStatusOverviewFields(rosterRow),
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
