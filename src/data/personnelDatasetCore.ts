import type {
  BackendEjournalImportSheet,
  BackendPersonnelRosterLatest,
  BackendPersonnelRosterVersion,
} from "../api";
import { parseDbColumns } from "../pages/ejournal/ejournalUtils";
import type {
  DbPreviewState,
  EjournalColumn,
  EjournalPreviewRow,
} from "../pages/ejournal/ejournalTypes";
import {
  buildRosterOnlyPreviewState,
  mergeRosterRowsIntoPreview,
} from "../pages/personnel/personnelRosterMerge";
import {
  findEjournalPersonnelSheet,
  getPersonDisplayName,
  getPersonExternalId,
  normalizePersonBirthKey,
  resolveMorningGeneralListColumnLabel,
  resolvePersonBirthDate,
} from "../pages/personnel/personnelUtils";
import { fillDownRosterUnitRows } from "../pages/personnel/rosterRowFill";

export type PersonnelDatasetVersion = {
  oosSheetId: string;
  oosStamp: string;
  rosterImportId: string;
  rosterSheetUpdatedAt: string;
  rosterRowCount: number;
};

export type PersonnelDataset = {
  rows: EjournalPreviewRow[];
  sheet: BackendEjournalImportSheet | null;
  columns: EjournalColumn[];
  total: number;
  rosterRows: EjournalPreviewRow[];
  rosterLabels: Record<string, string>;
  rosterColumns: EjournalColumn[];
  rosterUpdatedAt: string | null;
  version: PersonnelDatasetVersion;
  fingerprint: string;
  mergedAt: number;
  complete: boolean;
};

const sheetStamp = (sheet: BackendEjournalImportSheet | null | undefined) =>
  sheet
    ? `${sheet.updatedAt ?? ""}|${sheet.rowCount}|${sheet.columnCount}`
    : "";

export const buildPersonnelDatasetVersion = (
  sheet: BackendEjournalImportSheet | null | undefined,
  roster: BackendPersonnelRosterLatest | null | undefined,
): PersonnelDatasetVersion => ({
  oosSheetId: sheet?.id ?? "",
  oosStamp: sheetStamp(sheet),
  rosterImportId: roster?.importId ?? "",
  rosterSheetUpdatedAt: roster?.sheet?.updatedAt ?? "",
  rosterRowCount: roster?.sheet?.rowCount ?? roster?.rows?.length ?? 0,
});

/** Lightweight fingerprint input — no 2000+ roster row payload. */
export const buildPersonnelDatasetVersionFromRosterVersion = (
  sheet: BackendEjournalImportSheet | null | undefined,
  rosterVersion: BackendPersonnelRosterVersion | null | undefined,
): PersonnelDatasetVersion => ({
  oosSheetId: sheet?.id ?? "",
  oosStamp: sheetStamp(sheet),
  rosterImportId: rosterVersion?.importId ?? "",
  rosterSheetUpdatedAt: rosterVersion?.sheetUpdatedAt ?? "",
  rosterRowCount: rosterVersion?.rowCount ?? 0,
});

const DATASET_FINGERPRINT_SEP = "\u001f";

export const personnelDatasetFingerprint = (
  version: PersonnelDatasetVersion,
) =>
  [
    version.oosSheetId,
    version.oosStamp,
    version.rosterImportId,
    version.rosterSheetUpdatedAt,
    version.rosterRowCount,
  ].join(DATASET_FINGERPRINT_SEP);

export const rosterRowsFromPersonnelLatest = (
  latest: BackendPersonnelRosterLatest | null | undefined,
) => {
  if (!latest?.sheet || !Array.isArray(latest.rows)) {
    return [] as EjournalPreviewRow[];
  }
  return fillDownRosterUnitRows(
    latest.rows.map((row, rosterOrder) => ({
      __dbRowId: row.id,
      __rowNumber: row.excelRowNumber,
      __rosterOrder: rosterOrder,
      ...(row.values &&
      typeof row.values === "object" &&
      !Array.isArray(row.values)
        ? row.values
        : {}),
    })) as EjournalPreviewRow[],
  );
};

export const sortPersonnelRowsByRosterOrder = (
  rows: EjournalPreviewRow[],
): EjournalPreviewRow[] =>
  rows
    .map((row, sourceOrder) => ({ row, sourceOrder }))
    .sort((left, right) => {
      const leftOrder = Number(left.row.__rosterOrder);
      const rightOrder = Number(right.row.__rosterOrder);
      const leftInRoster = Number.isFinite(leftOrder);
      const rightInRoster = Number.isFinite(rightOrder);
      if (leftInRoster && rightInRoster) {
        return leftOrder - rightOrder || left.sourceOrder - right.sourceOrder;
      }
      if (leftInRoster) return -1;
      if (rightInRoster) return 1;
      return left.sourceOrder - right.sourceOrder;
    })
    .map(({ row }) => row);

const normalizePersonnelDatasetName = (value: string) =>
  value
    .replace(/\(\s*\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s*р\.?\s*н\.?\s*\)/gi, "")
    .replace(/[ʼ’']/g, "'")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const personnelDatasetDedupeKey = (row: EjournalPreviewRow) => {
  const name = normalizePersonnelDatasetName(getPersonDisplayName(row));
  const birthDate = normalizePersonBirthKey(resolvePersonBirthDate(row));
  if (name && birthDate) return `name-birth:${name}|${birthDate}`;
  const externalId = getPersonExternalId(row).trim();
  return externalId ? `id:${externalId}` : "";
};

const hasDatasetValue = (value: unknown) =>
  value != null && String(value).trim() !== "";

const mergeDuplicatePersonnelRows = (
  primary: EjournalPreviewRow,
  duplicate: EjournalPreviewRow,
) => {
  const merged = { ...primary };
  Object.entries(duplicate).forEach(([key, value]) => {
    if (!hasDatasetValue(merged[key]) && hasDatasetValue(value)) {
      merged[key] = value;
    }
  });
  const primaryOrder = Number(primary.__rosterOrder);
  const duplicateOrder = Number(duplicate.__rosterOrder);
  if (
    Number.isFinite(duplicateOrder) &&
    (!Number.isFinite(primaryOrder) || duplicateOrder < primaryOrder)
  ) {
    merged.__rosterOrder = duplicateOrder;
  }
  return merged;
};

export const dedupePersonnelDatasetRows = (
  rows: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  const result: EjournalPreviewRow[] = [];
  const indexByKey = new Map<string, number>();
  rows.forEach((row) => {
    const key = personnelDatasetDedupeKey(row);
    if (!key) {
      result.push(row);
      return;
    }
    const existingIndex = indexByKey.get(key);
    if (existingIndex == null) {
      indexByKey.set(key, result.length);
      result.push(row);
      return;
    }
    result[existingIndex] = mergeDuplicatePersonnelRows(
      result[existingIndex]!,
      row,
    );
  });
  return result;
};

const rosterMetadata = (
  latest: BackendPersonnelRosterLatest | null | undefined,
) => {
  const columns = parseDbColumns(latest?.sheet?.columns);
  return {
    rosterColumns: columns,
    rosterLabels: Object.fromEntries(
      columns.map((column) => [
        column.key,
        column.label?.trim() ||
          resolveMorningGeneralListColumnLabel(column.key) ||
          column.key,
      ]),
    ),
    rosterUpdatedAt:
      latest?.sheet?.updatedAt ?? latest?.createdAt ?? null,
  };
};

/** Same merge as the worker, but synchronous — for Node/server bundle. */
export const buildPersonnelDatasetSync = (
  preview: DbPreviewState | null,
  roster: BackendPersonnelRosterLatest | null,
  version?: PersonnelDatasetVersion,
): PersonnelDataset => {
  const resolvedVersion =
    version ?? buildPersonnelDatasetVersion(preview?.sheet ?? null, roster);
  const rosterRows = rosterRowsFromPersonnelLatest(roster);
  const fallback = buildRosterOnlyPreviewState(rosterRows, roster?.sheet ?? null);
  const base = preview ?? fallback;
  const mergedRows =
    base && rosterRows.length
      ? mergeRosterRowsIntoPreview(base, rosterRows)
      : (base?.rows ?? []);
  const rows = dedupePersonnelDatasetRows(
    sortPersonnelRowsByRosterOrder(mergedRows),
  );
  const metadata = rosterMetadata(roster);
  return {
    rows,
    sheet: base?.sheet ?? roster?.sheet ?? null,
    columns: base?.columns ?? [],
    total: rows.length,
    rosterRows,
    ...metadata,
    version: resolvedVersion,
    fingerprint: personnelDatasetFingerprint(resolvedVersion),
    mergedAt: Date.now(),
    complete: true,
  };
};

export { findEjournalPersonnelSheet };
