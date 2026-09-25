import type { BackendPersonnelRosterLatest } from "../../api";
import { loadSharedRosterLatest } from "../../data/sharedAppData";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { mapRosterLatestToPreviewRows, readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  countStaffSheetPersonsInArchive,
  countStaffSheetPersonsInRoster,
} from "../anketa-data/staffSheetPreview";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/morningGeneralListColumnLabels";
import { isPersonnelFromArchive, getRosterPersonName } from "../personnel/personnelRosterMerge";
import { isLikelyPersonnelRow } from "../personnel/personnelUtils";

const MAX_ROSTER_COLUMN = 45;

export type PersonnelV2StaffCell = {
  column: number;
  key: string;
  label: string;
  value: string;
};

export type PersonnelV2StaffPerson = {
  id: string;
  rowNumber: number;
  name: string;
  inArchive: boolean;
  cells: PersonnelV2StaffCell[];
};

export type PersonnelV2StaffSheet = {
  importId: string;
  importName: string;
  sourceFileName: string;
  sheetName: string;
  createdAt: string;
  rowCount: number;
  inStaff: number;
  inArchive: number;
  people: PersonnelV2StaffPerson[];
};

/** Непорожні колонки штатки 1–45. Іменовані дублі (`піб`, `взвод`) не додаємо. */
export const collectPersonnelV2StaffCells = (
  row: EjournalPreviewRow,
): PersonnelV2StaffCell[] => {
  const cells: PersonnelV2StaffCell[] = [];

  for (let column = 1; column <= MAX_ROSTER_COLUMN; column += 1) {
    const value = readRosterColumnValue(row, column).trim();
    if (!value) continue;
    cells.push({
      column,
      key: `column_${column}`,
      label: MORNING_GENERAL_LIST_COLUMN_LABELS[column] || `Колонка ${column}`,
      value,
    });
  }

  return cells;
};

const personFromRow = (
  row: EjournalPreviewRow,
  index: number,
): PersonnelV2StaffPerson | null => {
  if (
    !isLikelyPersonnelRow({
      ...row,
      __dbRowId: String(row.__dbRowId ?? `staff:${row.__rowNumber ?? index}`),
    })
  ) {
    return null;
  }
  const name = getRosterPersonName(row);
  if (!name) return null;
  return {
    id: String(row.__dbRowId ?? `staff:${row.__rowNumber ?? index}`),
    rowNumber: Number(row.__rowNumber) || 0,
    name,
    inArchive: isPersonnelFromArchive(row),
    cells: collectPersonnelV2StaffCells(row),
  };
};

export const buildPersonnelV2StaffSheet = (
  latest: BackendPersonnelRosterLatest,
): PersonnelV2StaffSheet => {
  const rows = mapRosterLatestToPreviewRows(latest);
  const people = rows
    .map((row, index) => personFromRow(row, index))
    .filter((person): person is PersonnelV2StaffPerson => Boolean(person))
    .sort((left, right) => left.rowNumber - right.rowNumber);

  return {
    importId: latest.importId,
    importName: latest.importName,
    sourceFileName: String(latest.sourceFileName ?? "").trim(),
    sheetName: latest.sheet?.name ?? "Загальний список",
    createdAt: latest.createdAt,
    rowCount: latest.sheet?.rowCount ?? latest.rows.length,
    inStaff: countStaffSheetPersonsInRoster(rows),
    inArchive: countStaffSheetPersonsInArchive(rows),
    people,
  };
};

/** Актуальна штатка з БД (`/ejournals/personnel/roster/latest`). */
export const loadPersonnelV2StaffSheet = async (options?: {
  force?: boolean;
  signal?: AbortSignal;
}): Promise<PersonnelV2StaffSheet> => {
  const latest = await loadSharedRosterLatest(options);
  if (!latest?.rows?.length) {
    throw new Error("У базі немає штатки. Спочатку імпортуйте файл «Штатка».");
  }
  return buildPersonnelV2StaffSheet(latest);
};
