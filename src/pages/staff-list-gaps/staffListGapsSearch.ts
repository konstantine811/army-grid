import type { StaffListColumn, StaffListRow } from "./staffListGapsParse";
import {
  defaultStaffListGapColumnIds,
  normalizeStaffListHeader,
  type StaffListSnapshot,
} from "./staffListGapsParse";
import {
  isAnketaDashOnlyEmptyValue,
  isReplaceableAnketaMissingValue,
} from "../anketa-data/anketaGaps";

export type StaffListEmptyCell = {
  rowId: string;
  columnId: string;
  rowNumber: number;
  columnIndex: number;
  header: string;
  pib: string;
};

export type StaffListGapSearchOptions = {
  skipKeys?: Set<string>;
};

export const staffListGapSkipKey = (cell: StaffListEmptyCell) =>
  `${cell.rowId}:${cell.columnId}`;

export const normalizeStaffListCellValue = (value: unknown) =>
  String(value ?? "")
    .replace(/[\u00a0\u200b-\u200d\ufeff]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const isStaffListCellEmpty = (row: StaffListRow, columnId: string) => {
  const normalized = normalizeStaffListCellValue(row.values[columnId]);
  if (!normalized) return true;
  if (isAnketaDashOnlyEmptyValue(normalized)) return true;
  return isReplaceableAnketaMissingValue(normalized);
};

export const countStaffListEmptyCells = (
  rows: StaffListRow[],
  columnIds: string[],
  options: StaffListGapSearchOptions = {},
) => {
  const skipKeys = options.skipKeys ?? new Set<string>();
  let count = 0;
  for (const row of rows) {
    for (const columnId of columnIds) {
      if (skipKeys.has(`${row.__rowId}:${columnId}`)) continue;
      if (isStaffListCellEmpty(row, columnId)) count += 1;
    }
  }
  return count;
};

export const summarizeStaffListGaps = (
  rows: StaffListRow[],
  columnIds: string[],
) => {
  let emptyCells = 0;
  const personsWithGaps = new Set<string>();
  for (const row of rows) {
    let rowHasGap = false;
    for (const columnId of columnIds) {
      if (!isStaffListCellEmpty(row, columnId)) continue;
      emptyCells += 1;
      rowHasGap = true;
    }
    if (rowHasGap) personsWithGaps.add(row.__rowId);
  }
  return {
    totalRows: rows.length,
    columns: columnIds.length,
    emptyCells,
    personsWithGaps: personsWithGaps.size,
  };
};

export const buildStaffListFocusedCell = (
  row: StaffListRow,
  column: StaffListColumn,
): StaffListEmptyCell => ({
  rowId: row.__rowId,
  columnId: column.id,
  rowNumber: row.__rowNumber,
  columnIndex: column.index,
  header: column.label,
  pib: row.pib,
});

export const findNextStaffListEmptyCell = (
  rows: StaffListRow[],
  columns: StaffListColumn[],
  columnIds: string[],
  current: StaffListEmptyCell | null,
  options: StaffListGapSearchOptions = {},
): StaffListEmptyCell | null => {
  if (!rows.length || !columnIds.length) return null;
  const skipKeys = options.skipKeys ?? new Set<string>();
  const columnById = new Map(columns.map((column) => [column.id, column]));
  const orderedColumnIds = columnIds
    .filter((id) => columnById.has(id))
    .sort(
      (left, right) =>
        (columnById.get(left)?.index ?? 0) - (columnById.get(right)?.index ?? 0),
    );
  if (!orderedColumnIds.length) return null;

  const startRowIndex = current
    ? rows.findIndex((row) => row.__rowId === current.rowId)
    : -1;
  const startColumnIndex = current
    ? orderedColumnIds.indexOf(current.columnId)
    : -1;

  for (let rowOffset = 0; rowOffset < rows.length; rowOffset += 1) {
    const rowIndex =
      startRowIndex < 0
        ? rowOffset
        : (startRowIndex + rowOffset) % rows.length;
    const row = rows[rowIndex]!;
    const columnStart =
      rowOffset === 0 && startColumnIndex >= 0 ? startColumnIndex + 1 : 0;

    for (
      let columnOffset = columnStart;
      columnOffset < orderedColumnIds.length;
      columnOffset += 1
    ) {
      const columnId = orderedColumnIds[columnOffset]!;
      const key = `${row.__rowId}:${columnId}`;
      if (skipKeys.has(key)) continue;
      if (!isStaffListCellEmpty(row, columnId)) continue;
      const column = columnById.get(columnId);
      if (!column) continue;
      if (
        current &&
        rowOffset === 0 &&
        columnOffset <= startColumnIndex &&
        rowIndex === startRowIndex
      ) {
        continue;
      }
      return buildStaffListFocusedCell(row, column);
    }
  }

  return null;
};

export const STAFF_LIST_GAP_COLUMNS_STORAGE_KEY = "staff-list-gaps-columns:v2";
const LEGACY_STAFF_LIST_GAP_COLUMNS_STORAGE_KEY = "staff-list-gaps-columns";

const isLegacyPhoneOnlyGapColumns = (
  snapshot: StaffListSnapshot,
  stored: string[],
  defaults: string[],
) => {
  if (stored.length !== 1 || defaults.length <= 1) return false;
  const column = snapshot.columns.find((item) => item.id === stored[0]);
  if (!column) return false;
  return /телефон|phone|моб/i.test(normalizeStaffListHeader(column.label));
};

/** Збережений вибір колонок або розширені дефолти (телефон + дата/місце народження). */
export const resolveStaffListGapColumnIds = (
  snapshot: StaffListSnapshot,
  stored: string[] = readStaffListGapColumnIds(),
) => {
  const defaults = defaultStaffListGapColumnIds(snapshot);
  const validStored = stored.filter((id) =>
    snapshot.columns.some((column) => column.id === id),
  );
  if (!validStored.length) return defaults;
  if (isLegacyPhoneOnlyGapColumns(snapshot, validStored, defaults)) {
    return defaults;
  }
  return validStored.sort((left, right) => {
    const leftIndex =
      snapshot.columns.find((column) => column.id === left)?.index ?? 0;
    const rightIndex =
      snapshot.columns.find((column) => column.id === right)?.index ?? 0;
    return leftIndex - rightIndex;
  });
};

export const readStaffListGapColumnIds = (): string[] => {
  try {
    const raw =
      sessionStorage.getItem(STAFF_LIST_GAP_COLUMNS_STORAGE_KEY) ??
      sessionStorage.getItem(LEGACY_STAFF_LIST_GAP_COLUMNS_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed)
      ? parsed.filter((item): item is string => typeof item === "string")
      : [];
  } catch {
    return [];
  }
};

export const writeStaffListGapColumnIds = (columnIds: string[]) => {
  try {
    sessionStorage.setItem(
      STAFF_LIST_GAP_COLUMNS_STORAGE_KEY,
      JSON.stringify(columnIds),
    );
  } catch {
    /* private mode */
  }
};
