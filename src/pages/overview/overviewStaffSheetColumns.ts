import type { MRT_ColumnDef } from "@/components/sci/SciDataTable";
import type { BackendPersonnelOverviewRow } from "../../api";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { STAFF_SHEET_EXPORT_COLUMN_NUMBERS } from "../excel-fill/staffSheet";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  FIGHTER_STATUS_FIELDS,
  getFighterStatusDirectValue,
  resolveFighterStatusTotalDays,
  type FighterStatusFieldKey,
} from "../personnel/fighterStatusImport";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/personnelUtils";

export type OverviewStaffSheetColumnDef = {
  id: string;
  header: string;
  columnNumber?: number;
  fighterKey?: FighterStatusFieldKey;
};

const rosterColumnId = (columnNumber: number) => `staff_${columnNumber}`;

const fighterColumnId = (key: FighterStatusFieldKey) => `staff_fighter_${key}`;

/** Уже показані окремими основними колонками таблиці Огляду. */
const OVERVIEW_PRIMARY_ROSTER_COLUMNS = new Set([
  1, 2, 7, 10, 13, 14, 21, 33,
]);
const OVERVIEW_PRIMARY_FIGHTER_FIELDS = new Set<FighterStatusFieldKey>([
  "fighter_status_direction",
  "fighter_status_exit_date",
  "fighter_status_return_date",
  "fighter_status_total_days",
  "fighter_status_value",
]);
const OVERVIEW_PRIORITY_ROSTER_COLUMNS = [31, 32] as const;

const duplicateRosterLabels = (() => {
  const counts = new Map<string, number>();
  for (const columnNumber of STAFF_SHEET_EXPORT_COLUMN_NUMBERS) {
    const label = MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber]?.trim();
    if (!label) continue;
    counts.set(label, (counts.get(label) ?? 0) + 1);
  }
  return new Set(
    [...counts.entries()]
      .filter(([, count]) => count > 1)
      .map(([label]) => label),
  );
})();

const rosterColumnHeader = (columnNumber: number) => {
  const label = MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber]?.trim() ?? "";
  if (!label) return `Колонка ${columnNumber}`;
  if (duplicateRosterLabels.has(label)) {
    return `${label} (кол. ${columnNumber})`;
  }
  return label;
};

const fighterColumnHeader = (label: string) => {
  const parts = label.split(" · ");
  return parts.length > 1 ? parts[1]! : label;
};

/** Усі колонки «1.ОС Загальний список» + «Статус бійців» для таблиці Огляду. */
export const OVERVIEW_STAFF_SHEET_COLUMN_DEFS: OverviewStaffSheetColumnDef[] = [
  ...STAFF_SHEET_EXPORT_COLUMN_NUMBERS.filter(
    (columnNumber) => !OVERVIEW_PRIMARY_ROSTER_COLUMNS.has(columnNumber),
  )
    .sort((left, right) => {
      const leftPriority = OVERVIEW_PRIORITY_ROSTER_COLUMNS.indexOf(
        left as (typeof OVERVIEW_PRIORITY_ROSTER_COLUMNS)[number],
      );
      const rightPriority = OVERVIEW_PRIORITY_ROSTER_COLUMNS.indexOf(
        right as (typeof OVERVIEW_PRIORITY_ROSTER_COLUMNS)[number],
      );
      if (leftPriority >= 0 || rightPriority >= 0) {
        if (leftPriority < 0) return 1;
        if (rightPriority < 0) return -1;
        return leftPriority - rightPriority;
      }
      return left - right;
    })
    .map((columnNumber) => ({
      id: rosterColumnId(columnNumber),
      header: rosterColumnHeader(columnNumber),
      columnNumber,
    })),
  ...FIGHTER_STATUS_FIELDS.filter(
    (field) => !OVERVIEW_PRIMARY_FIGHTER_FIELDS.has(field.key),
  ).map((field) => ({
      id: fighterColumnId(field.key),
      header: fighterColumnHeader(field.label),
      fighterKey: field.key,
    })),
];

/** Усі колонки Штатки приховані за замовчуванням — увімкнути через «Колонки». */
export const DEFAULT_OVERVIEW_STAFF_COLUMN_VISIBILITY: Record<string, boolean> =
  Object.fromEntries(
    OVERVIEW_STAFF_SHEET_COLUMN_DEFS.map((column) => [
      column.id,
      column.columnNumber === 31 || column.columnNumber === 32,
    ]),
  );

const staffSheetColumnsCache = new WeakMap<
  EjournalPreviewRow,
  Record<string, string>
>();

export const buildStaffSheetColumnsRecord = (
  rosterRow: EjournalPreviewRow | null | undefined,
): Record<string, string> => {
  if (!rosterRow) return {};
  const cached = staffSheetColumnsCache.get(rosterRow);
  if (cached) return cached;

  const record: Record<string, string> = {};
  for (const columnNumber of STAFF_SHEET_EXPORT_COLUMN_NUMBERS) {
    record[rosterColumnId(columnNumber)] = readRosterColumnValue(
      rosterRow,
      columnNumber,
    );
  }

  for (const field of FIGHTER_STATUS_FIELDS) {
    record[fighterColumnId(field.key)] =
      field.key === "fighter_status_total_days"
        ? resolveFighterStatusTotalDays(rosterRow)
        : getFighterStatusDirectValue(rosterRow, field.key);
  }

  staffSheetColumnsCache.set(rosterRow, record);
  return record;
};

const staffSheetCellValue = (
  row: BackendPersonnelOverviewRow,
  columnId: string,
) => row.staffSheetColumns?.[columnId]?.trim() || "—";

export const buildOverviewStaffSheetColumnDefs =
  (): Array<MRT_ColumnDef<BackendPersonnelOverviewRow>> =>
    OVERVIEW_STAFF_SHEET_COLUMN_DEFS.map((column) => ({
      id: column.id,
      header: column.header,
      columnMenuLabel: column.header,
      size: 170,
      accessorFn: (row) => staffSheetCellValue(row, column.id),
      exportValue: (row) => row.staffSheetColumns?.[column.id]?.trim() || "",
    }));
