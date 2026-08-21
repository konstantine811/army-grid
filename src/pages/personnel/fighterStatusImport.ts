import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import { hasRowData, valueToDisplay } from "../../excelRoundTrip";
import { formatValueForDisplay } from "../../shared/format";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { buildImportColumns } from "../ejournal/ejournalUtils";

export const FIGHTER_STATUS_FIELD_PREFIX = "fighter_status_";

export const FIGHTER_STATUS_FIELDS = [
  {
    key: "fighter_status_direction",
    label: "Статус бійців · Напрямок",
  },
  {
    key: "fighter_status_entry_date",
    label: "Статус бійців · Дата заходу",
  },
  {
    key: "fighter_status_exit_date",
    label: "Статус бійців · Дата виходу",
  },
  {
    key: "fighter_status_return_date",
    label: "Статус бійців · Дата повернення",
  },
  {
    key: "fighter_status_total_days",
    label: "Статус бійців · Днів",
  },
  {
    key: "fighter_status_value",
    label: "Статус бійців · Статус (200/300/500)",
  },
  {
    key: "fighter_status_weapon",
    label: "Статус бійців · Зброя",
  },
  {
    key: "fighter_status_communication",
    label: "Статус бійців · Засоби зв'язку",
  },
  {
    key: "fighter_status_note",
    label: "Статус бійців · Примітка",
  },
] as const;

export type FighterStatusFieldKey = (typeof FIGHTER_STATUS_FIELDS)[number]["key"];

export const FIGHTER_STATUS_FIELD_KEYS = {
  direction: "fighter_status_direction",
  entryDate: "fighter_status_entry_date",
  exitDate: "fighter_status_exit_date",
  returnDate: "fighter_status_return_date",
  totalDays: "fighter_status_total_days",
  status: "fighter_status_value",
  weapon: "fighter_status_weapon",
  communication: "fighter_status_communication",
  note: "fighter_status_note",
} as const satisfies Record<string, FighterStatusFieldKey>;

const normalizeStatusHeader = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0])
    .replace(/[ʼ’']/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

export const normalizeRosterMatchText = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0])
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const getStatusColumnIndex = (
  columns: Array<{ label: string; key: string; order: number }>,
  parts: string[],
) =>
  columns.findIndex((column) => {
    const label = normalizeStatusHeader(column.label || column.key);
    return parts.every((part) => label.includes(part));
  });

const readStatusCell = (row: { values: unknown[] }, index: number) =>
  index >= 0
    ? valueToDisplay(row.values[index] as Parameters<typeof valueToDisplay>[0]).trim()
    : "";

const resolveFighterStatusColumnIndexes = (
  columns: Array<{ label: string; key: string; order: number }>,
) => {
  const directionIndex = getStatusColumnIndex(columns, ["напрям"]);
  const nameIndex = getStatusColumnIndex(columns, ["прізвище"]);
  const callSignIndex = getStatusColumnIndex(columns, ["позив"]);
  const exitDateIndex = getStatusColumnIndex(columns, ["дата", "вих"]);
  const returnDateIndex =
    getStatusColumnIndex(columns, ["дата", "поверн"]) >= 0
      ? getStatusColumnIndex(columns, ["дата", "поверн"])
      : getStatusColumnIndex(columns, ["200", "300"]);
  const entryDateIndex =
    getStatusColumnIndex(columns, ["дата", "вход"]) >= 0
      ? getStatusColumnIndex(columns, ["дата", "вход"])
      : getStatusColumnIndex(columns, ["дата", "заход"]);
  const weaponIndex = getStatusColumnIndex(columns, ["збро"]);
  const communicationIndex = getStatusColumnIndex(columns, ["зв", "яз"]);
  const noteIndex = getStatusColumnIndex(columns, ["прим"]);

  const statusCandidates = columns
    .map((column, index) => ({
      index,
      label: normalizeStatusHeader(column.label || column.key),
    }))
    .filter(
      ({ label, index }) =>
        label.includes("статус") && (exitDateIndex < 0 || index > exitDateIndex),
    );
  const statusIndex = statusCandidates.length
    ? statusCandidates[statusCandidates.length - 1].index
    : -1;

  const totalDaysIndex = (() => {
    const byLabel = getStatusColumnIndex(columns, ["дн"]);
    if (byLabel >= 0) return byLabel;

    if (returnDateIndex >= 0) {
      const candidate = returnDateIndex + 1;
      const header = normalizeStatusHeader(
        columns[candidate]?.label || columns[candidate]?.key || "",
      );
      if (!header || header.length <= 2) return candidate;
    }

    if (statusIndex > 0) return statusIndex - 1;
    return -1;
  })();

  return {
    directionIndex,
    nameIndex,
    callSignIndex,
    entryDateIndex,
    exitDateIndex,
    returnDateIndex,
    totalDaysIndex,
    statusIndex,
    weaponIndex,
    communicationIndex,
    noteIndex,
  };
};

export const buildFighterStatusValues = (
  row: { values: unknown[] },
  indexes: ReturnType<typeof resolveFighterStatusColumnIndexes>,
) => ({
  fighter_status_direction: readStatusCell(row, indexes.directionIndex),
  fighter_status_entry_date: readStatusCell(row, indexes.entryDateIndex),
  fighter_status_exit_date: readStatusCell(row, indexes.exitDateIndex),
  fighter_status_return_date: readStatusCell(row, indexes.returnDateIndex),
  fighter_status_total_days: readStatusCell(row, indexes.totalDaysIndex),
  fighter_status_value: readStatusCell(row, indexes.statusIndex),
  fighter_status_weapon: readStatusCell(row, indexes.weaponIndex),
  fighter_status_communication: readStatusCell(row, indexes.communicationIndex),
  fighter_status_note: readStatusCell(row, indexes.noteIndex),
});

export const buildFighterStatusAdditions = (statusSheet: ExcelSheetSnapshot) => {
  const columns = buildImportColumns(statusSheet);
  const indexes = resolveFighterStatusColumnIndexes(columns);

  if (indexes.nameIndex < 0) return new Map<string, Record<string, unknown>>();

  const additions = new Map<string, Record<string, unknown>>();
  statusSheet.rows
    .filter((row) => hasRowData(row.values))
    .forEach((row) => {
      const name = readStatusCell(row, indexes.nameIndex);
      if (!name) return;

      const values = buildFighterStatusValues(row, indexes);
      const hasStatusData = Object.values(values).some(Boolean);
      if (!hasStatusData) return;

      const nameKey = normalizeRosterMatchText(name);
      additions.set(`name:${nameKey}`, values);

      const callSign = readStatusCell(row, indexes.callSignIndex);
      if (callSign) {
        additions.set(
          `name-call:${nameKey}:${normalizeRosterMatchText(callSign)}`,
          values,
        );
      }
    });

  return additions;
};

const getRosterValue = (rowValues: Record<string, unknown>, keyParts: string[]) => {
  const key = Object.keys(rowValues).find((item) =>
    keyParts.every((part) => item.toLocaleLowerCase("uk-UA").includes(part)),
  );
  return key
    ? valueToDisplay(rowValues[key] as Parameters<typeof valueToDisplay>[0]).trim()
    : "";
};

export const findFighterStatusAddition = (
  rowValues: Record<string, unknown>,
  additions: Map<string, Record<string, unknown>>,
) => {
  const name =
    getRosterValue(rowValues, ["піб"]) || getRosterValue(rowValues, ["прізвище"]);
  if (!name) return null;

  const nameKey = normalizeRosterMatchText(name);
  const callSign = getRosterValue(rowValues, ["позив"]);
  if (callSign) {
    const byNameAndCallSign = additions.get(
      `name-call:${nameKey}:${normalizeRosterMatchText(callSign)}`,
    );
    if (byNameAndCallSign) return byNameAndCallSign;
  }

  return additions.get(`name:${nameKey}`) ?? null;
};

export const findFighterStatusSheet = (
  sheets: ExcelSheetSnapshot[],
) =>
  sheets.find((sheet) =>
    /статус\s*бійц/i.test(sheet.sheetName.replace(/\s+/g, " ").trim()),
  );

export const getFighterStatusDirectValue = (
  row: EjournalPreviewRow | null | undefined,
  key: FighterStatusFieldKey,
) => {
  if (!row) return "";
  const rosterKey = `roster__${key}`;
  const direct = row[key] ?? row[rosterKey];
  return formatValueForDisplay(direct);
};

const parseFighterStatusDate = (value: string) => {
  const match = String(value ?? "")
    .trim()
    .match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += year >= 50 ? 1900 : 2000;
  if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900) {
    return null;
  }

  const date = new Date(year, month - 1, day);
  return Number.isNaN(date.getTime()) ? null : date;
};

export const diffFighterStatusDays = (exitDate: string, returnDate: string) => {
  const from = parseFighterStatusDate(exitDate);
  const to = parseFighterStatusDate(returnDate);
  if (!from || !to) return "";

  const days = Math.round((to.getTime() - from.getTime()) / 86400000);
  return Number.isFinite(days) ? String(days) : "";
};

/** Формат рапорту УБД: «з 14.06.2026-11.08.2026». */
export const buildFighterTaskPeriodText = (
  row: EjournalPreviewRow | null | undefined,
) => {
  const exitDate = getFighterStatusDirectValue(row, "fighter_status_exit_date");
  const returnDate = getFighterStatusDirectValue(row, "fighter_status_return_date");
  const entryDate = getFighterStatusDirectValue(row, "fighter_status_entry_date");
  const from = exitDate || entryDate;
  const to = returnDate;

  if (from && to) return `з ${from}-${to}`;
  if (from) return `з ${from}`;
  return "";
};

export const getFighterTaskPlace = (
  row: EjournalPreviewRow | null | undefined,
) => getFighterStatusDirectValue(row, "fighter_status_direction");

export const extractFighterStatusFieldRows = (
  row: EjournalPreviewRow | null | undefined,
  labels: Record<string, string> = {},
) =>
  FIGHTER_STATUS_FIELDS.map((field) => ({
    key: field.key,
    label: labels[field.key] || field.label,
    value: getFighterStatusDirectValue(row, field.key),
  })).filter((field) => field.value);

export const getRosterFighterStatusOverviewFields = (row: EjournalPreviewRow) => {
  const fighterExitDate = getFighterStatusDirectValue(
    row,
    "fighter_status_exit_date",
  );
  const fighterReturnDate = getFighterStatusDirectValue(
    row,
    "fighter_status_return_date",
  );
  const storedDays = getFighterStatusDirectValue(
    row,
    "fighter_status_total_days",
  );

  return {
    fighterDirection: getFighterStatusDirectValue(row, "fighter_status_direction"),
    fighterEntryDate: getFighterStatusDirectValue(row, "fighter_status_entry_date"),
    fighterExitDate,
    fighterReturnDate,
    fighterTotalDays:
      storedDays || diffFighterStatusDays(fighterExitDate, fighterReturnDate),
    fighterStatus: getFighterStatusDirectValue(row, "fighter_status_value"),
    fighterWeapon: getFighterStatusDirectValue(row, "fighter_status_weapon"),
    fighterCommunication: getFighterStatusDirectValue(
      row,
      "fighter_status_communication",
    ),
    fighterNote: getFighterStatusDirectValue(row, "fighter_status_note"),
  };
};
