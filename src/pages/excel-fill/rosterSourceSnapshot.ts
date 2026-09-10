import type { BackendPersonnelRosterLatest } from "../../api";
import type {
  CellValue,
  ExcelRow,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/morningGeneralListColumnLabels";

const MAX_ROSTER_COLUMN = 45;

const cellText = (value: unknown) => String(value ?? "").trim();

const normalizeKey = (value: string) =>
  value
    .replace(/[ʼ’']/g, "")
    .replace(/[()]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const labelHasClarification = (normalized: string) =>
  normalized.includes("уточн");

const ROSTER_COLUMN_KEY_ALIASES: Record<number, readonly string[]> = {
  31: ["column_31", "місце_перебування"],
  35: ["column_35", "місце_перебування_уточнення"],
  40: ["column_40", "місце_перебування_2"],
};

export const mapRosterLatestToPreviewRows = (
  latest: BackendPersonnelRosterLatest | null | undefined,
): EjournalPreviewRow[] => {
  if (!latest?.sheet) return [];
  return latest.rows.map((row) => ({
    __dbRowId: row.id,
    __rowNumber: row.excelRowNumber ?? undefined,
    ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
      ? row.values
      : {}),
  })) as EjournalPreviewRow[];
};

/** Read roster cell by 1-based «Загальний список» column index. */
export const readRosterColumnValue = (
  row: EjournalPreviewRow | null | undefined,
  columnNumber: number,
): string => {
  if (!row) return "";
  const direct = cellText(row[`column_${columnNumber}`]);
  if (direct) return direct;
  const rosterPrefixed = cellText(row[`roster__column_${columnNumber}`]);
  if (rosterPrefixed) return rosterPrefixed;

  const aliases = ROSTER_COLUMN_KEY_ALIASES[columnNumber];
  if (aliases) {
    for (const alias of aliases) {
      if (Object.prototype.hasOwnProperty.call(row, alias)) {
        return cellText(row[alias]);
      }
    }
    if (columnNumber === 35 || columnNumber === 40) {
      return "";
    }
  }

  const label = MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber];
  if (!label) return "";

  const wanted = normalizeKey(label);
  const wantedClarification = labelHasClarification(wanted);
  for (const [key, value] of Object.entries(row)) {
    if (key.startsWith("__")) continue;
    const bare = key.replace(/^roster__/i, "");
    if (/^column_\d+$/i.test(bare)) continue;
    const keyNorm = normalizeKey(bare).replace(/_/g, " ");
    if (keyNorm === wanted) {
      return cellText(value);
    }
  }
  for (const [key, value] of Object.entries(row)) {
    if (key.startsWith("__")) continue;
    const bare = key.replace(/^roster__/i, "");
    if (/^column_\d+$/i.test(bare)) continue;
    const keyNorm = normalizeKey(bare).replace(/_/g, " ");
    const keyClarification = labelHasClarification(keyNorm);
    if (wantedClarification !== keyClarification) continue;
    if (keyNorm.includes(wanted) || wanted.includes(keyNorm)) {
      const text = cellText(value);
      if (text) return text;
    }
  }
  return "";
};

const rowToValuesArray = (row: EjournalPreviewRow): CellValue[] => {
  const values: CellValue[] = new Array(MAX_ROSTER_COLUMN).fill("");
  for (let column = 1; column <= MAX_ROSTER_COLUMN; column += 1) {
    values[column - 1] = readRosterColumnValue(row, column);
  }
  return values;
};

/** Build Excel-like snapshot for fill / morning-report analyzers from personnel roster. */
export const rosterRowsToSourceSnapshot = (
  rows: EjournalPreviewRow[],
  label: string,
): ExcelWorkbookSnapshot | null => {
  if (!rows.length) return null;

  const headerRow: CellValue[] = new Array(MAX_ROSTER_COLUMN).fill("");
  for (const [columnNumber, columnLabel] of Object.entries(
    MORNING_GENERAL_LIST_COLUMN_LABELS,
  )) {
    headerRow[Number(columnNumber) - 1] = columnLabel;
  }

  const excelRows: ExcelRow[] = rows.map((row, index) => ({
    id: String(row.__dbRowId ?? index),
    excelRowNumber: Number(row.__rowNumber) || index + 2,
    values: rowToValuesArray(row),
    source: "merged" as const,
  }));

  const columnIndexes = Array.from({ length: MAX_ROSTER_COLUMN }, (_, index) => index);
  const sheet: ExcelSheetSnapshot = {
    sheetIndex: 0,
    sheetName: "1.ОС Загальний список",
    rawRows: [],
    headerRows: [headerRow],
    rows: excelRows,
    columnCount: MAX_ROSTER_COLUMN,
    columnIndexes,
    dataStartRow: 2,
  };

  const fileName = label.trim() || "personnel-roster.xlsx";
  const dummyFile = new File([], fileName);

  return {
    file: dummyFile,
    fileName,
    sheetName: sheet.sheetName,
    headerRows: sheet.headerRows,
    rows: excelRows,
    columnCount: MAX_ROSTER_COLUMN,
    columnIndexes,
    dataStartRow: 2,
    sheets: [sheet],
  };
};

export const rosterLatestToSourceSnapshot = (
  latest: BackendPersonnelRosterLatest | null | undefined,
) => {
  const rows = mapRosterLatestToPreviewRows(latest);
  const label = latest?.sourceFileName || latest?.importName || "Загальний список";
  return rosterRowsToSourceSnapshot(rows, label);
};
