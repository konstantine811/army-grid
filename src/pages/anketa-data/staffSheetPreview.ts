import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import type { StaffSheetEnrichmentEntry } from "./staffSheetEnrichment";

export const STAFF_SHEET_PREVIEW_COLUMNS = [
  { number: 1, label: "№" },
  { number: 2, label: "Підрозділ" },
  { number: 14, label: "ПІБ" },
  { number: 15, label: "Позивний" },
  { number: 10, label: "Анкета" },
  { number: 16, label: "Дата народж." },
  { number: 17, label: "Рік" },
  { number: 18, label: "Повних років" },
  { number: 19, label: "ІПН" },
  { number: 21, label: "Статус" },
] as const;

export type StaffSheetPreviewRow = {
  excelRowNumber: number;
  cells: string[];
};

const hasPersonName = (row: EjournalPreviewRow) =>
  readRosterColumnValue(row, 14).trim().length >= 3;

export const buildStaffSheetPreviewRows = (
  rows: EjournalPreviewRow[],
): StaffSheetPreviewRow[] =>
  [...rows]
    .filter(hasPersonName)
    .sort(
      (left, right) =>
        (Number(left.__rowNumber) || 0) - (Number(right.__rowNumber) || 0),
    )
    .map((row) => ({
      excelRowNumber: Number(row.__rowNumber) || 0,
      cells: STAFF_SHEET_PREVIEW_COLUMNS.map((column) =>
        readRosterColumnValue(row, column.number).trim(),
      ),
    }));

/** Підставляє зіставлені з ОС/анкетами значення (кол. 10, 16–19) у preview. */
const ENRICHMENT_PREVIEW_CELLS: Array<[cellIndex: number, valueIndex: number]> =
  [
    [4, 0],
    [5, 1],
    [6, 2],
    [7, 3],
    [8, 4],
  ];

export const buildStaffSheetPreviewRowsWithEnrichment = (
  rows: EjournalPreviewRow[],
  entries: StaffSheetEnrichmentEntry[],
): StaffSheetPreviewRow[] => {
  const byRow = new Map(
    entries.map((entry) => [entry.excelRowNumber, entry.values]),
  );
  return buildStaffSheetPreviewRows(rows).map((row) => {
    const values = byRow.get(row.excelRowNumber);
    if (!values) return row;
    const cells = [...row.cells];
    for (const [cellIndex, valueIndex] of ENRICHMENT_PREVIEW_CELLS) {
      const next = String(values[valueIndex] ?? "").trim();
      if (next) cells[cellIndex] = next;
    }
    return { ...row, cells };
  });
};
