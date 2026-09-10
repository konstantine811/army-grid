import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";

/** Рядок штатки: посада / звання / ПІБ — не порожній службовий рядок з однією «нова». */
export const isRosterStaffLineRow = (row: EjournalPreviewRow) =>
  [5, 8, 13, 14].some((columnNumber) =>
    readRosterColumnValue(row, columnNumber).trim(),
  );

const rosterBattalionFromRow = (row: EjournalPreviewRow) =>
  readRosterColumnValue(row, 1).trim().toLowerCase().replace(/ё/g, "е");

/** Propagate merged «Підрозділ» (col 2) down position/person rows in roster sheets. */
export const fillDownRosterUnitRows = (
  rows: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  let currentUnit = "";
  return rows.map((row) => {
    const battalion = rosterBattalionFromRow(row);
    if (battalion === "стара") {
      currentUnit = "";
      return row;
    }

    const unit = readRosterColumnValue(row, 2).trim();
    if (unit) {
      currentUnit = unit;
      return row;
    }
    if (!currentUnit || !isRosterStaffLineRow(row)) return row;
    return { ...row, column_2: currentUnit };
  });
};
