import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";

/** Propagate merged «Підрозділ» (col 2) down position/person rows in roster sheets. */
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
