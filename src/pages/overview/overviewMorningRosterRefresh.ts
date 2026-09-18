import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { isPersonnelFromArchive, pickArchiveOnlyRosterRows } from "../personnel/staffSheetArchive";

/** Google refresh contains only the main sheet; retain the loaded archive. */
export const mergeMorningRosterRefresh = (
  freshMainRows: EjournalPreviewRow[],
  loadedRows: EjournalPreviewRow[],
): EjournalPreviewRow[] => [
  ...freshMainRows,
  ...pickArchiveOnlyRosterRows(freshMainRows, loadedRows.filter(isPersonnelFromArchive)),
];
