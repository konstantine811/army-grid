import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

export const ROSTER_ARCHIVE_FLAG_KEY = "__rosterArchive";
export const ROSTER_ARCHIVE_SOURCE_KEY = "джерело";
export const ROSTER_ARCHIVE_SOURCE_VALUE = "Архів";

export const isPersonnelFromArchive = (
  row: EjournalPreviewRow | null | undefined,
) => {
  if (!row) return false;
  if (row[ROSTER_ARCHIVE_FLAG_KEY] === true) return true;
  if (row[ROSTER_ARCHIVE_FLAG_KEY] === "true") return true;
  const source = String(
    row[ROSTER_ARCHIVE_SOURCE_KEY] ??
      row[`roster__${ROSTER_ARCHIVE_SOURCE_KEY}`] ??
      "",
  )
    .trim()
    .toLocaleLowerCase("uk-UA");
  return source.includes("архів");
};
