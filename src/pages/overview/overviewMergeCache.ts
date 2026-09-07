import type { BackendPersonnelOverview } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

const overviewStamp = (
  overview: Pick<BackendPersonnelOverview, "importId" | "rows">,
) =>
  [
    overview.importId ?? "no-import",
    overview.rows.length,
    ...overview.rows
      .slice(0, 5)
      .map((row) => String(row.id ?? row.externalId ?? "").trim())
      .filter(Boolean),
  ].join(":");

/** Shared by IndexedDB cache keys and PostgreSQL overview-merge snapshots. */
export const overviewMergeFingerprint = (
  overview: Pick<BackendPersonnelOverview, "importId" | "rows">,
  rosterFingerprint: string,
  rosterRows: EjournalPreviewRow[] = [],
) => {
  const rosterStamp =
    rosterFingerprint.trim() ||
    `${rosterRows.length}:${String(rosterRows[0]?.__dbRowId ?? "").trim()}`;
  return `overview-merge:v1:${overviewStamp(overview)}:${rosterStamp}`;
};

export const overviewMergeCacheKey = (
  overview: Pick<BackendPersonnelOverview, "importId" | "rows">,
  rosterFingerprint: string,
  rosterRows: EjournalPreviewRow[],
  columns: Array<{ key?: string }>,
) =>
  `personnel:${overviewMergeFingerprint(
    overview,
    rosterFingerprint,
    rosterRows,
  )}:${columns.length}`;
