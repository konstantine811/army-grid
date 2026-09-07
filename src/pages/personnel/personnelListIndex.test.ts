import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { ROSTER_ARCHIVE_FLAG_KEY } from "./staffSheetArchiveMarker";
import { ROSTER_FIELD_PREFIX } from "./personnelRosterMerge";
import { buildPersonnelListIndex } from "./personnelListIndex";

const row = (
  name: string,
  extra: Partial<EjournalPreviewRow> = {},
): EjournalPreviewRow =>
  ({
    __dbRowId: String(extra.__dbRowId ?? `row:${name}`),
    прізвище: name,
    ...extra,
  }) as EjournalPreviewRow;

describe("buildPersonnelListIndex", () => {
  it("precomputes staff flags and counts", () => {
    const index = buildPersonnelListIndex([
      row("ПЕТРЕНКО Петро", {
        [`${ROSTER_FIELD_PREFIX}column_14`]: "ПЕТРЕНКО Петро",
      }),
      row("ІВАНЕНКО Іван", {
        __dbRowId: "roster:archive:ivan",
        [ROSTER_ARCHIVE_FLAG_KEY]: true,
      }),
    ]);

    expect(index.staffCounts).toEqual({ all: 2, in: 1, archive: 1 });
    expect(index.records[0]?.inStaff).toBe(true);
    expect(index.records[1]?.inArchive).toBe(true);
    expect(index.records[0]?.searchBase).toContain("петренко");
  });
});
