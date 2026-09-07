import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { ROSTER_ARCHIVE_FLAG_KEY } from "../personnel/staffSheetArchiveMarker";
import {
  countStaffSheetPersons,
  countStaffSheetPersonsInArchive,
  countStaffSheetPersonsInRoster,
} from "./staffSheetPreview";

const row = (
  name: string,
  extra: Partial<EjournalPreviewRow> = {},
): EjournalPreviewRow =>
  ({
    __rowNumber: 2,
    column_14: name,
    ...extra,
  }) as EjournalPreviewRow;

describe("staffSheetPreview counts", () => {
  it("ignores section headers and counts valid ПІБ like personnel page", () => {
    const rows = [
      row("ІВАНЕНКО Іван Іванович"),
      row("1 рота"),
      row("AB"),
    ];
    expect(countStaffSheetPersons(rows)).toBe(1);
  });

  it("splits roster and archive tabs", () => {
    const rows = [
      row("ПЕРШИЙ Боєць"),
      row("ДРУГИЙ Боєць", { [ROSTER_ARCHIVE_FLAG_KEY]: true }),
    ];
    expect(countStaffSheetPersonsInRoster(rows)).toBe(1);
    expect(countStaffSheetPersonsInArchive(rows)).toBe(1);
    expect(countStaffSheetPersons(rows)).toBe(2);
  });
});
