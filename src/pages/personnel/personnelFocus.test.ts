import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { findPersonnelRowByFocusTarget } from "./personnelFocus";

const row = (extra: Partial<EjournalPreviewRow>): EjournalPreviewRow =>
  ({
    __dbRowId: "db-row-1",
    id: "13188",
    прізвище: "МИНАЙЛЮК СЕРГІЙ ГРИГОРОВИЧ",
    ...extra,
  }) as EjournalPreviewRow;

describe("findPersonnelRowByFocusTarget", () => {
  const rows = [row({})];

  it("finds by string row id", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "db-row-1",
        externalId: "",
      }),
    ).toBe(rows[0]);
  });

  it("ignores overview synthetic row ids and matches by external id", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "roster:минайлюк сергій",
        externalId: "13188",
      }),
    ).toBe(rows[0]);
  });

  it("finds by external id when row id mismatches", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "stale-row",
        externalId: "13188",
      }),
    ).toBe(rows[0]);
  });

  it("finds by normalized numeric external id", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "",
        externalId: "13188.0",
      }),
    ).toBe(rows[0]);
  });
});
