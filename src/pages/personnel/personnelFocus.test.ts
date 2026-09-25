import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  findPersonnelRowByFocusTarget,
  findPersonnelRowByNameHint,
} from "./personnelFocus";

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
        search: "",
      }),
    ).toBe(rows[0]);
  });

  it("ignores overview synthetic row ids and matches by external id", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "roster:минайлюк сергій",
        externalId: "13188",
        search: "",
      }),
    ).toBe(rows[0]);
  });

  it("finds by external id when row id mismatches", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "stale-row",
        externalId: "13188",
        search: "",
      }),
    ).toBe(rows[0]);
  });

  it("finds by normalized numeric external id", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "",
        externalId: "13188.0",
        search: "",
      }),
    ).toBe(rows[0]);
  });

  it("finds by db row id passed as external id from overview", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "",
        externalId: "db-row-1",
        search: "",
      }),
    ).toBe(rows[0]);
  });

  it("falls back to search hint when ids do not match", () => {
    expect(
      findPersonnelRowByFocusTarget(rows, {
        rowId: "",
        externalId: "unknown-id",
        search: "МИНАЙЛЮК СЕРГІЙ ГРИГОРОВИЧ",
      }),
    ).toBe(rows[0]);
  });
});

describe("findPersonnelRowByNameHint", () => {
  const rows = [
    row({}),
    row({
      __dbRowId: "db-row-2",
      id: "99999",
      прізвище: "ГАЛУШКО ІВАН ПЕТРОВИЧ",
    }),
  ];

  it("finds by exact name", () => {
    expect(
      findPersonnelRowByNameHint(rows, "ГАЛУШКО ІВАН ПЕТРОВИЧ"),
    ).toBe(rows[1]);
  });

  it("finds by partial name when unique", () => {
    expect(findPersonnelRowByNameHint(rows, "ГАЛУШКО ІВАН")).toBe(rows[1]);
  });
});
