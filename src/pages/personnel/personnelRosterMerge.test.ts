import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  extractBirthDateFromPersonName,
  getRosterPersonBirthDate,
  isPersonnelInStaffRoster,
  mergeRosterRowsIntoPreview,
  ROSTER_FIELD_PREFIX,
} from "./personnelRosterMerge";

const oosRow = (
  name: string,
  extra: Record<string, unknown> = {},
): EjournalPreviewRow =>
  ({
    __dbRowId: String(extra.__dbRowId ?? `oos:${name}`),
    прізвище: name,
    дата_народження: extra.дата_народження ?? extra.birthDate ?? "",
    id: extra.id ?? "",
    рнокпп_за_наявності: extra.rnokpp ?? "",
    ...extra,
  }) as EjournalPreviewRow;

const rosterRow = (
  name: string,
  extra: Record<string, unknown> = {},
): EjournalPreviewRow =>
  ({
    __dbRowId: String(extra.__dbRowId ?? `roster:${name}`),
    column_14: extra.column_14 ?? name,
    column_16: extra.column_16 ?? extra.birthDate ?? "",
    column_19: extra.column_19 ?? extra.rnokpp ?? "",
    ...extra,
  }) as EjournalPreviewRow;

describe("extractBirthDateFromPersonName", () => {
  it("reads a date in parentheses from штатка ПІБ", () => {
    expect(
      extractBirthDateFromPersonName(
        "ШЕВЧЕНКО Олександр Володимирович (11.05.1981)",
      ),
    ).toBe("1981-05-11");
  });
});

describe("mergeRosterRowsIntoPreview", () => {
  it("does not glue a 1981 штатка row onto a 1985 namesake", () => {
    const preview = {
      rows: [
        oosRow("ШЕВЧЕНКО Олександр Володимирович", {
          birthDate: "07.09.1985",
          id: "2103004",
        }),
      ],
    };
    const roster = [
      rosterRow("ШЕВЧЕНКО Олександр Володимирович (11.05.1981)", {
        birthDate: "11.05.1981",
        rnokpp: "1111111111",
      }),
    ];

    const merged = mergeRosterRowsIntoPreview(preview, roster);
    const card = merged.find((row) => row.id === "2103004");
    const extras = merged.filter((row) => row.id !== "2103004");

    expect(card).toBeTruthy();
    expect(
      Object.keys(card ?? {}).some((key) => key.startsWith(ROSTER_FIELD_PREFIX)),
    ).toBe(false);
    expect(extras).toHaveLength(1);
    expect(getRosterPersonBirthDate(roster[0])).toBe("1981-05-11");
  });

  it("merges the same person by name and birth and does not add a duplicate card", () => {
    const preview = {
      rows: [
        oosRow("ШЕВЧЕНКО Олександр Володимирович", {
          birthDate: "07.09.1985",
          id: "2103004",
        }),
      ],
    };
    const roster = [
      rosterRow("ШЕВЧЕНКО Олександр Володимирович (07.09.1985)", {
        birthDate: "07.09.1985",
        rnokpp: "3129609236",
      }),
    ];

    const merged = mergeRosterRowsIntoPreview(preview, roster);
    expect(merged).toHaveLength(1);
    expect(merged[0][`${ROSTER_FIELD_PREFIX}column_19`]).toBe("3129609236");
  });

  it("merges by ІПН even if штатка ПІБ has a dirty date, and drops the extra card", () => {
    const preview = {
      rows: [
        oosRow("ШЕВЧЕНКО Олександр Володимирович", {
          birthDate: "07.09.1985",
          id: "2103004",
          rnokpp: "3129609236",
        }),
      ],
    };
    const roster = [
      rosterRow("ШЕВЧЕНКО Олександр Володимирович (11.05.1981)", {
        birthDate: "11.05.1981",
        rnokpp: "3129609236",
      }),
      rosterRow("ШЕВЧЕНКО Олександр Володимирович (11.05.1981)", {
        __dbRowId: "roster:dup",
        birthDate: "11.05.1981",
        rnokpp: "3129609236",
      }),
    ];

    const merged = mergeRosterRowsIntoPreview(preview, roster);
    expect(merged).toHaveLength(1);
    expect(merged[0].дата_народження).toBe("07.09.1985");
    expect(merged[0][`${ROSTER_FIELD_PREFIX}column_19`]).toBe("3129609236");
  });
});

describe("isPersonnelInStaffRoster", () => {
  it("treats a roster-only card as in staff", () => {
    expect(
      isPersonnelInStaffRoster({
        __dbRowId: "roster:name:шевченко",
        прізвище: "ШЕВЧЕНКО Олександр Володимирович",
      } as EjournalPreviewRow),
    ).toBe(true);
  });

  it("treats an EЖООС card enriched from штатка as in staff", () => {
    expect(
      isPersonnelInStaffRoster({
        __dbRowId: "clxyzpersonnelrow",
        прізвище: "ШЕВЧЕНКО Олександр Володимирович",
        [`${ROSTER_FIELD_PREFIX}column_14`]: "ШЕВЧЕНКО Олександр Володимирович",
      } as EjournalPreviewRow),
    ).toBe(true);
  });

  it("treats an EЖООС-only card as out of staff", () => {
    expect(
      isPersonnelInStaffRoster(
        oosRow("ШЕВЧЕНКО Олександр Володимирович", { id: "2103004" }),
      ),
    ).toBe(false);
  });
});
