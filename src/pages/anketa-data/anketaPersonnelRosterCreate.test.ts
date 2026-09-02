import { describe, expect, it } from "vitest";
import {
  buildPersonnelRosterValuesFromAnketa,
  buildPersonnelRowFromAnketa,
  isAnketaRowCreatableInPersonnel,
  mergeAnketaCreatedRowsIntoPreview,
  selectAnketaRowsMissingFromPersonnel,
} from "./anketaPersonnelRosterCreate";
import { formatAnketaBulkMergeReport } from "./anketaPersonMerge";
import { createEmptyAnketaRow } from "./anketaSheet";
import { getRosterValue } from "../personnel/personnelRosterMerge";
import { isLikelyPersonnelRow } from "../personnel/personnelUtils";

const anketa = (
  extra: Partial<ReturnType<typeof createEmptyAnketaRow>> = {},
) => {
  const row = createEmptyAnketaRow(1);
  return { ...row, fullName: "ШЕВЧЕНКО Тарас Григорович", ...extra };
};

describe("buildPersonnelRosterValuesFromAnketa", () => {
  it("fills morning-list columns and card keys so the person appears in the list", () => {
    const values = buildPersonnelRosterValuesFromAnketa(
      anketa({
        rank: "солдат",
        externalId: "2163435",
        birthDate: "21.03.1977",
        rnokpp: "2820405493",
        location: "ППД Вишневе",
        militaryId: "АА 123456",
        fullName: "БОРИШПОЛЕЦЬ Роман Юрійович (РАМА)",
      }),
    );

    expect(values.column_14).toBe("БОРИШПОЛЕЦЬ Роман Юрійович");
    expect(values.ПІБ).toBe("БОРИШПОЛЕЦЬ Роман Юрійович");
    expect(values.column_13).toBe("солдат");
    expect(values.column_16).toBe("21.03.1977");
    expect(values.column_19).toBe("2820405493");
    expect(values.column_31).toBe("ППД Вишневе");
    expect(values.column_15).toBe("РАМА");
    expect(values.позивний).toBe("РАМА");
    expect(getRosterValue(values as never, ["піб"])).toBe(
      "БОРИШПОЛЕЦЬ Роман Юрійович",
    );
  });

  it("maps named roster columns from the latest Штатка import", () => {
    const values = buildPersonnelRosterValuesFromAnketa(
      anketa({ rank: "сержант", rnokpp: "1234567890" }),
      [
        { key: "піб", label: "ПІБ", order: 0 },
        { key: "звання", label: "Звання", order: 1 },
        { key: "іпн", label: "ІПН", order: 2 },
      ],
    );

    expect(values.піб).toBe("ШЕВЧЕНКО Тарас Григорович");
    expect(values.звання).toBe("сержант");
    expect(values.іпн).toBe("1234567890");
  });
});

describe("selectAnketaRowsMissingFromPersonnel", () => {
  it("skips blank names and header-like rows", () => {
    expect(isAnketaRowCreatableInPersonnel(anketa({ fullName: "" }))).toBe(
      false,
    );
    expect(
      isAnketaRowCreatableInPersonnel(anketa({ fullName: "ПІБ" })),
    ).toBe(false);
    expect(
      selectAnketaRowsMissingFromPersonnel([
        anketa({ fullName: "" }),
        anketa({ fullName: "ПІБ" }),
        anketa({ fullName: "Іванов" }),
      ]),
    ).toEqual([]);
  });

  it("dedupes the same person twice in unmatched anketa rows", () => {
    const selected = selectAnketaRowsMissingFromPersonnel([
      anketa({ __rowId: "a", externalId: "111" }),
      anketa({ __rowId: "b", __rowNumber: 2, externalId: "111" }),
    ]);
    expect(selected).toHaveLength(1);
  });

  it("does not re-add a person already in the morning list", () => {
    const selected = selectAnketaRowsMissingFromPersonnel(
      [anketa({ externalId: "2163435" })],
      [{ id: "2163435", ПІБ: "ШЕВЧЕНКО Тарас Григорович" }],
    );
    expect(selected).toHaveLength(0);
  });
});

describe("formatAnketaBulkMergeReport", () => {
  it("shows how many people were added from unmatched anketa rows", () => {
    expect(
      formatAnketaBulkMergeReport({
        processed: 10,
        matched: 2,
        updated: 1,
        created: 7,
        skippedNoMatch: 0,
        skippedNoRowId: 0,
        skippedNoUpdates: 1,
        fieldCount: 3,
        phonesAdded: 0,
        errors: [],
      }),
    ).toContain("додано: 7");
  });
});

describe("buildPersonnelRowFromAnketa", () => {
  it("builds a list row that the personnel page will keep", () => {
    const row = buildPersonnelRowFromAnketa(
      anketa({
        rank: "солдат",
        externalId: "2163435",
        birthDate: "21.03.1977",
      }),
    );

    expect(isLikelyPersonnelRow(row)).toBe(true);
    expect(String(row.__dbRowId)).toMatch(/^anketa:/);
    expect(row.прізвище).toBe("ШЕВЧЕНКО Тарас Григорович");
  });
});

describe("mergeAnketaCreatedRowsIntoPreview", () => {
  it("appends missing people and skips those already in the list", () => {
    const created = buildPersonnelRowFromAnketa(
      anketa({ externalId: "99", fullName: "НОВИЙ Іван Петрович" }),
    );
    const already = buildPersonnelRowFromAnketa(
      anketa({ externalId: "2163435" }),
    );

    const merged = mergeAnketaCreatedRowsIntoPreview(
      [{ __dbRowId: "oos-1", прізвище: "ШЕВЧЕНКО Тарас Григорович", id: "2163435" }],
      [already, created],
    );

    expect(merged).toHaveLength(2);
    expect(merged.some((row) => row.прізвище === "НОВИЙ Іван Петрович")).toBe(
      true,
    );
  });
});
