import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizePersonBirthKey } from "../personnel/personnelUtils";
import { ROSTER_FIELD_PREFIX } from "../personnel/personnelRosterMerge";
import {
  buildAnketaInStaffRowIdSet,
  matchAnketaRowToPersonnelDetailed,
  orderAnketaRowsForGapSearch,
  type AnketaPersonnelMatch,
} from "./anketaPersonMatch";
import type { AnketaRow } from "./anketaSheet";

const buildIndex = (rows: EjournalPreviewRow[]) => {
  const byExternalId = new Map<string, AnketaPersonnelMatch>();
  const byRnokpp = new Map<string, AnketaPersonnelMatch>();
  const byNameBirth = new Map<string, AnketaPersonnelMatch>();
  const byName = new Map<string, AnketaPersonnelMatch[]>();
  const byShortName = new Map<string, AnketaPersonnelMatch[]>();

  for (const row of rows) {
    const birthDate = String(row.birthDate ?? row.column_16 ?? "");
    const match = {
      row,
      summary: {
        name: String(row.column_14 ?? ""),
        rank: "",
        externalId: String(row.id ?? ""),
        rnokpp: String(row.column_19 ?? ""),
        birthDate,
        callSign: "",
        positionIndex: "",
        serviceType: "",
        additionalInfo: "",
        birthPlace: "",
        sex: "",
        location: "",
        positionTitle: "",
        arrivedFrom: "",
        education: "",
        relatives: "",
        phones: [],
        phonesDisplay: [],
        militaryId: "",
        contractFrom: "",
        contractTo: "",
      },
      matchBy: "name" as const,
    };
    const nameKey = match.summary.name.toLocaleLowerCase("uk-UA");
    const list = byName.get(nameKey) ?? [];
    list.push(match);
    byName.set(nameKey, list);
    const parts = nameKey.split(" ").filter(Boolean);
    if (parts.length >= 2) {
      const shortKey = `${parts[0]} ${parts[1]}`;
      const shortList = byShortName.get(shortKey) ?? [];
      shortList.push(match);
      byShortName.set(shortKey, shortList);
    }
    const id = String(row.id ?? "");
    if (id) byExternalId.set(id, { ...match, matchBy: "externalId" });
    const rnokpp = String(row.column_19 ?? "").replace(/\D/g, "");
    if (rnokpp.length >= 8) {
      byRnokpp.set(rnokpp, { ...match, matchBy: "rnokpp" });
    }
    const birthKey = normalizePersonBirthKey(birthDate);
    if (birthKey) {
      byNameBirth.set(`${nameKey}|${birthKey}`, { ...match, matchBy: "nameBirth" });
    }
  }

  return { byExternalId, byRnokpp, byNameBirth, byName, byShortName };
};

describe("matchAnketaRowToPersonnelDetailed", () => {
  it("suggests similar people when patronymic differs", () => {
    const personnelRow = {
      __dbRowId: "1",
      id: "999",
      column_14: "ШЕВЧУК ОЛЕКСІЙ ВОЛОДИМИРОВИЧ",
      column_19: "2815806012",
    } as EjournalPreviewRow;

    const anketaRow = {
      __rowId: "a1",
      __rowNumber: 2,
      fullName: "ШЕВЧУК ОЛЕКСІЙ ЛЕОНІДОВИЧ",
      rnokpp: "3009104579",
      externalId: "22814",
    } as AnketaRow;

    const index = buildIndex([personnelRow]);
    const result = matchAnketaRowToPersonnelDetailed(anketaRow, index);

    expect(result.match).toBeNull();
    expect(result.similar).toHaveLength(1);
    expect(result.similar[0]?.summary.name).toContain("ВОЛОДИМИРОВИЧ");
  });

  it("does not link by foreign id when full name differs", () => {
    const personnelRow = {
      __dbRowId: "2",
      id: "22814",
      column_14: "100-Н ОЛЕКСІЙ ЛЕОНІДОВИЧ",
    } as EjournalPreviewRow;

    const anketaRow = {
      __rowId: "a1",
      __rowNumber: 2,
      fullName: "ШЕВЧУК ОЛЕКСІЙ ЛЕОНІДОВИЧ",
      externalId: "22814",
    } as AnketaRow;

    const index = buildIndex([personnelRow]);
    const result = matchAnketaRowToPersonnelDetailed(anketaRow, index);

    expect(result.match).toBeNull();
    expect(result.similar).toHaveLength(0);
  });

  it("disambiguates duplicate pib by id and birth date", () => {
    const personnelRows = [
      {
        __dbRowId: "1",
        id: "101",
        column_14: "ІВАНЕНКО ІВАН ІВАНОВИЧ",
        column_16: "01.01.1980",
        birthDate: "01.01.1980",
      },
      {
        __dbRowId: "2",
        id: "202",
        column_14: "ІВАНЕНКО ІВАН ІВАНОВИЧ",
        column_16: "02.02.1990",
        birthDate: "02.02.1990",
      },
    ] as EjournalPreviewRow[];

    const anketaRow = {
      __rowId: "a1",
      __rowNumber: 2,
      fullName: "ІВАНЕНКО ІВАН ІВАНОВИЧ",
      externalId: "202",
      birthDate: "02.02.1990",
    } as AnketaRow;

    const result = matchAnketaRowToPersonnelDetailed(
      anketaRow,
      buildIndex(personnelRows),
    );

    expect(result.match?.summary.externalId).toBe("202");
    expect(result.match?.matchBy).toBe("externalId");
  });

  it("disambiguates duplicate pib by birth date when id is missing", () => {
    const personnelRows = [
      {
        __dbRowId: "1",
        column_14: "ІВАНЕНКО ІВАН ІВАНОВИЧ",
        column_16: "01.01.1980",
        birthDate: "01.01.1980",
      },
      {
        __dbRowId: "2",
        column_14: "ІВАНЕНКО ІВАН ІВАНОВИЧ",
        column_16: "02.02.1990",
        birthDate: "02.02.1990",
      },
    ] as EjournalPreviewRow[];

    const anketaRow = {
      __rowId: "a1",
      __rowNumber: 2,
      fullName: "ІВАНЕНКО ІВАН ІВАНОВИЧ",
      birthDate: "02.02.1990",
    } as AnketaRow;

    const result = matchAnketaRowToPersonnelDetailed(
      anketaRow,
      buildIndex(personnelRows),
    );

    expect(result.match?.matchBy).toBe("nameBirth");
    expect(result.match?.summary.birthDate).toBe("02.02.1990");
  });

  it("links by spreadsheet id when surname and first name match", () => {
    const personnelRow = {
      __dbRowId: "4",
      id: "2163435",
      column_14: "ПОТАПОВ ОЛЕКСІЙ ОЛЕКСАНДРОВИЧ",
    } as EjournalPreviewRow;

    const anketaRow = {
      __rowId: "a2",
      __rowNumber: 3,
      fullName: "ПОТАПОВ Олексій Олександрович",
      externalId: "2163435",
    } as AnketaRow;

    const result = matchAnketaRowToPersonnelDetailed(
      anketaRow,
      buildIndex([personnelRow]),
    );

    expect(result.match?.summary.externalId).toBe("2163435");
    expect(result.match?.matchBy).toBe("externalId");
  });

  it("links when full name matches", () => {
    const personnelRow = {
      __dbRowId: "3",
      id: "555",
      column_14: "ШЕВЧУК ОЛЕКСІЙ ЛЕОНІДОВИЧ",
    } as EjournalPreviewRow;

    const anketaRow = {
      __rowId: "a1",
      __rowNumber: 2,
      fullName: "ШЕВЧУК ОЛЕКСІЙ ЛЕОНІДОВИЧ",
    } as AnketaRow;

    const result = matchAnketaRowToPersonnelDetailed(anketaRow, buildIndex([personnelRow]));

    expect(result.match?.summary.name).toContain("ШЕВЧУК");
    expect(result.match?.matchBy).toBe("name");
  });
});

describe("orderAnketaRowsForGapSearch", () => {
  it("puts in-staff rows first while preserving sheet order within each group", () => {
    const rows = [
      { __rowId: "a", __rowNumber: 2, fullName: "А" } as AnketaRow,
      { __rowId: "b", __rowNumber: 3, fullName: "Б" } as AnketaRow,
      { __rowId: "c", __rowNumber: 4, fullName: "В" } as AnketaRow,
    ];
    const ordered = orderAnketaRowsForGapSearch(
      rows,
      new Set(["c", "a"]),
    );
    expect(ordered.map((row) => row.__rowId)).toEqual(["a", "c", "b"]);
  });

  it("returns original rows when everyone is in the same group", () => {
    const rows = [
      { __rowId: "a", __rowNumber: 2, fullName: "А" } as AnketaRow,
      { __rowId: "b", __rowNumber: 3, fullName: "Б" } as AnketaRow,
    ];
    expect(
      orderAnketaRowsForGapSearch(rows, new Set(["a", "b"])).map(
        (row) => row.__rowId,
      ),
    ).toEqual(["a", "b"]);
    expect(
      orderAnketaRowsForGapSearch(rows, new Set()).map((row) => row.__rowId),
    ).toEqual(["a", "b"]);
  });
});

describe("buildAnketaInStaffRowIdSet", () => {
  it("includes only anketa rows matched to personnel in the staff roster", () => {
    const inStaffPerson = {
      __dbRowId: "roster:1",
      id: "101",
      column_14: "КОВАЛЬ ІВАН ПЕТРОВИЧ",
      [`${ROSTER_FIELD_PREFIX}column_14`]: "КОВАЛЬ ІВАН ПЕТРОВИЧ",
    } as EjournalPreviewRow;
    const archivePerson = {
      __dbRowId: "2",
      id: "202",
      column_14: "ШЕВЧЕНКО ТАРАС ГРИГОРОВИЧ",
    } as EjournalPreviewRow;
    const index = buildIndex([inStaffPerson, archivePerson]);
    const rows = [
      {
        __rowId: "a1",
        __rowNumber: 2,
        fullName: "КОВАЛЬ ІВАН ПЕТРОВИЧ",
        externalId: "101",
      } as AnketaRow,
      {
        __rowId: "a2",
        __rowNumber: 3,
        fullName: "ШЕВЧЕНКО ТАРАС ГРИГОРОВИЧ",
        externalId: "202",
      } as AnketaRow,
    ];

    expect([...buildAnketaInStaffRowIdSet(rows, index)]).toEqual(["a1"]);
  });
});
