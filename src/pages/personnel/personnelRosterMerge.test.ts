import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  applyPersonnelMergeDelta,
  buildPersonnelMergeDelta,
  buildStaffScopePreview,
  datasetCarriesSpreadsheetPersonIds,
  dropEmptyStaffBlankRows,
  extractBirthDateFromPersonName,
  getRosterPersonBirthDate,
  combineRosterRowSources,
  isPersonnelFromArchive,
  isPersonnelInStaffRoster,
  mergeRosterRowsIntoPreview,
  ROSTER_FIELD_PREFIX,
} from "./personnelRosterMerge";
import {
  formatPersonBirthDateDisplay,
  getPersonDisplayName,
  isLikelyPersonnelRow,
  resolvePersonBirthDate,
} from "./personnelUtils";

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

  it("reads a date with р.н. suffix from штатка ПІБ", () => {
    expect(
      extractBirthDateFromPersonName(
        "САВЧЕНКО Максим Олександрович (25.01.1997 р.н.)",
      ),
    ).toBe("1997-01-25");
  });
});

describe("formatPersonBirthDateDisplay", () => {
  it("extracts DD.MM.YYYY from birth date column noise", () => {
    expect(formatPersonBirthDateDisplay("25.01.1997 р.н.")).toBe(
      "25.01.1997",
    );
    expect(
      formatPersonBirthDateDisplay(
        "САВЧЕНКО Максим Олександрович (25.01.1997 р.н.)",
      ),
    ).toBe("25.01.1997");
  });
});

describe("resolvePersonBirthDate", () => {
  it("normalizes штатка birth date with р.н.", () => {
    const row = rosterRow("САВЧЕНКО Максим Олександрович", {
      column_16: "25.01.1997 р.н.",
    });
    expect(resolvePersonBirthDate(row)).toBe("25.01.1997");
  });

  it("prefers fresh staff column 16 over stale EJOOS birth date", () => {
    const row = {
      дата_народження: "07.09.1985",
      "roster__column_16": "11.05.1981",
    } as EjournalPreviewRow;
    expect(resolvePersonBirthDate(row)).toBe("11.05.1981");
  });

  it("prefers staff PIB with birth date in parentheses for display", () => {
    const row = {
      прізвище: "ШЕВЧЕНКО Олександр Володимирович",
      "roster__column_14":
        "ШЕВЧЕНКО Олександр Володимирович (11.05.1981 р.н.)",
    } as EjournalPreviewRow;
    expect(getPersonDisplayName(row)).toBe(
      "ШЕВЧЕНКО Олександр Володимирович (11.05.1981 р.н.)",
    );
  });

  it("returns empty string when row is null", () => {
    expect(resolvePersonBirthDate(null)).toBe("");
  });
});

describe("mergeRosterRowsIntoPreview", () => {
  it("adds an archive-only person to all personnel but not to current staff", () => {
    const merged = mergeRosterRowsIntoPreview({ rows: [] }, [
      rosterRow("АРХІВНИЙ Петро Іванович", {
        __rosterArchive: true,
        джерело: "Архів",
      }),
    ]);

    expect(merged).toHaveLength(1);
    expect(isPersonnelFromArchive(merged[0])).toBe(true);
    expect(isPersonnelInStaffRoster(merged[0])).toBe(false);
    expect(String(merged[0].__dbRowId)).toMatch(/^roster:archive:/);
  });

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

  it("does not add a second card when the only staff row has a different birth text", () => {
    const preview = {
      rows: [
        oosRow("КІЯНЕНКО Андрій Олександрович", {
          birthDate: "08.06.1992",
          __dbRowId: "roster:p:кіяненко",
        }),
      ],
    };
    const roster = [
      rosterRow("КІЯНЕНКО Андрій Олександрович", {
        birthDate: "92.06.2008",
        __dbRowId: "excel-row-2",
      }),
    ];

    const merged = mergeRosterRowsIntoPreview(preview, roster);

    expect(merged).toHaveLength(1);
    expect(merged[0]?.__dbRowId).toBe("roster:p:кіяненко");
    expect(merged.filter(isPersonnelInStaffRoster)).toHaveLength(1);
  });

  it("uses one staff row for only one duplicated OOS card", () => {
    const preview = {
      rows: [
        oosRow("ПЕТРЕНКО Іван Петрович", { id: "oos-1" }),
        oosRow("ПЕТРЕНКО Іван Петрович", { id: "oos-2" }),
      ],
    };
    const roster = [rosterRow("ПЕТРЕНКО Іван Петрович")];

    const merged = mergeRosterRowsIntoPreview(preview, roster);
    const staffRows = merged.filter(isPersonnelInStaffRoster);

    expect(merged).toHaveLength(2);
    expect(staffRows).toHaveLength(1);
    expect(staffRows[0]?.id).toBe("oos-1");
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

  it("does not glue Tsapenko onto Kravchuk when a stale card has Tsapenko's ІПН", () => {
    const preview = {
      rows: [
        oosRow("КРАВЧУК Богдан Сергійович", {
          birthDate: "22.07.1992",
          id: "2103182",
          rnokpp: "3380604650",
        }),
      ],
    };
    const roster = [
      rosterRow("ЦАПЕНКО Микола Володимирович", {
        birthDate: "22.07.1992",
        rnokpp: "3380604650",
      }),
    ];

    const merged = mergeRosterRowsIntoPreview(preview, roster);
    const kravchuk = merged.find((row) => row.id === "2103182");
    const tsapenko = merged.find(
      (row) => getPersonDisplayName(row) === "ЦАПЕНКО Микола Володимирович",
    );

    expect(merged).toHaveLength(2);
    expect(kravchuk).toBeTruthy();
    expect(
      Object.keys(kravchuk ?? {}).some((key) =>
        key.startsWith(ROSTER_FIELD_PREFIX),
      ),
    ).toBe(false);
    expect(tsapenko?.__dbRowId).toMatch(/^roster:/);
  });

  it("shows ПІБ from штатка when EЖООС прізвище is empty", () => {
    const preview = {
      rows: [
        oosRow("", {
          birthDate: "12.09.1992",
          id: "2103999",
          rnokpp: "3129609236",
        }),
      ],
    };
    const roster = [
      rosterRow("ЦАПЕНКО Микола Володимирович", {
        birthDate: "12.09.1992",
        rnokpp: "3129609236",
      }),
    ];

    const merged = mergeRosterRowsIntoPreview(preview, roster);
    expect(merged).toHaveLength(1);
    expect(getPersonDisplayName(merged[0])).toBe(
      "ЦАПЕНКО Микола Володимирович",
    );
    expect(isLikelyPersonnelRow(merged[0]!)).toBe(true);
    expect(isPersonnelInStaffRoster(merged[0]!)).toBe(true);
  });
});

describe("combineRosterRowSources", () => {
  it("adds people present only in staff import cache", () => {
    const db = [
      rosterRow("ШЕВЧЕНКО Олександр Володимирович", {
        birthDate: "07.09.1985",
      }),
    ];
    const imported = [
      rosterRow("ЦАПЕНКО Микола Володимирович", {
        __dbRowId: "import:367",
        birthDate: "22.07.1992",
      }),
    ];
    const combined = combineRosterRowSources(db, imported);
    expect(combined).toHaveLength(2);
    expect(
      combined.some((row) =>
        String(row.column_14 ?? "").includes("ЦАПЕНКО"),
      ),
    ).toBe(true);
  });

  it("keeps the primary source when the same sheet row changed", () => {
    const fresh = [
      rosterRow("НОВИЙ Військовослужбовець", {
        __rowNumber: 367,
      }),
    ];
    const stale = [
      rosterRow("СТАРИЙ Військовослужбовець", {
        __rowNumber: 367,
      }),
    ];

    const combined = combineRosterRowSources(fresh, stale);

    expect(combined).toHaveLength(1);
    expect(combined[0]?.column_14).toBe("НОВИЙ Військовослужбовець");
  });
});

describe("mergeRosterRowsIntoPreview roster-only dedupe", () => {
  it("keeps two different people who share the same mistaken ІПН", () => {
    const merged = mergeRosterRowsIntoPreview({ rows: [] }, [
      rosterRow("МАЗУР Дмитро Михайлович", {
        column_16: "27.12.1999",
        column_19: "3462502235",
        __rowNumber: 114,
      }),
      rosterRow("КОРОТКОВ Віталій Юрійович", {
        column_16: "19.10.1994",
        column_19: "3462502235",
        __rowNumber: 241,
      }),
    ]);

    expect(merged).toHaveLength(2);
    expect(merged.map((row) => getPersonDisplayName(row))).toEqual([
      "МАЗУР Дмитро Михайлович",
      "КОРОТКОВ Віталій Юрійович",
    ]);
    expect(merged.every(isPersonnelInStaffRoster)).toBe(true);
  });
});

describe("buildStaffScopePreview", () => {
  it("keeps the OOS id used for photos and hides people outside the roster", () => {
    const preview = buildStaffScopePreview(
      [
        rosterRow("ШЕВЧЕНКО Олександр Володимирович (07.09.1985)", {
          birthDate: "07.09.1985",
        }),
      ],
      null,
      {
        rows: [
          oosRow("ШЕВЧЕНКО Олександр Володимирович", {
            birthDate: "07.09.1985",
            id: "2103004",
          }),
          oosRow("КРАВЧУК Богдан Сергійович", {
            birthDate: "22.07.1992",
            id: "2103182",
          }),
        ],
        columns: [],
        sheet: null,
      },
    );

    expect(preview?.rows).toHaveLength(1);
    expect(preview?.rows[0]?.id).toBe("2103004");
    expect(isPersonnelInStaffRoster(preview?.rows[0])).toBe(true);
  });

  it("drops an empty duplicate when the staff line already has that name", () => {
    const rows = dropEmptyStaffBlankRows([
      oosRow("КІЯНЕНКО Андрій Олександрович", {
        id: "12",
        column_14: "КІЯНЕНКО Андрій Олександрович",
        column_5: "Командир батальйону",
      }),
      {
        __dbRowId: "roster:p:кіяненко",
        прізвище: "КІЯНЕНКО Андрій Олександрович",
        ПІБ: "КІЯНЕНКО Андрій Олександрович",
      } as EjournalPreviewRow,
    ]);

    expect(rows).toHaveLength(1);
    expect(rows[0]?.id).toBe("12");
  });

  it("lists the staff sheet order, so the first roster row stays first", () => {
    const preview = buildStaffScopePreview(
      [
        rosterRow("КІЯНЕНКО Андрій Олександрович", {
          __rosterOrder: 0,
          __rowNumber: 2,
        }),
        rosterRow("БАБЧЕНКО Олег Володимирович", {
          __rosterOrder: 2,
          __rowNumber: 4,
        }),
      ],
      null,
      {
        rows: [
          oosRow("БАБЧЕНКО Олег Володимирович", { id: "4246" }),
          oosRow("КІЯНЕНКО Андрій Олександрович", { id: "12" }),
        ],
        columns: [],
        sheet: null,
      },
    );

    expect(preview?.rows.map((row) => row.id)).toEqual(["12", "4246"]);
  });
});

describe("datasetCarriesSpreadsheetPersonIds", () => {
  it("accepts OOS person numbers and ignores fingerprint-only rows", () => {
    expect(
      datasetCarriesSpreadsheetPersonIds(
        [oosRow("БАБЧЕНКО Олег Володимирович", { id: "4246" })],
        1,
      ),
    ).toBe(true);
    expect(
      datasetCarriesSpreadsheetPersonIds(
        [
          {
            __dbRowId: "roster:babchenko",
            id: "p:бабченко олег володимирович:1972-08-16",
            прізвище: "БАБЧЕНКО Олег Володимирович",
          } as EjournalPreviewRow,
        ],
        1,
      ),
    ).toBe(false);
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

describe("personnel merge worker delta", () => {
  it("returns changed fields instead of full preview rows", () => {
    const source = [oosRow("КОВАЛЬ Іван Петрович")];
    const delta = buildPersonnelMergeDelta(
      { rows: source },
      [rosterRow("КОВАЛЬ Іван Петрович", { column_2: "1 рота" })],
    );

    expect(delta.patches).toHaveLength(1);
    expect(delta.appended).toHaveLength(0);
    expect("прізвище" in delta.patches[0].set).toBe(false);

    const merged = applyPersonnelMergeDelta(source, delta);
    expect(merged[0][`${ROSTER_FIELD_PREFIX}column_2`]).toBe("1 рота");
    expect(merged[0].прізвище).toBe("КОВАЛЬ Іван Петрович");
  });

  it("accepts the previous worker result shape during an update", () => {
    const source = [oosRow("СТАРИЙ Запис")];
    const legacyWorkerRows = [oosRow("НОВИЙ Запис")];

    expect(applyPersonnelMergeDelta(source, legacyWorkerRows)).toBe(
      legacyWorkerRows,
    );
  });
});
