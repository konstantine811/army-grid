import { describe, expect, it } from "vitest";
import { parseEjoosOos } from "./ejoosSyncPlan";
import {
  createOosRowResolver,
  filterOosStaffHistoryIndexes,
  findDuplicateOosById,
  findExistingOosPersonRow,
  findNextEmptyOosDataRow,
  findPossibleDuplicateOosByName,
  isOosSectionHeaderText,
  oosIdentityFromOp,
  oosPersonIdText,
} from "./ejoosOosText";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";

const grid = (
  rows: Array<Array<string | number | null>>,
): ExcelSheetSnapshot => ({
  sheetIndex: 1,
  sheetName: "2. ООС",
  rawRows: rows,
  headerRows: [],
  rows: [],
  columnCount: 3,
  columnIndexes: [0, 1, 2],
  dataStartRow: 6,
});

describe("filterOosStaffHistoryIndexes", () => {
  it("keeps real staff indexes and drops РОЗПОРЯДЖЕННЯ", () => {
    expect(
      filterOosStaffHistoryIndexes("РОЗПОРЯДЖЕННЯ\n2103520\n2103378"),
    ).toBe("2103520\n2103378");
    expect(filterOosStaffHistoryIndexes("РОЗПОРЯДЖЕННЯ")).toBe("");
  });
});

describe("findExistingOosPersonRow", () => {
  const rows: Array<Array<string | number | null>> = [
    [],
    [],
    [],
    [],
    [],
    ["звання", "ПІБ", "ID"],
    ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", 12840],
    ["", "#N/A", "#N/A"],
    [
      "",
      "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
      "",
    ],
    ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", 12840],
  ];
  const getCell = (row: number, column: number) =>
    rows[row - 1]?.[column - 1] ?? null;

  it("returns the first main-list card, not the one after the section header", () => {
    expect(
      findExistingOosPersonRow(getCell, {
        personId: "12840",
        fullName: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович",
        lastRow: rows.length,
      }),
    ).toBe(7);
  });

  it("matches a punctuated ПІБ to the same card", () => {
    expect(
      findExistingOosPersonRow(getCell, {
        personId: "",
        fullName: "ДОБРОВОЛЬСЬКИЙ, Володимир Миколайович.",
        lastRow: rows.length,
      }),
    ).toBe(7);
  });
});

describe("parseEjoosOos", () => {
  it("skips section headers and #N/A placeholder rows", () => {
    const sheet = grid([
      [],
      [],
      [],
      [],
      [],
      ["солдат", "ІНШИЙ", 1],
      [
        "",
        "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862",
        "",
      ],
      ["", "#N/A", null],
      ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", 12840],
    ]);
    const parsed = parseEjoosOos(sheet);
    expect(parsed.map((row) => row.excelRow)).toEqual([6, 9]);
    expect(isOosSectionHeaderText(sheet.rawRows[6][1] as string)).toBe(true);
  });
});

describe("createOosRowResolver identity aliases", () => {
  it("keeps nextPersonId and personId on the same reserved row", () => {
    let allocated = 0;
    const resolver = createOosRowResolver({
      getCell: () => "",
      lastRow: 10,
      allocateEmpty: () => {
        allocated += 1;
        return 20;
      },
    });
    const first = resolver.resolve(
      oosIdentityFromOp({
        personId: "",
        fullName: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович",
        payload: { nextPersonId: "12840" },
      }),
      { create: true },
    );
    const second = resolver.resolve(
      oosIdentityFromOp({
        personId: "12840",
        fullName: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович",
      }),
      { create: true },
    );
    expect(first).toBe(20);
    expect(second).toBe(20);
    expect(allocated).toBe(1);
  });

  it("updates an existing ID instead of allocating a new row", () => {
    let allocated = 0;
    const resolver = createOosRowResolver({
      getCell: (row, column) => {
        if (row !== 7) return "";
        if (column === 2) return "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович";
        if (column === 3) return "12840";
        return "";
      },
      lastRow: 10,
      allocateEmpty: () => {
        allocated += 1;
        return 20;
      },
    });
    expect(
      resolver.resolve(
        { personId: "12840", fullName: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович" },
        { create: true },
      ),
    ).toBe(7);
    expect(allocated).toBe(0);
  });

  it("matches numeric ID 24867 without treating it as an Excel date", () => {
    expect(oosPersonIdText(24867)).toBe("24867");
    expect(oosPersonIdText("24867")).toBe("24867");
    let allocated = 0;
    const resolver = createOosRowResolver({
      getCell: (row, column) => {
        if (row !== 7) return "";
        if (column === 2) return "ХУБАЄВ Ільяс Ільгамович";
        if (column === 3) return 24867;
        return "";
      },
      lastRow: 10,
      allocateEmpty: () => {
        allocated += 1;
        return 20;
      },
    });
    expect(
      findExistingOosPersonRow(
        (row, column) => {
          if (row !== 7) return "";
          if (column === 2) return "ХУБАЄВ Ільяс Ільгамович";
          if (column === 3) return 24867;
          return "";
        },
        { personId: "24867", fullName: "ХУБАЄВ Ільяс Ільгамович", lastRow: 10 },
      ),
    ).toBe(7);
    expect(
      resolver.resolve(
        { personId: "24867", fullName: "ХУБАЄВ Ільяс Ільгамович" },
        { create: true },
      ),
    ).toBe(7);
    expect(allocated).toBe(0);
  });

  it("does not overwrite a different ID that shares the same ПІБ", () => {
    expect(
      findExistingOosPersonRow(
        (row, column) => {
          if (row !== 7) return "";
          if (column === 2) return "ІВАНОВ Іван Іванович";
          if (column === 3) return "11111";
          return "";
        },
        { personId: "22222", fullName: "ІВАНОВ Іван Іванович", lastRow: 10 },
      ),
    ).toBe(0);
  });

  it("does not reuse a name-cached row when the new ID is different", () => {
    let next = 20;
    const resolver = createOosRowResolver({
      getCell: () => "",
      lastRow: 10,
      allocateEmpty: () => {
        next += 1;
        return next;
      },
    });
    const first = resolver.resolve(
      { personId: "11111", fullName: "ІВАНОВ Іван Іванович" },
      { create: true },
    );
    const second = resolver.resolve(
      { personId: "22222", fullName: "ІВАНОВ Іван Іванович" },
      { create: true },
    );
    expect(first).toBe(21);
    expect(second).toBe(22);
  });

  it("can fall back to a unique name only when the existing card has no ID", () => {
    expect(
      findExistingOosPersonRow(
        (row, column) => {
          if (row !== 7) return "";
          if (column === 2) return "ІВАНОВ Іван Іванович";
          return "";
        },
        { personId: "22222", fullName: "ІВАНОВ Іван Іванович", lastRow: 10 },
      ),
    ).toBe(7);
  });
});

describe("findNextEmptyOosDataRow", () => {
  it("uses blanks before the disposition section before appending", () => {
    const rows = [
      [],
      [],
      [],
      [],
      [],
      ["солдат", "ПЕРШИЙ", "1"],
      ["", "", ""],
      ["", "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862", ""],
      ["солдат", "РОЗПОРЯДЖЕННЯ", "2"],
    ];
    const getCell = (row: number, column: number) =>
      rows[row - 1]?.[column - 1] ?? "";
    expect(findNextEmptyOosDataRow(getCell, { lastRow: rows.length })).toBe(7);
  });

  it("can append when the main OOS block has no empty rows", () => {
    const rows = [
      [],
      [],
      [],
      [],
      [],
      ["солдат", "ПЕРШИЙ", "1"],
      ["солдат", "ДРУГИЙ", "2"],
      ["", "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА ВІЙСЬКОВОЇ ЧАСТИНИ А 4862", ""],
      ["солдат", "РОЗПОРЯДЖЕННЯ", "3"],
    ];
    const reserved = new Set([10]);
    const getCell = (row: number, column: number) =>
      rows[row - 1]?.[column - 1] ?? "";
    expect(
      findNextEmptyOosDataRow(getCell, {
        lastRow: rows.length,
        reserved,
        allowAppend: true,
      }),
    ).toBe(11);
    expect(findNextEmptyOosDataRow(getCell, { lastRow: rows.length })).toBe(0);
  });
});

describe("OOS duplicate checks", () => {
  it("groups two cards with the same ID", () => {
    expect(
      findDuplicateOosById([
        { excelRow: 100, personId: "12840", fullName: "A" },
        { excelRow: 101, personId: "12840", fullName: "A" },
        { excelRow: 102, personId: "1", fullName: "B" },
      ]).map((group) => group.map((row) => row.excelRow)),
    ).toEqual([[100, 101]]);
  });

  it("warns when the same name has a row without ID", () => {
    expect(
      findPossibleDuplicateOosByName([
        {
          excelRow: 100,
          personId: "12840",
          fullName: "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович",
        },
        {
          excelRow: 101,
          personId: "",
          fullName: "ДОБРОВОЛЬСЬКИЙ, Володимир Миколайович.",
        },
      ]).map((group) => group.map((row) => row.excelRow)),
    ).toEqual([[100, 101]]);
  });
});
