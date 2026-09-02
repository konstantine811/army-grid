import { describe, expect, it } from "vitest";
import {
  buildSheetGridFromXml,
  findTimesheetPersonRowsInGrid,
  mergeTimesheetGrids,
  pickTimesheetKeepRow,
  timesheetRowMatchesPerson,
} from "./ejoosTimesheetPersonRows";

const atrakhov = "АТРАХОВ Олександр Анатолійович";

describe("timesheetRowMatchesPerson", () => {
  it("matches a numeric ID in column H", () => {
    const row = ["", "764", "", "", "", "", atrakhov, 21155, "+", "+"];
    expect(timesheetRowMatchesPerson(row, "21155", atrakhov)).toBe(true);
  });

  it("matches leftover «згідно» by ПІБ even without a departure word", () => {
    const row = [
      "",
      "764",
      "",
      "",
      "",
      "",
      atrakhov,
      21155,
      "+",
      "йону від 05.08.2026 згідно",
    ];
    expect(timesheetRowMatchesPerson(row, "21155", atrakhov)).toBe(true);
  });

  it("does not match another person on the same index", () => {
    const row = [
      "",
      "764",
      "",
      "",
      "",
      "",
      "ХУБАЄВ Дмитро Алімбекович",
      24867,
      "+",
    ];
    expect(timesheetRowMatchesPerson(row, "21155", atrakhov)).toBe(false);
  });
});

describe("findTimesheetPersonRowsInGrid", () => {
  it("finds every named copy, including a leftover below used staff rows", () => {
    const grid: unknown[][] = [];
    grid[5] = ["header"];
    grid[930] = ["", "764", "", "", "", "", atrakhov, 21155, "−", "−"];
    grid[931] = [
      "",
      "764",
      "",
      "",
      "",
      "",
      "ХУБАЄВ Дмитро Алімбекович",
      24867,
      "+",
    ];
    grid[1032] = [
      "",
      "764",
      "",
      "",
      "",
      "",
      atrakhov,
      21155,
      "+",
      "йону від 05.08.2026 згідно",
    ];
    expect(
      findTimesheetPersonRowsInGrid(grid, "21155", atrakhov),
    ).toEqual([931, 1033]);
  });

  it("decodes shared-string ID/ПІБ from sheet XML", () => {
    const sst = [
      '<sst><si><t>АТРАХОВ Олександр Анатолійович</t></si>',
      "<si><t>21155</t></si></sst>",
    ].join("");
    const xml = [
      '<worksheet><sheetData><row r="1040">',
      '<c r="G1040" t="s"><v>0</v></c>',
      '<c r="H1040" t="s"><v>1</v></c>',
      "</row></sheetData></worksheet>",
    ].join("");
    expect(
      findTimesheetPersonRowsInGrid(
        buildSheetGridFromXml(xml, sst),
        "21155",
        atrakhov,
      ),
    ).toEqual([1040]);
  });

  it("finds a row written only in sheet XML (outside snapshot usedRange)", () => {
    const xml = [
      '<worksheet><sheetData>',
      '<row r="1033">',
      '<c r="B1033"><v>764</v></c>',
      `<c r="G1033" t="inlineStr"><is><t>${atrakhov}</t></is></c>`,
      '<c r="H1033"><v>21155</v></c>',
      '<c r="I1033" t="inlineStr"><is><t>+</t></is></c>',
      '<c r="J1033" t="inlineStr"><is><t>йону від 05.08.2026 згідно</t></is></c>',
      "</row></sheetData></worksheet>",
    ].join("");
    const fromXml = buildSheetGridFromXml(xml);
    const snapshot: unknown[][] = Array.from({ length: 1000 }, () => []);
    const merged = mergeTimesheetGrids(snapshot, fromXml);
    expect(
      findTimesheetPersonRowsInGrid(merged, "21155", atrakhov),
    ).toEqual([1033]);
  });
});

describe("pickTimesheetKeepRow", () => {
  it("prefers the closed вибув row, then planned, then first", () => {
    expect(
      pickTimesheetKeepRow([931, 1033], (row) => row === 1033),
    ).toBe(1033);
    expect(pickTimesheetKeepRow([931, 1033], () => false, 931)).toBe(931);
    expect(pickTimesheetKeepRow([931, 1033], () => false)).toBe(931);
  });
});
