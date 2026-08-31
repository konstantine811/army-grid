import { describe, expect, it } from "vitest";
import { parseEjoosOos } from "./ejoosSyncPlan";
import {
  findExistingOosPersonRow,
  isOosSectionHeaderText,
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
