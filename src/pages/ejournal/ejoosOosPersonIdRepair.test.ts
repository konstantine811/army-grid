import { describe, expect, it } from "vitest";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  buildOosPersonIdRepairWrites,
  oosPersonIdCellLooksCorrupted,
} from "./ejoosOosPersonIdRepair";

const workbook = (sheets: ExcelWorkbookSnapshot["sheets"]): ExcelWorkbookSnapshot => ({
  fileName: "ejoos.xlsx",
  sheets,
});

const sheet = (
  sheetName: string,
  rawRows: unknown[][],
): ExcelWorkbookSnapshot["sheets"][number] => ({
  sheetIndex: 0,
  sheetName,
  rawRows: rawRows as ExcelWorkbookSnapshot["sheets"][number]["rawRows"],
  headerRows: [],
  rows: [],
  columnCount: 20,
  columnIndexes: Array.from({ length: 20 }, (_, index) => index),
  dataStartRow: 6,
});

describe("ejoosOosPersonIdRepair", () => {
  it("detects date-shaped OOS ID cells", () => {
    expect(oosPersonIdCellLooksCorrupted("13.01.1959")).toBe(true);
    expect(oosPersonIdCellLooksCorrupted(new Date("1959-01-13"))).toBe(true);
    expect(oosPersonIdCellLooksCorrupted(24867)).toBe(false);
    expect(oosPersonIdCellLooksCorrupted("24867")).toBe(false);
  });

  it("builds writes for date IDs from SHPO by name", () => {
    const writes = buildOosPersonIdRepairWrites(
      workbook([
        sheet("1. ШПО", [
          [],
          [],
          [],
          [],
          [],
          [],
          ["2103378", "", "", "", "", "солдат", "ХУБАЄВ Ільяс Ільгамович", "24867"],
        ]),
        sheet("2. ООС", [
          [],
          [],
          [],
          [],
          [],
          ["солдат", "ХУБАЄВ Ільяс Ільгамович", "13.01.1959", "2103378"],
        ]),
      ]),
    );
    expect(writes).toEqual([{ row: 6, column: 3, value: "24867" }]);
  });

  it("replaces birth-date serial with journal ID from SHPO", () => {
    const writes = buildOosPersonIdRepairWrites(
      workbook([
        sheet("1. ШПО", [
          [],
          [],
          [],
          [],
          [],
          [],
          ["2103378", "", "", "", "", "солдат", "ХУБАЄВ Ільяс Ільгамович", "24867"],
        ]),
        sheet("2. ООС", [
          [],
          [],
          [],
          [],
          [],
          ["солдат", "ХУБАЄВ Ільяс Ільгамович", 21557, "2103378"],
        ]),
      ]),
    );
    expect(writes).toEqual([{ row: 6, column: 3, value: "24867" }]);
  });

  it("skips rows that already have the correct journal ID", () => {
    const writes = buildOosPersonIdRepairWrites(
      workbook([
        sheet("1. ШПО", [
          [],
          [],
          [],
          [],
          [],
          [],
          ["2103378", "", "", "", "", "солдат", "ХУБАЄВ Ільяс Ільгамович", "24867"],
        ]),
        sheet("2. ООС", [
          [],
          [],
          [],
          [],
          [],
          ["солдат", "ХУБАЄВ Ільяс Ільгамович", "24867", "2103378"],
        ]),
      ]),
    );
    expect(writes).toEqual([]);
  });

});
