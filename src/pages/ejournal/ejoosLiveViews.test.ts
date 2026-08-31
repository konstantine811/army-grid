import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { buildEjoosLiveView, parseExcluded } from "./ejoosLiveViews";

const sheet = (rawRows: Array<Array<string | number | null>>): ExcelSheetSnapshot =>
  ({
    sheetIndex: 0,
    sheetName: "3. Виключені",
    rawRows,
    headerRows: [],
    rows: [],
    columnCount: 32,
    columnIndexes: [],
    dataStartRow: 6,
  }) as ExcelSheetSnapshot;

describe("parseExcluded LiveView columns", () => {
  it("reads A/B/C and AB–AF, not G/H enlist fields", () => {
    const row = Array.from({ length: 32 }, () => "");
    row[0] = "солдат";
    row[1] = "БАСОВСЬКИЙ Юрій Михайлович";
    row[2] = "12521";
    row[6] = "01.02.2024";
    row[7] = "668-РС";
    row[27] = "05.08.2026";
    row[28] = "05.08.2026";
    row[29] = "123";
    row[30] = "а4784 військової частини";
    row[31] = "ПЕРЕВЕДЕННЯ\n_ А4784";
    const parsed = parseExcluded(
      sheet([[], [], [], [], [], row]),
    );
    expect(parsed).toHaveLength(1);
    expect(parsed[0].fullName).toBe("БАСОВСЬКИЙ Юрій Михайлович");
    expect(parsed[0].personId).toBe("12521");
    expect(parsed[0].rank).toBe("солдат");
    expect(parsed[0].excludeDate).toBe("05.08.2026");
    expect(parsed[0].orderNumber).toBe("123");
    expect(parsed[0].destination).toContain("а4784");
    expect(parsed[0].fullName).not.toBe("01.02.2024");
  });
});

describe("OOS duplicate checks", () => {
  const workbookSheet = (
    sheetName: string,
    rawRows: Array<Array<string | number | null>>,
  ): ExcelSheetSnapshot =>
    ({
      sheetIndex: 0,
      sheetName,
      rawRows,
      headerRows: [],
      rows: [],
      columnCount: 8,
      columnIndexes: [],
      dataStartRow: 6,
    }) as ExcelSheetSnapshot;

  const workbookOf = (oosRows: Array<Array<string | number | null>>): ExcelWorkbookSnapshot =>
    ({
      file: new File([], "ejoos.xlsx"),
      fileName: "ejoos.xlsx",
      sheetName: "2. ООС",
      headerRows: [],
      rows: [],
      columnCount: 3,
      columnIndexes: [],
      dataStartRow: 6,
      sheets: [
        workbookSheet("1. ШПО", []),
        workbookSheet("2. ООС", oosRows),
        workbookSheet("6. Табель", []),
      ],
    }) as ExcelWorkbookSnapshot;

  it("blocks export when the same OOS ID appears twice", () => {
    const view = buildEjoosLiveView({
      workbook: workbookOf([
        [],
        [],
        [],
        [],
        [],
        ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", "12840"],
        ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", "12840"],
      ]),
    });
    const blocker = view.checks.find((item) => item.id === "dup-oos-id");
    expect(blocker?.severity).toBe("error");
    expect(blocker?.detail).toMatch(/12840/);
  });

  it("warns when the same ПІБ has a second card without ID", () => {
    const view = buildEjoosLiveView({
      workbook: workbookOf([
        [],
        [],
        [],
        [],
        [],
        ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", "12840"],
        ["солдат", "ДОБРОВОЛЬСЬКИЙ Володимир Миколайович", ""],
      ]),
    });
    const warn = view.checks.find((item) => item.id === "dup-oos-name");
    expect(warn?.severity).toBe("warn");
  });
});
