import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import { parseExcluded } from "./ejoosLiveViews";

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
