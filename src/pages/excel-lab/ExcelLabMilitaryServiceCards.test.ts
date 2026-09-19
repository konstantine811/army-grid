import { describe, expect, it } from "vitest";
import { parseExcelLabMilitaryServiceCards } from "./ExcelLabMilitaryServiceCards";
import type { ExcelSheetSnapshot } from "@/excelRoundTrip";

describe("parseExcelLabMilitaryServiceCards", () => {
  it("picks alias from first non-empty column among 7, 8 and 9", () => {
    const sheet: ExcelSheetSnapshot = {
      sheetIndex: 0,
      sheetName: "ТПВ",
      rawRows: [],
      headerRows: [],
      columnCount: 32,
      columnIndexes: [],
      dataStartRow: 2,
      rows: [
        {
          id: "tpv-2",
          excelRowNumber: 2,
          source: "template",
          values: Object.assign(Array(32).fill(""), {
            6: "ПЕТРЕНКО Іван Іванович",
            7: null,
            8: "#N/A",
            9: "БАЛЯ",
            13: "ТПВ №4160",
            20: "активний",
          }),
        },
      ],
    };

    const cards = parseExcelLabMilitaryServiceCards([sheet]);
    expect(cards["ПЕТРЕНКО Іван Іванович"]?.ТПВ?.alias).toBe("БАЛЯ");
  });

  it("accepts numeric alias cells without calling trim on non-strings", () => {
    const sheet: ExcelSheetSnapshot = {
      sheetIndex: 0,
      sheetName: "ТПВ",
      rawRows: [],
      headerRows: [],
      columnCount: 32,
      columnIndexes: [],
      dataStartRow: 2,
      rows: [
        {
          id: "tpv-3",
          excelRowNumber: 3,
          source: "template",
          values: Object.assign(Array(32).fill(""), {
            6: "ІВАНОВ Іван Іванович",
            7: null,
            8: "#N/A",
            9: 12345,
            13: "ТПВ №100",
            20: "активний",
          }),
        },
      ],
    };

    const cards = parseExcelLabMilitaryServiceCards([sheet]);
    expect(cards["ІВАНОВ Іван Іванович"]?.ТПВ?.alias).toBe("12345");
  });
});
