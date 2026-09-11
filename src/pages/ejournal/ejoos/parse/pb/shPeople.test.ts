import { describe, expect, it } from "vitest";
import type { ExcelWorkbookSnapshot } from "../../../../excelRoundTrip";
import { parsePbShPeople } from "./shPeople";

describe("parsePbShPeople", () => {
  it("reads sh row with journal person id", () => {
    const workbook: ExcelWorkbookSnapshot = {
      fileName: "1ПБ_07092026.xlsx",
      sheets: [
        {
          sheetName: "sh",
          rawRows: [
            ["ID", "ПІБ", "Звання", "Індекс посади", "", "Статус"],
            ["21374", "Іванов І.І.", "солдат", "2103350", "", "21_МР"],
          ],
        },
      ],
    };

    expect(parsePbShPeople(workbook)).toEqual([
      {
        excelRow: 2,
        personId: "21374",
        fullName: "Іванов І.І.",
        rank: "солдат",
        positionIndex: "2103350",
        positionTitle: "",
        status: "21_МР",
        arrivedFrom: "",
      },
    ]);
  });
});
