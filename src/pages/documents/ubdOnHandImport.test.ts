import { describe, expect, it } from "vitest";
import type { BackendPersonDocument } from "../../api";
import {
  formatUbdOnHandSheetBreakdown,
  matchUbdOnHandAgainstDocuments,
  parseUbdOnHandWorkbook,
  type UbdOnHandImportRow,
} from "./ubdOnHandImport";

const doc = (
  id: string,
  fullName: string,
  type = "ubdReport",
): BackendPersonDocument =>
  ({
    id,
    type,
    personExternalId: `p:${fullName.toLocaleLowerCase("uk-UA")}`,
    title: "УБД",
    status: "document",
    fields: { fullName },
    createdAt: "2026-01-01T00:00:00Z",
    updatedAt: "2026-01-01T00:00:00Z",
  }) as BackendPersonDocument;

describe("ubdOnHandImport", () => {
  it("parses split-name and full-name sheets", () => {
    const rows = parseUbdOnHandWorkbook({
      file: new File([], "ubd.xlsx"),
      sheets: [
        {
          sheetName: "Загальна",
          rawRows: [
            ["", "", "УБД в наявності"],
            ["", "Призвище", "Власне ім’я", "По батькові (за наявності)", ""],
            ["1", "Анчиполевський", "Віталій", "Олексійович", ""],
            ["2", "Березюк", "Дмитро", "Миколайович", "Рота охорони"],
          ],
          headerRows: [],
          rows: [],
          columns: [],
        },
        {
          sheetName: "СЗЧ",
          rawRows: [
            ["", "УБД СЗЧ"],
            ["1", "Ангелов Юрій Олександрович"],
          ],
          headerRows: [],
          rows: [],
          columns: [],
        },
      ],
    } as never);

    expect(rows.map((row) => row.fullName)).toEqual(
      expect.arrayContaining([
        "Анчиполевський Віталій Олексійович",
        "Березюк Дмитро Миколайович",
        "Ангелов Юрій Олександрович",
      ]),
    );
    expect(rows.find((row) => row.fullName.startsWith("Березюк"))?.note).toBe(
      "Рота охорони",
    );
  });

  it("matches imported names to journal documents by ПІБ", () => {
    const imported: UbdOnHandImportRow[] = [
      {
        sheetName: "Загальна",
        rowNumber: 4,
        lastName: "Анчиполевський",
        firstName: "Віталій",
        patronymic: "Олексійович",
        fullName: "Анчиполевський Віталій Олексійович",
        nameKey: "анчиполевський віталій олексійович",
        note: "",
      },
      {
        sheetName: "Загальна",
        rowNumber: 99,
        lastName: "Невідомий",
        firstName: "Іван",
        patronymic: "Іванович",
        fullName: "Невідомий Іван Іванович",
        nameKey: "невідомий іван іванович",
        note: "",
      },
    ];

    const report = matchUbdOnHandAgainstDocuments(imported, [
      doc("d1", "Анчиполевський Віталій Олексійович"),
      doc("d2", "Інша Людина"),
    ]);

    expect(report.matched).toHaveLength(1);
    expect(report.matched[0]?.documents[0]?.document.id).toBe("d1");
    expect(report.unmatched).toHaveLength(1);
    expect(report.unmatched[0]?.fullName).toContain("Невідомий");
    expect(report.matchedDocumentIds.has("d1")).toBe(true);
  });

  it("summarizes imported rows by sheet name", () => {
    const imported: UbdOnHandImportRow[] = [
      {
        sheetName: "Загальна",
        rowNumber: 4,
        lastName: "А",
        firstName: "Б",
        patronymic: "В",
        fullName: "А Б В",
        nameKey: "а б в",
        note: "",
      },
      {
        sheetName: "СЗЧ",
        rowNumber: 2,
        lastName: "Г",
        firstName: "Д",
        patronymic: "Е",
        fullName: "Г Д Е",
        nameKey: "г д е",
        note: "",
      },
    ];
    expect(formatUbdOnHandSheetBreakdown(imported)).toBe(
      "Загальна: 1 · СЗЧ: 1",
    );
  });
});
