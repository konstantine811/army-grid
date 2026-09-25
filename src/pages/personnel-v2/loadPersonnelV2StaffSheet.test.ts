import { describe, expect, it } from "vitest";
import type { BackendPersonnelRosterLatest } from "../../api";
import { buildPersonnelV2StaffSheet } from "./loadPersonnelV2StaffSheet";

const roster = (
  rows: BackendPersonnelRosterLatest["rows"],
): BackendPersonnelRosterLatest => ({
  importId: "import-1",
  importName: "Штатка",
  sourceFileName: "Штатка.xlsx",
  createdAt: "2026-09-01T00:00:00.000Z",
  sheet: {
    id: "sheet-1",
    batchId: "import-1",
    name: "1.ОС Загальний список",
    sheetIndex: 0,
    columnCount: 21,
    rowCount: rows.length,
    createdAt: "2026-09-01T00:00:00.000Z",
    updatedAt: "2026-09-01T00:00:00.000Z",
  },
  rows,
});

describe("buildPersonnelV2StaffSheet", () => {
  it("reads staff names from the latest roster payload", () => {
    const sheet = buildPersonnelV2StaffSheet(
      roster([
        {
          id: "header",
          sheetId: "sheet-1",
          excelRowNumber: 1,
          values: { column_14: "ПІБ" },
          createdAt: "2026-09-01T00:00:00.000Z",
        },
        {
          id: "person-1",
          sheetId: "sheet-1",
          excelRowNumber: 2,
          values: {
            column_2: "1 рота",
            column_7: "Командир відділення",
            column_14: "ШЕВЧЕНКО Тарас Григорович",
            column_15: "Кобзар",
            column_21: "В строю",
            column_30: "додаткове поле",
          },
          createdAt: "2026-09-01T00:00:00.000Z",
        },
      ]),
    );

    expect(sheet.sheetName).toBe("1.ОС Загальний список");
    expect(sheet.inStaff).toBe(1);
    expect(sheet.people).toEqual([
      {
        id: "person-1",
        rowNumber: 2,
        name: "ШЕВЧЕНКО Тарас Григорович",
        inArchive: false,
        cells: [
          { column: 2, key: "column_2", label: "Підрозділ", value: "1 рота" },
          {
            column: 7,
            key: "column_7",
            label: "Повна посада",
            value: "Командир відділення",
          },
          {
            column: 14,
            key: "column_14",
            label: "ПІБ",
            value: "ШЕВЧЕНКО Тарас Григорович",
          },
          { column: 15, key: "column_15", label: "Позивний", value: "Кобзар" },
          { column: 21, key: "column_21", label: "Статус", value: "В строю" },
          {
            column: 30,
            key: "column_30",
            label: "Колонка 30",
            value: "додаткове поле",
          },
        ],
      },
    ]);
  });
});
