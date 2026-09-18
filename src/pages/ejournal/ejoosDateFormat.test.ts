import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  buildEjoosDateCellWrites,
  findAbsentDateColumns,
  findArrivalDateColumns,
} from "./ejoosDateFormat";

const ejoosSheet = (
  sheetName: string,
  rawRows: unknown[][],
): ExcelSheetSnapshot => ({
  sheetIndex: 0,
  sheetName,
  rawRows: rawRows as ExcelSheetSnapshot["rawRows"],
  headerRows: [],
  rows: [],
  columnCount: 20,
  columnIndexes: Array.from({ length: 20 }, (_, index) => index),
  dataStartRow: 6,
});

describe("ejoosDateFormat", () => {
  it("finds arrival date columns from headers", () => {
    const columns = findArrivalDateColumns(
      ejoosSheet("4. Тимчасово прибулі", [
        [],
        [],
        [],
        [],
        ["Звання", "ПІБ", "ID", "Індекс", "Звідки", "Підрозділ", "", "Дата прибуття"],
      ]),
    );
    expect(columns).toContain(8);
  });

  it("builds string writes for Excel serial arrival dates", () => {
    const writes = buildEjoosDateCellWrites(
      ejoosSheet("4. Тимчасово прибулі", [
        [],
        [],
        [],
        [],
        ["", "", "", "", "", "", "", "Дата прибуття"],
        ["солдат", "ІВАНЕНКО", "101", "2103001", "БРЕЗ", "", "", 46193],
      ]),
    );
    expect(writes).toEqual([
      { row: 6, column: 8, value: "20.06.2026" },
    ]);
  });

  it("finds absent date columns from headers and defaults", () => {
    const fromHeaders = findAbsentDateColumns(
      ejoosSheet("5. Тимчасово відсутні", [
        [],
        [],
        [],
        [],
        [
          "Звання",
          "ПІБ",
          "ID",
          "Індекс",
          "Підстава",
          "Місце",
          "Дата вибуття",
          "Дата наказу",
        ],
      ]),
    );
    expect(fromHeaders).toEqual(expect.arrayContaining([7, 8]));

    const defaults = findAbsentDateColumns(
      ejoosSheet("5. Тимчасово відсутні", [[], [], [], [], []]),
    );
    expect(defaults).toEqual([7, 8, 12, 13, 15]);
  });

  it("builds string writes for Excel serial absent dates", () => {
    const writes = buildEjoosDateCellWrites(
      ejoosSheet("5. Тимчасово відсутні", [
        [],
        [],
        [],
        [],
        ["", "", "", "", "", "", "Дата вибуття", "Дата наказу"],
        ["солдат", "ІВАНЕНКО", "101", "2103001", "СЗЧ", "", 46205, 46209],
      ]),
      { dateColumns: findAbsentDateColumns(
          ejoosSheet("5. Тимчасово відсутні", [
            [],
            [],
            [],
            [],
            ["", "", "", "", "", "", "Дата вибуття", "Дата наказу"],
            ["солдат", "ІВАНЕНКО", "101", "2103001", "СЗЧ", "", 46205, 46209],
          ]),
        ) },
    );
    expect(writes).toEqual([
      { row: 6, column: 7, value: "02.07.2026" },
      { row: 6, column: 8, value: "06.07.2026" },
    ]);
  });
});
