import { describe, expect, it } from "vitest";
import type { CellValue, ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  buildEjoosDataTableModel,
  findEjoosHeaderRowIndex,
} from "./ejoosSheetDataTable";

const sheet = (
  sheetName: string,
  rawRows: CellValue[][],
): ExcelSheetSnapshot => ({
  sheetIndex: 0,
  sheetName,
  rawRows,
  headerRows: [],
  rows: [],
  columnCount: 40,
  columnIndexes: Array.from({ length: 40 }, (_, index) => index),
  dataStartRow: 6,
});

describe("buildEjoosDataTableModel", () => {
  it("skips the field-number row and keeps real Excel row numbers", () => {
    const input = sheet("2. ООС", [
      [],
      ["ОБЛІК ОСОБОВОГО СКЛАДУ"],
      [],
      ["Звання", "Прізвище, імʼя, по батькові", "ID", "Індекс посади"],
      ["1", "2", "3", "4"],
      ["солдат", "ІВАНЕНКО Іван Іванович", "101", "2103001"],
      [],
      ["сержант", "ПЕТРЕНКО Петро Петрович", "202", "2103002"],
    ]);

    expect(findEjoosHeaderRowIndex(input)).toBe(3);
    const model = buildEjoosDataTableModel(input);
    expect(model.dataStartIndex).toBe(5);
    expect(model.nameColumnId).toBe("column_1");
    expect(model.rows.map((row) => row.__excelRowNumber)).toEqual([6, 8]);
    expect(model.rows.map((row) => row.__rowId)).toEqual([
      "2. ООС:6",
      "2. ООС:8",
    ]);
  });

  it("does not truncate OOS or Excluded to the old 800-row limit", () => {
    const data = Array.from({ length: 1_050 }, (_, index) => [
      "солдат",
      `ОСОБА ${index} Тестович`,
      String(index + 1),
      `210${String(index).padStart(4, "0")}`,
    ]);
    const model = buildEjoosDataTableModel(
      sheet("3. Виключені", [
        [],
        [],
        [],
        ["Звання", "ПІБ", "ID", "Індекс посади"],
        ["1", "2", "3", "4"],
        ...data,
      ]),
    );

    expect(model.rows).toHaveLength(1_050);
    expect(model.rows.at(-1)?.__excelRowNumber).toBe(1_055);
  });

  it("creates unique column IDs when headers are empty or repeated", () => {
    const model = buildEjoosDataTableModel(
      sheet("3. Виключені", [
        ["ПІБ", "", "Дата", "Дата"],
        ["ІВАНЕНКО Іван Іванович", "", "01.01.2026", "02.01.2026"],
      ]),
    );

    expect(model.columns.map((column) => column.id)).toEqual([
      "column_0",
      "column_1",
      "column_2",
      "column_3",
    ]);
    expect(new Set(model.columns.map((column) => column.id)).size).toBe(4);
  });
});
