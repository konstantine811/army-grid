import { describe, expect, it } from "vitest";
import type { CellValue, ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  appendArchiveOnlyToRosterRows,
  findArchiveSheet,
  isPersonnelFromArchive,
  parseArchiveSheetToRosterRows,
} from "./staffSheetArchive";

const archiveSheet = (rows: CellValue[][]): ExcelSheetSnapshot => ({
  sheetIndex: 2,
  sheetName: "Архів",
  rawRows: rows,
  headerRows: [rows[0] ?? []],
  rows: rows.slice(1).map((values, index) => ({
    id: `archive-${index + 2}`,
    excelRowNumber: index + 2,
    values,
    source: "template",
  })),
  columnCount: 29,
  columnIndexes: Array.from({ length: 29 }, (_, index) => index),
  dataStartRow: 2,
});

const archivePerson = (name: string): CellValue[] => {
  const row: CellValue[] = Array(29).fill("");
  row[1] = "1 ПР";
  row[4] = "стрілець";
  row[10] = "солдат";
  row[12] = name;
  row[13] = "01.02.1990";
  row[15] = "СОКІЛ";
  row[17] = "СЗЧ";
  row[21] = "БЗВП +";
  row[22] = "так";
  row[24] = "2ПР";
  row[25] = "н.п. Павлоград";
  row[26] = "примітка";
  return row;
};

const legacyArchivePerson = (name: string): CellValue[] => {
  const row: CellValue[] = Array(34).fill("");
  row[1] = "2 піхотна рота";
  row[6] = "старший солдат";
  row[13] = name;
  row[14] = 31_171;
  row[15] = 41;
  row[16] = "РУДИЙ";
  row[20] = "Відком. за межі ПБ";
  return row;
};

describe("staffSheetArchive", () => {
  it("finds the archive sheet by name", () => {
    const sheet = archiveSheet([[]]);
    expect(findArchiveSheet([{ ...sheet, sheetName: "Інше" }, sheet])).toBe(
      sheet,
    );
  });

  it("maps archive columns to the general roster shape", () => {
    const [row] = parseArchiveSheetToRosterRows(
      archiveSheet([[], archivePerson("АРХІВНИЙ Петро Іванович")]),
    );

    expect(row.column_14).toBe("АРХІВНИЙ Петро Іванович");
    expect(row.column_13).toBe("солдат");
    expect(row.column_15).toBe("СОКІЛ");
    expect(row.column_16).toBe("01.02.1990");
    expect(row.column_21).toBe("СЗЧ");
    expect(row.column_26).toBe("БЗВП +");
    expect(row.column_27).toBe("так");
    expect(row.column_29).toBe("2ПР");
    expect(row.column_31).toBe("н.п. Павлоград");
    expect(row.column_32).toBe("примітка");
    expect(isPersonnelFromArchive(row)).toBe(true);
  });

  it("reads legacy archive blocks where ПІБ is in column N", () => {
    const [row] = parseArchiveSheetToRosterRows(
      archiveSheet([[], legacyArchivePerson("СТАРИЙ Петро Іванович")]),
    );

    expect(row.column_14).toBe("СТАРИЙ Петро Іванович");
    expect(row.column_13).toBe("старший солдат");
    expect(row.column_15).toBe("РУДИЙ");
    expect(row.column_16).toBe("1985-05-04T00:00:00.000Z");
    expect(row.column_21).toBe("Відком. за межі ПБ");
  });

  it("adds only people absent from the general list", () => {
    const main = [{ column_14: "ДУБЛЬ Іван Іванович" }];
    const merged = appendArchiveOnlyToRosterRows(
      main,
      archiveSheet([
        [],
        archivePerson("ДУБЛЬ Іван Іванович"),
        archivePerson("ЛИШЕ АРХІВ Марія Іванівна"),
      ]),
    );

    expect(merged.map((row) => row.column_14)).toEqual([
      "ДУБЛЬ Іван Іванович",
      "ЛИШЕ АРХІВ Марія Іванівна",
    ]);
    expect(isPersonnelFromArchive(merged[1])).toBe(true);
  });
});
