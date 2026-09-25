import { describe, expect, it } from "vitest";
import {
  buildStaffListSnapshot,
  defaultStaffListGapColumnIds,
} from "./staffListGapsParse";
import {
  buildAnketaRowLookup,
  buildStaffListSyntheticAnketaRow,
  isStaffListSyntheticAnketaRow,
  resolveStaffListAnketaRow,
} from "./staffListGapsAnketaMatch";
import type { StaffListRow } from "./staffListGapsParse";
import {
  countStaffListEmptyCells,
  findNextStaffListEmptyCell,
  resolveStaffListGapColumnIds,
} from "./staffListGapsSearch";

describe("staffListGapsParse", () => {
  it("builds rows from a simple header sheet", () => {
    const snapshot = buildStaffListSnapshot({
      fileName: "test.xlsx",
      sheetName: "Sheet1",
      file: new File([], "test.xlsx"),
      headerRows: [
        [
          "№",
          "посада",
          "звання",
          "П.І.Б.",
          "Дата народження",
          "Місце народження",
          "Підрозділ",
          "№ телефону",
        ],
      ],
      rows: [
        {
          id: "r1",
          excelRowNumber: 2,
          values: [
            8,
            "інструктор",
            "сержант",
            "ГУГУЄВ Павло Сергійович",
            28389,
            "Біла Калitva",
            "_5 1ПБ",
            "",
          ],
          source: "template",
        },
        {
          id: "r2",
          excelRowNumber: 3,
          values: [
            237,
            "командир",
            "лейтенант",
            "ГАЛУШКО Олексій Олександрович",
            29605,
            "Харків",
            "_5 1ПБ",
            "0501234567",
          ],
          source: "template",
        },
      ],
      columnCount: 8,
      columnIndexes: [0, 1, 2, 3, 4, 5, 6, 7],
      dataStartRow: 2,
      sheets: [],
    });

    expect(snapshot.rows).toHaveLength(2);
    expect(snapshot.pibColumnId).toBeTruthy();
    expect(snapshot.rows[0]?.pib).toContain("ГУГУЄВ");
    expect(snapshot.rows[0]?.birthDate).toMatch(/\d{2}\.\d{2}\.\d{4}/);
    expect(defaultStaffListGapColumnIds(snapshot).length).toBeGreaterThanOrEqual(3);
  });
});

describe("staffListGapsSearch", () => {
  const snapshot = buildStaffListSnapshot({
    fileName: "test.xlsx",
    sheetName: "Sheet1",
    file: new File([], "test.xlsx"),
    headerRows: [["П.І.Б.", "№ телефону"]],
    rows: [
      {
        id: "r1",
        excelRowNumber: 2,
        values: ["ГУГУЄВ Павло Сергійович", ""],
        source: "template",
      },
      {
        id: "r2",
        excelRowNumber: 3,
        values: ["ГАЛУШКО Олексій Олександрович", "0501234567"],
        source: "template",
      },
    ],
    columnCount: 2,
    columnIndexes: [0, 1],
    dataStartRow: 2,
    sheets: [],
  });
  const phoneColumnId = defaultStaffListGapColumnIds(snapshot).find((id) => {
    const column = snapshot.columns.find((item) => item.id === id);
    return column && /телефон/i.test(column.label.toLocaleLowerCase("uk-UA"));
  })!;

  it("finds the first empty phone cell", () => {
    const next = findNextStaffListEmptyCell(
      snapshot.rows,
      snapshot.columns,
      [phoneColumnId],
      null,
    );
    expect(next?.pib).toContain("ГУГУЄВ");
    expect(next?.columnId).toBe(phoneColumnId);
  });

  it("finds dash placeholder in birth place when column is in gap scope", () => {
    const fullSnapshot = buildStaffListSnapshot({
      fileName: "test.xlsx",
      sheetName: "Sheet1",
      file: new File([], "test.xlsx"),
      headerRows: [
        ["П.І.Б.", "Дата народження", "Місце народження", "№ телефону"],
      ],
      rows: [
        {
          id: "r1",
          excelRowNumber: 2,
          values: ["КЛОКОВ Сергій Анатолійович", 28389, "-", "0501112233"],
          source: "template",
        },
      ],
      columnCount: 4,
      columnIndexes: [0, 1, 2, 3],
      dataStartRow: 2,
      sheets: [],
    });
    const phoneId = defaultStaffListGapColumnIds(fullSnapshot).find((id) =>
      /телефон/i.test(
        fullSnapshot.columns.find((column) => column.id === id)?.label ?? "",
      ),
    )!;
    const birthPlaceId = defaultStaffListGapColumnIds(fullSnapshot).find((id) =>
      /місце/i.test(
        fullSnapshot.columns.find((column) => column.id === id)?.label ?? "",
      ),
    )!;
    const gapColumnIds = resolveStaffListGapColumnIds(fullSnapshot, [phoneId]);
    expect(gapColumnIds).toContain(birthPlaceId);
    const next = findNextStaffListEmptyCell(
      fullSnapshot.rows,
      fullSnapshot.columns,
      gapColumnIds,
      null,
    );
    expect(next?.columnId).toBe(birthPlaceId);
  });
});

describe("staffListGapsAnketaMatch", () => {
  const staffRow: StaffListRow = {
    __rowId: "r2",
    __rowNumber: 2,
    __excelRowNumber: 3,
    pib: "ГАЛУШКО Олексій Олександрович",
    birthDate: "11.05.1981",
    values: {},
  };

  it("falls back to synthetic anketa row when sheet lookup misses", () => {
    const lookup = buildAnketaRowLookup([]);
    const resolved = resolveStaffListAnketaRow(lookup, staffRow);
    expect(isStaffListSyntheticAnketaRow(resolved)).toBe(true);
    expect(resolved.fullName).toBe(staffRow.pib);
    expect(resolved.birthDate).toBe(staffRow.birthDate);
  });

  it("builds synthetic row with stable id", () => {
    const row = buildStaffListSyntheticAnketaRow(staffRow);
    expect(row.__rowId).toBe("staff-list:r2");
  });
});
