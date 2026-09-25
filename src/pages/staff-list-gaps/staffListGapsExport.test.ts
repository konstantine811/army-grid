import { describe, expect, it } from "vitest";
import {
  buildStaffListExportFileName,
  staffListCellValueForExcelWrite,
} from "./staffListGapsExport";
import type { StaffListSnapshot } from "./staffListGapsParse";

describe("staffListGapsExport", () => {
  const snapshot: StaffListSnapshot = {
    fileName: "1 ПБ список.xlsx",
    sheetName: "Sheet1",
    sheetIndex: 0,
    columnIndexes: [0, 1, 2, 3, 4, 5, 6, 7],
    columns: [
      { id: "col_0", label: "№", index: 0 },
      { id: "phone_7", label: "№ телефону", index: 7 },
      { id: "birth_4", label: "Дата народження", index: 4 },
    ],
    rows: [],
    pibColumnId: "pib_3",
    birthDateColumnId: "birth_4",
  };

  it("builds export file name from original", () => {
    expect(buildStaffListExportFileName("list.xlsx")).toBe("list_filled.xlsx");
  });

  it("writes numbers and dates for excel cells", () => {
    expect(
      staffListCellValueForExcelWrite(snapshot.columns[0]!, snapshot, "237"),
    ).toBe(237);
    expect(
      staffListCellValueForExcelWrite(snapshot.columns[2]!, snapshot, "11.05.1981"),
    ).toBe("11.05.1981");
    expect(
      staffListCellValueForExcelWrite(snapshot.columns[1]!, snapshot, "0501234567"),
    ).toBe("0501234567");
  });
});
