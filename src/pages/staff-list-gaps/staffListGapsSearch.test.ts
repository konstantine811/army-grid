import { describe, expect, it } from "vitest";
import {
  countStaffListEmptyCells,
  findNextStaffListEmptyCell,
  isStaffListCellEmpty,
} from "./staffListGapsSearch";
import type { StaffListColumn, StaffListRow } from "./staffListGapsParse";

const row = (values: Record<string, string>): StaffListRow => ({
  __rowId: "r1",
  __rowNumber: 1,
  __excelRowNumber: 2,
  pib: "ТЕСТ Тест Тестович",
  birthDate: "01.01.1990",
  values,
});

const columns: StaffListColumn[] = [
  { id: "phone", label: "№ телефону", index: 7 },
];

describe("isStaffListCellEmpty", () => {
  it("treats blank and dash-only values as empty", () => {
    expect(isStaffListCellEmpty(row({ phone: "" }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: "   " }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: "-" }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: "—" }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: " – " }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: " - - " }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: "―" }), "phone")).toBe(true);
    expect(isStaffListCellEmpty(row({ phone: "немає" }), "phone")).toBe(true);
  });

  it("keeps real values as filled", () => {
    expect(isStaffListCellEmpty(row({ phone: "0501234567" }), "phone")).toBe(
      false,
    );
    expect(isStaffListCellEmpty(row({ phone: "АА-123456" }), "phone")).toBe(
      false,
    );
  });
});

describe("findNextStaffListEmptyCell", () => {
  it("finds cells with dash placeholder", () => {
    const rows = [row({ phone: "-" })];
    const next = findNextStaffListEmptyCell(rows, columns, ["phone"], null);
    expect(next?.columnId).toBe("phone");
    expect(countStaffListEmptyCells(rows, ["phone"])).toBe(1);
  });
});
