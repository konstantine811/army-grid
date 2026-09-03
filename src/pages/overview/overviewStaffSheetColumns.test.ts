import { describe, expect, it } from "vitest";
import { STAFF_SHEET_EXPORT_COLUMN_NUMBERS } from "../excel-fill/staffSheet";
import { FIGHTER_STATUS_FIELDS } from "../personnel/fighterStatusImport";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/personnelUtils";
import {
  buildStaffSheetColumnsRecord,
  DEFAULT_OVERVIEW_STAFF_COLUMN_VISIBILITY,
  OVERVIEW_STAFF_SHEET_COLUMN_DEFS,
} from "./overviewStaffSheetColumns";

describe("overviewStaffSheetColumns", () => {
  it("includes additional columns without fields already shown in Overview", () => {
    expect(OVERVIEW_STAFF_SHEET_COLUMN_DEFS.length).toBe(
      STAFF_SHEET_EXPORT_COLUMN_NUMBERS.length +
        FIGHTER_STATUS_FIELDS.length -
        13,
    );
    expect(
      OVERVIEW_STAFF_SHEET_COLUMN_DEFS.some(
        (column) => column.columnNumber === 1,
      ),
    ).toBe(false);
    expect(
      OVERVIEW_STAFF_SHEET_COLUMN_DEFS.some(
        (column) => column.columnNumber === 2,
      ),
    ).toBe(false);
    expect(
      OVERVIEW_STAFF_SHEET_COLUMN_DEFS.some(
        (column) => column.fighterKey === "fighter_status_value",
      ),
    ).toBe(false);
    expect(
      OVERVIEW_STAFF_SHEET_COLUMN_DEFS.slice(0, 2).map(
        (column) => column.header,
      ),
    ).toEqual(["Місце перебування (кол. 31)", "Примітки"]);
    expect(DEFAULT_OVERVIEW_STAFF_COLUMN_VISIBILITY.staff_31).toBe(true);
    expect(DEFAULT_OVERVIEW_STAFF_COLUMN_VISIBILITY.staff_32).toBe(true);
  });

  it("builds staff column values from roster row", () => {
    const record = buildStaffSheetColumnsRecord({
      column_2: "1 рота",
      column_13: "солдат",
      column_14: "Іванов І.І.",
      fighter_status_exit_date: "01.01.2026",
      fighter_status_return_date: "10.01.2026",
    });

    expect(record.staff_2).toBe("1 рота");
    expect(record.staff_13).toBe("солдат");
    expect(record.staff_14).toBe("Іванов І.І.");
    expect(record.staff_fighter_fighter_status_exit_date).toBe("01.01.2026");
    expect(record.staff_fighter_fighter_status_total_days).toBe("9");
  });

  it("uses roster column labels for headers", () => {
    const platoonColumn = OVERVIEW_STAFF_SHEET_COLUMN_DEFS.find(
      (column) => column.columnNumber === 3,
    );
    expect(platoonColumn?.header).toBe(MORNING_GENERAL_LIST_COLUMN_LABELS[3]);
  });
});
