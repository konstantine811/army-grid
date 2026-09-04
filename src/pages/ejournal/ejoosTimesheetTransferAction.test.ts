import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import { buildTimesheetLayout } from "./ejoosTimesheetLayout";
import {
  assertTimesheetTransferAction,
  resolveTimesheetTransferAction,
  TimesheetTransferError,
} from "./ejoosTimesheetTransferAction";

const sheetOf = (rows: string[][]): ExcelSheetSnapshot =>
  ({
    sheetName: "6. Табель",
    rawRows: rows,
    columnIndexes: [],
  }) as ExcelSheetSnapshot;

const emptyRow = () => Array.from({ length: 40 }, () => "");

const personRow = (
  index: string,
  name: string,
  id: string,
  days: string[] = [],
) => {
  const row = emptyRow();
  row[1] = index;
  row[6] = name;
  row[7] = id;
  days.forEach((mark, offset) => {
    row[8 + offset] = mark;
  });
  return row;
};

describe("resolveTimesheetTransferAction", () => {
  const lekay = {
    personId: "18801",
    fullName: "ЛЕКАЙ Володимир Петрович",
    fromPositionIndex: "2103777",
    payload: {
      shpoExcelRow: "40",
      oosExcelRow: "12",
      timesheetDestination: "в/ч А4784",
      destination: "А4784",
      orderNumber: "240",
      orderDate: "19.08.2026",
    },
  };

  it("moves an open monthly row found by personId", () => {
    const rows = Array.from({ length: 30 }, emptyRow);
    rows[5] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[19] = personRow("2103777", "ЛЕКАЙ Володимир Петрович", "18801");
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet);
    const action = resolveTimesheetTransferAction({
      ...lekay,
      layout,
      sheet,
    });
    expect(action.kind).toBe("MOVE_TO_HISTORY");
    expect(action.sourceRow).toBe(20);
    expect(action.createTimesheetHistory).toBe(false);
    expect(action.replaceInPlace).toBe(false);
  });

  it("moves an open monthly row found by exact canonical name", () => {
    const rows = Array.from({ length: 30 }, emptyRow);
    rows[5] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[19] = personRow("2103777", "ЛЕКАЙ Володимир Петрович (01.01.1990 р.н.)", "");
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet);
    const action = resolveTimesheetTransferAction({
      ...lekay,
      personId: "",
      layout,
      sheet,
    });
    expect(action.kind).toBe("MOVE_TO_HISTORY");
    expect(action.sourceRow).toBe(20);
  });

  it("patches a closed history row for the current transfer", () => {
    const rows = Array.from({ length: 30 }, emptyRow);
    rows[5] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[21] = personRow("2103777", "ЛЕКАЙ Володимир Петрович", "18801", [
      "+",
      "вибув у в/ч А4784 наказ №240",
    ]);
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet);
    const action = resolveTimesheetTransferAction({
      ...lekay,
      layout,
      sheet,
    });
    expect(action.kind).toBe("PATCH_HISTORY");
    expect(action.sourceRow).toBe(22);
    expect(action.replaceInPlace).toBe(true);
  });

  it("creates history in the source section when only the canonical slot exists", () => {
    const rows = Array.from({ length: 30 }, emptyRow);
    rows[5] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[19] = ["1 ПІХОТНА РОТА", "2103777", "Стрілець", "100915А", "5"];
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet);
    const action = resolveTimesheetTransferAction({
      ...lekay,
      layout,
      sheet,
    });
    expect(action.kind).toBe("CREATE_HISTORY_IN_SOURCE_SECTION");
    expect(action.sourceIndex).toBe("2103777");
    expect(action.sourceRow).toBe(20);
    expect(action.createTimesheetHistory).toBe(true);
  });

  it("throws when no Tab source can be resolved", () => {
    const rows = Array.from({ length: 20 }, emptyRow);
    rows[5] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[6] = personRow("2103001", "ІНШИЙ Солдат", "1");
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet);
    expect(() =>
      resolveTimesheetTransferAction({
        ...lekay,
        payload: { destination: "А4784" },
        layout,
        sheet,
      }),
    ).toThrow(TimesheetTransferError);
  });

  it("rejects the incomplete exclude hint that would write Виключені only", () => {
    expect(() =>
      assertTimesheetTransferAction({
        kind: "MOVE_TO_HISTORY",
        sourceIndex: "",
        sourceRow: 0,
        sourceSection: "",
        targetHistoryRow: 0,
        createTimesheetHistory: false,
        replaceInPlace: false,
      }),
    ).toThrow(/TIMESHEET_SOURCE_UNRESOLVED|Виключені не пишемо/);
    expect(() => assertTimesheetTransferAction(null)).toThrow(
      "EXTERNAL_TRANSFER_WITHOUT_TIMESHEET_ACTION",
    );
  });
});
