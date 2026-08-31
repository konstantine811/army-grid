import { describe, expect, it } from "vitest";
import {
  findDuplicateTimesheetExtras,
  isClosedTimesheetHistoryRow,
  pickCanonicalTimesheetRow,
  type TimesheetDuplicateScan,
} from "./ejoosTimesheetDuplicates";

const row = (
  excelRow: number,
  name: string,
  plus: number[],
  extra?: Partial<TimesheetDuplicateScan>,
): TimesheetDuplicateScan => ({
  excelRow,
  personId: extra?.personId ?? String(excelRow),
  fullName: name,
  positionIndex: extra?.positionIndex ?? "2103764",
  hasDepartureText: extra?.hasDepartureText ?? false,
  plusDays: plus,
});

describe("timesheet duplicates", () => {
  it("treats вибув without + as closed history, not a second active record", () => {
    expect(
      isClosedTimesheetHistoryRow(
        row(10, "ХУБАЄВ", [], { hasDepartureText: true }),
      ),
    ).toBe(true);
  });

  it("keeps one ХУБАЄВ and flags the copied + rows", () => {
    const extras = findDuplicateTimesheetExtras([
      row(931, "АТРАХОВ", [], { personId: "21155" }),
      row(932, "ХУБАЄВ", [29, 30, 31], { personId: "24867" }),
      row(933, "ХУБАЄВ", [29, 30, 31], { personId: "24867" }),
      row(934, "ХУБАЄВ", [29, 30, 31], { personId: "24867" }),
      row(935, "АТРАХОВ", [29, 30, 31], { personId: "21155" }),
    ]);
    const hubayevExtras = extras.filter((item) => item.extra.personId === "24867");
    expect(hubayevExtras).toHaveLength(2);
    expect(new Set(hubayevExtras.map((item) => item.keep.excelRow)).size).toBe(1);
  });

  it("prefers the current sh occupant on a duplicated staff index", () => {
    const keep = pickCanonicalTimesheetRow(
      [
        row(931, "АТРАХОВ", [], { personId: "21155" }),
        row(935, "АТРАХОВ", [31], { personId: "21155" }),
        row(932, "ХУБАЄВ", [31], { personId: "24867" }),
      ],
      { personId: "21155", fullName: "АТРАХОВ" },
    );
    expect(keep?.excelRow).toBe(935);
  });
});
