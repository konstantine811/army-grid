import { describe, expect, it } from "vitest";
import { DEFAULT_STATUS_RULES } from "../../ejoosRules";
import { mapPbStatusToEjoosWithRules } from "../../ejoosStatusMap";
import { planAbsentCloseFromShOps } from "./absentCloseFromSh";
import { planTimesheetDayFromShOps } from "./timesheetDayFromSh";

const mapStatus = (raw: string) =>
  mapPbStatusToEjoosWithRules(raw, DEFAULT_STATUS_RULES);

describe("absent close / timesheet day plan", () => {
  it("planAbsentCloseFromShOps closes open absence when sh is В СТРОЮ", () => {
    const absenceRowsClosedByMovement = new Set<number>();
    const ops = planAbsentCloseFromShOps({
      shPeople: [
        {
          excelRow: 10,
          personId: "501",
          fullName: "Тестов Т.Т.",
          positionIndex: "2103001",
          rank: "солдат",
          status: "В СТРОЮ",
        },
      ],
      archive: [],
      openById: new Map([
        [
          "501",
          {
            excelRow: 42,
            personId: "501",
            fullName: "Тестов Т.Т.",
            positionIndex: "2103001",
            rank: "солдат",
            ground: "СЗЧ",
            departDate: "05.08.2026",
            place: "додому",
          },
        ],
      ]),
      openByName: new Map(),
      arrivalById: new Map(),
      arrivalByName: new Map(),
      absenceRowsClosedByMovement,
      timesheetDayLabel: "10.08.2026",
      mapStatus,
      dateInLeadWindow: () => true,
      alreadyVacatedForAbsence: () => false,
      isSamePerson: (left, right) =>
        Boolean(left.personId && left.personId === right.personId),
      byPersonName: () => null,
      activeTimesheetRowOf: () => ({ excelRow: 100 }),
      augustAbsenceSpansFor: () => [],
    });
    expect(ops).toHaveLength(1);
    expect(ops[0]?.kind).toBe("absent_close");
    expect(absenceRowsClosedByMovement.has(42)).toBe(true);
  });

  it("planTimesheetDayFromShOps skips people with position events", () => {
    const ops = planTimesheetDayFromShOps({
      shPeople: [
        {
          excelRow: 11,
          personId: "502",
          fullName: "Посадов П.",
          positionIndex: "2103002",
          rank: "сержант",
          status: "В СТРОЮ",
        },
      ],
      ejoosDays: [],
      archiveAll: [],
      timesheetDay: 10,
      timesheetDayLabel: "10.08.2026",
      dayByIndex: new Map(),
      shpoByIndex: new Map(),
      timesheetScanByRow: new Map(),
      arrivalById: new Map(),
      arrivalByName: new Map(),
      mapStatus,
      alreadyVacatedForAbsence: () => false,
      positionEventForShPerson: () => true,
      isSamePerson: () => false,
      byPersonName: () => null,
      timesheetRowsOf: () => [],
      activeTimesheetRowOf: () => null,
      augustAbsenceSpansFor: () => [],
      staffEpisodePaintPayload: () => ({
        timesheetActiveFrom: "",
        timesheetPreserveHistory: "",
        timesheetAbsenceSpans: "",
        historyTimesheetExcelRow: "",
        historyTimesheetAbsenceSpans: "",
      }),
      findFalseHopExcludedRow: () => null,
    });
    expect(ops).toHaveLength(0);
  });
});
