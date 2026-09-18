import { describe, expect, it } from "vitest";
import type { PbArchivePeriod } from "../types/pb";
import type { EjoosAbsentRow } from "../types/workbookRows";
import { planAbsentArchiveOps } from "./absentFromArchive";

const mapStatus = (raw: string) => ({
  timesheetCode: /поран/i.test(raw) ? "ЛП" : raw === "В СТРОЮ" ? "+" : "",
});

const baseInput = () => ({
  archive: [] as PbArchivePeriod[],
  ejoosAbsents: [] as EjoosAbsentRow[],
  openById: new Map<string, EjoosAbsentRow>(),
  openByName: new Map<string, EjoosAbsentRow>(),
  shPersonById: new Map(),
  shPersonByName: new Map(),
  absenceRowsClosedByMovement: new Set<number>(),
  existingOps: [],
  mapStatus,
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) =>
    Boolean(left.personId && right.personId && left.personId === right.personId),
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    _fullName: string,
  ) => (personId ? map.get(personId) : null),
  laterArchivePeriodOf: () => null,
  alreadyVacatedForAbsence: () => false,
  inboundStaffPlacementThisMonth: () => false,
  activeTimesheetRowOf: () => ({ excelRow: 42 }),
  staffEpisodePaintPayload: () => ({
    timesheetActiveFrom: "03.09.2026",
    timesheetPreserveHistory: "",
    timesheetAbsenceSpans: "3-30:ЛП",
    historyTimesheetExcelRow: "",
    historyTimesheetAbsenceSpans: "",
  }),
});

describe("planAbsentArchiveOps timesheet refresh", () => {
  it("still emits absent_upsert when «Тимч. відсутні» already matches archive", () => {
    const period: PbArchivePeriod = {
      excelRow: 10,
      personId: "p1",
      fullName: "ДАВИДЕНКО Олександр Володимирович",
      rank: "солдат",
      absenceType: "ПОРАНЕННЯ",
      place: "госпіталь Петропavlivka",
      departDate: "03.09.2026",
      orderNumber: "259",
      orderDate: "03.09.2026",
      plannedReturn: "",
      returnDate: "",
      returnOrderNumber: "",
      returnOrderDate: "",
      periodNumber: "1",
      positionTitle: "",
    };
    const absentRow: EjoosAbsentRow = {
      excelRow: 5,
      personId: "p1",
      fullName: "ДАВИДЕНКО Олександр Володимирович",
      positionIndex: "101",
      rank: "солдат",
      ground: "ПОРАНЕННЯ",
      place: "госпіталь Петропavlivka",
      departDate: "03.09.2026",
      actualReturn: "",
    };
    const ops = planAbsentArchiveOps({
      ...baseInput(),
      archive: [period],
      ejoosAbsents: [absentRow],
      openById: new Map([["p1", absentRow]]),
    });

    expect(ops).toHaveLength(1);
    expect(ops[0].kind).toBe("absent_upsert");
    expect(ops[0].payload.timesheetAbsenceSpans).toBe("3-30:ЛП");
    expect(ops[0].payload.timesheetExcelRow).toBe("42");
    expect(ops[0].checkedDefault).toBe(true);
  });

  it("skips when absent row matches and tabell has nothing to paint", () => {
    const period: PbArchivePeriod = {
      excelRow: 10,
      personId: "p1",
      fullName: "ІВАНОВ Іван Іванович",
      rank: "солдат",
      absenceType: "ПОРАНЕННЯ",
      place: "шпиталь",
      departDate: "03.09.2026",
      orderNumber: "",
      orderDate: "",
      plannedReturn: "",
      returnDate: "",
      returnOrderNumber: "",
      returnOrderDate: "",
      periodNumber: "1",
      positionTitle: "",
    };
    const absentRow: EjoosAbsentRow = {
      excelRow: 5,
      personId: "p1",
      fullName: "ІВАНОВ Іван Іванович",
      positionIndex: "101",
      rank: "солдат",
      ground: "ПОРАНЕННЯ",
      place: "шпиталь",
      departDate: "03.09.2026",
      actualReturn: "",
    };
    const ops = planAbsentArchiveOps({
      ...baseInput(),
      archive: [period],
      ejoosAbsents: [absentRow],
      openById: new Map([["p1", absentRow]]),
      activeTimesheetRowOf: () => null,
      staffEpisodePaintPayload: () => ({
        timesheetActiveFrom: "",
        timesheetPreserveHistory: "",
        timesheetAbsenceSpans: "",
        historyTimesheetExcelRow: "",
        historyTimesheetAbsenceSpans: "",
      }),
    });

    expect(ops).toHaveLength(0);
  });
});
