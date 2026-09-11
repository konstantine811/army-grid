import { describe, expect, it } from "vitest";
import { planTempArrivalCloseOps } from "./tempArrivalClose";
import type { EjoosSyncOp } from "../types/syncOp";

describe("planTempArrivalCloseOps", () => {
  it("creates position_change with isTempArrivalPlacement when sh matches ШПО", () => {
    const ops = planTempArrivalCloseOps({
      shPeople: [
        {
          excelRow: 2,
          personId: "21374",
          fullName: "МАССАЙ Вадим",
          rank: "солдат",
          positionIndex: "2103350",
          positionTitle: "",
          status: "21_МР",
          arrivedFrom: "",
        },
      ],
      movementsAll: [],
      shpoByIndex: new Map([
        [
          "2103350",
          {
            excelRow: 7,
            positionIndex: "2103350",
            personId: "21374",
            fullName: "МАССАЙ Вадим",
            rank: "солдат",
          },
        ],
      ]),
      ejoosOos: [],
      dayByIndex: new Map([
        [
          "2103350",
          {
            excelRow: 7,
            personId: "21374",
            fullName: "МАССАЙ Вадим",
            rank: "солдат",
            positionIndex: "2103350",
            dayValue: "+",
          },
        ],
      ]),
      oosById: new Map(),
      oosByName: new Map(),
      journalMonthStartLabel: "01.09.2026",
      existingOps: [],
      arrivalOf: () => ({
        excelRow: 6,
        personId: "21374",
        fullName: "МАССАЙ Вадим",
        rank: "солдат",
        positionIndex: "",
        arriveDate: "13.05.2026",
        fromUnit: "БРЕЗ",
      }),
      isSamePerson: (left, right) =>
        Boolean(left.personId && left.personId === right.personId),
      eventInLeadWindow: () => true,
      movementEventTime: () => 0,
      isPositionIndex: (value) => /^\d{5,}$/.test(value),
      byPersonName: () => null,
      staffIndexTimesheetForPerson: () => null,
      staffAppointmentDateFor: () => "",
      timesheetEpisodeStartFor: () => "",
      inboundStaffDateFor: () => "",
      staffEpisodePaintPayload: () => ({
        timesheetActiveFrom: "07.09.2026",
        timesheetPreserveHistory: "",
        timesheetAbsenceSpans: "",
        historyTimesheetExcelRow: "",
        historyTimesheetAbsenceSpans: "",
      }),
    });

    expect(ops).toHaveLength(1);
    expect(ops[0]).toMatchObject({
      kind: "position_change",
      class: "ready",
      payload: {
        isTempArrivalPlacement: "1",
        arrivalExcelRow: "6",
        nextIndex: "2103350",
      },
    } satisfies Partial<EjoosSyncOp>);
  });

  it("skips when position_change already planned", () => {
    const existingOps: EjoosSyncOp[] = [
      {
        id: "x",
        kind: "position_change",
        class: "ready",
        sheet: "",
        personId: "21374",
        fullName: "МАССАЙ",
        positionIndex: "2103350",
        rank: "",
        before: "",
        after: "",
        sourceRef: "",
        why: "",
        confidence: "high",
        payload: {},
        checkedDefault: true,
      },
    ];
    const ops = planTempArrivalCloseOps({
      shPeople: [
        {
          excelRow: 2,
          personId: "21374",
          fullName: "МАССАЙ",
          rank: "солдат",
          positionIndex: "2103350",
          positionTitle: "",
          status: "",
          arrivedFrom: "",
        },
      ],
      movementsAll: [],
      shpoByIndex: new Map([
        [
          "2103350",
          {
            excelRow: 7,
            positionIndex: "2103350",
            personId: "21374",
            fullName: "МАССАЙ",
            rank: "солдат",
          },
        ],
      ]),
      ejoosOos: [],
      dayByIndex: new Map(),
      oosById: new Map(),
      oosByName: new Map(),
      journalMonthStartLabel: "01.09.2026",
      existingOps,
      arrivalOf: () => ({
        excelRow: 6,
        personId: "21374",
        fullName: "МАССАЙ",
        rank: "солдат",
        positionIndex: "",
        arriveDate: "",
        fromUnit: "БРЕЗ",
      }),
      isSamePerson: (left, right) =>
        Boolean(left.personId && left.personId === right.personId),
      eventInLeadWindow: () => true,
      movementEventTime: () => 0,
      isPositionIndex: (value) => /^\d{5,}$/.test(value),
      byPersonName: () => null,
      staffIndexTimesheetForPerson: () => null,
      staffAppointmentDateFor: () => "",
      timesheetEpisodeStartFor: () => "",
      inboundStaffDateFor: () => "",
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
