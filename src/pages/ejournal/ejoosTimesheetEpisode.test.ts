import { describe, expect, it } from "vitest";
import { isInternalStaffIndexHop } from "./ejoosMovementRules";
import { resolveTimesheetEpisodeStart } from "./ejoosTimesheetEpisode";
import {
  dayFromOrderLabel,
  timesheetMarkFromArchive,
} from "./ejoosTimesheetText";

const hop = (from: string, to: string) => ({
  type: "ПОСАДА",
  destination: "1ПБ",
  changeText: `${from} → ${to}`,
  previousIndex: from,
  nextIndex: to,
  note: "",
});

describe("isInternalStaffIndexHop", () => {
  it("treats 2103535 → 2103207 inside 1ПБ as a hop", () => {
    expect(isInternalStaffIndexHop(hop("2103535", "2103207"))).toBe(true);
  });

  it("treats a 1ПБ hop when Куди is the new position title", () => {
    expect(
      isInternalStaffIndexHop({
        type: "ПОСАДА",
        destination:
          "номер обслуги кулеметного взводу 1 піхотної роти 1 піхотного батальйону",
        changeText:
          "номер обслуги кулеметного взводу 1 піхотної роти 1 піхотного батальйону",
        previousIndex: "2103725",
        nextIndex: "2103255",
        note: "х",
      }),
    ).toBe(true);
  });

  it("does not treat disposition → staff as a hop", () => {
    expect(
      isInternalStaffIndexHop({
        type: "ПОСАДА",
        destination: "1ПБ",
        changeText: "розпорядження → 2103764",
        previousIndex: "розпорядження",
        nextIndex: "2103764",
        note: "",
      }),
    ).toBe(false);
  });
});

describe("resolveTimesheetEpisodeStart", () => {
  it("keeps 01.08 for a two-day position hop back to the same index", () => {
    expect(
      resolveTimesheetEpisodeStart({
        monthStartLabel: "01.08.2026",
        appointmentDate: "13.08.2026",
        inboundDate: "",
        hasMonthStartAbsence: false,
        hasDepartureEvidence: false,
        leftUnitThisMonth: false,
        ownUnitMoves: [hop("2103225", "2103157"), hop("2103157", "2103225")],
      }),
    ).toBe("01.08.2026");
  });

  it("keeps the appointment date after an actual departure", () => {
    expect(
      resolveTimesheetEpisodeStart({
        monthStartLabel: "01.08.2026",
        appointmentDate: "07.08.2026",
        inboundDate: "",
        hasMonthStartAbsence: false,
        hasDepartureEvidence: true,
        leftUnitThisMonth: false,
        ownUnitMoves: [hop("4907559", "2103764")],
      }),
    ).toBe("07.08.2026");
  });

  it("starts a temporary arrival's staff episode on the appointment date", () => {
    const activeFrom = resolveTimesheetEpisodeStart({
      monthStartLabel: "01.08.2026",
      appointmentDate: "12.08.2026",
      inboundDate: "10.08.2026",
      hasMonthStartAbsence: false,
      hasDepartureEvidence: false,
      leftUnitThisMonth: false,
      wasTemporaryArrival: true,
      ownUnitMoves: [hop("тимчасово прибулий", "2103340")],
    });
    expect(activeFrom).toBe("12.08.2026");
    const activeFromDay = dayFromOrderLabel(activeFrom);
    const mark = (day: number) =>
      timesheetMarkFromArchive(day, {
        activeFromDay,
        lastDay: 25,
        spans: [],
        fillBeforeActive: true,
      });
    expect(mark(11)).toBe("-");
    expect(mark(12)).toBe("+");
  });
});
