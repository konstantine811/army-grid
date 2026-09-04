import { describe, expect, it } from "vitest";
import { isInternalStaffIndexHop } from "./ejoosMovementRules";
import { resolveTimesheetEpisodeStart } from "./ejoosTimesheetEpisode";

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

  it("does not treat 3ПБ → 1ПБ as an internal 1ПБ index hop", () => {
    expect(
      isInternalStaffIndexHop({
        type: "ПОСАДА",
        destination: "_5 1ПБ",
        arrivedFrom: "_5 3ПБ",
        changeText:
          "гранатометник 2 піхотного відділення 3 піхотного взводу 2 піхотної роти",
        previousIndex: "5007925",
        nextIndex: "2103340",
        note: "наказ №233 від 12.08.2026",
      }),
    ).toBe(false);
  });

  it("does not treat 26 battalion → 1ПБ as an internal staff hop", () => {
    expect(
      isInternalStaffIndexHop({
        type: "ПОСАДА",
        destination: "_5 1ПБ",
        arrivedFrom: "26_БАТ ЗВ'ЯЗ",
        changeText:
          "гранатометник 1 піхотного відділення 3 піхотного взводу 3 піхотної роти",
        previousIndex: "6012407",
        nextIndex: "2103445",
        note: "наказ №236 від 15.08.2026",
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

  it("starts БУРМІСТР's 3ПБ → 1ПБ episode on 12.08", () => {
    const movement = {
      type: "ПОСАДА",
      destination: "_5 1ПБ",
      arrivedFrom: "_5 3ПБ",
      changeText:
        "гранатометник 2 піхотного відділення 3 піхотного взводу 2 піхотної роти",
      previousIndex: "5007925",
      nextIndex: "2103340",
      note: "наказ №233 від 12.08.2026",
    };
    expect(
      resolveTimesheetEpisodeStart({
        monthStartLabel: "01.08.2026",
        appointmentDate: "12.08.2026",
        inboundDate: "",
        hasMonthStartAbsence: false,
        hasDepartureEvidence: false,
        leftUnitThisMonth: false,
        ownUnitMoves: [movement],
      }),
    ).toBe("12.08.2026");
  });

  it("uses a temporary-arrival record as evidence but prefers the 1ПБ appointment date", () => {
    expect(
      resolveTimesheetEpisodeStart({
        monthStartLabel: "01.08.2026",
        appointmentDate: "12.08.2026",
        inboundDate: "10.08.2026",
        hasMonthStartAbsence: false,
        hasDepartureEvidence: false,
        leftUnitThisMonth: false,
        wasTemporaryArrival: true,
        ownUnitMoves: [],
      }),
    ).toBe("12.08.2026");
  });

  it("starts ТРОМБА's 26 battalion → 1ПБ episode on 15.08", () => {
    expect(
      resolveTimesheetEpisodeStart({
        monthStartLabel: "01.08.2026",
        appointmentDate: "15.08.2026",
        inboundDate: "",
        hasMonthStartAbsence: false,
        hasDepartureEvidence: false,
        leftUnitThisMonth: false,
        ownUnitMoves: [
          {
            type: "ПОСАДА",
            destination: "_5 1ПБ",
            arrivedFrom: "26_БАТ ЗВ'ЯЗ",
            changeText:
              "гранатометник 1 піхотного відділення 3 піхотного взводу 3 піхотної роти",
            previousIndex: "6012407",
            nextIndex: "2103445",
            note: "наказ №236 від 15.08.2026",
          },
        ],
      }),
    ).toBe("15.08.2026");
  });
});
