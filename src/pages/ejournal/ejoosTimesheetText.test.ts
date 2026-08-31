import { describe, expect, it } from "vitest";
import {
  archivePeriodOverlapsJournalTimesheet,
  archivePeriodTouchesJournalMonth,
  buildTimesheetAbsenceSpans,
  currentStatusConfirmsOpenAbsence,
  historyAbsenceSpansForClosedEpisode,
  timesheetMarkBeforeDeparture,
  timesheetMarkFromArchive,
  extractTimesheetDestinationFromPosition,
  formatTimesheetDeparture,
  archiveReturnContradictsCurrentSh,
} from "./ejoosTimesheetText";
import { formatTimesheetTransferMark } from "./ejoosExcludedColumns";
import { buildSheetRowPreviews } from "./ejoosSheetRowPreview";
import type { EjoosSyncOp } from "./ejoosSyncPlan";

const AUG_2026_START = Date.UTC(2026, 7, 1);
const AUG_2026_END = Date.UTC(2026, 7, 31);

describe("archivePeriodOverlapsJournalTimesheet", () => {
  it("NAVROTSKY: closed 2025 leave does not paint August 2026", () => {
    expect(
      archivePeriodOverlapsJournalTimesheet(
        Date.UTC(2025, 9, 9),
        Date.UTC(2025, 10, 7),
        AUG_2026_START,
        AUG_2026_END,
      ),
    ).toBe(false);
  });

  it("open 2025 archive without sh confirmation does not paint August", () => {
    expect(
      archivePeriodOverlapsJournalTimesheet(
        Date.UTC(2025, 9, 9),
        null,
        AUG_2026_START,
        AUG_2026_END,
      ),
    ).toBe(false);
  });

  it("open СЗЧ from 25.07 continues through August when sh still confirms", () => {
    expect(
      archivePeriodOverlapsJournalTimesheet(
        Date.UTC(2026, 6, 25),
        null,
        AUG_2026_START,
        AUG_2026_END,
        { carryOpen: true },
      ),
    ).toBe(true);
  });

  it("МЕДРОТА 31.07 → return 03.08 overlaps August days 1–2", () => {
    expect(
      archivePeriodOverlapsJournalTimesheet(
        Date.UTC(2026, 6, 31),
        Date.UTC(2026, 7, 3),
        AUG_2026_START,
        AUG_2026_END,
      ),
    ).toBe(true);
  });
});

describe("buildTimesheetAbsenceSpans / carry-in", () => {
  const mapCode = (absenceType: string) =>
    /сзч/i.test(absenceType) ? "СЗЧ" : /лік/i.test(absenceType) ? "лік" : "";

  it("paints 01–31 СЗЧ for open July absence when sh confirms", () => {
    const spans = buildTimesheetAbsenceSpans(
      [
        {
          departDate: "25.07.2026",
          returnDate: "",
          absenceType: "СЗЧ",
          departMs: Date.UTC(2026, 6, 25),
          returnMs: null,
        },
      ],
      {
        timesheetDay: 31,
        monthStartMs: AUG_2026_START,
        monthEndMs: AUG_2026_END,
        mapCode,
        hasReturn: () => false,
        confirmOpenCarry: () => true,
      },
    );
    expect(spans).toEqual([{ fromDay: 1, toDay: 31, code: "СЗЧ" }]);
  });

  it("does not paint August from stale 2025 open archive when sh is on duty", () => {
    const spans = buildTimesheetAbsenceSpans(
      [
        {
          departDate: "09.10.2025",
          returnDate: "",
          absenceType: "відпустка",
          departMs: Date.UTC(2025, 9, 9),
          returnMs: null,
        },
      ],
      {
        timesheetDay: 31,
        monthStartMs: AUG_2026_START,
        monthEndMs: AUG_2026_END,
        mapCode: () => "від",
        hasReturn: () => false,
        confirmOpenCarry: () => false,
      },
    );
    expect(spans).toEqual([]);
  });
});

describe("archivePeriodTouchesJournalMonth", () => {
  it("open July absence touches August only when carryOpen", () => {
    expect(
      archivePeriodTouchesJournalMonth(
        Date.UTC(2026, 6, 25),
        null,
        AUG_2026_START,
        AUG_2026_END,
      ),
    ).toBe(false);
    expect(
      archivePeriodTouchesJournalMonth(
        Date.UTC(2026, 6, 25),
        null,
        AUG_2026_START,
        AUG_2026_END,
        { carryOpen: true },
      ),
    ).toBe(true);
  });
});

describe("historical position row keeps prior absences", () => {
  it("01–04 лік, 05–08 +, 09 вибув", () => {
    const spans = historyAbsenceSpansForClosedEpisode(
      { timesheetAbsenceSpans: "1-4:лік" },
      9,
    );
    expect(timesheetMarkBeforeDeparture(1, 9, spans)).toBe("лік");
    expect(timesheetMarkBeforeDeparture(4, 9, spans)).toBe("лік");
    expect(timesheetMarkBeforeDeparture(5, 9, spans)).toBe("+");
    expect(timesheetMarkBeforeDeparture(8, 9, spans)).toBe("+");
  });
});

describe("disposition archive before order day", () => {
  it("10–11 ЗБ, 12 overlay, 13+ ЗБ, 01–09 +", () => {
    const spans = [{ fromDay: 10, toDay: 31, code: "ЗБ" }];
    const mark = (day: number) =>
      timesheetMarkFromArchive(day, {
        activeFromDay: 1,
        lastDay: 31,
        spans,
        fillBeforeActive: true,
      });
    expect(mark(1)).toBe("+");
    expect(mark(9)).toBe("+");
    expect(mark(10)).toBe("ЗБ");
    expect(mark(11)).toBe("ЗБ");
    expect(mark(13)).toBe("ЗБ");
  });
});

describe("currentStatusConfirmsOpenAbsence", () => {
  it("matches СЗЧ/СЗЧ and rejects on-duty leftover", () => {
    expect(currentStatusConfirmsOpenAbsence("СЗЧ", "СЗЧ")).toBe(true);
    expect(currentStatusConfirmsOpenAbsence("+", "СЗЧ")).toBe(false);
    expect(currentStatusConfirmsOpenAbsence("ЗБ", "СЗЧ")).toBe(false);
  });
});

describe("archiveReturnContradictsCurrentSh", () => {
  it("flags archive return while sh still СЗЧ", () => {
    expect(archiveReturnContradictsCurrentSh("СЗЧ", "СЗЧ", true)).toBe(true);
    expect(archiveReturnContradictsCurrentSh("СЗЧ", "СЗЧ", false)).toBe(false);
    expect(archiveReturnContradictsCurrentSh("+", "СЗЧ", true)).toBe(false);
  });
});

const BASOVSKYI_UNIT =
  "МЕХАНІЗОВАНОГО ВІДДІЛЕННЯ МЕХАНІЗОВАНОГО ВЗВОДУ МЕХАНІЗОВАНОЇ РОТИ МЕХАНІЗОВАНОГО БАТАЛЬЙОНУ";
const BASOVSKYI_POSITION = `Стрілець ${BASOVSKYI_UNIT}`;
const BASOVSKYI_UNIT_NOTE =
  "А4784 ВІЙСЬКОВОЇ ЧАСТИНИ А4655 ВІЙСЬКОВОЇ ЧАСТИНИ А1314";

describe("timesheet departure phrase strips position and picks до/у", () => {
  it("Basovskyi: drop job title, use до + subunit chain", () => {
    expect(extractTimesheetDestinationFromPosition(BASOVSKYI_POSITION)).toBe(
      BASOVSKYI_UNIT,
    );
    expect(formatTimesheetDeparture(BASOVSKYI_UNIT)).toBe(
      `вибув до ${BASOVSKYI_UNIT}`,
    );
  });

  it("uses у for розпорядження", () => {
    expect(formatTimesheetDeparture("розпорядження командира")).toBe(
      "вибув у розпорядження командира",
    );
  });

  it("does not treat a bare job title as a destination", () => {
    expect(extractTimesheetDestinationFromPosition("Стрілець")).toBe("");
  });

  it("prefers subunit chain over A#### even if timesheetDestination is the unit note", () => {
    expect(
      formatTimesheetTransferMark({
        timesheetDestination: BASOVSKYI_UNIT_NOTE,
        documentsDest: BASOVSKYI_UNIT_NOTE,
        destination: BASOVSKYI_UNIT_NOTE,
        changeText: BASOVSKYI_POSITION,
      }),
    ).toBe(`вибув до ${BASOVSKYI_UNIT}`);
  });

  it("keeps a manual timesheet destination that is not a military-unit code", () => {
    expect(
      formatTimesheetTransferMark({
        timesheetDestination: "розпорядження командира",
        changeText: BASOVSKYI_POSITION,
      }),
    ).toBe("вибув у розпорядження командира");
  });
});

describe("buildSheetRowPreviews", () => {
  it("shows Excluded / SHPO / timesheet cells for outbound ПЕРЕВ", () => {
    const op: EjoosSyncOp = {
      id: "excl-1",
      kind: "exclude_transfer",
      class: "ready",
      sheet: "Виключені → Табель → ШПО/ООС",
      personId: "12521",
      fullName: "БАСОВСЬКИЙ Юрій Михайлович",
      positionIndex: "2103200",
      rank: "солдат",
      before: "",
      after: "",
      sourceRef: "",
      why: "",
      confidence: "high",
      checkedDefault: true,
      payload: {
        fromRank: "солдат",
        fromName: "БАСОВСЬКИЙ Юрій Михайлович",
        fromPersonId: "12521",
        fromPositionIndex: "2103200",
        documentsDest: BASOVSKYI_UNIT_NOTE,
        destination: BASOVSKYI_UNIT_NOTE,
        timesheetDestination: BASOVSKYI_UNIT_NOTE,
        changeText: BASOVSKYI_POSITION,
        excludeDate: "05.08.2026",
        orderDate: "05.08.2026",
        orderNumber: "123",
        exclusionReason: "ПЕРЕВЕДЕННЯ",
        type: "ПЕРЕВ",
        shpoExcelRow: "104",
        oosExcelRow: "80",
        timesheetExcelRow: "110",
      },
    };
    const previews = buildSheetRowPreviews([op], 31);
    const excluded = previews.find((row) => row.sheetKey === "excluded");
    const timesheet = previews.find((row) => row.sheetKey === "timesheet-history");
    const shpo = previews.find((row) => row.sheetKey === "shpo");
    expect(excluded?.cells.find((cell) => cell.letter === "AE")?.value).toContain(
      "а4784",
    );
    expect(
      timesheet?.cells.some((cell) =>
        cell.value.includes(`вибув до ${BASOVSKYI_UNIT}`),
      ),
    ).toBe(true);
    expect(shpo?.cells.find((cell) => cell.letter === "A")?.kind).toBe("keep");
    expect(shpo?.cells.find((cell) => cell.letter === "G")?.kind).toBe("empty");
  });
});

