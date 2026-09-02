import { describe, expect, it } from "vitest";
import {
  archivePeriodOverlapsJournalTimesheet,
  archivePeriodTouchesJournalMonth,
  buildTimesheetAbsenceSpans,
  currentStatusConfirmsOpenAbsence,
  historyAbsenceSpansForClosedEpisode,
  timesheetHorizonFillDays,
  timesheetMarkBeforeDeparture,
  timesheetMarkFromArchive,
  isTimesheetAbsenceCode,
  extractTimesheetDestinationFromPosition,
  formatTimesheetDeparture,
  archiveReturnContradictsCurrentSh,
  findTimesheetMonthHeaderCell,
  isInternalStaffTimesheetDeparture,
} from "./ejoosTimesheetText";
import {
  formatExcludedDestination,
  formatTimesheetTransferMark,
} from "./ejoosExcludedColumns";
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

  it("caps an open absence at the 1PB report day, not the calendar month end", () => {
    const spans = buildTimesheetAbsenceSpans(
      [
        {
          departDate: "01.08.2026",
          returnDate: "",
          absenceType: "СЗЧ",
          departMs: Date.UTC(2026, 7, 1),
          returnMs: null,
        },
      ],
      {
        timesheetDay: 25,
        monthStartMs: AUG_2026_START,
        monthEndMs: AUG_2026_END,
        mapCode,
        hasReturn: () => false,
        confirmOpenCarry: () => true,
      },
    );
    expect(spans).toEqual([{ fromDay: 1, toDay: 25, code: "СЗЧ" }]);
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

  const vacationThenMedrota = [
    {
      excelRow: 436,
      departDate: "23.10.2025",
      returnDate: "",
      absenceType: "МЕДРОТА",
      departMs: Date.UTC(2025, 9, 23),
      returnMs: null,
    },
    {
      excelRow: 2650,
      departDate: "30.07.2026",
      returnDate: "09.08.2026",
      absenceType: "ВІДПУСТКА",
      departMs: Date.UTC(2026, 6, 30),
      returnMs: Date.UTC(2026, 7, 9),
    },
  ];
  const mapLeaveOrMed = (absenceType: string) =>
    /медрот|лік/i.test(absenceType)
      ? "лік"
      : /відпуст/i.test(absenceType)
        ? "від"
        : "";

  it("Сіряченко: stale 2025 МЕДРОТА does not paint лік after a later vacation", () => {
    const spans = buildTimesheetAbsenceSpans(vacationThenMedrota, {
      timesheetDay: 25,
      monthStartMs: AUG_2026_START,
      monthEndMs: AUG_2026_END,
      mapCode: mapLeaveOrMed,
      hasReturn: (value) => Boolean(value),
      confirmOpenCarry: () => true,
    });
    expect(spans).toEqual([{ fromDay: 1, toDay: 8, code: "від" }]);
    const days = Array.from({ length: 25 }, (_, index) =>
      timesheetMarkFromArchive(index + 1, {
        activeFromDay: 1,
        lastDay: 25,
        spans,
      }),
    );
    expect(days.slice(0, 8)).toEqual(Array(8).fill("від"));
    expect(days.slice(8)).toEqual(Array(17).fill("+"));
  });

  it("open 2025 МЕДРОТА still paints August when sh confirms and nothing later exists", () => {
    const spans = buildTimesheetAbsenceSpans(
      [vacationThenMedrota[0]],
      {
        timesheetDay: 25,
        monthStartMs: AUG_2026_START,
        monthEndMs: AUG_2026_END,
        mapCode: mapLeaveOrMed,
        hasReturn: () => false,
        confirmOpenCarry: () => true,
      },
    );
    expect(spans).toEqual([{ fromDay: 1, toDay: 25, code: "лік" }]);
  });

  it("vacation then real August treatment keeps both від and лік", () => {
    const spans = buildTimesheetAbsenceSpans(
      [
        {
          departDate: "30.07.2026",
          returnDate: "09.08.2026",
          absenceType: "ВІДПУСТКА",
          departMs: Date.UTC(2026, 6, 30),
          returnMs: Date.UTC(2026, 7, 9),
        },
        {
          departDate: "15.08.2026",
          returnDate: "20.08.2026",
          absenceType: "ЛІКУВАННЯ",
          departMs: Date.UTC(2026, 7, 15),
          returnMs: Date.UTC(2026, 7, 20),
        },
      ],
      {
        timesheetDay: 25,
        monthStartMs: AUG_2026_START,
        monthEndMs: AUG_2026_END,
        mapCode: mapLeaveOrMed,
        hasReturn: (value) => Boolean(value),
        confirmOpenCarry: () => false,
      },
    );
    expect(spans).toEqual([
      { fromDay: 1, toDay: 8, code: "від" },
      { fromDay: 15, toDay: 19, code: "лік" },
    ]);
    expect(
      timesheetMarkFromArchive(14, {
        activeFromDay: 1,
        lastDay: 25,
        spans,
      }),
    ).toBe("+");
    expect(
      timesheetMarkFromArchive(15, {
        activeFromDay: 1,
        lastDay: 25,
        spans,
      }),
    ).toBe("лік");
    expect(
      timesheetMarkFromArchive(20, {
        activeFromDay: 1,
        lastDay: 25,
        spans,
      }),
    ).toBe("+");
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

describe("isInternalStaffTimesheetDeparture", () => {
  it("treats вибув to a 1ПБ platoon/squad as an internal hop", () => {
    expect(
      isInternalStaffTimesheetDeparture(
        "вибув до 3 піхотного відділення 2 піхотного взводу",
      ),
    ).toBe(true);
  });

  it("keeps вибув to another battalion as real history", () => {
    expect(
      isInternalStaffTimesheetDeparture(
        "вибув у 2 піхотного батальйону від 05.08.2026",
      ),
    ).toBe(false);
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

describe("timesheetHorizonFillDays", () => {
  const codes = (pairs: Array<[number, string]>) => {
    const days: string[] = [];
    for (const [day, mark] of pairs) days[day] = mark;
    return days;
  };

  it("fills empty 21–24 with + when 20 and the report day are +", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes([
          [20, "+"],
          [25, "+"],
        ]),
        horizon: 25,
        reportCode: "+",
      }),
    ).toEqual([
      { day: 21, mark: "+" },
      { day: 22, mark: "+" },
      { day: 23, mark: "+" },
      { day: 24, mark: "+" },
    ]);
  });

  it("fills empty 21–25 with + when the last mark is + on 20", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes([[20, "+"]]),
        horizon: 25,
        reportCode: "+",
      }),
    ).toEqual([
      { day: 21, mark: "+" },
      { day: 22, mark: "+" },
      { day: 23, mark: "+" },
      { day: 24, mark: "+" },
      { day: 25, mark: "+" },
    ]);
  });

  it("КАН: 01–20 + and empty 21–25 still fills through 25 even with an old archive return", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes(
          Array.from({ length: 20 }, (_, index) => [index + 1, "+"] as [number, string]),
        ),
        horizon: 25,
        reportCode: "+",
        confirmedReturn: true,
      }),
    ).toEqual([
      { day: 21, mark: "+" },
      { day: 22, mark: "+" },
      { day: 23, mark: "+" },
      { day: 24, mark: "+" },
      { day: 25, mark: "+" },
    ]);
  });

  it("keeps open СЗЧ through 25 instead of a trailing + when there is no return", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes([
          [24, "СЗЧ"],
          [25, "+"],
        ]),
        horizon: 25,
        reportCode: "+",
      }),
    ).toEqual([{ day: 25, mark: "СЗЧ" }]);
  });

  it("keeps open ЛК through empty 21–24 and overwrites + on 25", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes([
          [20, "ЛК"],
          [25, "+"],
        ]),
        horizon: 25,
        reportCode: "+",
      }),
    ).toEqual([
      { day: 21, mark: "ЛК" },
      { day: 22, mark: "ЛК" },
      { day: 23, mark: "ЛК" },
      { day: 24, mark: "ЛК" },
      { day: 25, mark: "ЛК" },
    ]);
  });

  it("does not overwrite + on 25 when archive has a real return", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes([
          [24, "СЗЧ"],
          [25, "+"],
        ]),
        horizon: 25,
        reportCode: "+",
        confirmedReturn: true,
      }),
    ).toEqual([]);
  });

  it("does not invent + on days 1–24 when the row is empty", () => {
    expect(
      timesheetHorizonFillDays({
        dayCodes: [],
        horizon: 25,
        reportCode: "+",
      }),
    ).toEqual([{ day: 25, mark: "+" }]);
  });

  it("overwrites a leaked presentDays count on the horizon day with +", () => {
    expect(isTimesheetAbsenceCode("19")).toBe(false);
    expect(
      timesheetHorizonFillDays({
        dayCodes: codes([
          ...Array.from({ length: 24 }, (_, index) => [index + 1, "+"] as [number, string]),
          [25, "19"],
        ]),
        horizon: 25,
        reportCode: "+",
      }),
    ).toEqual([{ day: 25, mark: "+" }]);
  });
});

describe("currentStatusConfirmsOpenAbsence", () => {
  it("matches СЗЧ/СЗЧ and rejects on-duty leftover", () => {
    expect(currentStatusConfirmsOpenAbsence("СЗЧ", "СЗЧ")).toBe(true);
    expect(currentStatusConfirmsOpenAbsence("+", "СЗЧ")).toBe(false);
    expect(currentStatusConfirmsOpenAbsence("ЗБ", "СЗЧ")).toBe(false);
    expect(currentStatusConfirmsOpenAbsence("+", "лік")).toBe(false);
    expect(currentStatusConfirmsOpenAbsence("лік", "лік")).toBe(true);
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

  it("uses у в/ч for a bare military unit code", () => {
    expect(formatTimesheetDeparture("А7379")).toBe("вибув у в/ч А7379");
    expect(formatTimesheetDeparture("в/ч А7379")).toBe("вибув у в/ч А7379");
    expect(
      formatTimesheetTransferMark({
        timesheetDestination: "А7379",
        destination: "А7379",
      }),
    ).toBe("вибув у в/ч А7379");
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
        documentsDest: BASOVSKYI_POSITION,
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
        documentsDest: BASOVSKYI_POSITION,
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
    expect(excluded?.cells.find((cell) => cell.letter === "AE")?.value).toBe(
      formatExcludedDestination(BASOVSKYI_POSITION),
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

describe("formatExcludedDestination", () => {
  it("writes Виключені AE in lowercase", () => {
    expect(formatExcludedDestination(BASOVSKYI_POSITION)).toBe(
      BASOVSKYI_POSITION.toLocaleLowerCase("uk-UA"),
    );
    expect(
      formatExcludedDestination("  3 МЕХАНІЗОВАНИЙ БАТАЛЬЙОН  "),
    ).toBe("3 механізований батальйон");
  });
});

describe("findTimesheetMonthHeaderCell", () => {
  it("prefers I2 over a leftover Січень in column A", () => {
    const row2 = Array.from({ length: 12 }, () => "");
    row2[0] = "Січень 2026 р.";
    row2[8] = "Серпень 2026 р.";
    const found = findTimesheetMonthHeaderCell([[], row2]);
    expect(found?.matched).toMatch(/серпень/i);
    expect(found?.column).toBe(9);
  });

  it("reads an Excel date in I2 as the timesheet month", () => {
    const row2 = Array.from({ length: 12 }, () => "" as string | Date);
    row2[0] = "Січень 2026 р.";
    row2[8] = new Date(2026, 7, 1);
    const found = findTimesheetMonthHeaderCell([[], row2]);
    expect(found?.month).toBe(8);
    expect(found?.year).toBe(2026);
    expect(found?.column).toBe(9);
  });

  it("ignores leftover Січень in column A when I2 has no month", () => {
    const row2 = Array.from({ length: 12 }, () => "");
    row2[0] = "Січень 2026 р.";
    expect(findTimesheetMonthHeaderCell([[], row2])).toBeNull();
  });
});
