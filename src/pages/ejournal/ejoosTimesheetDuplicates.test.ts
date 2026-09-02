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
  firstDepartureDay: extra?.firstDepartureDay,
  departureText: extra?.departureText,
});

describe("timesheet duplicates", () => {
  it("treats вибув without + as closed history, not a second active record", () => {
    expect(
      isClosedTimesheetHistoryRow(
        row(10, "ХУБАЄВ", [], { hasDepartureText: true }),
      ),
    ).toBe(true);
  });

  it("keeps Atrakhov 05.08 вибув history plus the 07.08 active 2103764 row", () => {
    const history = row(1031, "АТРАХОВ Олександр Анатолійович", [1, 2, 3, 4], {
      personId: "21155",
      hasDepartureText: true,
      firstDepartureDay: 5,
      departureText: "вибув у 2 піхотного батальйону від 05.08.2026",
    });
    const active = row(764, "АТРАХОВ Олександр Анатолійович", [7, 8, 9, 10, 25], {
      personId: "21155",
    });
    const leftover = row(1033, "АТРАХОВ Олександр Анатолійович", [1, 2, 3, 4, 5], {
      personId: "21155",
    });
    const occupantByIndex = new Map([
      ["2103764", { personId: "21155", fullName: "АТРАХОВ Олександр Анатолійович" }],
    ]);
    expect(findDuplicateTimesheetExtras([history, active], occupantByIndex)).toEqual(
      [],
    );
    const extras = findDuplicateTimesheetExtras(
      [history, active, leftover],
      occupantByIndex,
    );
    expect(extras).toHaveLength(1);
    expect(extras[0].keep.excelRow).toBe(764);
    expect(extras[0].extra.excelRow).toBe(1033);
  });

  it("treats +…+ then вибув as history, not a duplicate of the new episode", () => {
    const history = row(100, "ПЕТРЕНКО", [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20], {
      personId: "111",
      hasDepartureText: true,
      firstDepartureDay: 21,
    });
    const active = row(101, "ПЕТРЕНКО", [21, 22, 23, 24, 25], {
      personId: "111",
    });
    expect(isClosedTimesheetHistoryRow(history)).toBe(true);
    expect(isClosedTimesheetHistoryRow(active)).toBe(false);
    const extras = findDuplicateTimesheetExtras([history, active]);
    expect(extras).toHaveLength(0);
    expect(pickCanonicalTimesheetRow([history, active])?.excelRow).toBe(101);
  });

  it("clears Рибчинський history copy after an internal position hop", () => {
    const active = row(117, "РИБЧИНСЬКИЙ Володимир Володимирович", [4, 5, 25], {
      personId: "16950",
      positionIndex: "2103206",
    });
    const history = row(198, "РИБЧИНСЬКИЙ Володимир Володимирович", [4, 5, 10], {
      personId: "16950",
      positionIndex: "2103195",
      hasDepartureText: true,
      firstDepartureDay: 11,
      departureText:
        "вибув до 3 піхотного відділення 2 піхотного взводу",
    });
    const extras = findDuplicateTimesheetExtras(
      [active, history],
      new Map([
        [
          "2103206",
          {
            personId: "16950",
            fullName: "РИБЧИНСЬКИЙ Володимир Володимирович",
          },
        ],
      ]),
    );
    expect(extras).toHaveLength(1);
    expect(extras[0].keep.excelRow).toBe(117);
    expect(extras[0].extra.excelRow).toBe(198);
    expect(extras[0].reason).toBe("internal_hop");
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

  it("flags leftover Atrakhov after apply: вибув + old «згідно» copy", () => {
    const written = row(1031, "АТРАХОВ", [1, 2, 3, 4, 5], {
      personId: "21155",
      hasDepartureText: true,
      firstDepartureDay: 6,
    });
    const leftover = row(1033, "АТРАХОВ", [1, 2, 3, 4, 5], {
      personId: "21155",
    });
    const extras = findDuplicateTimesheetExtras([written, leftover]);
    expect(extras).toHaveLength(1);
    expect(extras[0].keep.excelRow).toBe(1031);
    expect(extras[0].extra.excelRow).toBe(1033);
    expect(extras[0].reason).toBe("leftover_history");
  });

  it("flags an empty named copy under the filled staff row", () => {
    const staff = row(40, "АРУШАНЯН Норайр Рубенович", [18, 19, 20], {
      personId: "21692",
      positionIndex: "2103232",
    });
    const extra = row(120, "АРУШАНЯН Норайр Рубенович", [], {
      personId: "21692",
      positionIndex: "2103232",
    });
    const extras = findDuplicateTimesheetExtras(
      [staff, extra],
      new Map([
        ["2103232", { personId: "21692", fullName: "АРУШАНЯН Норайр Рубенович" }],
      ]),
    );
    expect(extras).toHaveLength(1);
    expect(extras[0].keep.excelRow).toBe(40);
    expect(extras[0].extra.excelRow).toBe(120);
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

  it("does not treat another person's СЗЧ row as a same_index extra", () => {
    const staff = row(405, "КАПУСТА Сергій Тарасович", [1, 2, 3, 25], {
      personId: "1",
      positionIndex: "2103520",
    });
    const szch = row(544, "ДІДЕНКО Ілля Андрійович", [1, 2, 3], {
      personId: "22898",
      positionIndex: "2103520",
    });
    const extras = findDuplicateTimesheetExtras(
      [staff, szch],
      new Map([
        [
          "2103520",
          { personId: "9905", fullName: "МАКСИМЕНКО Олексій Євгенійович" },
        ],
      ]),
    );
    expect(extras).toEqual([]);
  });

  it("does not glue different people who share a broken Excel ID", () => {
    const extras = findDuplicateTimesheetExtras([
      row(681, "ЛІСОВСЬКИЙ Микола Сергійович", [1], {
        personId: "22895",
        positionIndex: "2103607",
      }),
      row(264, "ЗДЕРОК Дмитро Володимирович", [1, 25], {
        personId: "22895",
        positionIndex: "2103313",
      }),
      row(1172, "ВНУКОВ Максим Анатолійович", [], {
        personId: "22895",
        positionIndex: "2103313",
        hasDepartureText: true,
        firstDepartureDay: 5,
      }),
      row(1173, "ЛЕБЕДЕВ Дмитро Сергійович", [], {
        personId: "22895",
        positionIndex: "379",
        hasDepartureText: true,
      }),
      row(1175, "КРАХМАЛЬОВ Дмитро Олегович", [], {
        personId: "22895",
        positionIndex: "30.07.1960",
      }),
    ]);
    expect(extras).toEqual([]);
  });

  it("still flags a same-name leftover copy when IDs collide with other people", () => {
    const history = row(1172, "ВНУКОВ Максим Анатолійович", [1, 2], {
      personId: "22895",
      hasDepartureText: true,
      firstDepartureDay: 5,
    });
    const copy = row(1300, "ВНУКОВ Максим Анатолійович", [1, 2], {
      personId: "22895",
    });
    const extras = findDuplicateTimesheetExtras([history, copy]);
    expect(extras).toHaveLength(1);
    expect(extras[0].keep.excelRow).toBe(1172);
    expect(extras[0].extra.excelRow).toBe(1300);
    expect(extras[0].reason).toBe("leftover_history");
  });
});
