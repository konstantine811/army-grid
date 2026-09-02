import { describe, expect, it } from "vitest";
import {
  formatBirthDateDisplay,
  isBirthdayInMonth,
  listCurrentMonthBirthdays,
  listOnDutyMissingBirthDates,
  parseBirthDateParts,
  pickUsableBirthDate,
} from "./socPassportBirthdays";
import type { SocPerson } from "./socPassportTypes";

const person = (overrides: Partial<SocPerson>): SocPerson =>
  ({
    id: overrides.name ?? "id",
    name: "ТЕСТ Іван Іванович",
    normalizedName: "тест іван іванович",
    shortName: "тест іван",
    callsign: "СОКІЛ",
    position: "",
    positionIndex: "",
    rank: "солдат",
    staffRank: "",
    rankGroup: "soldier",
    serviceType: "mobilized",
    sex: "male",
    birthDate: "12.09.1990",
    age: 35,
    ageBand: "31-40",
    birthPlace: "",
    region: "unknown",
    regionLabel: "",
    regionOccupied: false,
    nationality: "ukraine",
    marital: "unknown",
    childrenUnder18: 0,
    children3plus: false,
    relativesServing: false,
    relativesAbroad: false,
    relativesHostile: false,
    relativesRaw: "",
    extraRaw: "",
    arrivedFrom: "",
    calledBy: "",
    arrivalSource: "unknown",
    hasUbd: false,
    ubdNumber: "",
    ubdRosterStatus: null,
    staticCombatExitOverride: false,
    oosDislocation: "КСП Павлоград",
    combatDutyEvidence: [],
    isIdp: false,
    morningStatus: "В СТРОЮ",
    morningAbsenceNotes: "",
    morningDestination: "1 рота",
    morningLocation: "ППД Вишневе",
    isTransiter: false,
    bzvpStatus: "",
    brezAssignment: "",
    onStaff: true,
    onList: true,
    present: true,
    inDisposition: false,
    morningExitCount: 0,
    jbdExitCount: 0,
    exitCount: 0,
    exitBand: "none",
    match: {
      oos: true,
      morning: true,
      exits: false,
      jbdExits: false,
      bplaExits: false,
      ubdRoster: false,
      tempArrival: false,
    },
    parseNotes: [],
    ...overrides,
  }) as SocPerson;

const september = new Date(Date.UTC(2026, 8, 2));

describe("parseBirthDateParts", () => {
  it("reads dotted and ISO dates", () => {
    expect(parseBirthDateParts("12.09.1990")).toEqual({
      day: 12,
      month: 9,
      year: 1990,
    });
    expect(parseBirthDateParts("1990-09-12")).toEqual({
      day: 12,
      month: 9,
      year: 1990,
    });
    expect(formatBirthDateDisplay("1990-09-12")).toBe("12.09.1990");
  });
});

describe("listCurrentMonthBirthdays", () => {
  it("keeps people in formation with a birthday this month", () => {
    const rows = listCurrentMonthBirthdays(
      [
        person({ name: "ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", birthDate: "27.09.1986" }),
        person({ name: "ЛІТНІЙ Петро", birthDate: "03.07.1988" }),
      ],
      september,
    );
    expect(rows).toEqual([
      expect.objectContaining({
        name: "ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ",
        callsign: "СОКІЛ",
        birthDate: "27.09.1986",
        location: "ППД Вишневе",
        day: 27,
      }),
    ]);
    expect(isBirthdayInMonth("27.09.1986", september)).toBe(true);
  });

  it("drops transients from destination or notes", () => {
    const rows = listCurrentMonthBirthdays(
      [
        person({
          name: "ТРАНЗИТ ПІДРОЗДІЛ",
          isTransiter: true,
          morningDestination: "ТРАНЗИТЕР",
        }),
        person({
          name: "ТРАНЗИТ ПРИМІТКА",
          morningAbsenceNotes: "ТРАНЗИТЕР",
        }),
        person({
          name: "НЕ В СТРОЮ",
          morningStatus: "ВІДПУСТКА",
        }),
      ],
      september,
    );
    expect(rows).toEqual([]);
  });

  it("falls back to OOS location when morning location is empty", () => {
    const rows = listCurrentMonthBirthdays(
      [
        person({
          morningLocation: "",
          oosDislocation: "Полігон ЧЛ",
        }),
      ],
      september,
    );
    expect(rows[0]?.location).toBe("Полігон ЧЛ");
  });
});

describe("listOnDutyMissingBirthDates", () => {
  it("lists people in formation whose EJOOS birth date is missing", () => {
    const rows = listOnDutyMissingBirthDates([
      person({ name: "Є ДАТА", birthDate: "12.09.1990" }),
      person({
        name: "НЕМАЄ В ООС",
        birthDate: "",
        match: {
          oos: false,
          morning: true,
          exits: false,
          jbdExits: false,
          bplaExits: false,
          ubdRoster: false,
          tempArrival: false,
        },
      }),
      person({ name: "ПОРОЖНЯ ДАТА", birthDate: "" }),
      person({
        name: "ТРАНЗИТЕР",
        birthDate: "",
        isTransiter: true,
        morningDestination: "ТРАНЗИТЕР",
      }),
    ]);

    expect(rows.map((row) => row.name)).toEqual([
      "НЕМАЄ В ООС",
      "ПОРОЖНЯ ДАТА",
    ]);
    expect(rows[0]?.note).toBe("немає в ООС ЕЖООС");
    expect(rows[1]?.note).toBe("немає дати в ООС і в штатці");
  });
});

describe("pickUsableBirthDate", () => {
  it("prefers a parseable OOS date over the morning-report date", () => {
    expect(pickUsableBirthDate("27.09.1986", "22.07.1992")).toBe("27.09.1986");
  });

  it("falls back to the staff-sheet date when OOS is empty or not a birth date", () => {
    expect(pickUsableBirthDate("", "22.07.1992")).toBe("22.07.1992");
    expect(pickUsableBirthDate("46233", "22.07.1992")).toBe("22.07.1992");
  });
});

describe("ЦАПЕНКО from morning report", () => {
  it("does not list a July birthday in September even with a staff-sheet date", () => {
    const rows = listCurrentMonthBirthdays(
      [
        person({
          name: "ЦАПЕНКО Микола Володимирович",
          birthDate: "22.07.1992",
          morningStatus: "В строю",
        }),
      ],
      september,
    );
    expect(rows).toEqual([]);
  });

  it("adds a September birthday taken from the morning report", () => {
    const rows = listCurrentMonthBirthdays(
      [
        person({
          name: "ЦАПЕНКО Микола Володимирович",
          birthDate: "12.09.1992",
          morningStatus: "В строю",
          match: {
            oos: false,
            morning: true,
            exits: false,
            jbdExits: false,
            bplaExits: false,
            ubdRoster: false,
            tempArrival: false,
          },
        }),
      ],
      september,
    );
    expect(rows).toEqual([
      expect.objectContaining({
        name: "ЦАПЕНКО Микола Володимирович",
        birthDate: "12.09.1992",
        day: 12,
      }),
    ]);
  });
});
