import { describe, expect, it } from "vitest";
import {
  buildBchsStayPlaceStats,
  extractBchsAwayPeopleFromDbRows,
  filterBchsNovaPeople,
  isBchsBattleReadyPerson,
  isBchsOnCombatExit,
  summarizeBchsPersonnelStay,
} from "./bchsCalc";
import type { BchsPersonnelAwayPerson } from "./bchsTypes";

const person = (
  overrides: Partial<BchsPersonnelAwayPerson> = {},
): BchsPersonnelAwayPerson => ({
  battalion: "нова",
  rosterUnit: "1 рота",
  position: "стрілець",
  shpkFact: "",
  rankCategory: "солдат",
  rankTitle: "солдат",
  fullName: "Іваненко Іван",
  callsign: "",
  status: "в строю",
  roleType: "піхота",
  combatReadiness: "",
  bzvpStatus: "",
  brezAssignment: "",
  destination: "",
  treatmentNote: "",
  mobilizationContract: "",
  medicalPlace: "",
  medicalNote: "",
  ...overrides,
});

describe("BCHS stay / БГ / exit stats", () => {
  it("groups місце перебування and puts empty last", () => {
    const stats = buildBchsStayPlaceStats([
      person({ medicalPlace: "ППД Вишневе" }),
      person({ medicalPlace: "ппд  вишневе" }),
      person({ medicalPlace: "КСП Павлоград" }),
      person({ medicalPlace: "" }),
    ]);

    expect(stats.map((item) => [item.label, item.value])).toEqual([
      ["ППД Вишневе", 2],
      ["КСП Павлоград", 1],
      ["Не вказано", 1],
    ]);
  });

  it("drops nova rows without ПІБ", () => {
    expect(
      filterBchsNovaPeople([
        person({ fullName: "Іваненко Іван" }),
        person({ fullName: "" }),
        person({ fullName: "   " }),
        person({ fullName: "-", battalion: "нова" }),
        person({ fullName: "Петренко Петро", battalion: "стара" }),
      ]),
    ).toEqual([person({ fullName: "Іваненко Іван" })]);
  });

  it("counts БГ vs не БГ only among нова", () => {
    const summary = summarizeBchsPersonnelStay([
      person({ fullName: "А", combatReadiness: "БГ" }),
      person({ fullName: "Б", combatReadiness: "не бг" }),
      person({ fullName: "В", combatReadiness: "", battalion: "стара" }),
    ]);

    expect(summary.total).toBe(2);
    expect(summary.battleReady).toBe(1);
    expect(summary.notBattleReady).toBe(1);
  });

  it("detects who is on exit from place text and fighter dates", () => {
    expect(
      isBchsOnCombatExit(person({ medicalPlace: "на виході" })),
    ).toBe(true);
    expect(
      isBchsOnCombatExit(person({ medicalPlace: "на виконанні БЗ" })),
    ).toBe(true);
    expect(
      isBchsOnCombatExit(
        person({
          fighterExitDate: "14.06.2026",
          fighterReturnDate: "",
        }),
      ),
    ).toBe(true);
    expect(
      isBchsOnCombatExit(
        person({
          fighterExitDate: "14.06.2026",
          fighterReturnDate: "11.08.2026",
        }),
      ),
    ).toBe(false);
    expect(
      isBchsOnCombatExit(person({ status: "вихідний" })),
    ).toBe(false);
    expect(isBchsBattleReadyPerson(person({ combatReadiness: "БГ" }))).toBe(
      true,
    );
  });

  it("reads stay place, БГ and fighter exit from штатка row keys", () => {
    const [row] = extractBchsAwayPeopleFromDbRows([
      {
        battalion: "нова",
        column_31: "",
        column_40: "ППД Вишневе",
        column_23: "",
        column_42: "БГ",
        fighter_status_exit_date: "02.09.2026",
        fighter_status_return_date: "",
        fullName: "Петренко Петро",
      },
    ]);

    expect(row.medicalPlace).toBe("ППД Вишневе");
    expect(row.combatReadiness).toBe("БГ");
    expect(row.fighterExitDate).toBe("02.09.2026");
    expect(isBchsOnCombatExit(row)).toBe(true);
  });
});
