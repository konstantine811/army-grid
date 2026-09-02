import { describe, expect, it } from "vitest";
import type { UbdBasisOrderOption } from "./ubdBasisOrdersData";
import {
  findUbdBasisOrdersForLocationPeriod,
  pickUbdBasisOrderForTaskPeriod,
  resolveUbdBasisForTask,
} from "./ubdBasisOrders";

const directory: UbdBasisOrderOption[] = [
  {
    number: "№123/дск",
    date: "01.08.2026",
    location: "Петропавлівка",
    validFrom: "01.08.2026",
    validTo: "10.08.2026",
  },
  {
    number: "№145/дск",
    date: "11.08.2026",
    location: "Петропавлівка",
    validFrom: "11.08.2026",
    validTo: "25.08.2026",
  },
  {
    number: "№151/дск",
    date: "15.08.2026",
    location: "Новоселівка",
    validFrom: "15.08.2026",
    validTo: "30.08.2026",
  },
];

describe("БР за локацією + датою", () => {
  it("підставляє БР за датою виходу і локацією, не лише за місцем", () => {
    expect(
      pickUbdBasisOrderForTaskPeriod(
        "18.08.2026",
        directory,
        "н.п. Петропавлівка",
      )?.number,
    ).toBe("№145/дск");
    expect(
      pickUbdBasisOrderForTaskPeriod(
        "05.08.2026",
        directory,
        "Петропавлівка",
      )?.number,
    ).toBe("№123/дск");
  });

  it("знаходить усі БР, якщо період перекриває зміну розпорядження", () => {
    const matches = findUbdBasisOrdersForLocationPeriod(
      "Петропавлівка",
      "з 05.08.2026-20.08.2026",
      directory,
    );
    expect(matches.map((item) => item.number)).toEqual(["№123/дск", "№145/дск"]);

    const resolved = resolveUbdBasisForTask(
      "з 05.08.2026-20.08.2026",
      "Петропавлівка",
      directory,
    );
    expect(resolved?.number).toBe("№123/дск, №145/дск");
  });

  it("не бере БР іншої локації", () => {
    expect(
      pickUbdBasisOrderForTaskPeriod(
        "18.08.2026",
        directory,
        "Новоселівка",
      )?.number,
    ).toBe("№151/дск");
  });
});
