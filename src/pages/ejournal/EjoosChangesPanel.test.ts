import { describe, expect, it } from "vitest";
import { parsePastedPersonList } from "./EjoosChangesPanel";

describe("parsePastedPersonList", () => {
  it("splits a pasted roster into names", () => {
    expect(
      parsePastedPersonList(
        "БАСОВСЬКИЙ Юрій Михайлович\nВІТКОВ Віталій Анатолійович\n\nГУМЕНЮК Вадим Анатолійович",
      ),
    ).toEqual([
      "БАСОВСЬКИЙ Юрій Михайлович",
      "ВІТКОВ Віталій Анатолійович",
      "ГУМЕНЮК Вадим Анатолійович",
    ]);
  });
});
