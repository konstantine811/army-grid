import { describe, expect, it } from "vitest";
import { capitalizeReportPosition } from "./reportPosition";

describe("capitalizeReportPosition", () => {
  it("capitalizes the first Ukrainian letter", () => {
    expect(capitalizeReportPosition("стрілець")).toBe("Стрілець");
    expect(capitalizeReportPosition("водій-електрик")).toBe("Водій-електрик");
  });

  it("keeps an already capitalized position", () => {
    expect(
      capitalizeReportPosition(
        "Оператор відділення радіоелектронної боротьби 1 піхотного батальйону",
      ),
    ).toBe(
      "Оператор відділення радіоелектронної боротьби 1 піхотного батальйону",
    );
  });

  it("trims empty values", () => {
    expect(capitalizeReportPosition("  ")).toBe("");
    expect(capitalizeReportPosition("")).toBe("");
  });
});
