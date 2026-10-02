import { describe, expect, it } from "vitest";
import { compareSortValues } from "./SciDataTable";

describe("compareSortValues", () => {
  it("sorts exit day counts as numbers, not as years", () => {
    const values = ["39", "33", "9", "5", "64", "27", "19", "15", "8"];
    const sorted = [...values].sort((left, right) =>
      compareSortValues(right, left),
    );
    expect(sorted).toEqual(["64", "39", "33", "27", "19", "15", "9", "8", "5"]);
  });

  it("still sorts dotted dates chronologically", () => {
    expect(compareSortValues("01.09.2026", "15.09.2026")).toBeLessThan(0);
    expect(compareSortValues("15.09.2026", "01.09.2026")).toBeGreaterThan(0);
  });
});
