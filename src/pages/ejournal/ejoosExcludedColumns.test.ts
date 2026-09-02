import { describe, expect, it } from "vitest";
import {
  collectExcludedPositionDateWrites,
  formatExcludedPositionDates,
  isExcludedWrapColumn,
} from "./ejoosExcludedColumns";

describe("formatExcludedPositionDates", () => {
  it("puts space-separated dates on their own lines, earliest first", () => {
    expect(formatExcludedPositionDates("28.06.2026 16.05.2026")).toBe(
      "16.05.2026\n28.06.2026",
    );
    expect(formatExcludedPositionDates("28.06.2026 28.05.2026")).toBe(
      "28.05.2026\n28.06.2026",
    );
    expect(formatExcludedPositionDates("16.05.2026 09.02.2026")).toBe(
      "09.02.2026\n16.05.2026",
    );
  });

  it("re-sorts multiline dates from newest-first to earliest-first", () => {
    expect(
      formatExcludedPositionDates("10.03.2026\n17.09.2025\n05.08.2025"),
    ).toBe("05.08.2025\n17.09.2025\n10.03.2026");
  });

  it("leaves a single date on one line", () => {
    expect(formatExcludedPositionDates("16.05.2026")).toBe("16.05.2026");
  });

  it("dedupes and pads unpadded dates", () => {
    expect(formatExcludedPositionDates("5.8.2025 05.08.2025 17.09.2025")).toBe(
      "05.08.2025\n17.09.2025",
    );
  });

  it("returns empty for blank values", () => {
    expect(formatExcludedPositionDates("")).toBe("");
    expect(formatExcludedPositionDates(null)).toBe("");
    expect(formatExcludedPositionDates(undefined)).toBe("");
  });
});

describe("collectExcludedPositionDateWrites", () => {
  it("rewrites only cells that are packed or out of chronological order", () => {
    const rows: Array<Array<unknown> | undefined> = Array.from(
      { length: 8 },
      () => [],
    );
    rows[5] = [, , , , "28.06.2026 16.05.2026"];
    rows[6] = [, , , , "16.05.2026\n28.06.2026"];
    rows[7] = [, , , , "10.03.2026\n17.09.2025\n05.08.2025"];

    expect(collectExcludedPositionDateWrites(rows)).toEqual([
      {
        row: 6,
        column: 5,
        value: "16.05.2026\n28.06.2026",
        wrapText: true,
      },
      {
        row: 8,
        column: 5,
        value: "05.08.2025\n17.09.2025\n10.03.2026",
        wrapText: true,
      },
    ]);
  });
});

describe("isExcludedWrapColumn", () => {
  it("wraps appointment dates so multiline cells show every line", () => {
    expect(isExcludedWrapColumn(5)).toBe(true);
    expect(isExcludedWrapColumn(26)).toBe(true);
    expect(isExcludedWrapColumn(32)).toBe(true);
    expect(isExcludedWrapColumn(4)).toBe(false);
  });
});
