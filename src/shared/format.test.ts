import { describe, expect, it } from "vitest";
import {
  formatApiDate,
  formatApiDateTime,
  formatUkDateForExcelWrite,
} from "./format";

describe("formatApiDateTime", () => {
  it("shows Kyiv wall time for a UTC instant in summer (UTC+3)", () => {
    expect(formatApiDateTime("2026-09-02T09:18:53.000Z")).toBe(
      "02.09.2026, 12:18:53",
    );
  });

  it("keeps an invalid value as text", () => {
    expect(formatApiDateTime("not-a-date")).toBe("not-a-date");
  });
});

describe("formatApiDate", () => {
  it("shows the Kyiv calendar date, not UTC", () => {
    expect(formatApiDate("2026-09-02T21:10:00.000Z")).toBe("03.09.2026");
  });
});

describe("formatUkDateForExcelWrite", () => {
  it("converts Excel serial numbers and strings to dd.MM.yyyy", () => {
    expect(formatUkDateForExcelWrite(46193)).toBe("20.06.2026");
    expect(formatUkDateForExcelWrite("46193")).toBe("20.06.2026");
    expect(formatUkDateForExcelWrite("13.05.2026")).toBe("13.05.2026");
  });
});
