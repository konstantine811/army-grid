import { describe, expect, it } from "vitest";
import {
  formatUaDateTyping,
  isoDateToUaLabel,
  isCompleteUaDate,
  uaDateToIso,
  uaLabelToIsoDate,
} from "./ejoosAsOfDate";

describe("ejoosAsOfDate", () => {
  it("formats typing mask with dots", () => {
    expect(formatUaDateTyping("25082026")).toBe("25.08.2026");
    expect(formatUaDateTyping("25.08.2026")).toBe("25.08.2026");
    expect(formatUaDateTyping("25abc08")).toBe("25.08");
  });

  it("converts UA label and ISO both ways", () => {
    expect(uaLabelToIsoDate("25.08.2026")).toBe("2026-08-25");
    expect(isoDateToUaLabel("2026-08-25")).toBe("25.08.2026");
  });

  it("validates complete dates", () => {
    expect(isCompleteUaDate("25.08.2026")).toBe(true);
    expect(isCompleteUaDate("25.08")).toBe(false);
    expect(uaDateToIso("07.09.2026")).toBe("2026-09-07");
  });
});
