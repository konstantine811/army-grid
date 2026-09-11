import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  formatUaDateTyping,
  isAsOfPickerDayCommit,
  isoDateToUaLabel,
  isCompleteUaDate,
  normalizeAsOfIso,
  readRememberedAsOfIso,
  uaDateToIso,
  uaLabelToIsoDate,
  writeRememberedAsOfIso,
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

  it("does not commit a native picker month/year flip until a day is chosen", () => {
    expect(isAsOfPickerDayCommit("2026-08-25", "2026-09-25")).toBe(false);
    expect(isAsOfPickerDayCommit("2026-08-31", "2026-09-30")).toBe(false);
    expect(isAsOfPickerDayCommit("2026-08-25", "2025-08-25")).toBe(false);
    expect(isAsOfPickerDayCommit("2026-09-25", "2026-09-10")).toBe(true);
    expect(isAsOfPickerDayCommit("", "2026-09-10")).toBe(true);
    expect(isAsOfPickerDayCommit("2026-09-10", "bad")).toBe(false);
  });

  describe("remembered as-of date", () => {
    beforeEach(() => {
      const localStorageMock = {
        store: {} as Record<string, string>,
        getItem(key: string) {
          return localStorageMock.store[key] ?? null;
        },
        setItem(key: string, value: string) {
          localStorageMock.store[key] = value;
        },
        removeItem(key: string) {
          delete localStorageMock.store[key];
        },
        clear() {
          localStorageMock.store = {};
        },
      };
      vi.stubGlobal("localStorage", localStorageMock);
    });

    afterEach(() => {
      vi.unstubAllGlobals();
    });

    it("normalizes UA and ISO labels", () => {
      expect(normalizeAsOfIso("10.09.2026")).toBe("2026-09-10");
      expect(normalizeAsOfIso("2026-09-10")).toBe("2026-09-10");
      expect(normalizeAsOfIso("10.09")).toBe("");
    });

    it("persists the last applied date for the next visit", () => {
      expect(readRememberedAsOfIso()).toBe("");
      expect(writeRememberedAsOfIso("10.09.2026")).toBe("2026-09-10");
      expect(readRememberedAsOfIso()).toBe("2026-09-10");
    });
  });
});
