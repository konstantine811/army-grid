import { describe, expect, it, vi, beforeEach, afterEach } from "vitest";
import {
  getLatestStaffSheetSyncSlot,
  shouldAutoSyncStaffSheetRoster,
  STAFF_SHEET_SYNC_EVENING_HOUR,
  STAFF_SHEET_SYNC_LUNCH_HOUR,
  STAFF_SHEET_SYNC_MORNING_HOUR,
  writeStaffSheetLastSyncAt,
} from "./staffSheet";

const kyivTime = (
  year: number,
  month: number,
  day: number,
  hour: number,
  minute = 0,
) => {
  const probe = new Date(Date.UTC(year, month - 1, day, hour, minute));
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Kyiv",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(probe);
  const actualHour = Number(
    parts.find((part) => part.type === "hour")?.value ?? "0",
  );
  const actualMinute = Number(
    parts.find((part) => part.type === "minute")?.value ?? "0",
  );
  const deltaMinutes = hour * 60 + minute - (actualHour * 60 + actualMinute);
  return new Date(probe.getTime() + deltaMinutes * 60_000);
};

describe("staff sheet auto sync schedule", () => {
  beforeEach(() => {
    vi.stubGlobal("localStorage", {
      store: {} as Record<string, string>,
      getItem(key: string) {
        return this.store[key] ?? null;
      },
      setItem(key: string, value: string) {
        this.store[key] = value;
      },
      removeItem(key: string) {
        delete this.store[key];
      },
      clear() {
        this.store = {};
      },
    });
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("uses morning, lunch and evening Kyiv slots", () => {
    const morning = kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_MORNING_HOUR + 1);
    expect(getLatestStaffSheetSyncSlot(morning).getTime()).toBe(
      kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_MORNING_HOUR).getTime(),
    );

    const lunch = kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_LUNCH_HOUR + 1);
    expect(getLatestStaffSheetSyncSlot(lunch).getTime()).toBe(
      kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_LUNCH_HOUR).getTime(),
    );

    const evening = kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_EVENING_HOUR + 1);
    expect(getLatestStaffSheetSyncSlot(evening).getTime()).toBe(
      kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_EVENING_HOUR).getTime(),
    );
  });

  it("requires sync again after each slot boundary", () => {
    writeStaffSheetLastSyncAt(
      kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_MORNING_HOUR + 1).toISOString(),
    );
    expect(
      shouldAutoSyncStaffSheetRoster(
        kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_LUNCH_HOUR + 1),
      ),
    ).toBe(true);

    writeStaffSheetLastSyncAt(
      kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_LUNCH_HOUR + 1).toISOString(),
    );
    expect(
      shouldAutoSyncStaffSheetRoster(
        kyivTime(2026, 3, 3, STAFF_SHEET_SYNC_EVENING_HOUR + 1),
      ),
    ).toBe(true);
  });
});
