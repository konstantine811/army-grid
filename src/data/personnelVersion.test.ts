import { describe, expect, it, vi, beforeEach } from "vitest";
import { api } from "../api";
import { loadPersonnelVersionProbe } from "./personnelVersion";

describe("personnelVersion", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
  });

  it("deduplicates concurrent version probes", async () => {
    const probe = {
      fingerprint: "fp-live",
      version: {
        oosSheetId: "sheet-1",
        oosStamp: "2026-01-01|100|20",
        rosterImportId: "roster-1",
        rosterSheetUpdatedAt: "2026-01-01",
        rosterRowCount: 651,
      },
      rosterVersion: null,
      snapshot: null,
    };
    const request = vi
      .spyOn(api, "getPersonnelVersion")
      .mockImplementation(
        () =>
          new Promise((resolve) => {
            setTimeout(() => resolve(probe), 10);
          }),
      );

    const [first, second] = await Promise.all([
      loadPersonnelVersionProbe(),
      loadPersonnelVersionProbe(),
    ]);

    expect(first).toEqual(probe);
    expect(second).toEqual(probe);
    expect(request).toHaveBeenCalledTimes(1);
  });
});
