import { describe, expect, it, vi, beforeEach } from "vitest";
import type { BackendPersonnelBootstrap } from "../../api";
import { api } from "../../api";
import { loadPersonnelBootstrapMeta } from "../../data/personnelVersion";
import {
  fetchMergedOverviewSnapshot,
  fetchOverviewAssetsSnapshot,
  fetchOverviewStaffSnapshot,
} from "./overviewServerSnapshots";

vi.mock("../../data/personnelVersion", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../data/personnelVersion")>();
  return {
    ...actual,
    loadPersonnelBootstrapMeta: vi.fn(actual.loadPersonnelBootstrapMeta),
  };
});

const peekDataCacheMock = vi.fn<(key: string) => unknown>(() => null);

vi.mock("../../data/idbDataCache", async (importOriginal) => {
  const actual = await importOriginal<typeof import("../../data/idbDataCache")>();
  return {
    ...actual,
    peekDataCache: (key: string) => peekDataCacheMock(key),
    writeDataCache: vi.fn(async () => undefined),
  };
});

const bootstrap = (
  partial: Partial<BackendPersonnelBootstrap> = {},
): BackendPersonnelBootstrap => ({
  fingerprint: "fp-fresh",
  version: {
    oosSheetId: "",
    oosStamp: "",
    rosterImportId: "",
    rosterSheetUpdatedAt: "",
    rosterRowCount: 1,
  },
  rosterVersion: null,
  dataset: { available: true, rowCount: 1, updatedAt: "2026-01-01T00:00:00Z" },
  staff: { available: true, rowCount: 1, updatedAt: "2026-01-01T00:00:00Z" },
  assets: { available: true, rowCount: 0, updatedAt: null },
  merge: {
    available: true,
    rowCount: 0,
    updatedAt: null,
    fingerprint: "merge-fp",
  },
  ...partial,
});

describe("overviewServerSnapshots", () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    peekDataCacheMock.mockReset();
    peekDataCacheMock.mockReturnValue(null);
    vi.mocked(loadPersonnelBootstrapMeta).mockResolvedValue(null);
  });

  it("skips staff GET when local cache matches bootstrap updatedAt", async () => {
    const staff = { rows: [{ id: "1" }], importId: "x" };
    const getStaff = vi.spyOn(api, "getOverviewStaff");
    const rebuild = vi.spyOn(api, "rebuildOverviewStaff");
    vi.mocked(loadPersonnelBootstrapMeta).mockResolvedValue(bootstrap());

    peekDataCacheMock.mockImplementation((key) => {
      if (key === "personnel:overview-staff:v1:fp-fresh") return staff;
      if (key === "personnel:overview-staff:v1:meta:fp-fresh") {
        return {
          serverUpdatedAt: "2026-01-01T00:00:00Z",
          rowCount: 1,
        };
      }
      return null;
    });

    const result = await fetchOverviewStaffSnapshot({
      fingerprint: "fp-fresh",
    });

    expect(result).toBe(staff);
    expect(getStaff).not.toHaveBeenCalled();
    expect(rebuild).not.toHaveBeenCalled();
  });

  it("rebuilds staff snapshot on cache miss when allowed", async () => {
    vi.spyOn(api, "getOverviewStaff").mockResolvedValue(null);
    const rebuilt = { rows: [{ id: "2" }], importId: "y" };
    vi.spyOn(api, "rebuildOverviewStaff").mockResolvedValue(rebuilt as never);

    const result = await fetchOverviewStaffSnapshot({
      fingerprint: "fp-2",
      allowRebuild: true,
    });

    expect(result).toBe(rebuilt);
  });

  it("does not rebuild staff snapshot on cache miss by default", async () => {
    vi.spyOn(api, "getOverviewStaff").mockResolvedValue(null);
    const rebuild = vi.spyOn(api, "rebuildOverviewStaff");

    const result = await fetchOverviewStaffSnapshot({
      fingerprint: "fp-empty",
    });

    expect(result).toBeNull();
    expect(rebuild).not.toHaveBeenCalled();
  });

  it("skips staff GET when bootstrap reports snapshot missing", async () => {
    vi.mocked(loadPersonnelBootstrapMeta).mockResolvedValue({
      fingerprint: "fp-miss",
      version: {
        oosSheetId: "",
        oosStamp: "",
        rosterImportId: "",
        rosterSheetUpdatedAt: "",
        rosterRowCount: 0,
      },
      rosterVersion: null,
      dataset: { available: false, rowCount: 0, updatedAt: null },
      staff: { available: false, rowCount: 0, updatedAt: null },
      assets: { available: false, rowCount: 0, updatedAt: null },
      merge: {
        available: false,
        rowCount: 0,
        updatedAt: null,
        fingerprint: null,
      },
    });
    const getStaff = vi.spyOn(api, "getOverviewStaff");
    const rebuild = vi.spyOn(api, "rebuildOverviewStaff");

    const result = await fetchOverviewStaffSnapshot({
      fingerprint: "fp-miss",
    });

    expect(getStaff).not.toHaveBeenCalled();
    expect(rebuild).not.toHaveBeenCalled();
    expect(result).toBeNull();
  });

  it("rebuilds assets snapshot on cache miss", async () => {
    vi.spyOn(api, "getOverviewAssets").mockResolvedValue(null);
    const rebuilt = {
      fingerprint: "fp-3",
      questionnairePresence: {},
      questionnaireSourceIds: {},
      documents: {},
    };
    vi.spyOn(api, "rebuildOverviewAssets").mockResolvedValue(rebuilt);

    const result = await fetchOverviewAssetsSnapshot({
      fingerprint: "fp-3",
    });

    expect(result).toBe(rebuilt);
  });

  it("rebuilds merged overview on cache miss", async () => {
    vi.spyOn(api, "getMergedPersonnelOverview").mockResolvedValue(null);
    const rebuilt = { rows: [{ id: "m1" }], importId: "merge" };
    vi.spyOn(api, "rebuildMergedPersonnelOverview").mockResolvedValue(
      rebuilt as never,
    );

    const result = await fetchMergedOverviewSnapshot({
      mergeFingerprint: "merge-fp",
    });

    expect(result).toBe(rebuilt);
  });
});
