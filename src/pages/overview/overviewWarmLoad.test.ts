import { describe, expect, it } from "vitest";
import type { BackendPersonnelBootstrap } from "../../api";
import {
  canSkipOverviewDatasetReload,
  canUseOverviewWarmCacheOnly,
  isOverviewStaffCacheFresh,
} from "./overviewWarmLoad";

const bootstrap = (
  partial: Partial<BackendPersonnelBootstrap> = {},
): BackendPersonnelBootstrap => ({
  fingerprint: "fp-1",
  version: {
    oosSheetId: "",
    oosStamp: "",
    rosterImportId: "",
    rosterSheetUpdatedAt: "",
    rosterRowCount: 651,
  },
  rosterVersion: null,
  dataset: { available: true, rowCount: 651, updatedAt: "2026-01-01T00:00:00Z" },
  staff: { available: true, rowCount: 651, updatedAt: "2026-01-01T00:00:00Z" },
  assets: { available: true, rowCount: 0, updatedAt: null },
  merge: {
    available: true,
    rowCount: 0,
    updatedAt: null,
    fingerprint: "merge-fp",
  },
  ...partial,
});

describe("overviewWarmLoad", () => {
  it("skips dataset reload when warm cache matches live version", () => {
    expect(
      canSkipOverviewDatasetReload({
        datasetOnly: true,
        cachedDataset: {
          fingerprint: "fp-1",
          rows: [{} as never],
        } as never,
        versionProbe: {
          fingerprint: "fp-1",
          version: bootstrap().version,
          rosterVersion: null,
          snapshot: {
            fingerprint: "fp-1",
            rowCount: 651,
            updatedAt: "2026-01-01T00:00:00Z",
            matchesLive: true,
          },
        },
      }),
    ).toBe(true);
  });

  it("does not treat staff cache as fresh without meta", () => {
    expect(isOverviewStaffCacheFresh("fp-1", bootstrap())).toBe(false);
  });

  it("uses warm cache only on staff tab without force", () => {
    expect(
      canUseOverviewWarmCacheOnly({
        force: false,
        datasetOnly: true,
        cachedDataset: { rows: [{} as never], fingerprint: "fp-1" } as never,
      }),
    ).toBe(true);
    expect(
      canUseOverviewWarmCacheOnly({
        force: true,
        datasetOnly: true,
        cachedDataset: { rows: [{} as never], fingerprint: "fp-1" } as never,
      }),
    ).toBe(false);
  });
});
