import { describe, expect, it, vi } from "vitest";
import { api, type BackendPersonnelOverview } from "../../api";
import {
  loadPersonnelOverviewInBatches,
  mergeOverviewPages,
  overviewHasMorePages,
} from "./overviewBatchLoad";

const page = (
  ids: string[],
  total: number,
): BackendPersonnelOverview => ({
  importId: "imp",
  importName: "ЕЖООС",
  metrics: {
    total,
    onDuty: total,
    businessTrip: 0,
    leave: 0,
    medical: 0,
    awol: 0,
    other: 0,
  },
  units: ["1ПБ"],
  rows: ids.map((id) => ({
    id,
    externalId: id,
    name: `Особа ${id}`,
    rank: "солдат",
    unit: "1ПБ",
    status: "ON_DUTY",
    statusLabel: "На службі",
    validFrom: null,
    days: null,
    plannedReturn: null,
    place: "",
    updatedAt: "",
  })),
  critical: [],
  todayChanges: {
    total: 0,
    onDuty: 0,
    businessTrip: 0,
    leave: 0,
    medical: 0,
    awol: 0,
    other: 0,
  },
  todayUpdates: 0,
});

describe("mergeOverviewPages", () => {
  it("appends the next batch and keeps the first rows", () => {
    const first = page(["a", "b"], 4);
    const second = page(["c", "d"], 4);
    expect(mergeOverviewPages(first, second).rows.map((row) => row.id)).toEqual([
      "a",
      "b",
      "c",
      "d",
    ]);
  });

  it("does not duplicate when the API ignores offset and resends everyone", () => {
    const first = page(["a", "b"], 2);
    const all = page(["a", "b"], 2);
    expect(mergeOverviewPages(first, all).rows).toHaveLength(2);
  });
});

describe("overviewHasMorePages", () => {
  it("stops when the first response already has every row", () => {
    const all = page(["a", "b", "c"], 3);
    expect(overviewHasMorePages(all, all.rows.length, 20)).toBe(false);
  });

  it("asks for the next pack while the list is still short", () => {
    const first = page(["a", "b"], 5);
    expect(overviewHasMorePages(first, first.rows.length, 2)).toBe(true);
  });
});

describe("loadPersonnelOverviewInBatches", () => {
  it("bypasses the server cache only for an explicit refresh", async () => {
    const getOverview = vi
      .spyOn(api, "getPersonnelOverview")
      .mockResolvedValue(page(["a"], 1));

    await loadPersonnelOverviewInBatches({ force: true });

    expect(getOverview).toHaveBeenCalledWith({
      limit: 200,
      offset: 0,
      force: true,
      signal: undefined,
    });
    getOverview.mockRestore();
  });
});
