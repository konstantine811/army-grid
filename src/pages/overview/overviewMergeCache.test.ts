import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverview } from "../../api";
import {
  overviewMergeCacheKey,
  overviewMergeFingerprint,
} from "./overviewMergeCache";

const overview = (
  rows: Array<{ id: string; externalId?: string }>,
  importId = "import-1",
): Pick<BackendPersonnelOverview, "importId" | "rows"> => ({
  importId,
  rows: rows.map((row) => ({
    id: row.id,
    externalId: row.externalId ?? row.id,
    name: "ТЕСТ",
    rank: "солдат",
    unit: "1 рота",
    status: "ON_DUTY",
    statusLabel: "На службі",
    validFrom: null,
    days: null,
    plannedReturn: null,
    place: "",
    updatedAt: "",
  })),
});

describe("overviewMergeCacheKey", () => {
  it("changes when overview row count changes", () => {
    const rosterFingerprint = "oos|roster|fp";
    const first = overviewMergeCacheKey(
      overview([{ id: "a" }, { id: "b" }]),
      rosterFingerprint,
      [],
      [],
    );
    const second = overviewMergeCacheKey(
      overview([{ id: "a" }, { id: "b" }, { id: "c" }]),
      rosterFingerprint,
      [],
      [],
    );
    expect(first).not.toBe(second);
  });

  it("is stable for the same inputs", () => {
    const input = overview([{ id: "a" }], "ej-1");
    const keyA = overviewMergeCacheKey(input, "fp-1", [{ __dbRowId: "r1" }], [
      { key: "column_1" },
    ]);
    const keyB = overviewMergeCacheKey(input, "fp-1", [{ __dbRowId: "r1" }], [
      { key: "column_1" },
    ]);
    expect(keyA).toBe(keyB);
    expect(keyA.startsWith("personnel:overview-merge:v1:")).toBe(true);
  });
});

describe("overviewMergeFingerprint", () => {
  it("matches between client cache key prefix and server snapshot id", () => {
    const input = overview([{ id: "a" }, { id: "b" }], "ej-1");
    const fingerprint = overviewMergeFingerprint(input, "dataset-fp");
    const cacheKey = overviewMergeCacheKey(input, "dataset-fp", [], []);
    expect(cacheKey.startsWith(`personnel:${fingerprint}:`)).toBe(true);
  });
});
