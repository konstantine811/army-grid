import { describe, expect, it } from "vitest";
import type {
  BackendPersonnelRosterLatest,
  BackendPersonnelRosterVersion,
} from "../api";
import {
  isRosterVersionCurrent,
  ROSTER_VERSION_CHECK_TTL_MS,
  shouldCheckRosterVersion,
} from "./sharedAppData";

const cachedRoster = {
  importId: "import-2",
  importName: "Штатка",
  createdAt: "2026-09-05T08:00:00.000Z",
  sheet: {
    id: "sheet-2",
    batchId: "import-2",
    name: "Загальний список",
    sheetIndex: 0,
    columnCount: 40,
    rowCount: 2946,
    columns: [],
    createdAt: "2026-09-05T08:00:00.000Z",
    updatedAt: "2026-09-05T08:00:00.000Z",
  },
  rows: [],
} as BackendPersonnelRosterLatest;

describe("isRosterVersionCurrent", () => {
  it("accepts only the same server import, row count and sheet timestamp", () => {
    const version: BackendPersonnelRosterVersion = {
      importId: "import-2",
      createdAt: "2026-09-05T08:00:00.000Z",
      sheetUpdatedAt: "2026-09-05T08:00:00.000Z",
      rowCount: 2946,
    };
    expect(isRosterVersionCurrent(cachedRoster, version)).toBe(true);
    expect(
      isRosterVersionCurrent(cachedRoster, {
        ...version,
        importId: "import-1",
        rowCount: 2286,
      }),
    ).toBe(false);
  });

  it("does not request the roster version on every page transition", () => {
    const checkedAt = 1_000_000;
    expect(shouldCheckRosterVersion(checkedAt, checkedAt + 5_000)).toBe(false);
    expect(
      shouldCheckRosterVersion(
        checkedAt,
        checkedAt + ROSTER_VERSION_CHECK_TTL_MS,
      ),
    ).toBe(true);
  });
});
