import { describe, expect, it } from "vitest";
import type {
  BackendEjournalImportSheet,
  BackendPersonnelRosterLatest,
} from "../api";
import {
  buildPersonnelDatasetVersion,
  personnelDatasetFingerprint,
  personnelDatasetToPreview,
  rosterRowsFromPersonnelLatest,
  sortPersonnelRowsByRosterOrder,
  type PersonnelDataset,
} from "./personnelDataset";

const sheet = (
  id: string,
  updatedAt: string,
  rowCount = 10,
): BackendEjournalImportSheet => ({
  id,
  batchId: `batch-${id}`,
  name: "2. ООС",
  sheetIndex: 1,
  columnCount: 20,
  rowCount,
  columns: [],
  createdAt: updatedAt,
  updatedAt,
});

const roster = (
  updatedAt = "2026-09-05T08:00:00.000Z",
): BackendPersonnelRosterLatest => ({
  importId: "roster-1",
  importName: "Штатка",
  createdAt: updatedAt,
  sheet: { ...sheet("staff", updatedAt, 2), name: "Загальний список" },
  rows: [
    {
      id: "r1",
      sheetId: "staff",
      excelRowNumber: 2,
      createdAt: updatedAt,
      values: { column_2: "1 рота", column_14: "ПЕРШИЙ Боєць" },
    },
    {
      id: "r2",
      sheetId: "staff",
      excelRowNumber: 3,
      createdAt: updatedAt,
      values: { column_2: "", column_14: "ДРУГИЙ Боєць" },
    },
  ],
});

describe("personnelDataset", () => {
  it("changes the fingerprint whenever OOS or roster version changes", () => {
    const first = personnelDatasetFingerprint(
      buildPersonnelDatasetVersion(
        sheet("oos", "2026-09-05T08:00:00.000Z"),
        roster(),
      ),
    );
    const same = personnelDatasetFingerprint(
      buildPersonnelDatasetVersion(
        sheet("oos", "2026-09-05T08:00:00.000Z"),
        roster(),
      ),
    );
    const changedOos = personnelDatasetFingerprint(
      buildPersonnelDatasetVersion(
        sheet("oos", "2026-09-05T09:00:00.000Z"),
        roster(),
      ),
    );
    const changedRoster = personnelDatasetFingerprint(
      buildPersonnelDatasetVersion(
        sheet("oos", "2026-09-05T08:00:00.000Z"),
        roster("2026-09-05T10:00:00.000Z"),
      ),
    );

    expect(same).toBe(first);
    expect(changedOos).not.toBe(first);
    expect(changedRoster).not.toBe(first);
  });

  it("normalizes merged roster units once for both pages", () => {
    const rows = rosterRowsFromPersonnelLatest(roster());

    expect(rows).toHaveLength(2);
    expect(rows[0].column_2).toBe("1 рота");
    expect(rows[1].column_2).toBe("1 рота");
  });

  it("exposes the exact same row array through the Personnel preview", () => {
    const rows = [{ __dbRowId: "person-1", ПІБ: "СПІЛЬНИЙ Боєць" }];
    const dataset: PersonnelDataset = {
      rows,
      sheet: sheet("oos", "2026-09-05T08:00:00.000Z"),
      columns: [],
      total: 1,
      rosterRows: [],
      rosterLabels: {},
      rosterColumns: [],
      rosterUpdatedAt: null,
      version: buildPersonnelDatasetVersion(null, null),
      fingerprint: "",
      mergedAt: 1,
      complete: true,
    };

    expect(personnelDatasetToPreview(dataset)?.rows).toBe(rows);
  });

  it("keeps staff rows in roster order instead of surname order", () => {
    const rows = sortPersonnelRowsByRosterOrder([
      { __dbRowId: "alphabetical-first", __rosterOrder: 1, ПІБ: "АБРАМЕНКО" },
      { __dbRowId: "staff-first", __rosterOrder: 0, ПІБ: "ЯКОВЕНКО" },
      { __dbRowId: "outside-roster", ПІБ: "БОНДАРЕНКО" },
    ]);

    expect(rows.map((row) => row.__dbRowId)).toEqual([
      "staff-first",
      "alphabetical-first",
      "outside-roster",
    ]);
  });
});
