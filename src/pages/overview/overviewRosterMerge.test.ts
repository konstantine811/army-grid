import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverview } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { getRosterPersonName, isNovaRosterRow } from "../personnel/personnelRosterMerge";
import {
  mergeRosterRowsIntoOverview,
  summarizeNovaStaffFromRoster,
  summarizeStaffFromRoster,
} from "./overviewRosterMerge";

const emptyOverview = (
  rows: BackendPersonnelOverview["rows"] = [],
): BackendPersonnelOverview => ({
  importId: "imp",
  importName: "ЕЖООС",
  rows,
  metrics: {
    total: rows.length,
    onDuty: 0,
    businessTrip: 0,
    leave: 0,
    medical: 0,
    awol: 0,
    other: 0,
  },
  units: [],
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

const ejoosRow = (name: string, id = "") => ({
  id: id || `ejoos:${name}`,
  externalId: id,
  name,
  rank: "солдат",
  unit: "1 рота",
  status: "ON_DUTY",
  statusLabel: "На службі",
  validFrom: null,
  days: null,
  plannedReturn: null,
  place: "",
  updatedAt: "",
});

describe("getRosterPersonName", () => {
  it("prefers a real ПІБ over звання in column 13", () => {
    expect(
      getRosterPersonName({
        column_13: "солдат",
        column_14: "ШЕВЧЕНКО Олександр Володимирович",
      } as EjournalPreviewRow),
    ).toBe("ШЕВЧЕНКО Олександр Володимирович");
  });

  it("reads ПІБ from column 13 when column 14 is a callsign", () => {
    expect(
      getRosterPersonName({
        column_13: "ШЕВЧЕНКО Олександр Володимирович",
        column_14: "Сокіл",
      } as EjournalPreviewRow),
    ).toBe("ШЕВЧЕНКО Олександр Володимирович");
  });
});

describe("isNovaRosterRow", () => {
  it("treats battalion column_1 = нова as nova", () => {
    expect(isNovaRosterRow({ column_1: "нова" } as EjournalPreviewRow)).toBe(true);
    expect(isNovaRosterRow({ column_1: "стара" } as EjournalPreviewRow)).toBe(
      false,
    );
  });
});

describe("mergeRosterRowsIntoOverview", () => {
  it("counts named nova roster rows in staff and skips empty ПІБ", () => {
    const overview = emptyOverview([
      ejoosRow("КОВАЛЬ Іван Петрович", "2103001"),
    ]);
    const rosterRows = [
      {
        __dbRowId: "r1",
        __rowNumber: 2,
        column_1: "нова",
        column_14: "КОВАЛЬ Іван Петрович",
      },
      {
        __dbRowId: "r2",
        __rowNumber: 3,
        column_1: "нова",
        column_2: "2 рота",
      },
      {
        __dbRowId: "r3",
        __rowNumber: 4,
        column_1: "стара",
        column_14: "ПЕТРЕНКО Олег Іванович",
      },
    ] as EjournalPreviewRow[];

    const merged = mergeRosterRowsIntoOverview(overview, rosterRows);
    const staff = merged.rows.filter((row) => row.inNovaStaff);

    expect(staff).toHaveLength(1);
    expect(staff.some((row) => row.name === "КОВАЛЬ Іван Петрович")).toBe(true);
    expect(staff.some((row) => row.name === "Без ПІБ")).toBe(false);
    expect(merged.rows.some((row) => row.name === "ПЕТРЕНКО Олег Іванович")).toBe(
      true,
    );
    expect(
      merged.rows.find((row) => row.name === "ПЕТРЕНКО Олег Іванович")
        ?.inNovaStaff,
    ).toBe(false);
    expect(
      merged.rows.find((row) => row.name === "ПЕТРЕНКО Олег Іванович")?.battalion,
    ).toBe("стара");
    expect(merged.rows.filter((row) => row.inStaff)).toHaveLength(2);
  });

  it("summarizes all battalions when not limited to нова", () => {
    const rosterRows = [
      { __dbRowId: "r1", column_1: "нова", column_14: "КОВАЛЬ Іван" },
      { __dbRowId: "r2", column_1: "стара", column_14: "ПЕТРЕНКО Олег" },
    ] as EjournalPreviewRow[];

    const all = summarizeStaffFromRoster(rosterRows);
    expect(all.people).toBe(2);
    expect(all.battalions).toEqual(["нова", "стара"]);
    expect(summarizeStaffFromRoster(rosterRows, undefined, "стара").people).toBe(
      1,
    );
  });

  it("keeps a nova person whose name is only in fullName", () => {
    const rosterRows = [
      {
        __dbRowId: "r1",
        __rowNumber: 2,
        column_1: "нова",
        fullName: "СИДОРЕНКО Василь Іванович",
      },
    ] as EjournalPreviewRow[];

    const summary = summarizeNovaStaffFromRoster(rosterRows);
    expect(summary).toMatchObject({ positions: 1, people: 1, vacant: 0 });

    const merged = mergeRosterRowsIntoOverview(emptyOverview(), rosterRows);
    const staff = merged.rows.filter((row) => row.inNovaStaff);
    expect(staff).toHaveLength(1);
    expect(staff[0]?.name).toMatch(/СИДОРЕНКО/i);
  });

  it("counts vacant nova rows as positions, not people", () => {
    const summary = summarizeNovaStaffFromRoster([
      { __dbRowId: "r1", column_1: "нова", column_14: "КОВАЛЬ Іван" },
      { __dbRowId: "r2", column_1: "нова" },
      { __dbRowId: "r3", column_1: "стара", column_14: "ПЕТРЕНКО Олег" },
    ] as EjournalPreviewRow[]);

    expect(summary.positions).toBe(2);
    expect(summary.people).toBe(1);
    expect(summary.vacant).toBe(1);
  });
});
