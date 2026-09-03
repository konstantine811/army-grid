import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverview } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { getRosterPersonName, getRosterUnit, isNovaRosterRow } from "../personnel/personnelRosterMerge";
import {
  buildStaffOverviewRowsFromPersonnel,
  buildStaffOverviewRowsFromRoster,
  collectRosterUnitOptions,
  fillDownRosterUnitRows,
  mergeRosterRowsIntoOverview,
  overviewStatusFilterLabel,
  rosterRowToOverviewRow,
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

describe("getRosterUnit", () => {
  it("reads підрозділ from column 2, not місце перебування", () => {
    expect(
      getRosterUnit({
        column_2: "2 піхотна рота",
        column_31: "в/ч А1363",
        column_29: "AB339559",
      } as EjournalPreviewRow),
    ).toBe("2 піхотна рота");
  });
});

describe("collectRosterUnitOptions", () => {
  it("returns unique staff units for the selected battalion", () => {
    const rosterRows = [
      { column_1: "нова", column_2: "1 піхотна рота", column_14: "КОВАЛЬ" },
      { column_1: "нова", column_2: "2 піхотна рота", column_14: "ПЕТРЕНКО" },
      { column_1: "нова", column_2: "2 піхотна рота" },
      { column_1: "стара", column_2: "Штаб", column_14: "СИДОРЕНКО" },
    ] as EjournalPreviewRow[];

    expect(collectRosterUnitOptions(rosterRows)).toEqual([
      "1 піхотна рота",
      "2 піхотна рота",
      "Штаб",
    ]);
    expect(collectRosterUnitOptions(rosterRows, undefined, "нова")).toEqual([
      "1 піхотна рота",
      "2 піхотна рота",
    ]);
  });
});

describe("rosterRowToOverviewRow", () => {
  it("uses column 2 for unit", () => {
    const row = rosterRowToOverviewRow(
      {
        column_1: "нова",
        column_2: "Мінометний взвод",
        column_14: "КОВАЛЬ Іван",
        column_31: "3765688",
      } as EjournalPreviewRow,
      {},
      { inNovaStaff: true, battalion: "нова" },
    );
    expect(row?.unit).toBe("Мінометний взвод");
  });
});

describe("buildStaffOverviewRowsFromRoster", () => {
  it("builds overview rows directly from staff roster without ejoos merge", () => {
    const rosterRows = [
      {
        __dbRowId: "r1",
        column_1: "нова",
        column_2: "2 піхотна рота",
        column_13: "солдат",
        column_14: "КОВАЛЬ Іван Петрович",
        column_21: "В строю",
      },
      {
        __dbRowId: "r2",
        column_1: "нова",
        column_2: "2 піхотна рота",
        column_14: "",
      },
    ] as EjournalPreviewRow[];

    const rows = buildStaffOverviewRowsFromRoster(rosterRows);
    expect(rows).toHaveLength(1);
    expect(rows[0]?.name).toBe("КОВАЛЬ Іван Петрович");
    expect(rows[0]?.unit).toBe("2 піхотна рота");
    expect(rows[0]?.fromEjoos).toBe(false);
    expect(rows[0]?.statusLabel).toBe("В строю");
  });
});

describe("fillDownRosterUnitRows", () => {
  it("inherits a merged Google Sheet unit for following position rows", () => {
    const rows = fillDownRosterUnitRows([
      {
        __dbRowId: "1",
        column_2: "2 піхотна рота",
        column_5: "Командир",
        column_14: "Іванов Іван",
      },
      {
        __dbRowId: "2",
        column_2: "",
        column_5: "Стрілець",
        column_14: "Петренко Петро",
      },
      {
        __dbRowId: "3",
        column_2: "",
        column_5: "Кулеметник",
        column_14: "",
      },
      {
        __dbRowId: "4",
        column_2: "3 піхотна рота",
        column_5: "Командир",
      },
    ] as EjournalPreviewRow[]);

    expect(rows.map((row) => row.column_2)).toEqual([
      "2 піхотна рота",
      "2 піхотна рота",
      "2 піхотна рота",
      "3 піхотна рота",
    ]);
  });
});

describe("buildStaffOverviewRowsFromPersonnel", () => {
  it("uses the same merged in-staff cards as Особовий склад", () => {
    const personnelRows = [
      {
        __dbRowId: "oos-1",
        прізвище: "КОВАЛЬ Іван Петрович",
        звання: "старший солдат",
        roster__column_1: "нова",
        roster__column_2: "2 піхотна рота",
        roster__column_13: "солдат",
        roster__column_14: "КОВАЛЬ Іван Петрович (зі штатки)",
        roster__column_21: "В строю",
      },
      {
        __dbRowId: "oos-2",
        прізвище: "ПЕТРЕНКО Олег Іванович",
      },
    ] as EjournalPreviewRow[];

    const rows = buildStaffOverviewRowsFromPersonnel(personnelRows);

    expect(rows).toHaveLength(2);
    expect(rows[0]).toMatchObject({
      name: "КОВАЛЬ Іван Петрович",
      rank: "старший солдат",
      unit: "2 піхотна рота",
      battalion: "нова",
      inStaff: true,
    });
    expect(rows[1]).toMatchObject({
      name: "ПЕТРЕНКО Олег Іванович",
      inStaff: false,
    });
  });
});

describe("overviewStatusFilterLabel", () => {
  it("groups all medical wording variants under Лікування", () => {
    const rows = [
      { status: "MEDICAL", statusLabel: "Лікування" },
      { status: "MEDICAL", statusLabel: "На лікуванні" },
      { status: "MEDICAL", statusLabel: "лікування після поранення" },
      { status: "ON_DUTY", statusLabel: "В строю" },
    ] as never[];

    expect(rows.map(overviewStatusFilterLabel)).toEqual([
      "Лікування",
      "Лікування",
      "Лікування",
      "В строю",
    ]);
  });

  it("keeps the original label for general on-duty statuses", () => {
    expect(
      overviewStatusFilterLabel({
        status: "ON_DUTY",
        statusLabel: "Новоприбулий",
      } as never),
    ).toBe("Новоприбулий");
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

  it("overwrites ejoos unit with roster column 2", () => {
    const overview = emptyOverview([
      ejoosRow("КОВАЛЬ Іван Петрович", "2103001"),
    ]);
    const rosterRows = [
      {
        __dbRowId: "r1",
        column_1: "нова",
        column_2: "2 піхотна рота",
        column_14: "КОВАЛЬ Іван Петрович",
        column_31: "в/ч А1363",
      },
    ] as EjournalPreviewRow[];

    const merged = mergeRosterRowsIntoOverview(overview, rosterRows);
    expect(merged.rows[0]?.unit).toBe("2 піхотна рота");
    expect(merged.units).toContain("2 піхотна рота");
    expect(merged.units).not.toContain("в/ч А1363");
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
