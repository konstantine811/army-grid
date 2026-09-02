import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverview } from "../api";
import type { EjournalPreviewRow } from "../pages/ejournal/ejournalTypes";
import { runHeavyJobSync } from "./heavyJobs";

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

describe("runHeavyJobSync", () => {
  it("extracts named nova people for BCHS", () => {
    const result = runHeavyJobSync({
      type: "bchsExtractPeople",
      rows: [
        { column_1: "нова", column_14: "КОВАЛЬ Іван Петрович" },
        { column_1: "нова", column_14: "" },
        { column_1: "стара", column_14: "ПЕТРЕНКО Олег Іванович" },
      ],
    });

    expect(result.novaCount).toBe(1);
    expect(result.people.some((person) => person.fullName.includes("КОВАЛЬ"))).toBe(
      true,
    );
  });

  it("merges roster rows into overview the same way as the page helper", () => {
    const merged = runHeavyJobSync({
      type: "mergeOverview",
      overview: emptyOverview([
        {
          id: "ejoos:2103001",
          externalId: "2103001",
          name: "КОВАЛЬ Іван Петрович",
          rank: "солдат",
          unit: "1 рота",
          status: "ON_DUTY",
          statusLabel: "На службі",
          validFrom: null,
          days: null,
          plannedReturn: null,
          place: "",
          updatedAt: "",
        },
      ]),
      rosterRows: [
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
      ] as EjournalPreviewRow[],
    });

    expect(merged.rows.filter((row) => row.inNovaStaff)).toHaveLength(1);
  });

  it("returns preview rows unchanged when there is no roster", () => {
    const rows = [
      { __dbRowId: "p1", column_14: "КОВАЛЬ Іван Петрович" },
    ] as EjournalPreviewRow[];
    const merged = runHeavyJobSync({
      type: "mergePersonnel",
      preview: { rows },
      rosterRows: [],
    });
    expect(merged).toBe(rows);
  });
});
