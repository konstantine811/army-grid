import { describe, expect, it } from "vitest";
import { mergeMorningRosterRefresh } from "./overviewMorningRosterRefresh";
import { buildStaffOverviewRowsFromRoster } from "./overviewRosterMerge";
import { buildBchsMorningSections } from "./overviewRotaBchsMorningExport";

const archivePerson = {
  __rosterArchive: true, __rowNumber: 901698,
  column_1: "нова", column_2: "3 піхотна рота",
  column_3: "2 піхотний взвод", column_5: "Кулеметник",
  column_13: "солдат", column_14: "КРОШКА Іван Григорович",
  column_15: "КРОХА", column_21: "Лікування", column_31: "Шпиталь",
};

describe("morning roster refresh", () => {
  it("retains an archive-only person in the hospital table after refreshing the main sheet", () => {
    const fresh = { ...archivePerson, __rosterArchive: false, column_14: "ІВАНЕНКО Іван Іванович" };
    const rows = mergeMorningRosterRefresh([fresh], [archivePerson]);
    const overview = buildStaffOverviewRowsFromRoster(rows);
    const hospital = buildBchsMorningSections(overview, "3 піхотна рота", rows)
      .find((section) => section.id === "hospital");
    expect(hospital?.people.map((person) => person.name)).toContain(archivePerson.column_14);
  });

  it("prefers the fresh main entry when the same person is also in the archive", () => {
    const fresh = { ...archivePerson, __rosterArchive: false, column_21: "В строю", column_31: "ППД" };
    expect(mergeMorningRosterRefresh([fresh], [archivePerson])).toEqual([fresh]);
  });

  it("does not restore stale main-sheet rows missing from the fresh source", () => {
    expect(mergeMorningRosterRefresh([], [{ ...archivePerson, __rosterArchive: false }])).toEqual([]);
  });
});
