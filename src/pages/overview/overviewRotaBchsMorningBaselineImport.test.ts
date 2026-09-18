import { existsSync } from "node:fs";
import { describe, expect, it } from "vitest";
import XlsxPopulate from "xlsx-populate";
import { writeRotaBchsMorningWorkbook } from "./overviewRotaBchsMorningExport";
import { buildBchsMorningStatusChangeRows } from "./overviewRotaBchsMorningStatusChanges";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  inferUnitLabelFromBchsSheetName,
  parseBchsBaselineDateFromFileName,
  parseBchsMorningBaselineWorkbook,
} from "./overviewRotaBchsMorningBaselineImport";
import type { BchsMorningSnapshotPerson } from "./overviewRotaBchsMorningSnapshot";

const USER_BCHS_FIXTURE =
  "/Volumes/KINGSTON/army_work/Ранковий звіт/12.08.2026/бук/ПБ 3ПР БЧС 12.09.2026.xlsx";

const row = (
  overrides: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: "1",
  externalId: "1",
  validFrom: null,
  days: null,
  plannedReturn: null,
  place: "",
  updatedAt: "",
  name: "БІЛИЦЬКИЙ Сергій Володимирович",
  unit: "3 піхотна рота",
  rank: "солдат",
  status: "ON_DUTY",
  statusLabel: "В строю",
  staffStatus: "В строю",
  staffStatusLabel: "В строю",
  staffSheetColumns: {
    staff_3: "1 піхотний взвод",
    staff_5: "стрілець",
    staff_13: "солдат",
    staff_21: "В строю",
    staff_31: "ППД",
  },
  ...overrides,
});

const staffColumnsForSnapshotPerson = (person: BchsMorningSnapshotPerson) => {
  switch (person.bucket) {
    case "mission":
      return {
        staff_21: person.status || "В строю",
        staff_31: "На виконанні",
      };
    case "awol":
      return {
        staff_21: "СЗЧ",
        staff_31: person.location || "ППД",
      };
    case "leave":
      return {
        staff_21: person.status || "Відпустка",
        staff_31: person.location || "",
      };
    case "hospital":
    case "medPoint":
      return {
        staff_21: person.status || "Лік.Хвор",
        staff_31: person.location || "",
      };
    default:
      return {
        staff_21: person.status || "В строю",
        staff_31: person.location || "ППД",
      };
  }
};

describe("overviewRotaBchsMorningBaselineImport", () => {
  it("parses date and unit hints from file metadata", () => {
    expect(parseBchsBaselineDateFromFileName("ПБ 3ПР БЧС 12.09.2026.xlsx")).toBe(
      "2026-09-12",
    );
    expect(inferUnitLabelFromBchsSheetName("3ПР")).toBe("3 піхотна рота");
  });

  it("reads people back from a generated morning BCHS workbook", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/pb-rota-bchs-morning-template.xlsx",
    );
    writeRotaBchsMorningWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      staffCount: 113,
      rows: [
        row(),
        row({
          id: "2",
          externalId: "2",
          name: "КУЗЬМЕНКО Руслан Олегович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "В строю",
            staff_31: "На виконанні",
          },
        }),
      ],
    });

    const snapshot = parseBchsMorningBaselineWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      fileName: "ПБ 3ПР БЧС 12.09.2026.xlsx",
    });

    expect(snapshot.unitLabel).toBe("3 піхотна рота");
    expect(snapshot.date).toBe("2026-09-12");
    expect(snapshot.people.map((person) => person.name)).toEqual(
      expect.arrayContaining([
        "БІЛИЦЬКИЙ Сергій Володимирович",
        "КУЗЬМЕНКО Руслан Олегович",
      ]),
    );
    expect(snapshot.people.some((person) => person.bucket === "mission")).toBe(true);
  });

  it("keeps mission bucket when the same person is also listed in a platoon table", async () => {
    if (!existsSync(USER_BCHS_FIXTURE)) return;
    const wb = await XlsxPopulate.fromFileAsync(USER_BCHS_FIXTURE);
    const snapshot = parseBchsMorningBaselineWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      fileName: "ПБ 3ПР БЧС 12.09.2026.xlsx",
    });
    const missionPeople = snapshot.people.filter(
      (person) => person.bucket === "mission",
    );
    expect(missionPeople.length).toBeGreaterThanOrEqual(7);
    expect(
      missionPeople.some((person) =>
        person.name.includes("ЩЕЛКУНОВ"),
      ),
    ).toBe(true);
    const returnedPeople = snapshot.people.filter((person) =>
      /ЩЕЛКУНОВ|ШУРЧКОВ/.test(person.name),
    );
    expect(returnedPeople).toHaveLength(2);
    const returnedNames = new Set(returnedPeople.map((person) => person.name));
    const currentRows = snapshot.people.map((person, index) =>
      returnedNames.has(person.name)
        ? row({
            externalId: `returned-${index}`,
            name: person.name,
            staffSheetColumns: {
              ...row().staffSheetColumns,
              staff_21: "В строю",
              staff_31: "ППД Вишневе",
            },
          })
        : row({
            externalId: `same-${index}`,
            name: person.name,
            staffSheetColumns: {
              ...row().staffSheetColumns,
              ...staffColumnsForSnapshotPerson(person),
            },
          }),
    );
    const changes = buildBchsMorningStatusChangeRows(
      currentRows,
      snapshot,
      "3 піхотна рота",
    );
    const returnedChanges = changes.filter((change) =>
      returnedNames.has(change.name),
    );
    expect(returnedChanges).toHaveLength(2);
    expect(
      changes.filter((change) => !returnedNames.has(change.name)),
    ).toEqual([]);
    for (const change of returnedChanges) {
      expect(change.previousState).toBe("На виконанні");
      expect(change.newState).toBe("В строю · ППД Вишневе");
      expect(change.changeNote).toContain("Прийшов з виконання");
    }
  });
});
