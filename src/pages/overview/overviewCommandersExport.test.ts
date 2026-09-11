import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  buildOverviewCommanderExportRows,
  buildOverviewCommandersExportFileName,
  buildOverviewCommandersExportSheets,
  groupOverviewCommandersByUnit,
  isCompanyCommanderPosition,
  isDeputyCommanderPosition,
  resolveOverviewCommanderRole,
} from "./overviewCommandersExport";

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
    staff_4: "1 відділення",
    staff_5: "стрілець",
    staff_13: "солдат",
    staff_14: "БІЛИЦЬКИЙ Сергій Володимирович",
    staff_15: "НАЛІМ",
    staff_21: "В строю",
    staff_31: "ППД Вишневе",
  },
  ...overrides,
});

describe("overviewCommandersExport", () => {
  it("detects company, platoon and section commanders and skips deputies", () => {
    expect(isCompanyCommanderPosition("командир роти")).toBe(true);
    expect(isCompanyCommanderPosition("Командир 3 піхотної роти")).toBe(true);
    expect(isCompanyCommanderPosition("командир ротного взводу")).toBe(false);
    expect(isDeputyCommanderPosition("Заступник командира роти")).toBe(true);
    expect(isCompanyCommanderPosition("Заступник командира роти")).toBe(false);

    expect(
      resolveOverviewCommanderRole(
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "командир роти",
          },
        }),
      ),
    ).toBe("Командир роти");
    expect(
      resolveOverviewCommanderRole(
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "Командир взводу",
          },
        }),
      ),
    ).toBe("Командир взводу");
    expect(
      resolveOverviewCommanderRole(
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "командир ротного взводу",
          },
        }),
      ),
    ).toBe("Командир взводу");
    expect(
      resolveOverviewCommanderRole(
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "командир відділення",
          },
        }),
      ),
    ).toBe("Командир відділення");
    expect(
      resolveOverviewCommanderRole(
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "Заступник командира взводу",
          },
        }),
      ),
    ).toBeNull();
    expect(resolveOverviewCommanderRole(row())).toBeNull();
  });

  it("builds a sorted commander table and drops the dead", () => {
    const commanders = buildOverviewCommanderExportRows([
      row({
        id: "s2",
        externalId: "s2",
        name: "ВІДДІЛЕННЯ Другий",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_3: "2 піхотний взвод",
          staff_4: "2",
          staff_5: "командир відділення",
          staff_14: "ВІДДІЛЕННЯ Другий",
        },
      }),
      row({
        id: "p1",
        externalId: "p1",
        name: "ВЗВОДНИЙ Перший",
        unit: "1 піхотна рота",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_5: "Командир 1 піхотного взводу",
          staff_13: "лейтенант",
          staff_14: "ВЗВОДНИЙ Перший",
          staff_15: "СОКІЛ",
        },
      }),
      row({
        id: "c3",
        externalId: "c3",
        name: "РОТНИЙ Третій",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_3: "ж",
          staff_5: "командир роти",
          staff_13: "капітан",
          staff_14: "РОТНИЙ Третій",
        },
      }),
      row({
        id: "dead",
        externalId: "dead",
        name: "ЗАГИБЛИЙ Командир",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_5: "командир взводу",
          staff_14: "ЗАГИБЛИЙ Командир",
          staff_21: "Загиблий",
        },
      }),
      row({
        id: "dep",
        externalId: "dep",
        name: "ЗАСТУПНИК Петро",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_5: "Заступник командира роти",
          staff_14: "ЗАСТУПНИК Петро",
        },
      }),
    ]);

    expect(commanders.map((person) => person.name)).toEqual([
      "ВЗВОДНИЙ Перший",
      "РОТНИЙ Третій",
      "ВІДДІЛЕННЯ Другий",
    ]);
    expect(commanders[0]).toMatchObject({
      role: "Командир взводу",
      callsign: "СОКІЛ",
      unit: "1 піхотна рота",
    });
    expect(commanders[1]).toMatchObject({
      role: "Командир роти",
      rank: "капітан",
      unit: "3 піхотна рота",
    });
    expect(commanders[2]).toMatchObject({
      role: "Командир відділення",
      position: "Командир 2 відділення",
      section: "2 відділення",
    });
  });

  it("puts every rota table one after another on a single sheet", () => {
    const groups = groupOverviewCommandersByUnit(
      buildOverviewCommanderExportRows([
        row({
          id: "s1",
          externalId: "s1",
          name: "ВІДДІЛЕННЯ Перший",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_4: "1",
            staff_5: "командир відділення",
            staff_14: "ВІДДІЛЕННЯ Перший",
          },
        }),
        row({
          id: "s2",
          externalId: "s2",
          name: "ВІДДІЛЕННЯ Другий",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_4: "2",
            staff_5: "командир відділення",
            staff_14: "ВІДДІЛЕННЯ Другий",
          },
        }),
        row({
          id: "p2",
          externalId: "p2",
          name: "ВЗВОДНИЙ Другий",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
            staff_5: "Командир взводу",
            staff_14: "ВЗВОДНИЙ Другий",
          },
        }),
        row({
          id: "p1",
          externalId: "p1",
          name: "ВЗВОДНИЙ Перший",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "Командир взводу",
            staff_14: "ВЗВОДНИЙ Перший",
          },
        }),
        row({
          id: "c3",
          externalId: "c3",
          name: "РОТНИЙ Третій",
          unit: "1 піхотна рота",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "ж",
            staff_5: "командир роти",
            staff_14: "РОТНИЙ Третій",
          },
        }),
        row({
          id: "c1",
          externalId: "c1",
          name: "РОТНИЙ Перший",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "ж",
            staff_5: "командир роти",
            staff_14: "РОТНИЙ Перший",
          },
        }),
      ]),
    );

    expect(groups.map((group) => group.unit)).toEqual([
      "1 піхотна рота",
      "3 піхотна рота",
    ]);
    expect(groups[0]?.companyCommanders.map((person) => person.name)).toEqual([
      "РОТНИЙ Третій",
    ]);
    expect(groups[1]?.companyCommanders.map((person) => person.name)).toEqual([
      "РОТНИЙ Перший",
    ]);
    expect(groups[1]?.platoons.map((platoon) => platoon.title)).toEqual([
      "1 піхотний взвод",
      "2 піхотний взвод",
    ]);
    expect(groups[1]?.platoons[0]?.commanders.map((person) => person.name)).toEqual([
      "ВЗВОДНИЙ Перший",
    ]);
    expect(groups[1]?.platoons[0]?.sections.map((section) => section.title)).toEqual([
      "1 відділення",
      "2 відділення",
    ]);

    const sheets = buildOverviewCommandersExportSheets(
      groups.flatMap((group) => [
        ...group.companyCommanders,
        ...group.platoons.flatMap((platoon) => [
          ...platoon.commanders,
          ...platoon.sections.flatMap((section) => section.commanders),
        ]),
      ]),
    );
    expect(sheets.map((sheet) => sheet.sheet)).toEqual(["Командири"]);
    expect(
      sheets[0]?.data
        .map((excelRow) => String(excelRow[0]?.value ?? ""))
        .filter(Boolean),
    ).toEqual([
      "1 піхотна рота",
      "№",
      "Командування роти",
      "1",
      "3 піхотна рота",
      "№",
      "Командування роти",
      "1",
      "1 піхотний взвод",
      "2",
      "1 відділення",
      "3",
      "2 відділення",
      "4",
      "2 піхотний взвод",
      "5",
    ]);
  });

  it("names the file after the selected unit when present", () => {
    expect(
      buildOverviewCommandersExportFileName(
        "3 піхотна рота",
        new Date("2026-09-10T12:00:00"),
      ),
    ).toBe("Командири 3 піхотна рота 10.09.2026.xlsx");
    expect(
      buildOverviewCommandersExportFileName(
        null,
        new Date("2026-09-10T12:00:00"),
      ),
    ).toBe("Командири 10.09.2026.xlsx");
  });
});
