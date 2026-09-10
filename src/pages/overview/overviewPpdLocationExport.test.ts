import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { buildStaffSheetColumnsRecord } from "./overviewStaffSheetColumns";
import {
  buildOverviewPpdLocationExportSheets,
  filterOverviewPpdVyshneveRows,
  filterOverviewPolygonRows,
  isOverviewPpdVyshneveLocation,
  isOverviewPolygonLocation,
  matchesOverviewRotaUnitFilter,
  resolveOverviewPpdLocationExportRows,
  resolveOverviewPpdLocationParts,
} from "./overviewPpdLocationExport";

const row = (
  name: string,
  base = "",
  detail = "",
  unit = "3 піхотна рота",
  extra: Record<string, string> = {},
) =>
  ({
    id: name,
    externalId: name,
    name,
    status: "ON_DUTY",
    statusLabel: "В строю",
    unit,
    staffSheetColumns: {
      staff_14: name,
      staff_15: "ПОЗ",
      staff_21: "В строю",
      staff_23: "",
      staff_31: base,
      staff_35: detail,
      staff_32: "прим.",
      ...extra,
    },
  }) as unknown as BackendPersonnelOverviewRow;

describe("overviewPpdLocationExport", () => {
  it("reads base from col 31 and detail from col 35", () => {
    expect(resolveOverviewPpdLocationParts(row("A", "ППД Вишневе", "ППД"))).toEqual({
      base: "ППД Вишневе",
      detail: "ППД",
      baseColumn: 31,
      detailColumn: 35,
    });
    expect(resolveOverviewPpdLocationParts(row("B", "ППД Вишневе", ""))).toEqual({
      base: "ППД Вишневе",
      detail: "",
      baseColumn: 31,
      detailColumn: 35,
    });
  });

  it("matches sheet 1 and sheet 2 rules", () => {
    expect(isOverviewPpdVyshneveLocation(row("A", "ППД Вишневе", ""))).toBe(true);
    expect(isOverviewPpdVyshneveLocation(row("B", "ППД Вишневе", "ППД"))).toBe(true);
    expect(isOverviewPpdVyshneveLocation(row("C", "ППД Вишневе", "ПОЛІГОН"))).toBe(
      false,
    );
    expect(isOverviewPpdVyshneveLocation(row("D", "Полігон Д", "ПОЛІГОН"))).toBe(
      false,
    );

    expect(isOverviewPolygonLocation(row("E", "ППД Вишневе", "ПОЛІГОН"))).toBe(true);
    expect(isOverviewPolygonLocation(row("F", "ППД Вишневе", "ППД"))).toBe(false);
    expect(isOverviewPolygonLocation(row("G", "Полігон Д", "ПОЛІГОН"))).toBe(false);
  });

  it("matches 3 rota filter label against 3 піхотна рота rows", () => {
    expect(
      matchesOverviewRotaUnitFilter("3 піхотна рота", "3 рота"),
    ).toBe(true);

    const visible = [
      row("ІВАНОВ", "ППД Вишневе", ""),
      row("ПЕТРЕНКО", "ППД Вишневе", "ПОЛІГОН"),
    ];
    const rows = resolveOverviewPpdLocationExportRows({
      rows: visible,
      allRows: [
        ...visible,
        row("ЧУЖИЙ", "ППД Вишневе", "", "2 піхотна рота"),
      ],
      columns: [],
      filters: [{ id: "unit", label: "Підрозділ", values: ["3 рота"] }],
    });

    expect(rows.map((item) => item.name)).toEqual(["ІВАНОВ", "ПЕТРЕНКО"]);
    expect(filterOverviewPpdVyshneveRows(rows).map((item) => item.name)).toEqual([
      "ІВАНОВ",
    ]);
    expect(filterOverviewPolygonRows(rows).map((item) => item.name)).toEqual([
      "ПЕТРЕНКО",
    ]);
  });

  it("does not copy col 31 into empty col 35", () => {
    const staffSheetColumns = buildStaffSheetColumnsRecord({
      column_31: "ППД Вишневе",
      column_14: "ГУК Володимир",
      місце_перебування: "ППД Вишневе",
    } as EjournalPreviewRow);
    const overviewRow = row("ГУК Володимир", "", "");
    overviewRow.staffSheetColumns = {
      ...overviewRow.staffSheetColumns,
      ...staffSheetColumns,
    };
    expect(staffSheetColumns.staff_35).toBe("");
    expect(isOverviewPpdVyshneveLocation(overviewRow)).toBe(true);
    expect(isOverviewPolygonLocation(overviewRow)).toBe(false);
  });

  it("reads polygon from roster col 35 even when staffSheetColumns omit it", () => {
    const overviewRow = row("ГНІЦЕВИЧ Влад", "ППД Вишневе", "");
    const rosterRow = {
      column_31: "ППД Вишневе",
      column_35: "ПОЛІГОН",
      column_14: "ГНІЦЕВИЧ Влад",
    } as EjournalPreviewRow;
    const sheets = buildOverviewPpdLocationExportSheets(
      [overviewRow],
      [rosterRow],
    );

    expect((sheets[0]?.data.length ?? 0) - 1).toBe(0);
    expect((sheets[1]?.data.length ?? 0) - 1).toBe(1);
  });

  it("reads staff_35 clarification from roster columns", () => {
    const staffSheetColumns = buildStaffSheetColumnsRecord({
      column_31: "ППД Вишневе",
      column_35: "ПОЛІГОН",
      column_14: "КОВАЛЕНКО Костян",
    } as EjournalPreviewRow);
    const overviewRow = row("КОВАЛЕНКО Костян", "", "");
    overviewRow.staffSheetColumns = {
      ...overviewRow.staffSheetColumns,
      ...staffSheetColumns,
    };
    expect(staffSheetColumns.staff_35).toBe("ПОЛІГОН");
    expect(isOverviewPolygonLocation(overviewRow)).toBe(true);
  });

  it("builds two sheets with required columns", () => {
    const sheets = buildOverviewPpdLocationExportSheets([
      row("ІВАНОВ Іван", "ППД Вишневе", "ППД"),
      row("КОВАЛЕНКО Костян", "ППД Вишневе", "ПОЛІГОН", "3 піхотна рота", {
        staff_15: "КРОХА",
        staff_23: "БГ",
      }),
    ]);

    expect((sheets[0]?.data.length ?? 0) - 1).toBe(1);
    expect((sheets[1]?.data.length ?? 0) - 1).toBe(1);
    expect((sheets[1]?.data[1] as Array<{ value: string }>).map((cell) => cell.value)).toEqual([
      "КОВАЛЕНКО Костян",
      "КРОХА",
      "В строю",
      "БГ",
      "прим.",
      "ППД Вишневе — ПОЛІГОН",
    ]);
  });
});
