import { describe, expect, it } from "vitest";
import JSZip from "jszip";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  buildTimesheetLayout,
  extractSheetFormulasByCell,
  resolveCanonicalTimesheetSlot,
  resolveHistoryTimesheetRow,
  stampTimesheetHistoryInserts,
  takeHistoryTimesheetRow,
  withTimesheetHistoryInsert,
  type PendingHistoryInsert,
} from "./ejoosTimesheetLayout";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

if (typeof window === "undefined") {
  (globalThis as { window?: typeof globalThis }).window = globalThis;
}

const emptyRow = () => Array.from({ length: 40 }, () => "");

const sheetOf = (rows: string[][]): ExcelSheetSnapshot =>
  ({
    sheetName: "6. Табель",
    rawRows: rows,
    columnIndexes: [],
  }) as ExcelSheetSnapshot;

const fillHistory = (rows: string[][], start: number, end: number) => {
  for (let row = start; row <= end; row += 1) {
    rows[row - 1] = ["", `2199${row}`, "Стрілець", "100915А", "5"];
  }
};

const twoSectionSheet = () => {
  const rows = Array.from({ length: 500 }, emptyRow);
  rows[5] = ["УПРАВЛІННЯ", "УПРАВЛІННЯ"];
  rows[6] = [
    "УПРАВЛІННЯ",
    "2103119",
    "Командир батальйону",
    "210003",
    "24",
    "старший лейтенант",
    "СИДОРЕНКО",
    "1443",
  ];
  fillHistory(rows, 9, 37);
  rows[37] = ["На продовольчому забезпеченні в УПРАВЛІННІ перебуває", "1"];
  rows[351] = ["3 ПІХОТНА РОТА", "3 ПІХОТНА РОТА"];
  rows[352] = [
    "3 ПІХОТНА РОТА",
    "2103461",
    "Старший стрілець",
    "100915А",
    "5",
    "",
    "",
    "",
  ];
  rows[477] = ["3 ПІХОТНА РОТА", "2103999", "Стрілець", "100915А", "5"];
  fillHistory(rows, 479, 497);
  rows[497] = ["На продовольчому забезпеченні в 3 ПІХОТНІЙ РОТІ перебуває", "65"];
  return sheetOf(rows);
};

const twoSectionFormulas = {
  I38: 'COUNTIF(I7:I8,"+")',
  I498: 'COUNTIF(I353:I478,"+")',
};

describe("multi-section timesheet history inserts", () => {
  it("plans separate inserts for УПРАВЛІННЯ and 3 рота footers", () => {
    const sheet = twoSectionSheet();
    const layout = buildTimesheetLayout(sheet, { formulas: twoSectionFormulas });
    const reserved = new Set<number>();
    const insertCountBySection = new Map<string, number>();
    const pending: PendingHistoryInsert[] = [];
    const management = resolveCanonicalTimesheetSlot({
      index: "2103119",
      layout,
    });
    const company = resolveCanonicalTimesheetSlot({
      index: "2103461",
      layout,
    });
    const first = takeHistoryTimesheetRow(
      resolveHistoryTimesheetRow({
        sourceSlot: management,
        layout,
        sheet,
        reserved,
        insertCountBySection,
      }),
      pending,
    );
    const second = takeHistoryTimesheetRow(
      resolveHistoryTimesheetRow({
        sourceSlot: company,
        layout,
        sheet,
        reserved,
        insertCountBySection,
      }),
      pending,
    );
    expect(first.row).toBe(38);
    expect(second.row).toBe(498);
    expect(pending.map((item) => ({ footer: item.footerRow, target: item.targetRow }))).toEqual([
      { footer: 38, target: 38 },
      { footer: 498, target: 498 },
    ]);
    expect(pending[0]?.sectionKey).not.toBe(pending[1]?.sectionKey);
    const writes: ZipCellWrite[] = [
      withTimesheetHistoryInsert({ row: first.row, column: 7, value: "ВІЄРА" }, first),
      withTimesheetHistoryInsert(
        { row: second.row, column: 7, value: "ПОЧЕПЕЦЬКИЙ" },
        second,
      ),
    ];
    stampTimesheetHistoryInserts(writes, pending);
    expect(writes[0]).toMatchObject({
      row: 38,
      insertRowsBefore: true,
      insertRowCount: 1,
      insertedRow: true,
    });
    expect(writes[1]).toMatchObject({
      row: 498,
      insertRowsBefore: true,
      insertRowCount: 1,
      insertedRow: true,
    });
    expect(writes[0]?.insertGroup).not.toBe(writes[1]?.insertGroup);
  });

  it("inserts one row per footer and keeps a lower canonical write on the shifted staff row", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(1, 1).value("6. Табель");
    sheet.cell(6, 1).value("'УПРАВЛІННЯ");
    sheet.cell(6, 2).value("УПРАВЛІННЯ");
    sheet.cell(7, 1).value("'УПРАВЛІННЯ");
    sheet.cell(7, 2).value("2103119");
    sheet.cell(7, 7).value("СИДОРЕНКО").style({ fill: "5b9bd5" });
    for (let row = 9; row <= 37; row += 1) {
      sheet.cell(row, 2).value(`2199${row}`);
    }
    sheet.cell(38, 1).value("На продовольчому забезпеченні в УПРАВЛІННІ перебуває");
    sheet.cell(38, 9).formula('COUNTIF(I7:I8,"+")');
    sheet.cell(352, 1).value("'3 ПІХОТНА РОТА");
    sheet.cell(352, 2).value("3 ПІХОТНА РОТА");
    sheet.cell(353, 1).value("'3 ПІХОТНА РОТА");
    sheet.cell(353, 2).value("2103461");
    sheet.cell(353, 3).value("Старший стрілець");
    sheet.cell(353, 7).style({ fill: "ffff00" });
    sheet.cell(478, 2).value("2103999");
    for (let row = 479; row <= 497; row += 1) {
      sheet.cell(row, 2).value(`2198${row}`);
    }
    sheet
      .cell(498, 1)
      .value("На продовольчому забезпеченні в 3 ПІХОТНІЙ РОТІ перебуває");
    sheet.cell(498, 9).formula('COUNTIF(I353:I478,"+")');
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const snapshot = twoSectionSheet();
    const layout = buildTimesheetLayout(snapshot, { formulas: twoSectionFormulas });
    const reserved = new Set<number>();
    const insertCountBySection = new Map<string, number>();
    const pending: PendingHistoryInsert[] = [];
    const management = resolveCanonicalTimesheetSlot({
      index: "2103119",
      layout,
    });
    const company = resolveCanonicalTimesheetSlot({
      index: "2103461",
      layout,
    });
    const historyA = takeHistoryTimesheetRow(
      resolveHistoryTimesheetRow({
        sourceSlot: management,
        layout,
        sheet: snapshot,
        reserved,
        insertCountBySection,
      }),
      pending,
    );
    const historyB = takeHistoryTimesheetRow(
      resolveHistoryTimesheetRow({
        sourceSlot: company,
        layout,
        sheet: snapshot,
        reserved,
        insertCountBySection,
      }),
      pending,
    );
    const writes: ZipCellWrite[] = [
      withTimesheetHistoryInsert(
        {
          row: historyA.row,
          column: 7,
          value: "ВІЄРА",
          styleSourceRow: 7,
          heightSourceRow: 7,
        },
        historyA,
      ),
      withTimesheetHistoryInsert(
        {
          row: historyB.row,
          column: 7,
          value: "ПОЧЕПЕЦЬКИЙ",
          styleSourceRow: 353,
          heightSourceRow: 353,
        },
        historyB,
      ),
      {
        row: 353,
        column: 2,
        value: "2103461",
      },
      {
        row: 353,
        column: 7,
        value: "ГОГА",
        styleSourceRow: 353,
        heightSourceRow: 353,
      },
    ];
    stampTimesheetHistoryInserts(writes, pending);
    expect(writes.filter((write) => write.insertRowsBefore).map((write) => ({
      row: write.row,
      count: write.insertRowCount,
    }))).toEqual([
      { row: 38, count: 1 },
      { row: 498, count: 1 },
    ]);

    const next = await applyInlineStringWritesToWorkbook(blob, /табель/i, writes);
    const written = await module.default.fromDataAsync(await next.arrayBuffer());
    const timesheet = written.sheet("6. Табель");
    expect(String(timesheet.cell(38, 7).value() ?? "")).toMatch(/вієра/i);
    expect(String(timesheet.cell(39, 1).value() ?? "")).toMatch(
      /на продовольчому забезпеченні в управлінні/i,
    );
    expect(String(timesheet.cell(353, 2).value() ?? "")).not.toBe("2103461");
    expect(String(timesheet.cell(354, 2).value() ?? "")).toBe("2103461");
    expect(String(timesheet.cell(354, 7).value() ?? "")).toMatch(/гога/i);
    expect(String(timesheet.cell(499, 7).value() ?? "")).toMatch(/почепецьк/i);
    expect(String(timesheet.cell(500, 1).value() ?? "")).toMatch(
      /на продовольчому забезпеченні в 3 піхотній роті/i,
    );
    expect(String(timesheet.cell(353, 1).value() ?? "")).toMatch(/3 піхотна рота/i);

    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const sheetXml = (await zip.file("xl/worksheets/sheet1.xml")?.async("string")) || "";
    const formulas = extractSheetFormulasByCell(sheetXml);
    const companyCountIf = [...formulas.values()].find((formula) =>
      /COUNTIF/i.test(formula) && /478|479/.test(formula),
    );
    expect(companyCountIf).toBeTruthy();
    expect(companyCountIf).not.toMatch(/I353:I478/);
    expect(companyCountIf).toMatch(/COUNTIF\(\s*I354:I479/i);
  }, 30_000);

  it("rebases original writes after two history inserts in the same section", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(1, 1).value("6. Табель");
    sheet.cell(352, 1).value("'3 ПІХОТНА РОТА");
    sheet.cell(352, 2).value("3 ПІХОТНА РОТА");
    sheet.cell(353, 2).value("2103461");
    sheet.cell(353, 7).value("ДЖЕРЕЛО").style({ fill: "ffff00" });
    for (let row = 479; row <= 497; row += 1) {
      sheet.cell(row, 2).value(`2198${row}`);
    }
    sheet
      .cell(498, 1)
      .value("На продовольчому забезпеченні в 3 ПІХОТНІЙ РОТІ перебуває");
    sheet.cell(498, 9).formula('COUNTIF(I353:I478,"+")');
    sheet.cell(499, 1).value("'РБпАК");
    sheet.cell(499, 2).value("РБпАК");
    sheet.cell(499, 7).value("СТАРИЙ 499").style({ fill: "5b9bd5" });
    sheet.cell(500, 2).value("2105001");
    sheet.cell(500, 7).value("СТАРИЙ 500");
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const snapshot = twoSectionSheet();
    const layout = buildTimesheetLayout(snapshot, { formulas: twoSectionFormulas });
    const reserved = new Set<number>();
    const insertCountBySection = new Map<string, number>();
    const pending: PendingHistoryInsert[] = [];
    const company = resolveCanonicalTimesheetSlot({
      index: "2103461",
      layout,
    });
    const historyA = takeHistoryTimesheetRow(
      resolveHistoryTimesheetRow({
        sourceSlot: company,
        layout,
        sheet: snapshot,
        reserved,
        insertCountBySection,
      }),
      pending,
    );
    const historyB = takeHistoryTimesheetRow(
      resolveHistoryTimesheetRow({
        sourceSlot: company,
        layout,
        sheet: snapshot,
        reserved,
        insertCountBySection,
      }),
      pending,
    );
    expect(historyA.row).toBe(498);
    expect(historyB.row).toBe(499);
    const writes: ZipCellWrite[] = [
      withTimesheetHistoryInsert(
        {
          row: historyA.row,
          column: 7,
          value: "ІСТОРІЯ A",
          styleSourceRow: 353,
          heightSourceRow: 353,
        },
        historyA,
      ),
      withTimesheetHistoryInsert(
        {
          row: historyB.row,
          column: 7,
          value: "ІСТОРІЯ B",
          styleSourceRow: 353,
          heightSourceRow: 353,
        },
        historyB,
      ),
      {
        row: 499,
        column: 7,
        value: "ПІСЛЯ ЗСУВУ 499",
        styleSourceRow: 499,
        heightSourceRow: 499,
      },
      {
        row: 500,
        column: 7,
        value: "ПІСЛЯ ЗСУВУ 500",
        styleSourceRow: 500,
        heightSourceRow: 500,
      },
    ];
    stampTimesheetHistoryInserts(writes, pending);
    expect(writes.filter((write) => write.insertedRow).map((write) => write.row)).toEqual([
      498, 499,
    ]);
    expect(writes.find((write) => write.insertRowsBefore)).toMatchObject({
      row: 498,
      insertRowCount: 2,
    });
    const nextRowWrite = writes.find(
      (write) => write.value === "ПІСЛЯ ЗСУВУ 499",
    );
    expect(nextRowWrite?.insertedRow).toBeFalsy();

    const next = await applyInlineStringWritesToWorkbook(blob, /табель/i, writes);
    expect(writes.find((write) => write.value === "ІСТОРІЯ A")?.row).toBe(498);
    expect(writes.find((write) => write.value === "ІСТОРІЯ B")?.row).toBe(499);
    expect(nextRowWrite?.row).toBe(501);
    expect(nextRowWrite?.styleSourceRow).toBe(501);
    expect(nextRowWrite?.heightSourceRow).toBe(501);
    const shifted500 = writes.find((write) => write.value === "ПІСЛЯ ЗСУВУ 500");
    expect(shifted500?.row).toBe(502);
    expect(shifted500?.styleSourceRow).toBe(502);
    expect(shifted500?.heightSourceRow).toBe(502);

    const written = await module.default.fromDataAsync(await next.arrayBuffer());
    const timesheet = written.sheet("6. Табель");
    expect(String(timesheet.cell(498, 7).value() ?? "")).toBe("ІСТОРІЯ A");
    expect(String(timesheet.cell(499, 7).value() ?? "")).toBe("ІСТОРІЯ B");
    expect(String(timesheet.cell(500, 1).value() ?? "")).toMatch(
      /на продовольчому забезпеченні в 3 піхотній роті/i,
    );
    expect(String(timesheet.cell(501, 1).value() ?? "")).toMatch(/рбпак/i);
    expect(String(timesheet.cell(501, 7).value() ?? "")).toBe("ПІСЛЯ ЗСУВУ 499");
    expect(String(timesheet.cell(502, 7).value() ?? "")).toBe("ПІСЛЯ ЗСУВУ 500");
    expect(String(timesheet.cell(499, 7).value() ?? "")).not.toBe("ПІСЛЯ ЗСУВУ 499");
  }, 30_000);
});
