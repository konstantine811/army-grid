import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { expandSharedFormulas } from "./ejoosWorkbookSanitize";
import { copyTimesheetRowStylesWithZip } from "./ejoosExcludeTransferZip";
import {
  applyInlineStringWritesToWorkbook,
  shiftSheetRowsDown,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

if (typeof window === "undefined") {
  (globalThis as { window?: typeof globalThis }).window = globalThis;
}

const cellStyleId = (sheetXml: string, ref: string) => {
  const match = sheetXml.match(
    new RegExp(`<c\\b([^>]*\\br="${ref}"(?![0-9A-Za-z])[^>]*)`, "i"),
  );
  return match?.[1].match(/\bs="(\d+)"/i)?.[1];
};

describe("timesheet occupied styles", () => {
  it("copies s= from an occupied row onto a vacant history row", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(1, 1).value("6. Табель");
    sheet.cell(7, 2).value("2103378");
    sheet
      .cell(7, 7)
      .value("НОВІКОВ Олександр Сергійович")
      .style({
        fill: "5b9bd5",
        border: true,
        verticalAlignment: "center",
        horizontalAlignment: "center",
      });
    sheet.cell(7, 9).value("+").style({
      border: true,
      verticalAlignment: "center",
      horizontalAlignment: "center",
    });
    sheet.cell(8, 2).value("2103435");
    sheet.cell(8, 7).value("ЯМКОВИЙ Руслан Костянтинович");
    sheet.cell(8, 9).value("вибув на А0409 (скасовано)");
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await copyTimesheetRowStylesWithZip(blob, [
      { sourceRow: 7, targetRow: 8 },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    expect(sheetXml).toBeTruthy();
    expect(sheetXml?.match(/<dimension\b/gi)?.length ?? 0).toBeLessThan(2);
    const occupied = cellStyleId(sheetXml || "", "G7");
    const history = cellStyleId(sheetXml || "", "G8");
    expect(occupied).toBeTruthy();
    expect(history).toBe(occupied);
  });

  it("shiftSheetRowsDown updates shared formula refs", () => {
    const sheetXml = [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      "<sheetData>",
      '<row r="200"><c r="K200"><f>SUM(I201:J201)</f></c></row>',
      '<row r="201">',
      '<c r="I201"><f t="shared" ref="I201:J201" si="0">A201+B201</f></c>',
      '<c r="J201"><f t="shared" si="0"/></c>',
      "</row>",
      "</sheetData>",
      "</worksheet>",
    ].join("");
    const shifted = shiftSheetRowsDown(sheetXml, 201, 1);
    expect(shifted).toMatch(/<f[^>]*ref="I202:J202"/);
    expect(shifted).toMatch(/SUM\(I202:J202\)/);
    expect(shifted).toMatch(/r="I202"/);
    expect(shifted).toMatch(/r="J202"/);
    expect(shifted).not.toMatch(/ref="I201:J201"/);
  });

  it("shiftSheetRowsDown keeps valid row open tags", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(200, 1).value("'2 ПІХОТНА РОТА");
    sheet.cell(201, 1).value("'3 ПІХОТНА РОТА");
    const blob = (await workbook.outputAsync("blob")) as Blob;
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    const before = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const row201 = before?.match(/<row\b[^>]*\br="201"[^>]*>[\s\S]*?<\/row>/i)?.[0];
    expect(row201).toMatch(/^<row r="201"/);
    const shifted = shiftSheetRowsDown(before || "", 201, 1);
    expect(shifted).toMatch(/<row r="202"/);
    expect(shifted).not.toMatch(/<\/row>\s+r="/);
  });

  it("expandSharedFormulas materializes slave cells", () => {
    const sheetXml = [
      "<sheetData>",
      '<row r="201">',
      '<c r="I201"><f t="shared" ref="I201:J201" si="0">I201+J201</f></c>',
      '<c r="J201"><f t="shared" si="0"/></c>',
      "</row>",
      "</sheetData>",
    ].join("");
    const expanded = expandSharedFormulas(sheetXml);
    expect(expanded).not.toMatch(/\bt="shared"/i);
    expect(expanded).toMatch(/<f>I201\+J201<\/f>/);
    expect(expanded).toMatch(/<f>J201\+K201<\/f>/);
  });

  it("insertRowsBefore drops shared formulas that Excel would repair away", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(200, 9).formula("I201+J201");
    sheet.cell(201, 9).formula("I201+J201");
    sheet.cell(201, 10).formula("I201+J201");
    sheet.cell(201, 1).value("'3 ПІХОТНА РОТА");
    sheet.cell(201, 2).value("2103700");
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(blob, /табель/i, [
      { row: 201, column: 2, value: "2103700", insertRowsBefore: true },
      { row: 201, column: 7, value: "ТКАЧУК" },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    expect(sheetXml).toBeTruthy();
    expect(sheetXml).not.toMatch(/\bt="shared"/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("survives insertRowsBefore and re-read via xlsx-populate", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(1, 1).value("6. Табель");
    sheet.cell(199, 1).value("'1 ПІХОТНА РОТА");
    sheet.cell(199, 2).value("На продовольчому в 1 ПІХОТНІЙ РОТІ");
    sheet.cell(200, 1).value("'2 ПІХОТНА РОТА");
    sheet.cell(200, 2).value("2103700");
    sheet.cell(200, 7).value("ТКАЧУК");
    sheet.cell(201, 1).value("'3 ПІХОТНА РОТА");
    sheet.cell(201, 2).value("3 ПІХОТНА РОТА");
    const blob = (await workbook.outputAsync("blob")) as Blob;
    const beforeZip = await JSZip.loadAsync(await blob.arrayBuffer());
    const beforeXml = await beforeZip.file("xl/worksheets/sheet1.xml")?.async("string");
    expect(beforeXml).toMatch(/<row r="201"/);

    const writes: ZipCellWrite[] = [
      { row: 201, column: 2, value: "2103700", insertRowsBefore: true },
      { row: 201, column: 7, value: "ТКАЧУК (вибув)" },
    ];
    const next = await applyInlineStringWritesToWorkbook(
      blob,
      /табель/i,
      writes,
    );
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    expect(sheetXml).toBeTruthy();
    expect(sheetXml).toContain("</worksheet>");
    if (sheetXml && sheetXml.length > 836) {
      expect(sheetXml.slice(Math.max(0, 836 - 80), 836 + 80)).not.toMatch(
        /\/>\s*>\s*<\//,
      );
    }
    try {
      await module.default.fromDataAsync(await next.arrayBuffer());
    } catch (loadErr) {
      const snippet = sheetXml?.slice(736, 936) ?? "";
      throw new Error(
        `${loadErr instanceof Error ? loadErr.message : loadErr}\nworksheet@${736}: ${snippet}`,
      );
    }
    expect(sheetXml).toMatch(/r="G201"/);
    expect(sheetXml).toMatch(/r="B202"/);
  });
});
