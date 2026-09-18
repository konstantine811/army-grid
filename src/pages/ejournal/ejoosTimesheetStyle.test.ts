import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { expandSharedFormulas } from "./ejoosWorkbookSanitize";
import { copyTimesheetRowStylesWithZip } from "./ejoosExcludeTransferZip";
import { EJOOS_GENERAL_NUM_FMT_ID, EJOOS_TEXT_NUM_FMT_ID } from "./ejoosStaffIndexFormat";
import {
  applyInlineStringWritesToWorkbook,
  restyleOosDataRows,
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
    expect(history).toBeTruthy();
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xfList = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
      (match) => match[0],
    );
    const occupiedXf = xfList[Number(occupied)] ?? "";
    const historyXf = xfList[Number(history)] ?? "";
    expect(historyXf.match(/\bfillId="(\d+)"/i)?.[1]).toBe(
      occupiedXf.match(/\bfillId="(\d+)"/i)?.[1],
    );
    expect(historyXf).toMatch(/horizontal="center"/i);
    expect(stylesXml).toMatch(/Times New Roman/i);
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
    const setFormula = (row: number, column: number, formula: string) =>
      (
        sheet.cell(row, column) as unknown as {
          formula: (value: string) => unknown;
        }
      ).formula(formula);
    setFormula(200, 9, "I201+J201");
    setFormula(201, 9, "I201+J201");
    setFormula(201, 10, "I201+J201");
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

  it("ignores wrapText and bold on SHPO writes", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("1. ШПО");
    sheet.cell(7, 6).value("солдат").style({
      bold: true,
      wrapText: true,
      horizontalAlignment: "left",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(
      blob,
      /шпо|штатно.?посад/i,
      [
        {
          row: 8,
          column: 6,
          value: "молодший сержант",
          styleSourceRow: 7,
          styleSourceColumn: 6,
          copyNeighborStyle: true,
          wrapText: true,
        },
      ],
    );
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const styleId = cellStyleId(sheetXml || "", "F8");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xf = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)]
      .map((match) => match[0])[Number(styleId)] ?? "";
    expect(xf).toMatch(/horizontal="center"/i);
    expect(xf).not.toMatch(/wrapText="1"/i);
    const fontId = xf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    expect(fonts[Number(fontId)] ?? "").not.toMatch(/<b\b/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("ignores wrapText on timesheet writes", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(7, 9).value("вибув").style({
      bold: true,
      wrapText: true,
      horizontalAlignment: "left",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(blob, /табель/i, [
      {
        row: 8,
        column: 9,
        value: "вибув на А0409 (скасовано)",
        styleSourceRow: 7,
        styleSourceColumn: 9,
        wrapText: true,
      },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const styleId = cellStyleId(sheetXml || "", "I8");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xf = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)]
      .map((match) => match[0])[Number(styleId)] ?? "";
    expect(xf).not.toMatch(/wrapText="1"/i);
    const fontId = xf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    expect(fonts[Number(fontId)] ?? "").not.toMatch(/<b\b/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("applies Times New Roman center on timesheet writes while preserving fill", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(7, 9).value("вибув на А0409").style({
      fill: "ffff00",
      bold: true,
      horizontalAlignment: "left",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(
      blob,
      /табель/i,
      [
        {
          row: 8,
          column: 9,
          value: "вибув на А0409 (скасовано)",
          styleSourceRow: 7,
          styleSourceColumn: 9,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        },
      ],
    );
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    expect(stylesXml).toMatch(/Times New Roman/i);
    const sourceStyleId = cellStyleId(sheetXml || "", "I7");
    const writtenStyleId = cellStyleId(sheetXml || "", "I8");
    expect(sourceStyleId).toBeTruthy();
    expect(writtenStyleId).toBeTruthy();
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xfList = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
      (match) => match[0],
    );
    const sourceXf = xfList[Number(sourceStyleId)] ?? "";
    const writtenXf = xfList[Number(writtenStyleId)] ?? "";
    expect(writtenXf).toMatch(/horizontal="center"/i);
    expect(writtenXf).toMatch(/vertical="center"/i);
    expect(writtenXf).not.toMatch(/wrapText="1"/i);
    const sourceFillId = sourceXf.match(/\bfillId="(\d+)"/i)?.[1];
    const writtenFillId = writtenXf.match(/\bfillId="(\d+)"/i)?.[1];
    expect(writtenFillId).toBe(sourceFillId);
    const fontId = writtenXf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    const font = fonts[Number(fontId)] ?? "";
    expect(font).toMatch(/Times New Roman/i);
    expect(font).not.toMatch(/<b\b/i);
    expect(font).toMatch(/sz val="12"/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("uses fixed 12pt instead of header Times New Roman 14", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet.cell(1, 1).value("Заголовок").style({
      fontFamily: "Times New Roman",
      fontSize: 14,
      bold: false,
      horizontalAlignment: "center",
      verticalAlignment: "center",
    });
    sheet.cell(7, 9).value("+").style({
      fontFamily: "Times New Roman",
      fontSize: 11,
      bold: false,
      horizontalAlignment: "center",
      verticalAlignment: "center",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(blob, /табель/i, [
      {
        row: 8,
        column: 9,
        value: "-",
        styleSourceRow: 7,
        styleSourceColumn: 9,
        copyNeighborStyle: false,
        keepNeighborStyle: true,
      },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const styleId = cellStyleId(sheetXml || "", "I8");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xfList = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
      (match) => match[0],
    );
    const xf = xfList[Number(styleId)] ?? "";
    const fontId = xf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    const font = fonts[Number(fontId)] ?? "";
    expect(font).toMatch(/Times New Roman/i);
    expect(font).toMatch(/sz val="12"/i);
    expect(font).not.toMatch(/sz val="14"/i);
    expect(font).not.toMatch(/<b\b/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("uses Times New Roman on absent sheet writes", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("5. Тимчасово відсутні");
    sheet.cell(7, 2).value("ІВАНОВ").style({
      fontFamily: "Calibri",
      horizontalAlignment: "left",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(
      blob,
      /відсутн/i,
      [
        {
          row: 8,
          column: 2,
          value: "ПЕТРЕНКО",
          styleSourceRow: 7,
          styleSourceColumn: 2,
          keepNeighborStyle: true,
        },
      ],
    );
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    expect(stylesXml).toMatch(/Times New Roman/i);
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const styleId = cellStyleId(sheetXml || "", "B8");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xf = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)]
      .map((match) => match[0])[Number(styleId)] ?? "";
    const fontId = xf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    expect(fonts[Number(fontId)] ?? "").toMatch(/Times New Roman/i);
    expect(fonts[Number(fontId)] ?? "").toMatch(/sz val="12"/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("never keeps bold on OOS writes even with copyNeighborStyle", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("2. ООС");
    sheet.cell(7, 7).value("НОВІКОВ").style({
      bold: true,
      wrapText: false,
      horizontalAlignment: "left",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(blob, /оос/i, [
      {
        row: 8,
        column: 7,
        value: "ЯМКОВИЙ",
        styleSourceRow: 7,
        copyNeighborStyle: true,
      },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const styleId = cellStyleId(sheetXml || "", "G8");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xf = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)]
      .map((match) => match[0])[Number(styleId)] ?? "";
    const fontId = xf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    expect(fonts[Number(fontId)] ?? "").not.toMatch(/<b\b/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });

  it("applies Times New Roman center non-bold wrap on OOS writes", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("3. ООС");
    sheet.cell(7, 4).value("2103378").style({ bold: true, wrapText: false });
    sheet.cell(7, 7).value("НОВІКОВ").style({
      bold: true,
      horizontalAlignment: "left",
      verticalAlignment: "top",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(blob, /оос/i, [
      {
        row: 8,
        column: 7,
        value: "ЯМКОВИЙ\nРуслан",
        styleSourceRow: 7,
        copyNeighborStyle: false,
        wrapText: true,
      },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    expect(sheetXml).toMatch(/r="G8"/);
    expect(stylesXml).toMatch(/Times New Roman/i);
    const styleId = cellStyleId(sheetXml || "", "G8");
    expect(styleId).toBeTruthy();
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xfList = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
      (match) => match[0],
    );
    expect(Number(styleId)).toBeLessThan(xfList.length);
    const xf = xfList[Number(styleId)] ?? "";
    expect(xf).toMatch(/horizontal="center"/i);
    expect(xf).toMatch(/vertical="center"/i);
    expect(xf).toMatch(/wrapText="1"/i);
    const fontId = xf.match(/\bfontId="(\d+)"/i)?.[1];
    const fonts = [
      ...(stylesXml?.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "").matchAll(
        /<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi,
      ),
    ].map((match) => match[0]);
    const font = fonts[Number(fontId)] ?? "";
    expect(font).toMatch(/Times New Roman/i);
    expect(font).not.toMatch(/<b\b/i);
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

  it("writes timesheet staff index as text, not custom date format", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("6. Табель");
    sheet
      .cell(7, 2)
      .value("2103378")
      .style({ numberFormat: "MMMM yyyy", horizontalAlignment: "center" });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const next = await applyInlineStringWritesToWorkbook(blob, /табель/i, [
      {
        row: 8,
        column: 2,
        value: "2110786",
        styleSourceRow: 7,
        styleSourceColumn: 2,
        keepNeighborStyle: true,
      },
    ]);
    const zip = await JSZip.loadAsync(await next.arrayBuffer());
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    const cell = sheetXml?.match(/<c\b[^>]*\br="B8"[^>]*>[\s\S]*?<\/c>/i)?.[0];
    expect(cell).toBeTruthy();
    expect(cell).toMatch(/\bt="(?:inlineStr|s)"/i);
    expect(cell).not.toMatch(/<v>2110786<\/v>/);
    const styleId = cellStyleId(sheetXml || "", "B8");
    const cellXfsBody =
      stylesXml?.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    const xf = [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)]
      .map((match) => match[0])[Number(styleId)] ?? "";
    expect(xf).toMatch(new RegExp(`numFmtId="${EJOOS_TEXT_NUM_FMT_ID}"`, "i"));
    expect(xf).toMatch(/applyNumberFormat="1"/i);
    await module.default.fromDataAsync(await next.arrayBuffer());
  });
});

describe("OOS РНОКПП number format", () => {
  const xfForCell = (stylesXml: string, styleId: string | undefined) => {
    const cellXfsBody =
      stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
    return (
      [...cellXfsBody.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
        (match) => match[0],
      )[Number(styleId)] ?? ""
    );
  };

  it("writes РНОКПП as General even if the neighbor style is Custom @", async () => {
    const module = await import(
      "xlsx-populate/browser/xlsx-populate-no-encryption"
    );
    const workbook = await module.default.fromBlankAsync();
    const sheet = workbook.sheet(0);
    sheet.name("2. ООС");
    sheet.cell(7, 3).value("21643").style({
      numberFormat: "@",
      horizontalAlignment: "center",
      verticalAlignment: "center",
    });
    sheet.cell(7, 22).value("3142223156").style({
      numberFormat: "@",
      horizontalAlignment: "center",
      verticalAlignment: "center",
    });
    const blob = (await workbook.outputAsync("blob")) as Blob;

    const written = await applyInlineStringWritesToWorkbook(blob, /оос/i, [
      {
        row: 8,
        column: 22,
        value: 1234567890,
        copyNeighborStyle: true,
        styleSourceRow: 7,
        wrapText: true,
      },
    ]);
    const restyled = await restyleOosDataRows(written, [7, 8]);
    const zip = await JSZip.loadAsync(await restyled.arrayBuffer());
    const sheetXml = await zip.file("xl/worksheets/sheet1.xml")?.async("string");
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    expect(sheetXml && stylesXml).toBeTruthy();
    const rnokppXf = xfForCell(stylesXml || "", cellStyleId(sheetXml || "", "V8"));
    expect(rnokppXf).toMatch(
      new RegExp(`numFmtId="${EJOOS_GENERAL_NUM_FMT_ID}"`, "i"),
    );
    expect(rnokppXf).toMatch(/applyNumberFormat="1"/i);
    const restyledV7 = xfForCell(
      stylesXml || "",
      cellStyleId(sheetXml || "", "V7"),
    );
    expect(restyledV7).toMatch(
      new RegExp(`numFmtId="${EJOOS_GENERAL_NUM_FMT_ID}"`, "i"),
    );
    await module.default.fromDataAsync(await restyled.arrayBuffer());
  });
});
