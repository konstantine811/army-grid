import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { copyTimesheetRowStylesWithZip } from "./ejoosExcludeTransferZip";

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
});
