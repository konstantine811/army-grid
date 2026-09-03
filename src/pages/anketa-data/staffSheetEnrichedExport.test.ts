import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { rosterRowsFromStaffSheetPayload } from "./staffSheetImport";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { buildStaffSheetRosterImportPayload } from "../excel-fill/staffSheet";
import {
  resetStaffSheetExportTemplateCache,
  writeStaffSheetExportWorkbook,
} from "./staffSheetExportWorkbook";

if (typeof window === "undefined") {
  (globalThis as { window?: typeof globalThis }).window = globalThis;
}

if (typeof FileReader === "undefined") {
  class NodeFileReader {
    result: ArrayBuffer | null = null;
    onload: ((event: { target: NodeFileReader }) => void) | null = null;
    onerror: ((event: { target: NodeFileReader }) => void) | null = null;
    onloadend: ((event: { target: NodeFileReader }) => void) | null = null;
    readyState = 0;
    readAsArrayBuffer(blob: Blob) {
      void blob.arrayBuffer().then((buffer) => {
        this.result = buffer;
        this.readyState = 2;
        const event = { target: this };
        this.onload?.(event);
        this.onloadend?.(event);
      });
    }
    readAsBinaryString() {}
    readAsDataURL() {}
    readAsText() {}
    abort() {}
    addEventListener() {}
    removeEventListener() {}
    dispatchEvent() {
      return false;
    }
  }
  globalThis.FileReader = NodeFileReader as unknown as typeof FileReader;
}

const templatePath = path.join(
  process.cwd(),
  "public/templates/staffSheetExportTemplate.xlsx",
);

const gvizFirstDataRow = [
  "нова",
  "ж",
  "ж",
  "ж",
  "Командир батальйону",
  "0210003",
  "командир 1 піхотного батальйону",
  "підполковник",
  "Оф.",
  "",
  "",
  "",
  "старший лейтенант",
  "КІЯНЕНКО Андрій Олександрович",
  "МАРІК",
];

const gvizHeader = [
  "№",
  "Підрозділ",
  "Взвод",
  "Відділення",
  "Посада",
  "ВОС",
  "Повна посада",
  "ШПК факт",
  "ШПК факт",
  "Анкета",
  "Військовий квиток",
  "Мобілізація/контракт",
  "Звання ",
  "ПІБ",
  "Позивний",
];

describe("staffSheetEnrichedExport roster values", () => {
  it("does not prefix column labels to gviz row values", () => {
    const table = [gvizHeader, gvizFirstDataRow];
    const payload = buildStaffSheetRosterImportPayload(table, {
      source: "gviz",
      sourceLabel: "test",
      includeAllRows: true,
    });
    const rows = rosterRowsFromStaffSheetPayload(payload);
    const row2 = rows.find((row) => row.__rowNumber === 2);
    expect(row2).toBeTruthy();
    expect(readRosterColumnValue(row2!, 14)).toBe(
      "КІЯНЕНКО Андрій Олександрович",
    );
    expect(readRosterColumnValue(row2!, 5)).toBe("Командир батальйону");
    expect(readRosterColumnValue(row2!, 1)).toBe("нова");
    expect(readRosterColumnValue(row2!, 13)).toBe("старший лейтенант");
  });
});

describe("staffSheetExportWorkbook template styles", () => {
  it("keeps yellow header and bordered data cells from the staff sheet template", async () => {
    resetStaffSheetExportTemplateCache();
    const templateData = readFileSync(templatePath).buffer;
    const table = [gvizHeader, gvizFirstDataRow];
    const payload = buildStaffSheetRosterImportPayload(table, {
      source: "gviz",
      sourceLabel: "test",
      includeAllRows: true,
    });
    const rows = rosterRowsFromStaffSheetPayload(payload);

    const buffer = await writeStaffSheetExportWorkbook(rows, [], {
      templateData,
      download: false,
    });

    const XlsxPopulate = (
      await import("xlsx-populate/browser/xlsx-populate-no-encryption")
    ).default;
    const workbook = await XlsxPopulate.fromDataAsync(buffer);
    const sheet =
      workbook
        .sheets()
        .find((item: { name: () => string }) =>
          /загальний\s*список/i.test(item.name()),
        ) ?? workbook.sheet(0);

    const fillRgb = (cell: { style: (name: string) => unknown }) => {
      const fill = cell.style("fill") as { color?: { rgb?: string } } | string;
      if (typeof fill === "string") return fill.toLowerCase();
      return fill.color?.rgb?.replace(/^ff/i, "").toLowerCase() ?? "";
    };

    expect(fillRgb(sheet.cell(1, 1))).toBe("ffff00");
    expect(fillRgb(sheet.cell(1, 10))).toBe("ffc000");
    expect(fillRgb(sheet.cell(1, 22))).toBe("92d050");
    expect(sheet.cell(1, 1).style("bold")).toBe(true);
    expect(sheet.cell(2, 14).style("leftBorderStyle")).toBe("thin");
    expect(sheet.cell(2, 14).value()).toBe("КІЯНЕНКО Андрій Олександрович");
    expect(sheet.row(1).height()).toBe(46.5);
  });
});
