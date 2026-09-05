import { readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";
import { rosterRowsFromStaffSheetPayload } from "./staffSheetImport";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  buildStaffSheetRosterImportPayload,
  parseStaffSheetGvizResponse,
} from "../excel-fill/staffSheet";
import {
  resetStaffSheetExportTemplateCache,
  writeStaffSheetAnketaVkOverlay,
} from "./staffSheetExportWorkbook";
import type { StaffSheetEnrichmentEntry } from "./staffSheetEnrichment";

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

describe("staffSheet gviz stacked labels", () => {
  it("reconstructs Kiyanenko, vacant row, and Babchenko from gviz labels", () => {
    const gvizSnippet = {
      table: {
        cols: [
          { label: "№ нова нова" },
          { label: "Підрозділ ж ж" },
          { label: "Взвод ж ж" },
          { label: "Відділення ж ж" },
          {
            label:
              "Посада Командир батальйону Заступник командира батальйону",
          },
          { label: "ВОС 0210003 0210003" },
          {
            label:
              "Повна посада командир 1 піхотного батальйону заступник командира 1 піхотного батальйону",
          },
          { label: "ШПК факт підполковник майор" },
          { label: "ШПК факт Оф. Оф." },
          { label: "Анкета " },
          { label: "Військовий квиток " },
          { label: "Мобілізація/контракт " },
          { label: "Звання  старший лейтенант " },
          { label: "ПІБ КІЯНЕНКО Андрій Олександрович " },
          { label: "Позивний МАРІК " },
        ],
        rows: [
          {
            c: [
              { v: "нова" },
              { v: "ж" },
              { v: "ж" },
              { v: "ж" },
              {
                v: "Заступник командира батальйону з психологічної підтримки персоналу",
              },
              { v: "3420003" },
              {
                v: "заступник командира 1 піхотного батальйону з психологічної підтримки персоналу",
              },
              { v: "майор" },
              { v: "Оф." },
              { v: "так" },
              null,
              null,
              { v: "майор" },
              { v: "БАБЧЕНКО Олег Володимирович" },
              { v: "ШЕФ" },
            ],
          },
        ],
      },
    };
    const table = parseStaffSheetGvizResponse(JSON.stringify(gvizSnippet));
    const payload = buildStaffSheetRosterImportPayload(table, {
      source: "gviz",
      sourceLabel: "test",
      includeAllRows: true,
    });
    const rows = rosterRowsFromStaffSheetPayload(payload);
    const row2 = rows.find((row) => row.__rowNumber === 2);
    const row3 = rows.find((row) => row.__rowNumber === 3);
    const row4 = rows.find((row) => row.__rowNumber === 4);
    expect(row2).toBeTruthy();
    expect(row3).toBeTruthy();
    expect(row4).toBeTruthy();
    expect(readRosterColumnValue(row2!, 14)).toBe(
      "КІЯНЕНКО Андрій Олександрович",
    );
    expect(readRosterColumnValue(row2!, 11)).toBe("");
    expect(readRosterColumnValue(row4!, 14)).toBe(
      "БАБЧЕНКО Олег Володимирович",
    );
  });
});

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
    expect(readRosterColumnValue(row2!, 11)).toBe("");
  });
});

describe("staffSheetAnketaVkOverlay", () => {
  const prepareBaseWithPib = async (
    pib: string,
    excelRow: number,
    patch?: (sheet: {
      cell: (row: number, col: number) => { value: (v?: unknown) => unknown };
    }) => void,
  ) => {
    const baseData = readFileSync(templatePath).buffer;
    const XlsxPopulate = (
      await import("xlsx-populate/browser/xlsx-populate-no-encryption")
    ).default;
    const baseWorkbook = await XlsxPopulate.fromDataAsync(baseData);
    const baseSheet = baseWorkbook.sheet(0);
    baseSheet.cell(excelRow, 14).value(pib);
    patch?.(baseSheet);
    const output = await baseWorkbook.outputAsync();
    return output.arrayBuffer();
  };

  it("updates only columns 10 and 11 in the base workbook", async () => {
    resetStaffSheetExportTemplateCache();
    const baseData = await prepareBaseWithPib("Іванов Іван Іванович", 3);
    const entries: StaffSheetEnrichmentEntry[] = [
      {
        excelRowNumber: 3,
        pib: "Іванов Іван Іванович",
        values: ["так", "МО 312448", "", "", "", ""],
      },
    ];

    const buffer = await writeStaffSheetAnketaVkOverlay(baseData, entries, {
      download: false,
    });

    const XlsxPopulate = (
      await import("xlsx-populate/browser/xlsx-populate-no-encryption")
    ).default;
    const workbook = await XlsxPopulate.fromDataAsync(buffer);
    const sheet = workbook.sheet(0);

    expect(sheet.cell(1, 10).value()).toBe("Анкета");
    expect(sheet.cell(3, 10).value()).toBe("так");
    expect(sheet.cell(3, 11).value()).toBe("МО 312448");
    expect(sheet.cell(2, 10).value() ?? "").toBe("");
    expect(sheet.cell(1, 1).value()).toBe("№");
  }, 15_000);

  it("matches by PIB when excelRowNumber points to a vacant row", async () => {
    resetStaffSheetExportTemplateCache();
    const pib = "ПЕТРЕНКО Павло Володимирович";
    const baseData = await prepareBaseWithPib(pib, 4);
    const entries: StaffSheetEnrichmentEntry[] = [
      {
        excelRowNumber: 3,
        pib,
        values: ["так", "АВ 985186", "", "", "", ""],
      },
    ];

    const buffer = await writeStaffSheetAnketaVkOverlay(baseData, entries, {
      download: false,
    });

    const XlsxPopulate = (
      await import("xlsx-populate/browser/xlsx-populate-no-encryption")
    ).default;
    const workbook = await XlsxPopulate.fromDataAsync(buffer);
    const sheet = workbook.sheet(0);

    expect(sheet.cell(3, 10).value() ?? "").toBe("");
    expect(sheet.cell(3, 11).value() ?? "").toBe("");
    expect(sheet.cell(4, 10).value()).toBe("так");
    expect(sheet.cell(4, 11).value()).toBe("АВ 985186");
  });

  it("clears duplicate column header text in column 11", async () => {
    resetStaffSheetExportTemplateCache();
    const baseData = await prepareBaseWithPib(
      "Іванов Іван Іванович",
      2,
      (sheet) => {
        sheet.cell(2, 11).value("Військовий квиток");
      },
    );

    const buffer = await writeStaffSheetAnketaVkOverlay(
      baseData,
      [
        {
          excelRowNumber: 2,
          pib: "Іванов Іван Іванович",
          values: ["", "", "", "", "", ""],
        },
      ],
      { download: false },
    );

    const XlsxPopulate = (
      await import("xlsx-populate/browser/xlsx-populate-no-encryption")
    ).default;
    const workbook = await XlsxPopulate.fromDataAsync(buffer);
    const sheet = workbook.sheet(0);
    expect(sheet.cell(2, 11).value() ?? "").toBe("");
  });
});
