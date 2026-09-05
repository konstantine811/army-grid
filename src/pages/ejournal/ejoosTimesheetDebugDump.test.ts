import { describe, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  buildTimesheetDebugDump,
  formatTimesheetDebugDump,
} from "./ejoosTimesheetDebugDump";

const sheetOf = (rows: string[][]): ExcelSheetSnapshot => ({
  sheetIndex: 0,
  sheetName: "6. Табель",
  rawRows: rows,
  headerRows: [],
  rows: [],
  columnCount: Math.max(0, ...rows.map((row) => row.length)),
  columnIndexes: [],
  dataStartRow: 1,
});

describe("ejoosTimesheetDebugDump", () => {
  it("prints parsed Табель structure (demo)", () => {
    const rows = Array.from({ length: 60 }, () => Array(40).fill(""));
    rows[5] = ["", "", "", "", "", "", "", "", "Серпень 2026 р."];
    rows[33] = ["'УПРАВЛІННЯ"];
    rows[34] = ["", "2103279", "", "", "", "солдат", "КРИВЦОВ", "3136"];
    rows[35] = [
      "",
      "2103198",
      "",
      "",
      "",
      "солдат",
      "ВІЄРА ДА КОСТА Аліссон Енріке",
      "24553",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "+",
      "вибув до 4 штурмового батальйону",
      "-",
      "-",
    ];
    rows[46] = ["'1 ПІХОТНА РОТА"];
    rows[47] = ["", "2103144", "", "", "", "лейтенант", "ІВАНОВ", "10001"];
    rows[48] = ["'2 ПІХОТНИЙ ВЗВОД"];
    rows[49] = ["", "2103101", "", "", "", "солдат", "ВЗВОДНИК", "10002"];
    rows[50] = ["'1 ПІХОТНЕ ВІДДІЛЕННЯ"];
    rows[51] = ["", "2103199", "", "", "", "солдат", "КУЛЕМЕТНИК", "10003"];
    rows[52] = ["'3 ПІХОТНИЙ ВЗВОД"];
    rows[53] = ["", "2103102", "", "", "", "солдат", "ІНШИЙ ВЗВОД", "10004"];
    rows[54] = ["", "", "", "", "", "", "", ""];
    rows[55] = ["На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває 14"];
    rows[56] = ["'2 ПІХОТНА РОТА"];

    const sheet = sheetOf(rows);
    const positionTitle =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";

    const dump = buildTimesheetDebugDump(sheet, {
      positionTitle,
      sourceRow: 35,
      maxRow: 58,
    });

    console.log("\n" + formatTimesheetDebugDump(dump) + "\n");
  });
});
