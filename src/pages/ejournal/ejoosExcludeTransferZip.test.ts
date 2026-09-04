import { describe, expect, it } from "vitest";
import { nextTimesheetHistoryRow } from "./ejoosExcludeTransferZip";
import { buildSheetGridFromXml } from "./ejoosTimesheetPersonRows";

describe("nextTimesheetHistoryRow", () => {
  it("places history immediately after the last occupied XML row", () => {
    const xml = [
      "<worksheet><sheetData>",
      '<row r="214"><c r="B214"><v>2103230</v></c>',
      '<c r="G214" t="inlineStr"><is><t>ЛЕКАЙ Володимир Петрович</t></is></c></row>',
      '<row r="215"><c r="B215" s="339"/><c r="G215" s="413"/><c r="H215" s="322"/></row>',
      '<row r="216"><c r="B216"><v>10509</v></c></row>',
      "</sheetData></worksheet>",
    ].join("");

    const row = nextTimesheetHistoryRow(
      buildSheetGridFromXml(xml),
      214,
      new Set(),
    );

    expect(row).toBe(215);
  });
});
