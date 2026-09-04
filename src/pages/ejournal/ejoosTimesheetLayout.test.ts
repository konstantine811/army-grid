import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  TimesheetLayoutError,
  assertCanonicalSlotFreeFor,
  assertHistoryRowsOutsideCanonical,
  buildTimesheetLayout,
  parseCountIfStaffRange,
  resolveCanonicalTimesheetSlot,
  resolveHistoryTimesheetRow,
} from "./ejoosTimesheetLayout";

const sheetOf = (rows: string[][]): ExcelSheetSnapshot =>
  ({
    sheetName: "6. Табель",
    rawRows: rows,
    columnIndexes: [],
  }) as ExcelSheetSnapshot;

const emptyRow = () => Array.from({ length: 40 }, () => "");

describe("ejoosTimesheetLayout", () => {
  const companySheet = () => {
    const rows = Array.from({ length: 80 }, emptyRow);
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
    rows[7] = [
      "УПРАВЛІННЯ",
      "2103120",
      "Заступник командира батальйону",
      "210003",
      "23",
      "",
      "",
      "",
    ];
    rows[8] = ["На продовольчому забезпеченні в УПРАВЛІННІ перебуває", "9"];
    rows[45] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[46] = [
      "1 ПІХОТНА РОТА",
      "2103144",
      "Командир роти",
      "210003",
      "19",
      "лейтенант",
      "МАРЧЕНКО",
      "1",
    ];
    rows[55] = ["1 ПІХОТНА РОТА", "1 ПІХОТНИЙ ВЗВОД"];
    rows[69] = ["1 ПІХОТНА РОТА", "2 піхотне відділення"];
    rows[72] = [
      "1 ПІХОТНА РОТА",
      "2103167",
      "Гранатометник",
      "103061А",
      "5",
      "солдат",
      "ГОГА Владислав Петрович",
      "13651",
    ];
    rows[73] = [
      "1 ПІХОТНА РОТА",
      "2103168",
      "Кулеметник",
      "101627А",
      "5",
      "",
      "",
      "",
    ];
    rows[77] = ["", "", "", "", "", "", "", ""];
    rows[78] = [
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
      "12",
    ];
    rows[79] = ["2 ПІХОТНА РОТА", "2 ПІХОТНА РОТА"];
    return sheetOf(rows);
  };

  it("parses staff bounds from footer COUNTIF", () => {
    expect(parseCountIfStaffRange('COUNTIF(I353:I478,"+")')).toEqual({
      start: 353,
      end: 478,
    });
    expect(parseCountIfStaffRange('COUNTIF($I$353:$I$478,"+")')).toEqual({
      start: 353,
      end: 478,
    });
  });

  it("anchors GOGA to canonical index 2103167 inside 1 рота / 1 взвод / 2 відділення", () => {
    const layout = buildTimesheetLayout(companySheet());
    const slot = resolveCanonicalTimesheetSlot({
      index: "2103167",
      layout,
    });
    expect(slot.row).toBe(73);
    expect(slot.section).toMatch(/1 ПІХОТНА РОТА/i);
    expect(slot.platoon).toMatch(/1 ПІХОТНИЙ ВЗВОД/i);
    expect(slot.squad).toMatch(/2 піхотне відділення/i);
    expect(slot.position).toMatch(/гранатометник/i);
    expect(slot.vos).toBe("103061А");
    expect(slot.tariff).toBe("5");
    expect(slot.occupantName).toMatch(/гога/i);
  });

  it("keeps vacant staff slots as canonical even without a person", () => {
    const layout = buildTimesheetLayout(companySheet());
    const slot = resolveCanonicalTimesheetSlot({
      index: "2103168",
      layout,
    });
    expect(slot.row).toBe(74);
    expect(slot.occupied).toBe(false);
    expect(slot.section).toMatch(/1 ПІХОТНА РОТА/i);
  });

  it("finds history only between last staff row and food-summary of the same section", () => {
    const sheet = companySheet();
    const layout = buildTimesheetLayout(sheet);
    const source = resolveCanonicalTimesheetSlot({
      index: "2103167",
      layout,
    });
    const history = resolveHistoryTimesheetRow({
      sourceSlot: source,
      layout,
      sheet,
    });
    expect(history.row).toBeGreaterThan(source.row);
    expect(history.row).toBeLessThan(79);
    expect(history.row).toBe(75);
    expect(history.insertBefore).toBe(false);
  });

  it("inserts a history row before the footer instead of falling into the next company", () => {
    const sheet = companySheet();
    const layout = buildTimesheetLayout(sheet);
    const source = resolveCanonicalTimesheetSlot({
      index: "2103119",
      layout,
    });
    expect(source.section).toMatch(/управлінн/i);
    const history = resolveHistoryTimesheetRow({
      sourceSlot: source,
      layout,
      sheet,
    });
    expect(history.row).toBe(9);
    expect(history.insertBefore).toBe(true);
    expect(history.row).toBeLessThan(46);
  });

  it("blocks overwrite when another person occupies the canonical slot", () => {
    const layout = buildTimesheetLayout(companySheet());
    const slot = resolveCanonicalTimesheetSlot({
      index: "2103167",
      layout,
    });
    expect(() =>
      assertCanonicalSlotFreeFor(slot, {
        personId: "99999",
        fullName: "АТРАХОВ",
      }),
    ).toThrow(/зайнятий/i);
    expect(() =>
      assertCanonicalSlotFreeFor(slot, {
        personId: "13651",
        fullName: "ГОГА Владислав Петрович",
      }),
    ).not.toThrow();
  });

  it("treats later history copy of the same index as non-canonical", () => {
    const rows = Array.from({ length: 20 }, emptyRow);
    rows[6] = ["УПРАВЛІННЯ", "УПРАВЛІННЯ"];
    rows[7] = [
      "УПРАВЛІННЯ",
      "2103764",
      "Стрілець",
      "100915А",
      "5",
      "солдат",
      "АТРАХОВ",
      "111",
    ];
    rows[8] = [
      "УПРАВЛІННЯ",
      "2103764",
      "Стрілець",
      "100915А",
      "5",
      "солдат",
      "ХУБАЄВ",
      "222",
    ];
    rows[8][19] = "вибув до 2 штурмової роти";
    rows[9] = ["На продовольчому забезпеченні в УПРАВЛІННІ перебуває"];
    const layout = buildTimesheetLayout(sheetOf(rows), {
      formulas: { I10: 'COUNTIF(I8:I8,"+")' },
    });
    const slot = resolveCanonicalTimesheetSlot({
      index: "2103764",
      layout,
    });
    expect(slot.row).toBe(8);
    expect(slot.occupantName).toMatch(/атрахов/i);
    const company = layout.sections.find((section) =>
      /управлінн/i.test(section.label),
    );
    expect(company?.canonicalEndRow).toBe(8);
    expect(company?.historyStartRow).toBe(9);
  });

  it("flags two empty staff slots with the same index", () => {
    const rows = Array.from({ length: 16 }, emptyRow);
    rows[6] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[7] = ["1 ПІХОТНА РОТА", "2103000", "Стрілець", "100915А", "5"];
    rows[8] = ["1 ПІХОТНА РОТА", "2103000", "Стрілець", "100915А", "5"];
    rows[9] = ["На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває"];
    const layout = buildTimesheetLayout(sheetOf(rows), {
      formulas: { I10: 'COUNTIF(I8:I9,"+")' },
    });
    expect(layout.issues[0]?.code).toBe("TIMESHEET_SLOT_AMBIGUOUS");
    expect(() =>
      resolveCanonicalTimesheetSlot({ index: "2103000", layout }),
    ).toThrow(/кілька штатних рядків/i);
  });

  it("does not treat history copies as canonical even when they look like staff", () => {
    const rows = Array.from({ length: 500 }, emptyRow);
    rows[351] = ["3 ПІХОТНА РОТА", "3 ПІХОТНА РОТА"];
    rows[352] = [
      "3 ПІХОТНА РОТА",
      "2103461",
      "Старший стрілець",
      "100915А",
      "5",
      "старший солдат",
      "",
      "",
    ];
    rows[477] = [
      "3 ПІХОТНА РОТА",
      "2103999",
      "Стрілець",
      "100915А",
      "5",
      "",
      "",
      "",
    ];
    for (let row = 478; row <= 493; row += 1) {
      rows[row] = ["", `21034${row}`, "Стрілець", "100915А", "5"];
    }
    rows[492] = [
      "3 ПІХОТНА РОТА",
      "2103461",
      "Старший стрілець",
      "100915А",
      "5",
      "старший солдат",
      "ПОЧЕПЕЦЬКИЙ",
      "111",
    ];
    rows[493] = ["", "2103416", "Стрілець", "100915А", "5"];
    rows[495] = ["", "2103431", "Стрілець", "100915А", "5"];
    rows[496] = [
      "3 ПІХОТНА РОТА",
      "2111771",
      "Стрілець",
      "100915А",
      "5",
      "солдат",
      "КОВАЛЬЧУК",
      "222",
    ];
    rows[497] = [
      "На продовольчому забезпеченні в 3 ПІХОТНІЙ РОТІ перебуває",
      "65",
    ];
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet, {
      formulas: { I498: 'COUNTIF(I353:I478,"+")' },
    });
    const section = layout.sections.find((item) =>
      /3 ПІХОТНА РОТА/i.test(item.label),
    );
    expect(section).toMatchObject({
      headerRow: 352,
      canonicalStartRow: 353,
      canonicalEndRow: 478,
      historyStartRow: 479,
      historyEndRow: 497,
      footerRow: 498,
    });
    const slot = resolveCanonicalTimesheetSlot({
      index: "2103461",
      layout,
    });
    expect(slot.row).toBe(353);
    expect(slot.occupied).toBe(false);
    const history = resolveHistoryTimesheetRow({
      sourceSlot: slot,
      layout,
      sheet,
    });
    expect(history).toEqual({ row: 495, insertBefore: false });
    expect(() =>
      assertHistoryRowsOutsideCanonical(layout, [history.row]),
    ).not.toThrow();
  });

  it("does not take section from a misplaced history row", () => {
    const rows = Array.from({ length: 80 }, emptyRow);
    rows[5] = ["УПРАВЛІННЯ", "УПРАВЛІННЯ"];
    rows[6] = [
      "УПРАВЛІННЯ",
      "2103119",
      "Командир батальйону",
      "210003",
      "24",
      "",
      "",
      "",
    ];
    rows[33] = [
      "УПРАВЛІННЯ",
      "2103167",
      "Гранатометник",
      "103061А",
      "5",
      "солдат",
      "ВІЄРА",
      "12345",
    ];
    rows[34] = ["На продовольчому забезпеченні в УПРАВЛІННІ перебуває", "1"];
    rows[45] = ["1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[72] = [
      "1 ПІХОТНА РОТА",
      "2103167",
      "Гранатометник",
      "103061А",
      "5",
      "",
      "",
      "",
    ];
    rows[78] = [
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
      "12",
    ];
    const sheet = sheetOf(rows);
    const layout = buildTimesheetLayout(sheet, {
      formulas: {
        I35: 'COUNTIF(I7:I8,"+")',
        I79: 'COUNTIF(I47:I74,"+")',
      },
    });
    const slot = resolveCanonicalTimesheetSlot({
      index: "2103167",
      layout,
    });
    expect(slot.row).toBe(73);
    expect(slot.section).toMatch(/1 ПІХОТНА РОТА/i);
    const history = resolveHistoryTimesheetRow({
      sourceSlot: slot,
      sourceRow: 34,
      layout,
      sheet,
    });
    expect(history.row).toBeGreaterThan(74);
    expect(history.row).toBeLessThan(79);
    expect(() =>
      resolveHistoryTimesheetRow({
        sourceRow: 34,
        layout,
        sheet,
      }),
    ).toThrow(TimesheetLayoutError);
  });
});
