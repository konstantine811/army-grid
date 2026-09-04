import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  companyUnitPhrasesFromPosition,
  extractUnitPhrasesFromPosition,
  findTimesheetAppendRowForUnit,
  findTimesheetUnitSectionBounds,
  timesheetAppendNeedsRowInsert,
  timesheetRowInExpectedUnitSection,
  unitDescriptorsMatch,
} from "./ejoosTimesheetUnitSections";

const sheetOf = (rows: string[][]): ExcelSheetSnapshot => ({
  sheetName: "6. Табель",
  rawRows: rows,
  columnIndexes: [],
});

describe("ejoosTimesheetUnitSections", () => {
  it("matches company headers across cases", () => {
    expect(unitDescriptorsMatch("1 ПІХОТНА РОТА", "1 піхотної роти")).toBe(true);
    expect(unitDescriptorsMatch("'1 ПІХОТНА РОТА", "1 піхотної роти")).toBe(true);
    expect(unitDescriptorsMatch("'УПРАВЛІННЯ", "управління")).toBe(true);
    expect(unitDescriptorsMatch("2 піхотне відділення", "2 піхотного відділення")).toBe(
      true,
    );
    expect(unitDescriptorsMatch("1 ПІХОТНА РОТА", "2 ПІХОТНА РОТА")).toBe(false);
    expect(
      unitDescriptorsMatch(
        "'ВЗВОД  ЛОГІСТИЧНО-ЕВАКУАЦІЙНИХ БЕЗПІЛОТНИХ НАЗЕМНИХ СИСТЕМ",
        "взводу логістично-евакуаційних безпілотних наземних систем",
      ),
    ).toBe(true);
    expect(
      unitDescriptorsMatch(
        "ВЗВОД ЗВ'ЯЗКУ",
        "взводу логістично-евакуаційних безпілотних наземних систем",
      ),
    ).toBe(false);
  });

  it("finds section bounds for apostrophe headers", () => {
    const rows = Array.from({ length: 85 }, () => Array(40).fill(""));
    rows[46] = ["'1 ПІХОТНА РОТА"];
    rows[47] = ["", "2103144", "", "", "", "лейтенант", "ІВАНОВ", "10001"];
    rows[78] = ["'2 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const bounds = findTimesheetUnitSectionBounds({ sheet }, "1 піхотної роти");
    expect(bounds?.headerRow).toBe(47);
    expect(bounds?.endRow).toBe(78);
  });

  it("extracts unit chain from a long position title", () => {
    const phrases = extractUnitPhrasesFromPosition(
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону",
    );
    expect(phrases.some((item) => /2.*відділен/i.test(item))).toBe(true);
    expect(phrases.some((item) => /1.*рот/i.test(item))).toBe(true);
  });

  it("appends at the end of 1 ПІХОТНА РОТА, not under УПРАВЛІННЯ", () => {
    const rows = Array.from({ length: 210 }, () => Array(40).fill(""));
    rows[45] = ["'УПРАВЛІННЯ", "2103198", "", "", "", "солдат", "ХИБНИЙ", "99999"];
    rows[46] = ["'1 ПІХОТНА РОТА"];
    rows[47] = ["", "2103144", "", "", "", "лейтенант", "ІВАНОВ", "10001"];
    rows[197] = ["", "2103190", "", "", "", "солдат", "КРИВЦОВ", "3136"];
    rows[198] = ["", "", "", "", "", "", "", ""];
    rows[199] = ["На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває"];
    const sheet = sheetOf(rows);
    const bounds = findTimesheetUnitSectionBounds({ sheet }, "1 піхотної роти");
    expect(bounds?.headerRow).toBe(47);
    expect(bounds?.endRow).toBe(199);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 46,
      positionTitle:
        "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону",
      staffIndex: "2103198",
    });
    expect(appendRow).toBe(199);
    expect(appendRow).not.toBe(46);
  });

  it("detects a history row misplaced under УПРАВЛІННЯ", () => {
    const rows = Array.from({ length: 60 }, () => Array(40).fill(""));
    rows[45] = ["'УПРАВЛІННЯ"];
    rows[46] = [
      "",
      "2103198",
      "",
      "",
      "",
      "солдат",
      "ВІЄРА ДА КОСТА",
      "12345",
    ];
    rows[47] = ["'1 ПІХОТНА РОТА"];
    rows[48] = ["", "2103144", "", "", "", "лейтенант", "ІВАНОВ", "10001"];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";
    expect(timesheetRowInExpectedUnitSection(sheet, 46, title)).toBe(false);
    expect(timesheetRowInExpectedUnitSection(sheet, 48, title)).toBe(true);
  });

  it("appends departure history at end of whole company, not platoon", () => {
    const rows = Array.from({ length: 60 }, () => Array(40).fill(""));
    rows[46] = ["'1 ПІХОТНА РОТА"];
    rows[47] = ["", "2103100", "", "", "", "лейтенант", "РОТНИЙ", "10001"];
    rows[48] = ["'2 ПІХОТНИЙ ВЗВОД"];
    rows[49] = ["", "2103101", "", "", "", "солдат", "ВЗВОДНИК", "10002"];
    rows[50] = ["'1 ПІХОТНЕ ВІДДІЛЕННЯ"];
    rows[51] = [
      "",
      "2103198",
      "",
      "",
      "",
      "солдат",
      "КУЛЕМЕТНИК",
      "10003",
    ];
    rows[52] = ["'3 ПІХОТНИЙ ВЗВОД"];
    rows[53] = ["", "2103102", "", "", "", "солдат", "ІНШИЙ ВЗВОД", "10004"];
    rows[54] = ["", "", "", "", "", "", "", ""];
    rows[55] = ["На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває"];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 1 піхотного відділення 2 піхотного взводу 1 піхотної роти";

    const companyRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 51,
      positionTitle: title,
      placementScope: "company",
    });
    expect(companyRow).toBe(55);
    expect(companyRow).toBeGreaterThan(52);
  });

  it("finds company section when header, platoon, squad and summary are in column B", () => {
    const rows = Array.from({ length: 220 }, () => Array(40).fill(""));
    rows[45] = [
      "'УПА БЕЗПІЛОТНИХ СИСТ",
      "2111765",
      "600723A",
      "10",
      "",
      "солдат",
      "ПОПЕРЕДНІЙ",
      "10000",
    ];
    rows[46] = [
      "'1 ПІХОТНА РОТА",
      "1 ПІХОТНА РОТА",
    ];
    rows[47] = [
      "'1 ПІХОТНА РОТА",
      "2103144",
      "210003",
      "19",
      "",
      "лейтенант",
      "ІВАНОВ",
      "10001",
    ];
    rows[89] = [
      "'1 ПІХОТНА РОТА",
      "2103180",
      "100915A",
      "5",
      "",
      "солдат",
      "ПЕРШИЙ",
      "10002",
    ];
    rows[90] = ["'1 ПІХОТНА РОТА", "2 ПІХОТНИЙ ВЗВОД"];
    rows[91] = [
      "'1 ПІХОТНА РОТА",
      "2103181",
      "790037A",
      "4",
      "",
      "солдат",
      "ДРУГИЙ",
      "10003",
    ];
    rows[103] = ["'1 ПІХОТНА РОТА", "2 піхотне відділення"];
    rows[104] = [
      "'1 ПІХОТНА РОТА",
      "2103198",
      "193540A",
      "4",
      "",
      "солдат",
      "КУЛЕМЕТНИК",
      "10004",
    ];
    rows[197] = ["", "2103190", "", "", "", "солдат", "КРИВЦОВ", "3136"];
    rows[198] = ["", "", "", "", "", "", "", ""];
    rows[199] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[200] = ["'2 ПІХОТНА РОТА", "2 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";

    const bounds = findTimesheetUnitSectionBounds({ sheet }, "1 піхотної роти");
    expect(bounds?.headerRow).toBe(47);
    expect(bounds?.endRow).toBe(199);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 45,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(199);
  });

  it("does not truncate company section on duplicate rota label in col A", () => {
    const rows = Array.from({ length: 210 }, () => Array(40).fill(""));
    rows[46] = ["'1 ПІХОТНА РОТА"];
    rows[47] = ["", "2103100", "", "", "", "лейтенант", "РОТНИЙ", "10001"];
    rows[191] = ["'1 ПІХОТНА РОТА", "2103242", "Вогнеметник", "193540A", "4", "солдат", "ТЕНЬКА", "22724"];
    rows[197] = ["", "2103190", "", "", "", "солдат", "КРИВЦОВ", "3136"];
    rows[198] = ["", "", "", "", "", "", "", ""];
    rows[199] = ["На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває"];
    rows[200] = ["'2 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const bounds = findTimesheetUnitSectionBounds({ sheet }, "1 піхотної роти");
    expect(bounds?.endRow).toBe(199);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 900,
      positionTitle:
        "вогнеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти",
      placementScope: "company",
    });
    expect(appendRow).toBe(199);
    expect(appendRow).toBeLessThan(200);
  });

  it("appends after departed block before next rota, not at sheet tail", () => {
    const rows = Array.from({ length: 950 }, () => Array(40).fill(""));
    rows[46] = ["'1 ПІХОТНА РОТА"];
    rows[47] = ["", "2103100", "", "", "", "лейтенант", "РОТНИЙ", "10001"];
    for (let row = 172; row <= 191; row += 1) {
      rows[row - 1] = ["", `2103${row}`, "", "", "", "солдат", `ВИБУВ ${row}`, "10000"];
    }
    rows[197] = ["", "2103190", "", "", "", "солдат", "КРИВЦОВ", "3136"];
    rows[198] = ["", "", "", "", "", "", "", ""];
    rows[199] = ["На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває"];
    rows[200] = ["'2 ПІХОТНА РОТА"];
    rows[900] = ["", "2103198", "", "", "", "солдат", "ВІЄРА", "12345"];
    const sheet = sheetOf(rows);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 900,
      positionTitle:
        "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти",
      placementScope: "company",
    });
    expect(appendRow).toBe(199);
    expect(appendRow).toBe(199);
  });

  it("finds company section by repeated col A when col B has no rota title", () => {
    const rows = Array.from({ length: 220 }, () => Array(40).fill(""));
    rows[46] = [
      "'1 ПІХОТНА РОТА",
      "2103144",
      "210003",
      "19",
      "",
      "лейтенант",
      "ІВАНОВ",
      "10001",
    ];
    rows[90] = ["'1 ПІХОТНА РОТА", "2 ПІХОТНИЙ ВЗВОД"];
    rows[103] = ["'1 ПІХОТНА РОТА", "2 піхотне відділення"];
    rows[104] = [
      "'1 ПІХОТНА РОТА",
      "2103198",
      "193540A",
      "4",
      "",
      "солдат",
      "КУЛЕМЕТНИК",
      "10004",
    ];
    rows[196] = [
      "'1 ПІХОТНА РОТА",
      "2103190",
      "",
      "",
      "",
      "солдат",
      "КРИВЦОВ",
      "3136",
    ];
    rows[197] = ["'1 ПІХОТНА РОТА", "", "", "", "", "", "", ""];
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[199] = ["'2 ПІХОТНА РОТА", "2 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";

    const bounds = findTimesheetUnitSectionBounds({ sheet }, "1 піхотної роти");
    expect(bounds?.headerRow).toBe(47);
    expect(bounds?.endRow).toBe(198);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 105,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(198);
  });

  it("stacks append rows when the pre-summary slot is already reserved", () => {
    const rows = Array.from({ length: 220 }, () => Array(40).fill(""));
    rows[46] = [
      "'1 ПІХОТНА РОТА",
      "2103144",
      "210003",
      "19",
      "",
      "лейтенант",
      "ІВАНОВ",
      "10001",
    ];
    rows[196] = [
      "'1 ПІХОТНА РОТА",
      "2103190",
      "",
      "",
      "",
      "солдат",
      "КРИВЦОВ",
      "3136",
    ];
    rows[197] = ["'1 ПІХОТНА РОТА", "", "", "", "", "", "", ""];
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";
    const reserved = new Set<number>([198]);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 46,
      positionTitle: title,
      placementScope: "company",
      reserved,
    });
    expect(appendRow).toBe(199);
  });

  it("inserts before summary without targeting the next company header", () => {
    const rows = Array.from({ length: 210 }, () => Array(40).fill(""));
    rows[46] = ["'1 ПІХОТНА РОТА", "1 ПІХОТНА РОТА"];
    rows[47] = ["'1 ПІХОТНА РОТА", "2103144", "", "", "", "лейтенант", "ІВАНОВ", "10001"];
    rows[195] = ["'1 ПІХОТНА РОТА", "2103255", "", "", "", "солдат", "ТЕНЬКА", "3136"];
    rows[196] = ["'1 ПІХОТНА РОТА", "2103190", "", "", "", "солдат", "КРИВЦОВ", "3136"];
    rows[197] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[199] = ["'2 ПІХОТНА РОТА", "2 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";
    const bounds = findTimesheetUnitSectionBounds({ sheet }, "1 піхотної роти");
    expect(bounds?.endRow).toBe(197);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 47,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(198);
    expect(timesheetAppendNeedsRowInsert({ sheet }, appendRow, bounds!)).toBe(true);
  });

  it("does not fall back to УПРАВЛІННЯ when source row is misplaced", () => {
    const rows = Array.from({ length: 220 }, () => Array(40).fill(""));
    rows[45] = [
      "'УПРАВЛІННЯ",
      "2103198",
      "",
      "",
      "",
      "солдат",
      "ВІЄРА",
      "24553",
    ];
    rows[46] = [
      "'1 ПІХОТНА РОТА",
      "2103144",
      "210003",
      "19",
      "",
      "лейтенант",
      "ІВАНОВ",
      "10001",
    ];
    rows[196] = [
      "'1 ПІХОТНА РОТА",
      "2103190",
      "",
      "",
      "",
      "солдат",
      "КРИВЦОВ",
      "3136",
    ];
    rows[197] = ["'1 ПІХОТНА РОТА", "", "", "", "", "", "", ""];
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    const sheet = sheetOf(rows);
    const title =
      "кулеметник 2 піхотного відділення 2 піхотного взводу 1 піхотної роти 1 піхотного батальйону";

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 46,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(198);
    expect(appendRow).toBeGreaterThan(46);
  });

  it("finds 2nd company with a single person when summary omits col A rota label", () => {
    const rows = Array.from({ length: 250 }, () => Array(40).fill(""));
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[199] = ["'2 ПІХОТНА РОТА", "2103200", "", "", "", "солдат", "КЛУБАНЬ", "10001"];
    rows[200] = ["", "На продовольчому забезпеченні в 2 ПІХОТНІЙ РОТІ перебуває"];
    const sheet = sheetOf(rows);
    const title =
      "старший стрілець - оператор безпілотних літальних апаратів 2 піхотного відділення 1 піхотного взводу 2 піхотної роти 1 піхотного батальйону";

    const bounds = findTimesheetUnitSectionBounds({ sheet }, "2 піхотної роти");
    expect(bounds?.headerRow).toBe(200);
    expect(bounds?.endRow).toBe(200);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 200,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(201);
    expect(timesheetAppendNeedsRowInsert({ sheet }, appendRow, bounds!)).toBe(true);
  });

  it("inserts before next rota when company 2 has no summary row", () => {
    const rows = Array.from({ length: 250 }, () => Array(40).fill(""));
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[199] = ["'2 ПІХОТНА РОТА", "2103200", "", "", "", "солдат", "КЛУБАНЬ", "10001"];
    rows[200] = ["'3 ПІХОТНА РОТА", "3 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const title =
      "старший стрілець - оператор безпілотних літальних апаратів 2 піхотного відділення 1 піхотного взводу 2 піхотної роти 1 піхотного батальйону";

    const bounds = findTimesheetUnitSectionBounds({ sheet }, "2 піхотної роти");
    expect(bounds?.headerRow).toBe(200);
    expect(bounds?.endRow).toBe(200);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 200,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(201);
    expect(timesheetAppendNeedsRowInsert({ sheet }, appendRow, bounds!)).toBe(true);
  });

  it("stacks departure rows when summary slot is already reserved", () => {
    const rows = Array.from({ length: 250 }, () => Array(40).fill(""));
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольchому забезпеченні в 1 PІХOTNIJ ROTI перebуває",
    ];
    rows[199] = ["'2 ПІХОТНА РОТА", "2103200", "", "", "", "солдат", "КЛУБАНЬ", "10001"];
    rows[200] = ["", "На продовольчому забезпеченні в 2 ПІХОТНІЙ РОТІ перебуває"];
    rows[201] = ["'3 ПІХОТНА РОТА", "3 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const title =
      "старший стрілець - оператор безпілотних літальних апаратів 2 піхотного відділення 1 піхотного взводу 2 піхотної роти 1 піхотного батальйону";
    const reserved = new Set<number>([201]);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 200,
      positionTitle: title,
      placementScope: "company",
      reserved,
    });
    expect(appendRow).toBe(201);
  });

  it("does not treat summary col B as company header", () => {
    const rows = Array.from({ length: 250 }, () => Array(40).fill(""));
    rows[199] = ["'2 ПІХОТНА РОТА", "2103200", "", "", "", "солдат", "КЛУБАНЬ", "10001"];
    rows[200] = ["", "На продовольчому забезпеченні в 2 ПІХОТНІЙ РОТІ перебуває"];
    const sheet = sheetOf(rows);
    const bounds = findTimesheetUnitSectionBounds({ sheet }, "2 піхотної роти");
    expect(bounds?.headerRow).toBe(200);
    expect(bounds?.division).not.toMatch(/продоволь/i);
  });

  it("finds 2nd company section by repeated col A without col B rota header", () => {
    const rows = Array.from({ length: 250 }, () => Array(40).fill(""));
    rows[45] = ["'1 ПІХОТНА РОТА", "2103144", "", "", "", "лейтенант", "РОТНИЙ", "10001"];
    rows[198] = [
      "'1 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 1 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[199] = ["'2 ПІХОТНА РОТА", "2103200", "", "", "", "солдат", "КЛУБАНЬ", "10001"];
    rows[240] = ["'2 ПІХОТНА РОТА", "2103400", "", "", "", "солдат", "ОСТАННІЙ", "10002"];
    rows[241] = [
      "'2 ПІХОТНА РОТА",
      "На продовольчому забезпеченні в 2 ПІХОТНІЙ РОТІ перебуває",
    ];
    rows[242] = ["'3 ПІХОТНА РОТА", "3 ПІХОТНА РОТА"];
    const sheet = sheetOf(rows);
    const title =
      "старший стрілець - оператор безпілотних літальних апаратів 2 піхотного відділення 1 піхотного взводу 2 піхотної роти 1 піхотного батальйону";

    const bounds = findTimesheetUnitSectionBounds({ sheet }, "2 піхотної роти");
    expect(bounds?.headerRow).toBe(200);
    expect(bounds?.endRow).toBe(241);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 200,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBe(242);
    expect(timesheetAppendNeedsRowInsert({ sheet }, appendRow, bounds!)).toBe(true);
  });

  it("appends departure at end of unnumbered logistics platoon section", () => {
    const title =
      "водій-електрик-моторист відділення важких логістично-евакуаційних безпілотних наземних систем взводу логістично-евакуаційних безпілотних наземних систем 1 піхотного батальйону";
    const phrases = extractUnitPhrasesFromPosition(title);
    expect(phrases.some((item) => /взводу.*логістично/i.test(item))).toBe(true);
    expect(companyUnitPhrasesFromPosition(title).some((item) => /взводу/i.test(item))).toBe(
      true,
    );

    const rows = Array.from({ length: 80 }, () => Array(40).fill(""));
    rows[20] = [
      "'1 ПІХОТНИЙ БАТАЛЬЙОН",
      "ВЗВОД  ЛОГІСТИЧНО-ЕВАКУАЦІЙНИХ БЕЗПІЛОТНИХ НАЗЕМНИХ СИСТЕМ",
    ];
    rows[21] = [
      "",
      "2103401",
      "",
      "",
      "",
      "солдат",
      "ВОДІЙ",
      "10001",
    ];
    rows[22] = ["", "", "", "", "", "", "", ""];
    rows[23] = ["", "2103402", "", "", "", "солдат", "ІНШИЙ", "10002"];
    const sheet = sheetOf(rows);

    const bounds = findTimesheetUnitSectionBounds(
      { sheet },
      "взводу логістично-евакуаційних безпілотних наземних систем",
    );
    expect(bounds?.headerRow).toBe(21);
    expect(bounds?.endRow).toBeGreaterThanOrEqual(22);

    const appendRow = findTimesheetAppendRowForUnit(sheet, {
      sourceRow: 21,
      positionTitle: title,
      placementScope: "company",
    });
    expect(appendRow).toBeGreaterThanOrEqual(22);
  });
});
