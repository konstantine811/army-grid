// @ts-expect-error xlsx-populate does not ship type declarations.
import XlsxPopulate from "xlsx-populate";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  buildBchsMorningSections,
  buildRotaBchsMorningExportFileName,
  computeBchsMorningSummaryFieldCounts,
  resolveBchsMorningSheetName,
  resolveBchsMorningStaffUnit,
  resolveReportStatus,
  splitPlatoonTablePeople,
  writeRotaBchsMorningWorkbook,
} from "./overviewRotaBchsMorningExport";

const row = (
  overrides: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: "1",
  externalId: "1", validFrom: null, days: null, plannedReturn: null, place: "", updatedAt: "",
  name: "БІЛИЦЬКИЙ Сергій Володимирович",
  unit: "3 піхотна рота",
  rank: "солдат",
  status: "ON_DUTY",
  statusLabel: "В строю",
  staffStatus: "В строю",
  staffStatusLabel: "В стroю",
  staffSheetColumns: {
    staff_3: "1 піхотний взвод",
    staff_5: "стрілець",
    staff_13: "солдат",
    staff_15: "НАЛІМ",
    staff_21: "В строю",
    staff_31: "ППД Вишневе",
  },
  ...overrides,
});

describe("overviewRotaBchsMorningExport", () => {
  it("builds rota sheet and staff labels", () => {
    expect(resolveBchsMorningSheetName("3 піхотна рота")).toBe("3ПР");
    expect(resolveBchsMorningStaffUnit("3 піхотна рота")).toBe("1ПБ 3ПР");
    expect(
      buildRotaBchsMorningExportFileName(
        "3 піхотна рота",
        new Date("2026-09-09T12:00:00"),
      ),
    ).toBe("ПБ 3ПР БЧС 09.09.2026.xlsx");
  });

  it("splits in-rota and medical people into left/right sections", () => {
    const sections = buildBchsMorningSections(
      [
        row(),
        row({
          id: "2",
          name: "КУЗЬМЕНКО Руслан Олегович",
          status: "MEDICAL",
          statusLabel: "Лікування",
          staffSheetColumns: {
            staff_3: "1 піхотний взвод",
            staff_5: "стрілець",
            staff_13: "солдат",
            staff_15: "КОМЕРС",
            staff_21: "Лікування",
            staff_31: "Шпиталь",
            staff_32: "16.10 продовжує лікування по пораненню",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.some((section) => section.title === "1 Піхотний взвод")).toBe(
      true,
    );
    expect(
      sections.some((section) =>
        section.title.includes("Шпиталь/Мед. рота"),
      ),
    ).toBe(true);
    expect(
      sections.find((section) => section.title === "1 Піхотний взвод")?.people,
    ).toHaveLength(1);
  });

  it("writes only short Посада, not Повна посада", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          positionTitle: "стрілець 1 піхотного відділення 1 піхотного взводу",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "стрілець",
            staff_7: "стрілець 1 піхотного відділення 1 піхотного взводу",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.find((section) => section.id === "platoon1")?.people[0]?.position).toBe(
      "стрілець",
    );
  });

  it("puts командир взводу first, then командир відділення with section number", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "СТРІЛЕЦЬ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_4: "2 відділення",
            staff_5: "стрілець",
          },
        }),
        row({
          name: "ВІДДІЛЕННЯ Два",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_4: "2 відділення",
            staff_5: "Командир відділення",
          },
        }),
        row({
          name: "ВІДДІЛЕННЯ Один",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_4: "1 відділення",
            staff_5: "Командир відділення",
          },
        }),
        row({
          name: "ВЗВОДНИЙ Іван Іванович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "Командир взводу",
          },
        }),
      ],
      "3 піхотна рота",
    );

    const people = sections.find((section) => section.id === "platoon1")?.people ?? [];
    expect(people.map((person) => person.name)).toEqual([
      "ВЗВОДНИЙ Іван Іванович",
      "ВІДДІЛЕННЯ Один",
      "ВІДДІЛЕННЯ Два",
      "СТРІЛЕЦЬ Петро Петрович",
    ]);
    expect(people.map((person) => person.position)).toEqual([
      "Командир взводу",
      "Командир 1 відділення",
      "Командир 2 відділення",
      "стрілець",
    ]);
    expect(splitPlatoonTablePeople(people)).toEqual({
      commanders: [expect.objectContaining({ name: "ВЗВОДНИЙ Іван Іванович" })],
      sectionCommanders: [
        expect.objectContaining({ name: "ВІДДІЛЕННЯ Один", position: "Командир 1 відділення" }),
        expect.objectContaining({ name: "ВІДДІЛЕННЯ Два", position: "Командир 2 відділення" }),
      ],
      rest: [expect.objectContaining({ name: "СТРІЛЕЦЬ Петро Петрович" })],
    });
  });

  it("keeps командир взводу and командир відділення in the vzvod table in any status", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "ВЗВОДНИЙ У відпустці",
          status: "LEAVE",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "Командир взводу",
            staff_21: "Відпустка",
          },
        }),
        row({
          name: "ВІДДІЛЕННЯ У шпиталі",
          status: "MEDICAL",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_4: "2 відділення",
            staff_5: "Командир відділення",
            staff_21: "Лікування",
            staff_31: "Шпиталь",
          },
        }),
        row({
          name: "СТРІЛЕЦЬ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "стрілець",
          },
        }),
      ],
      "3 піхотна рота",
    );

    const groups = splitPlatoonTablePeople(
      sections.find((section) => section.id === "platoon1")?.people ?? [],
    );
    expect(groups.commanders.map((person) => person.name)).toEqual(["ВЗВОДНИЙ У відпустці"]);
    expect(groups.sectionCommanders.map((person) => person.name)).toEqual(["ВІДДІЛЕННЯ У шпиталі"]);
    expect(groups.rest.map((person) => person.name)).toEqual(["СТРІЛЕЦЬ Петро Петрович"]);
    expect(sections.find((section) => section.id === "leave")?.people).toHaveLength(1);
    expect(sections.find((section) => section.id === "hospital")?.people).toHaveLength(1);
  });

  it("writes the real rank, not командир взводу or командир відділення", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "ВЗВОДНИЙ Іван Іванович",
          rank: "молодший лейтенант",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "Командир взводу",
            staff_13: "молодший лейтенант",
          },
        }),
        row({
          name: "ВІДДІЛЕННЯ Петро Петрович",
          rank: "сержант",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "Командир відділення",
            staff_13: "сержант",
          },
        }),
      ],
      "3 піхотна рота",
    );

    const people = sections.find((section) => section.id === "platoon1")?.people ?? [];
    expect(people.map((person) => person.rank)).toEqual([
      "молодший лейтенант",
      "сержант",
    ]);
    expect(people.map((person) => person.position)).toEqual([
      "Командир взводу",
      "Командир відділення",
    ]);
  });

  it("writes workbook cells with template-based styles", async () => {
    const wb = await XlsxPopulate.fromBlankAsync();
    const sheet = wb.sheet(0);
    sheet.name("3ПР");
    for (let rowNumber = 1; rowNumber <= 12; rowNumber += 1) {
      for (let column = 1; column <= 16; column += 1) {
        const cell = sheet.cell(rowNumber, column);
        cell.style("fontSize", 16);
        cell.style("fontColor", { rgb: "FF000000" });
        if (rowNumber === 2) {
          cell.style("fill", { type: "solid", color: { rgb: "FFBED7EE" } });
        }
      }
    }

    writeRotaBchsMorningWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      staffCount: 113,
      rows: [
        row(),
        row({
          id: "2",
          name: "КОВАЛЬ Іван Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_15: "КРОХА",
          },
        }),
      ],
    });

    expect(sheet.name()).toBe("3ПР");
    expect(sheet.cell(2, 1).value()).toBe("1 Піхотний взвод");
    expect(sheet.cell(3, 1).value()).toBe("№");
    expect(sheet.column(1).width()).toBeGreaterThanOrEqual(6);
    expect(sheet.cell(4, 1).value()).toBe("Командир взводу");
    expect(sheet.cell(5, 1).value()).toBe("Командир відділення");
    expect(sheet.cell(6, 1).style("numberFormat")).toBe("General");
    expect(sheet.cell(6, 1).value()).toBe(1);
    expect(sheet.cell(6, 3).value()).toBe("БІЛИЦЬКИЙ Сергій Володимирович");
    expect(sheet.cell(6, 4).value()).toBe("НАЛІМ");
    expect(sheet.cell(6, 6).value()).toBe("В строю");
    expect(sheet.cell(6, 8).value()).toBe("1ПБ 3ПР");
    expect(sheet.cell(7, 1).value()).toBe(2);
    expect(sheet.cell(7, 3).value()).toBe("КОВАЛЬ Іван Петрович");
    expect(sheet.cell(6, 3).style("fontColor")).toMatchObject({
      rgb: "FF000000",
    });
  });

  it("fills real template without leaving sample ШТАТ rows below export", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/pb-rota-bchs-morning-template.xlsx",
    );
    writeRotaBchsMorningWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      staffCount: 113,
      rows: [row()],
    });
    const sheet = wb.sheet("3ПР");
    expect(sheet.cell(4, 1).value()).toBe("Командир взводу");
    expect(sheet.cell(6, 3).value()).toBe("БІЛИЦЬКИЙ Сергій Володимирович");
    expect(sheet.cell(50, 8).value()).toBeNull();
    expect(wb.sheets().map((entry: { name: () => string }) => entry.name())).toEqual(["3ПР"]);
  });

  it("rebuilds all totals and removes stale template tables, merges and formulas", async () => {
    const wb = await XlsxPopulate.fromFileAsync("public/templates/pb-rota-bchs-morning-template.xlsx");
    const make = (status: string, position = "стрілець", staff_3 = row().staffSheetColumns?.staff_3) => row({
      staffSheetColumns: { ...row().staffSheetColumns, staff_21: status, staff_5: position, staff_3 },
    });
    const rows = [
      make("В строю", "командир роти", "ж"),
      ...Array.from({ length: 22 }, () => make("Відком. за межі ПБ")),
      make("Прикомандирований"), make("Лікувальна відпустка"),
      make("500"), make("Навч."), make("На виході"), make("Вик. БЗ"),
      make("СЗЧ"), make("В строю БГ"), make("В строю"),
    ];
    writeRotaBchsMorningWorkbook(wb, { unitLabel: "3 піхотна рота", rows, staffCount: 113 });
    const sheet = wb.sheet("3ПР");
    expect(sheet.cell("V10").value()).toBe(22);
    expect(sheet.cell("U21").value()).toBe("Взводні (наявність)");
    expect(sheet.cell("V21").value()).toBe(0);
    expect(sheet.cell("V22").value()).toBeNull();
    expect(sheet.cell("V19").value()).toBe(1);
    expect(sheet.cell("V12").value()).toBe(1);
    expect(sheet.cell("V7").value()).toBe(rows.length);
    expect(sheet.cell("V27").value()).toBe(3);
    expect(sheet.cell("V25").value()).toBe(0);
    expect(sheet.cell("V26").value()).toBe(4);
    expect(sheet.cell("T68").value()).toBeNull();
    expect(sheet.cell("B73").value()).toBeNull();
    expect(sheet.cell("B73").style("bottomBorder").style).toBeUndefined();
    const sections = buildBchsMorningSections(rows, "3 піхотна рота");
    let nextRow = 2;
    for (const section of sections.filter(section => section.side === "right")) {
      expect(sheet.cell(nextRow + 2 + section.people.length, 17).value()).toBe(section.people.length);
      nextRow += section.people.length + 4;
    }
    const saved = await XlsxPopulate.fromDataAsync(await wb.outputAsync());
    expect(saved.sheet("3ПР").cell("V10").value()).toBe(22);
    expect(saved.sheet("3ПР").cell("V10").formula()).toBeUndefined();
    expect(saved.sheet("3ПР").range("B7:G7").merged()).toBe(false);
  });

  it("keeps 113 authorized positions separate from 85 listed people", async () => {
    const wb = await XlsxPopulate.fromBlankAsync();
    writeRotaBchsMorningWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      rows: Array.from({ length: 85 }, (_, index) => row({ id: String(index) })),
      staffCount: 113,
    });
    const sheet = wb.sheet("3ПР");
    expect(sheet.cell("V6").value()).toBe(113);
    expect(sheet.cell("V7").value()).toBe(85);
  });

  it.each(["1 піхотна рота", "2 піхотна рота", "3 піхотна рота"])("writes valid scoped print area for %s after removing template sheets", async (unitLabel) => {
    const wb = await XlsxPopulate.fromFileAsync("public/templates/pb-rota-bchs-morning-template.xlsx");
    writeRotaBchsMorningWorkbook(wb, { unitLabel, rows: [row()], staffCount: 113 });
    const zip = await JSZip.loadAsync(await wb.outputAsync());
    const xml = await zip.file("xl/workbook.xml")!.async("string");
    const names = [...xml.matchAll(/<definedName\s([^>]*)>([^<]*)<\/definedName>/g)];
    expect(names).toHaveLength(1);
    expect(names[0][1]).toContain('name="_xlnm.Print_Area"');
    expect(names[0][1]).toContain('localSheetId="0"');
    expect(names[0][2]).toBe(`'${resolveBchsMorningSheetName(unitLabel)}'!$A$1:$AA$38`);
    expect(xml).not.toContain('localSheetId="-1"');
    expect(xml).not.toContain("2 ЗП");
    expect(xml).not.toContain("_FilterDatabase");
    const reloaded = await XlsxPopulate.fromDataAsync(await wb.outputAsync());
    expect(reloaded.sheet(0).definedName("_xlnm.Print_Area").address()).toBe("A1:AA38");
  });

  it("matches the reference table order, seven columns and split subtotals using current people", async () => {
    const wb = await XlsxPopulate.fromFileAsync("public/templates/pb-rota-bchs-morning-template.xlsx");
    const make = (status: string, location = "ППД", position = "стрілець", staff_3 = row().staffSheetColumns?.staff_3) => row({
      staffSheetColumns: { ...row().staffSheetColumns, staff_21: status, staff_31: location, staff_5: position, staff_3 },
    });
    const rows = [
      make("В строю", "ППД", "командир роти", "ж"), make("В строю"), make("Новоприбулий"), make("СЗЧ"),
      make("Лік.Пор", "Шпиталь"), make("Лік.Хвор", "Шпиталь"),
      make("Лік.Пор", "Мед. пункт"), make("Лік.Хвор", "Мед. пункт"),
      make("Лік.Відп."), make("Лікувальна відпустка"), make("Відп."),
      make("Приком."), make("Відком. за межі ПБ"), make("500"),
      make("Навч."), make("Відряд."), make("На виході"), make("Вик. БЗ"), make("Загиблі"),
    ];
    writeRotaBchsMorningWorkbook(wb, { unitLabel: "3 піхотна рота", rows, staffCount: 113 });
    const sheet = wb.sheet(0);
    const sections = buildBchsMorningSections(rows, "3 піхотна рота");
    expect(sections.filter(section => section.side === "left").map(section => section.id)).toEqual(["management", "platoon1", "newcomers", "awol"]);
    const splitCounts: Record<string, number[]> = { hospital: [1, 1, 2], medPoint: [1, 1, 2], leave: [2, 1, 3], detached: [1, 1, 2], training: [1, 1, 2] };
    for (const side of ["right", "extra"] as const) {
      let start = side === "right" ? 2 : 34;
      const column = side === "right" ? 11 : 20;
      const numberColumn = column - 1;
      for (const section of sections.filter(section => section.side === side)) {
        expect(sheet.cell(start, numberColumn).value()).toBe(section.title);
        expect(sheet.cell(start + 1, numberColumn).value()).toBe("№");
        if (section.people.length) {
          expect(sheet.cell(start + 2, numberColumn).value()).toBe(1);
          expect(sheet.cell(start + 1 + section.people.length, numberColumn).value()).toBe(
            section.people.length,
          );
        }
        expect(sheet.cell(start + 1, column + 6).value()).toBe("ШТАТ");
        if (section.people.length) expect(sheet.cell(start + 2, column + 6).value()).toBe("1ПБ 3ПР");
        const totalRow = start + 2 + section.people.length;
        expect(sheet.cell(totalRow, column + 6).value()).toBe(section.people.length);
        if (splitCounts[section.id]) {
          expect([2, 4, 6].map(offset => sheet.cell(totalRow, column + offset).value())).toEqual(splitCounts[section.id]);
        }
        start = totalRow + 2;
      }
    }
    expect(sheet.cell("V11").value()).toBe(4);
    expect(sheet.cell("V12").value()).toBe(3);
    expect(sheet.cell("V7").value()).toBe(rows.length);
    expect(sheet.range("U5:V5").merged()).toBe(true);
    expect(sheet.cell("V6").value()).toBe(113);
  });

  it("uses location for missions and the dedicated readiness column for BG, without an out section", async () => {
    const make = (status: string, location: string, readiness: string, extra: Record<string, string> = {}) => row({
      staffSheetColumns: { ...row().staffSheetColumns, staff_21: status, staff_31: location, staff_23: readiness, ...extra },
    });
    const rows = [
      make("В строю", "На виконання", "БГ"),
      make("В строю", "  НА ВИКОНАННІ  ", ""),
      make("В строю", "ППД", " БГ "),
      make("В строю БГ", "ППД", "Не БГ"),
      make("Вик. БЗ", "ППД", ""),
      make("В строю", "ППД", "", { staff_32: "На виконання" }),
      make("На виході", "ППД", ""),
      make("В строю", "ППД", "", { staff_42: "БГ" }),
      make("Відпустка", "На виконання", "БГ"),
      make("СЗЧ", "На виконання", "БГ"),
    ];
    const sections = buildBchsMorningSections(rows, "3 піхотна рота");
    expect(sections.find(section => section.id === "mission")?.people).toHaveLength(2);
    expect(sections.find(section => section.id === "platoon1")?.people.map(person => person.name)).toEqual(
      expect.arrayContaining([rows[0].name, rows[1].name]),
    );
    expect(sections.some(section => section.id === "out")).toBe(false);
    expect(
      new Set(
        sections.flatMap((section) => section.people.map((person) => person.sourceOrder)),
      ).size,
    ).toBe(rows.length);
    const wb = await XlsxPopulate.fromFileAsync("public/templates/pb-rota-bchs-morning-template.xlsx");
    writeRotaBchsMorningWorkbook(wb, { unitLabel: "3 піхотна рота", rows, staffCount: 113 });
    const saved = await XlsxPopulate.fromDataAsync(await wb.outputAsync());
    const sheet = saved.sheet(0);
    expect(sheet.cell("V23").value()).toBe(2);
    expect(sheet.cell("V25").value()).toBe(3);
    expect(sheet.cell("V26").value()).toBe(4);
    expect(sheet.cell("V27").value()).toBe(6);
    expect(sheet.cell("V7").value()).toBe(10);
    expect(sheet.cell("U24").value()).toBeUndefined();
    expect(sheet.cell("V24").value()).toBeUndefined();
    expect(sheet.column(1).width()).toBeGreaterThanOrEqual(6);
    expect(sheet.column(10).width()).toBeGreaterThanOrEqual(6);
    expect(sheet.column(19).width()).toBeGreaterThanOrEqual(6);
    expect(sheet.cell("S34").value()).toBe("На виконанні");
    expect(sheet.cell("S35").value()).toBe("№");
    expect(sheet.cell("S36").style("numberFormat")).toBe("General");
    expect(sheet.cell("S36").value()).toBe(1);
    expect(sheet.cell("S37").value()).toBe(2);
    expect(sheet.cell("Y36").value()).toBe("На виконання");
    expect(sheet.cell("U36").value()).toBe(rows[0].name);
    expect(sheet.cell("U37").value()).toBe(rows[1].name);
    expect(sheet.cell("T38").value()).toBe("Всього:");
    expect(sheet.cell("Z38").value()).toBe(sheet.cell("V23").value());
  });

  it("retains the mission table with a zero total when nobody is assigned", async () => {
    const wb = await XlsxPopulate.fromFileAsync("public/templates/pb-rota-bchs-morning-template.xlsx");
    writeRotaBchsMorningWorkbook(wb, { unitLabel: "3 піхотна рота", rows: [row()], staffCount: 113 });
    const saved = await XlsxPopulate.fromDataAsync(await wb.outputAsync());
    const sheet = saved.sheet(0);
    expect(sheet.cell("S34").value()).toBe("На виконанні");
    expect(sheet.cell("S35").value()).toBe("№");
    expect(sheet.cell("U35").value()).toBe("ПІБ");
    expect(sheet.cell("T36").value()).toBe("Всього:");
    expect(sheet.cell("Z36").value()).toBe(0);
    expect(sheet.cell("V23").value()).toBe(0);
  });

  it("counts support, BG and in-ranks from staff columns 22/23/21", () => {
    const rows = [
      row({
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "В строю",
          staff_22: "Забезпечення",
          staff_23: "",
        },
      }),
      row({
        name: "ДРУГИЙ Іван Іванович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "В строю",
          staff_22: "Забезпечення",
          staff_23: "БГ",
        },
      }),
      row({
        name: "ТРЕТІЙ Петро Петрович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "Відпустка",
          staff_22: "Забезпечення",
          staff_23: "БГ",
        },
      }),
      row({
        name: "ЧЕТВЕРТИЙ Олег Олегович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_3: "ж",
          staff_21: "В строю",
          staff_22: "Піхота",
          staff_23: "",
        },
      }),
    ];

    expect(computeBchsMorningSummaryFieldCounts(rows)).toEqual({
      management: 1,
      support: 2,
      platoonLeaders: 0,
      detached: 0,
      attached: 0,
      battleReady: 1,
      inRanks: 3,
    });
  });

  it("routes detached and attached into one section by status", () => {
    const rows = [
      row({
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "Відком. за межі ПБ",
        },
      }),
      row({
        name: "ДРУГИЙ Іван Іванович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_21: "Прикомандирований",
        },
      }),
    ];
    const sections = buildBchsMorningSections(rows, "3 піхотна рота");
    expect(sections.find((section) => section.id === "detached")?.people).toHaveLength(2);
    const counts = computeBchsMorningSummaryFieldCounts(rows);
    expect(counts.detached).toBe(1);
    expect(counts.attached).toBe(1);
  });

  it("routes management table by Взвод = ж, not by position title", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "КОМАНДИР РОТИ Іван Іванович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "ж",
            staff_5: "командир роти",
          },
        }),
        row({
          name: "СТРІЛЕЦЬ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "командир роти",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.find((section) => section.id === "management")?.people).toHaveLength(1);
    expect(sections.find((section) => section.id === "platoon1")?.people).toHaveLength(1);
  });

  it("routes platoon tables by Взвод column only, not Відділення", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "ДРУГИЙ ВЗВОД Іван Іванович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
            staff_4: "1 відділення",
            staff_5: "стрілець",
          },
        }),
        row({
          name: "БЕЗ ВЗВОДУ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "",
            staff_4: "1 відділення",
            staff_5: "стрілець",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.find((section) => section.id === "platoon2")?.people).toHaveLength(1);
    expect(sections.find((section) => section.id === "platoon1")?.people).toHaveLength(1);
  });

  it("keeps in-service mission people in their vzvod and in На виконанні", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "НА ВИКОНАННІ Перший",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_21: "В строю",
            staff_31: "На виконанні",
          },
        }),
        row({
          name: "НА ВИКОНАННІ Другий",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
            staff_21: "В строю",
            staff_31: "На виконання",
            staff_22: "Забезпечення",
          },
        }),
        row({
          name: "В ППД Третій",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_21: "В строю",
            staff_31: "ППД",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.find((section) => section.id === "mission")?.people).toHaveLength(2);
    expect(sections.find((section) => section.id === "platoon1")?.people.map((person) => person.name)).toEqual([
      "НА ВИКОНАННІ Перший",
      "В ППД Третій",
    ]);
    expect(sections.find((section) => section.id === "platoon2")?.people).toBeUndefined();
    expect(sections.find((section) => section.id === "support")?.people.map((person) => person.name)).toEqual([
      "НА ВИКОНАННІ Другий",
    ]);
    expect(
      new Set(
        sections.flatMap((section) => section.people.map((person) => person.sourceOrder)),
      ).size,
    ).toBe(3);
  });

  it("lists every Командир взводу in Взводні after Забезпечення", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "ВЗВОДНИЙ Перший",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "Командир взводу",
          },
        }),
        row({
          name: "ВЗВОДНИЙ Другий",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
            staff_5: "командир 2 піхотного взводу",
            staff_31: "На виконанні",
          },
        }),
        row({
          name: "ВІДПУСТКА Взводний",
          status: "LEAVE",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "3 піхотний взвод",
            staff_5: "Командир взводу",
            staff_21: "Відпустка",
          },
        }),
        row({
          name: "СТРІЛЕЦЬ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "стрілець",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.filter((section) => section.side === "left").map((section) => section.id)).toEqual([
      "platoonLeaders",
      "platoon1",
      "platoon2",
      "platoon3",
    ]);
    expect(sections.find((section) => section.id === "platoonLeaders")?.title).toBe("Взводні");
    expect(sections.find((section) => section.id === "platoonLeaders")?.people.map((person) => person.name)).toEqual([
      "ВЗВОДНИЙ Перший",
      "ВЗВОДНИЙ Другий",
      "ВІДПУСТКА Взводний",
    ]);
    expect(sections.find((section) => section.id === "platoon1")?.people.map((person) => person.name)).toEqual([
      "ВЗВОДНИЙ Перший",
      "СТРІЛЕЦЬ Петро Петрович",
    ]);
    expect(sections.find((section) => section.id === "platoon2")?.people).toHaveLength(1);
    expect(sections.find((section) => section.id === "leave")?.people).toHaveLength(1);
  });

  it("counts only in-service commanders in Взводні (наявність)", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/pb-rota-bchs-morning-template.xlsx",
    );
    writeRotaBchsMorningWorkbook(wb, {
      unitLabel: "3 піхотна рота",
      staffCount: 113,
      rows: [
        row({
          name: "ВЗВОДНИЙ Перший",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "Командир взводу",
          },
        }),
        row({
          name: "ВЗВОДНИЙ Другий",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
            staff_5: "Командир взводу",
            staff_21: "Відпустка",
          },
        }),
        row({
          name: "ВЗВОДНИЙ На виконанні",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "3 піхотний взвод",
            staff_5: "Командир взводу",
            staff_21: "В строю",
            staff_31: "На виконанні",
          },
        }),
        row({
          name: "СТРІЛЕЦЬ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_5: "стрілець",
          },
        }),
      ],
    });
    const sheet = wb.sheet("3ПР");
    expect(sheet.cell("U21").value()).toBe("Взводні (наявність)");
    expect(sheet.cell("V21").value()).toBe(2);
    expect(sheet.cell("V20").value()).toBe(0);
  });

  it("keeps Забезпечення out of vzvod tables, and Управління out of both", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "УПРАВЛІННЯ Забезпечення",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "ж",
            staff_5: "діловод",
            staff_22: "Забезпечення",
          },
        }),
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "стрілець",
            staff_22: "Забезпечення",
          },
        }),
        row({
          name: "МЕДИК Іван Іванович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
            staff_5: "бойовий медик",
            staff_22: "Забезпечення",
          },
        }),
        row({
          name: "СТРІЛЕЦЬ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
            staff_5: "стрілець",
            staff_22: "Піхота",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.find((section) => section.id === "management")?.people.map((person) => person.name)).toEqual([
      "УПРАВЛІННЯ Забезпечення",
    ]);
    expect(sections.find((section) => section.id === "support")?.people.map((person) => person.name)).toEqual([
      "БІЛИЦЬКИЙ Сергій Володимирович",
      "МЕДИК Іван Іванович",
    ]);
    expect(sections.find((section) => section.id === "platoon1")?.people.map((person) => person.name)).toEqual([
      "СТРІЛЕЦЬ Петро Петрович",
    ]);
    expect(sections.find((section) => section.id === "platoon2")?.people).toBeUndefined();
    expect(
      new Set(
        sections.flatMap((section) => section.people.map((person) => person.sourceOrder)),
      ).size,
    ).toBe(4);
  });

  it("routes support table rows without a vzvod by Тип В\\С", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "",
            staff_22: "Забезпечення",
          },
        }),
      ],
      "3 піхотна рота",
    );

    expect(sections.find((section) => section.id === "support")?.people).toHaveLength(1);
  });

  it("routes leave table only from status and location columns", () => {
    const sections = buildBchsMorningSections(
      [
        row({
          name: "ЩЕЛКУНОВ Денис Дмитрович",
          status: "LEAVE",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "В строю",
            staff_31: "На виконанні",
            staff_32: "Відпустка лікув. раніше",
          },
        }),
        row({
          name: "ПРИХОДКІН Гліб Сергійович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "Відпустка",
          },
        }),
        row({
          name: "МЕДИЧНА Лікувальна Відпустка",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "Лікування",
            staff_31: "Відпустка лікув.",
          },
        }),
      ],
      "1 піхотна рота",
    );

    expect(sections.find((section) => section.id === "leave")?.people.map((person) => person.name)).toEqual([
      "ПРИХОДКІН Гліб Сергійович",
      "МЕДИЧНА Лікувальна Відпустка",
    ]);
  });

  it("maps medical note to Лік.Пор / Лік.Хвор", () => {
    expect(
      resolveReportStatus(
        row({
          status: "MEDICAL",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "Лікування",
            staff_32: "по пораненню",
          },
        }),
      ),
    ).toBe("Лік.Пор");
    expect(
      resolveReportStatus(
        row({
          status: "MEDICAL",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_21: "Лікування",
            staff_32: "по хворобі",
          },
        }),
      ),
    ).toBe("Лік.Хвор");
  });
});
