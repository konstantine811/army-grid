import { describe, expect, it } from "vitest";
import { isBchsWoundedByExcelNote } from "../bchs/bchsCalc";
import {
  buildImportantOverviewExportFileName,
  buildImportantOverviewSummary,
  buildImportantOverviewSummarySheetData,
  buildImportantOverviewExportSheetData,
  buildImportantOverviewExportSheetOptions,
  buildImportantOverviewExportSheets,
  buildImportantOverviewExportTitle,
  buildOverviewExportFilterHeaderText,
  buildOverviewExportSheetData,
  buildOverviewExportSheetOptions,
  formatOverviewExportFilterValue,
} from "./overviewExport";

describe("overviewExport", () => {
  it("formats default filter labels", () => {
    expect(formatOverviewExportFilterValue("unit", "ALL")).toBe("Усі підрозділи");
    expect(formatOverviewExportFilterValue("status", "ALL")).toBe("Усі статуси");
  });

  it("builds filter header text for unit and status", () => {
    expect(
      buildOverviewExportFilterHeaderText({
        unit: "2 піхотна рота",
        status: "В строю",
      }),
    ).toBe("Підрозділ: 2 піхотна рота · Статус: В строю");
  });

  it("writes active table filters into the export header", () => {
    const activeFilters = [
      { id: "unit", label: "Підрозділ", values: ["2 піхотна рота"] },
      { id: "status", label: "Статус", values: ["Лікування"] },
    ];

    expect(
      buildOverviewExportFilterHeaderText({ activeFilters }),
    ).toBe("Підрозділ: 2 піхотна рота · Статус: Лікування");
    expect(
      buildImportantOverviewExportTitle({ activeFilters }),
    ).toBe("Підрозділ: 2 піхотна рота · Статус: Лікування");
  });

  it("prepends filter row to exported sheet", () => {
    const sheet = buildOverviewExportSheetData(
      {
        columns: [
          { id: "name", label: "ПІБ", value: (row) => row.name },
          { id: "unit", label: "Підрозділ", value: (row) => row.unit },
        ],
        rows: [{ id: "1", name: "Іванов", unit: "2 піхотна рота" } as never],
      },
      { unit: "2 піхотна рота", status: "ALL" },
    );

    expect(sheet).toHaveLength(3);
    expect(sheet[0]?.[0]).toMatchObject({
      value: "Підрозділ: 2 піхотна рота · Статус: Усі статуси",
      columnSpan: 2,
    });
    expect(sheet[1]?.map((cell) => cell && "value" in cell && cell.value)).toEqual([
      "ПІБ",
      "Підрозділ",
    ]);
    expect(sheet[0]?.[0]).toMatchObject({
      backgroundColor: "#234F3E",
      textColor: "#FFFFFF",
      height: 28,
    });
    expect(sheet[1]?.[0]).toMatchObject({
      backgroundColor: "#39735C",
      textColor: "#FFFFFF",
      height: 34,
    });
    expect(sheet[2]?.[0]).toMatchObject({
      borderStyle: "thin",
      height: 30,
    });
  });

  it("sets readable widths and freezes headers", () => {
    const options = buildOverviewExportSheetOptions({
      columns: [
        { id: "person", label: "ПІБ", value: (row) => row.name },
        {
          id: "positionTitle",
          label: "Повна посада",
          value: (row) => row.positionTitle ?? "",
        },
        { id: "fighterTotalDays", label: "Днів", value: () => "5" },
      ],
      rows: [],
    });

    expect(options).toMatchObject({
      stickyRowsCount: 2,
      stickyColumnsCount: 1,
      showGridLines: false,
      columns: [{ width: 30 }, { width: 42 }, { width: 10 }],
    });
    expect(options).not.toHaveProperty("zoomScale");
  });

  it("exports the fixed important columns independently of table visibility", () => {
    const sheet = buildImportantOverviewExportSheetData(
      [
        {
          name: "Іванов Іван",
          staffSheetColumns: {
            staff_5: "Кулеметник",
            staff_15: "СОКІЛ",
            staff_31: "На виконанні",
            staff_32: "Прибув 09.08.2026",
          },
        } as never,
      ],
      { unit: "2 піхотна рота", status: "На виконанні" },
    );

    expect(buildImportantOverviewExportTitle({
      unit: "2 піхотна рота",
      status: "На виконанні",
    })).toBe("2 піхотна рота — На виконанні");
    expect(sheet[1]?.map((cell) => cell && "value" in cell && cell.value)).toEqual([
      "Посада",
      "ШПК факт",
      "ПІБ",
      "Позивний",
      "Тип В\\С",
      "Статус БГ",
      "Відрядження (БРЕЗ)",
      "Обмеження",
      "Місце перебування",
      "Примітки",
      "Напрямок",
    ]);
    expect(sheet[2]?.[0]).toMatchObject({ value: "Кулеметник" });
    expect(sheet[2]?.[2]).toMatchObject({ value: "Іванов Іван" });
    expect(sheet[2]?.[8]).toMatchObject({ value: "На виконанні" });
    expect(buildImportantOverviewExportSheetOptions()).toMatchObject({
      stickyRowsCount: 2,
      stickyColumnsCount: 3,
      orientation: "landscape",
    });
  });

  it("splits multiple selected statuses into separate workbook tabs", () => {
    const sheets = buildImportantOverviewExportSheets(
      [
        {
          name: "Іванов",
          status: "MEDICAL",
          statusLabel: "На лікуванні",
        },
        {
          name: "Петренко",
          status: "LEAVE",
          statusLabel: "Відпустка",
        },
      ] as never[],
      {
        activeFilters: [
          { id: "unit", label: "Підрозділ", values: ["2 рота"] },
          {
            id: "status",
            label: "Статус",
            values: ["Лікування", "Відпустка"],
          },
        ],
      },
      [
        {
          name: "Іванов",
          unit: "2 рота",
          status: "MEDICAL",
          statusLabel: "На лікуванні",
        },
        {
          name: "Петренко",
          unit: "2 рота",
          status: "LEAVE",
          statusLabel: "Відпустка",
        },
        {
          name: "Слюсар",
          unit: "2 рота",
          status: "MEDICAL",
          statusLabel: "Лікування після поранення",
        },
        {
          name: "Сидоренко",
          unit: "2 рота",
          status: "ON_DUTY",
          statusLabel: "В строю",
          staffSheetColumns: { staff_31: "На виконанні" },
        },
        {
          name: "Коваленко",
          unit: "3 рота",
          status: "ON_DUTY",
          statusLabel: "В строю",
          staffSheetColumns: { staff_31: "На виконанні" },
        },
      ] as never[],
    );

    expect(sheets.map((sheet) => sheet.sheet)).toEqual([
      "Лікування",
      "Відпустка",
      "На виконанні",
    ]);
    expect(sheets[0]?.data).toHaveLength(4);
    expect(sheets[1]?.data).toHaveLength(3);
    expect(sheets[0]?.data[0]?.[0]).toMatchObject({
      value: "Підрозділ: 2 рота · Статус: Лікування",
    });
    expect(sheets[1]?.data[2]?.[2]).toMatchObject({ value: "Петренко" });
    expect(sheets[2]?.data[2]?.[2]).toMatchObject({ value: "Сидоренко" });
    expect(sheets[2]?.data).toHaveLength(3);
    expect(sheets[2]?.data[0]?.[0]).toMatchObject({
      value: "Підрозділ: 2 рота · Місце перебування: На виконанні",
    });
  });

  it("names the important workbook from unit, statuses and export date", () => {
    expect(
      buildImportantOverviewExportFileName(
        {
          activeFilters: [
            { id: "unit", label: "Підрозділ", values: ["2 рота"] },
            {
              id: "status",
              label: "Статус",
              values: ["Лікування", "Відпустка"],
            },
          ],
        },
        new Date(2026, 8, 3),
      ),
    ).toBe(
      "1ПБ-2 рота-Статуси-Лікування, Відпустка-На виконанні-03.09.2026.xlsx",
    );
  });

  it("builds a company summary from General List rows", () => {
    const filters = {
      activeFilters: [
        { id: "unit", label: "Підрозділ", values: ["2 рота"] },
      ],
    };
    const rosterRows = [
      {
        __dbRowId: "1",
        column_2: "2 рота",
        column_14: "Іванов Іван",
        column_21: "Лікування",
        column_22: "Упр. взводу",
        column_31: "На виконанні",
      },
      {
        __dbRowId: "2",
        column_2: "2 рота",
        column_14: "Петренко Петро",
        column_21: "Відпустка",
        column_22: "Забезпечення",
        column_31: "В наявності",
      },
      {
        __dbRowId: "3",
        column_2: "2 рота",
        column_14: "",
      },
      {
        __dbRowId: "4",
        column_2: "3 рота",
        column_14: "Сидоренко Сидір",
      },
    ] as never[];

    const summary = buildImportantOverviewSummary(rosterRows, filters);
    expect(summary).toMatchObject({
      unitLabel: "2 рота",
      staff: 3,
      listed: 2,
      treatment: 1,
      treatmentWounded: 0,
      treatmentIllness: 1,
      vacation: 1,
      absent: 2,
      management: 1,
      support: 1,
      onExit: 1,
      available: 0,
      inRanks: 0,
    });
    expect(buildImportantOverviewSummarySheetData(summary)).toHaveLength(23);
    expect(
      buildImportantOverviewSummarySheetData(summary).some(
        (row) => row[0] && "value" in row[0] && row[0].value === "300",
      ),
    ).toBe(false);

    const sheets = buildImportantOverviewExportSheets(
      [],
      filters,
      [],
      rosterRows,
    );
    expect(sheets[0]?.sheet).toBe("Підрахунок");
  });

  it("counts В наявності from В строю status, not the stay-place text", () => {
    const summary = buildImportantOverviewSummary(
      [
        {
          __dbRowId: "1",
          column_2: "2 рота",
          column_14: "Іванов Іван Іванович",
          column_21: "В строю",
          column_31: "Полігон Д",
        },
        {
          __dbRowId: "2",
          column_2: "2 рота",
          column_14: "Петренко Петро Петрович",
          column_21: "Лікування",
          column_31: "В наявності",
        },
      ] as never[],
      {
        activeFilters: [
          { id: "unit", label: "Підрозділ", values: ["2 рота"] },
        ],
      },
    );

    expect(summary.available).toBe(1);
  });

  it("splits Лікування by a universal wounding note regex", () => {
    const filters = {
      activeFilters: [
        { id: "unit", label: "Підрозділ", values: ["2 рота"] },
      ],
    };
    const rosterRows = [
      {
        __dbRowId: "1",
        column_2: "2 рота",
        column_14: "Іванов Іван Іванович",
        column_21: "Лікування",
        column_32: "ПРИБУВ 14.06.26 Полі Б\nпо пораненню",
      },
      {
        __dbRowId: "2",
        column_2: "2 рота",
        column_14: "Петренко Петро Петрович",
        column_21: "Лікування",
        column_32: "повернули після СЗЧ 24.05.26\nпо пораненню",
      },
      {
        __dbRowId: "3",
        column_2: "2 рота",
        column_14: "Сидоренко Сидір Сидорович",
        column_21: "Лікування",
        column_32: "по хворобі\nприбув після СЗЧ 10.08.2026",
      },
      {
        __dbRowId: "4",
        column_2: "2 рота",
        column_14: "Коваленко Костянтин",
        column_21: "Лікування",
        column_32: "на навчання по зв'язку з Зимою\n20.02.26\nна лікування 10.08.2026",
      },
      {
        __dbRowId: "5",
        column_2: "2 рота",
        column_14: "Мельник Максим",
        column_21: "Лікування",
        column_32: "поранення\nВодій квадроциклу",
      },
      {
        __dbRowId: "6",
        column_2: "3 рота",
        column_14: "Інший Іван",
        column_21: "Лікування",
        column_32: "по пораненню",
      },
    ] as never[];

    const summary = buildImportantOverviewSummary(rosterRows, filters);
    expect(summary.treatment).toBe(5);
    expect(summary.treatmentWounded).toBe(3);
    expect(summary.treatmentIllness).toBe(2);
  });

  it("also sees поранення from an open 1PB archive period for a treatment person", () => {
    const filters = {
      activeFilters: [
        { id: "unit", label: "Підрозділ", values: ["2 рота"] },
      ],
    };
    const rosterRows = [
      {
        __dbRowId: "1",
        column_2: "2 рота",
        column_14: "Іванов Іван Іванович",
        column_21: "Лікування",
      },
      {
        __dbRowId: "2",
        column_2: "2 рота",
        column_14: "Петренко Петро Петрович",
        column_21: "Лікування",
      },
      {
        __dbRowId: "3",
        column_2: "3 рота",
        column_14: "Сидоренко Сидір Сидорович",
        column_21: "Лікування",
      },
    ] as never[];
    const archivePeriods = [
      {
        fullName: "ІВАНОВ Іван Іванович",
        absenceType: "ЛІКУВАННЯ",
        place: "по пораненню",
        returnDate: "",
      },
      {
        fullName: "Сидоренко Сидір Сидорович",
        absenceType: "300",
        place: "",
        returnDate: "",
      },
      {
        fullName: "Іванов Іван Іванович",
        absenceType: "300",
        place: "",
        returnDate: "01.09.2026",
      },
    ] as never[];

    const summary = buildImportantOverviewSummary(
      rosterRows,
      filters,
      archivePeriods,
    );

    expect(summary.treatmentWounded).toBe(1);
    expect(summary.treatmentIllness).toBe(1);
  });

  it("matches wounding notes in any line of the cell", () => {
    const wounded = [
      "по пораненню",
      "ПРИБУВ 14.06.26 Полі Б\nпо пораненню",
      "повернули після СЗЧ 24.05.26\nпо пораненню",
      "поранення\nВодій квадроциклу",
      "лікування після поранення",
      "300",
    ];
    const notWounded = [
      "по хворобі\nприбув після СЗЧ 10.08.2026",
      "на лікування 10.08.2026",
      "ТРАНЗИТЕР",
      "НРК/ мультикам",
      "вибув до Бобра з МР 10.12.2025",
    ];
    for (const note of wounded) {
      expect(isBchsWoundedByExcelNote(note), note).toBe(true);
    }
    for (const note of notWounded) {
      expect(isBchsWoundedByExcelNote(note), note).toBe(false);
    }
  });
});
