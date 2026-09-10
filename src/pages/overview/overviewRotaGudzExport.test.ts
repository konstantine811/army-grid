import XlsxPopulate from "xlsx-populate";
import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  buildRotaGudzExportFileName,
  buildRotaGudzPersonRows,
  isPlatoonCommanderPosition,
  isRotaGudzExitPerson,
  isRotaGudzGuardPerson,
  isSectionCommanderPosition,
  extractRotaGudzBzvpDate,
  resolveRotaGudzNote,
  resolveRotaGudzCallsign,
  resolveRotaGudzPersonName,
  resolveRotaGudzRankLabel,
  resolveSelectedOverviewUnit,
  sanitizeRotaGudzWorkbookBuffer,
  writeRotaGudzReportWorkbook,
} from "./overviewRotaGudzExport";

const row = (
  overrides: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: "1",
  name: "БІЛИЦЬКИЙ Сергій Володимирович",
  unit: "3 піхотна рота",
  rank: "солдат",
  status: "ON_DUTY",
  statusLabel: "В строю",
  staffStatus: "В строю",
  staffStatusLabel: "В строю",
  staffSheetColumns: {
    staff_3: "1 піхотний взвод",
    staff_13: "солдат",
    staff_15: "НАЛІМ",
    staff_19: "3031803439",
    staff_26: "31.05.2026",
    staff_31: "На виконанні",
    staff_33: "вихід",
  },
  ...overrides,
});

describe("overviewRotaGudzExport", () => {
  it("maps exit note from direction", () => {
    expect(resolveRotaGudzNote(row())).toBe("вихід");
    expect(isRotaGudzExitPerson(buildRotaGudzPersonRows([row()])[0]!)).toBe(
      true,
    );
  });

  it("maps guard note", () => {
    const guard = row({
      staffSheetColumns: {
        ...row().staffSheetColumns,
        staff_31: "Охорона ППД",
        staff_33: "",
      },
    });
    expect(resolveRotaGudzNote(guard)).toBe("охорона");
    expect(isRotaGudzGuardPerson(buildRotaGudzPersonRows([guard])[0]!)).toBe(
      true,
    );
  });

  it("reads BZVP date or raw course column text", () => {
    expect(extractRotaGudzBzvpDate("31.05.2026")).toBe("31.05.2026");
    expect(extractRotaGudzBzvpDate("Закінчення курсу 15.06.2026")).toBe(
      "15.06.2026",
    );
    expect(extractRotaGudzBzvpDate("БЗВП +")).toBe("БЗВП +");
    expect(extractRotaGudzBzvpDate("має")).toBe("має");
    expect(extractRotaGudzBzvpDate("Закінчення курсу")).toBe(
      "Закінчення курсу",
    );

    const person = row({
      staffSheetColumns: {
        ...row().staffSheetColumns,
        staff_24: "99999",
        staff_25: "88888",
        staff_27: "77777",
        staff_26: "БЗВП +",
      },
    });
    expect(buildRotaGudzPersonRows([person])[0]?.bzvp).toBe("БЗВП +");
  });

  it("skips birth date values in course column", () => {
    const person = row({
      staffSheetColumns: {
        ...row().staffSheetColumns,
        staff_26: "25.01.1997 р.н.",
      },
    });
    expect(buildRotaGudzPersonRows([person])[0]?.bzvp).toBe("");
  });

  it("groups platoon label from staff column 3", () => {
    const people = buildRotaGudzPersonRows([row()]);
    expect(people[0]?.platoon).toBe("1 взвод");
  });

  it("requires exactly one selected unit", () => {
    expect(
      resolveSelectedOverviewUnit([
        { id: "unit", label: "Підрозділ", values: ["3 рота"] },
      ]),
    ).toBe("3 рота");
    expect(
      resolveSelectedOverviewUnit([
        { id: "unit", label: "Підрозділ", values: ["2 рота", "3 рота"] },
      ]),
    ).toBeNull();
  });

  it("builds file name like the reference sample", () => {
    expect(
      buildRotaGudzExportFileName(
        "3 РОТА",
        new Date("2026-08-10T12:00:00"),
      ),
    ).toBe("3 РОТА станом на 10.08..xlsx");
  });

  it("merges platoon title row across columns", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/gudz-rota-report-template.xlsx",
    );
    writeRotaGudzReportWorkbook(wb, {
      unitLabel: "3 РОТА",
      rows: [
        row({
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "1 піхотний взвод",
          },
        }),
        row({
          name: "ПЕТРОВ Петро Петрович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_3: "2 піхотний взвод",
          },
        }),
      ],
    });
    const sheet = wb.sheet("ОС 3 РОТА");
    let secondPlatoonHeaderRow = 0;
    for (let rowNumber = 8; rowNumber <= 40; rowNumber += 1) {
      if (sheet.cell(rowNumber, 1).value() === "2 взвод") {
        secondPlatoonHeaderRow = rowNumber;
        break;
      }
    }
    expect(secondPlatoonHeaderRow).toBeGreaterThan(0);
    expect(sheet.cell(secondPlatoonHeaderRow, 1).value()).toBe("2 взвод");
    expect(sheet.cell(secondPlatoonHeaderRow + 1, 1).value()).toBe("КОМАНДИР");
    expect(sheet.cell(secondPlatoonHeaderRow + 2, 1).value()).toBe(2);
    expect(sheet.cell(secondPlatoonHeaderRow + 2, 4).value()).toBe(
      "ПЕТРОВ Петро Петрович",
    );
  });

  it("survives save and reopen without column shift", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/gudz-rota-report-template.xlsx",
    );
    const mergePlan = writeRotaGudzReportWorkbook(wb, {
      unitLabel: "3 РОТА",
      rows: [
        row({
          name: "ТЕСТОВИЙ Тест Тестович",
          staffSheetColumns: {
            staff_3: "1 піхотний взвод",
            staff_13: "солдат",
            staff_15: "МИР",
            staff_19: "3015305252",
            staff_26: "31.05.2026",
            staff_31: "КСП",
          },
        }),
      ],
    });
    const buffer = await sanitizeRotaGudzWorkbookBuffer(
      await wb.outputAsync(),
      mergePlan,
    );
    const reopened = await XlsxPopulate.fromDataAsync(buffer);
    const sheet = reopened.sheet("ОС 3 РОТА");
    const dataRow = 10;
    expect(sheet.cell(dataRow, 4).value()).toBe("ТЕСТОВИЙ Тест Тестович");
    expect(sheet.cell(dataRow, 5).value()).toBe("МИР");
    expect(String(sheet.cell(dataRow, 6).value())).toBe("3015305252");
  });

  it("writes platoon rows with №, rank, name and callsign in separate columns", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/gudz-rota-report-template.xlsx",
    );
    writeRotaGudzReportWorkbook(wb, {
      unitLabel: "3 РОТА",
      rows: [
        row({
          name: "ТЕСТОВИЙ Тест Тестович",
          staffSheetColumns: {
            staff_3: "1 піхотний взвод",
            staff_13: "солдат",
            staff_15: "МИР",
            staff_19: "3015305252",
            staff_26: "46173",
            staff_31: "КСП Павлоград",
          },
        }),
      ],
    });
    const sheet = wb.sheet("ОС 3 РОТА");
    expect(sheet.cell(9, 1).value()).toBe("КОМАНДИР");
    const dataRow = 10;
    expect(sheet.cell(dataRow, 1).value()).toBe(1);
    expect(sheet.cell(dataRow, 2).value()).toBe(1);
    expect(sheet.cell(dataRow, 3).value()).toBe("солдат");
    expect(sheet.cell(dataRow, 4).value()).toBe("ТЕСТОВИЙ Тест Тестович");
    expect(sheet.cell(dataRow, 5).value()).toBe("МИР");
    expect(String(sheet.cell(dataRow, 6).value())).toBe("3015305252");
    expect(sheet.cell(dataRow, 8).value()).toBe("КСП Павлоград");
  });

  it("keeps row 26 aligned after platoon headers and many rows", async () => {
    const surnames = [
      "БІЛИЦЬКИЙ",
      "БАБЕНКО",
      "КОВАЛЕНКО",
      "ШЕВЧЕНКО",
      "МЕЛЬНИК",
      "КОVALENKO",
    ].map((value) => value.replace("KOVALENKO", "КОВАЛЬ"));
    const rows = Array.from({ length: 30 }, (_, index) => {
      const number = index + 1;
      const surname = surnames[(number - 1) % surnames.length]!;
      return row({
        id: String(number),
        name: `${surname} Іван Іванович`,
        staffSheetColumns: {
          staff_3: number <= 15 ? "1 піхотний взвод" : "2 піхотний взвод",
          staff_13: "солдат",
          staff_15: number === 26 ? "МИР" : `ПОЗ${number}`,
          staff_19: number === 26 ? "3015305252" : String(3_000_000_000 + number),
          staff_26: "46173",
          staff_31: number === 26 ? "КСП Павлоград" : "вихід",
        },
      });
    });
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/gudz-rota-report-template.xlsx",
    );
    writeRotaGudzReportWorkbook(wb, { unitLabel: "3 РОТА", rows });
    const sheet = wb.sheet("ОС 3 РОТА");
    let person26Row = 0;
    for (let rowNumber = 3; rowNumber <= 80; rowNumber += 1) {
      if (sheet.cell(rowNumber, 1).value() === 26) {
        person26Row = rowNumber;
        break;
      }
    }
    expect(person26Row).toBeGreaterThan(0);
    expect(sheet.cell(person26Row, 4).value()).toBe("БАБЕНКО Іван Іванович");
    expect(sheet.cell(person26Row, 5).value()).toBe("МИР");
    expect(String(sheet.cell(person26Row, 6).value())).toBe("3015305252");
    expect(sheet.cell(person26Row, 8).value()).toBe("КСП Павлоград");
  });

  it("uses staff column 14 when overview name is empty", () => {
    const people = buildRotaGudzPersonRows([
      row({
        name: "",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_14: "САВЧЕНКО Максим Олександрович",
        },
      }),
    ]);
    expect(people[0]?.name).toBe("САВЧЕНКО Максим Олександрович");
  });

  it("rejects row number in name and reads shifted PIB from column 13", () => {
    expect(
      resolveRotaGudzPersonName(
        row({
          name: "26",
          staffSheetColumns: {
            staff_14: "26",
            staff_13: "БАБЕНКО Дмитро Олександрович",
          },
        }),
      ),
    ).toBe("БАБЕНКО Дмитро Олександрович");

    expect(
      buildRotaGudzPersonRows([
        row({
          name: "26",
          staffSheetColumns: {
            staff_14: "26",
            staff_13: "БАБЕНКО Дмитро Олександрович",
            staff_15: "МАЛЕНЬКИЙ",
          },
        }),
      ]),
    ).toMatchObject([
      {
        name: "БАБЕНКО Дмитро Олександрович",
        callsign: "МАЛЕНЬКИЙ",
      },
    ]);
  });

  it("keeps compound rank titles like молодший сержант", () => {
    const people = buildRotaGudzPersonRows([
      row({
        name: "ЯЩЕНКО Юрій Дмитрович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_13: "молодший сержант",
          staff_15: "ФУДЖІ",
        },
      }),
      row({
        name: "ГАПОН Андрій Вікторович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_13: "старший солдат",
          staff_15: "КОК",
        },
      }),
      row({
        name: "БАРІНОВ Денис Валерійович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_13: "молодший лейтенант",
          staff_15: "СКОВОРОДА",
        },
      }),
    ]);
    expect(people[0]?.rank).toBe("молодший сержант");
    expect(people[1]?.rank).toBe("старший солдат");
    expect(people[2]?.rank).toBe("молодший лейтенант");
    expect(
      buildRotaGudzPersonRows([
        row({
          name: "ГАЙДАР Дмитро Вікторович",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_13: "рядовий",
            staff_15: "ФАРА",
          },
        }),
      ])[0]?.rank,
    ).toBe("рядовий");
  });

  it("does not treat callsign as rank when staff column 13 is shifted", () => {
    const people = buildRotaGudzPersonRows([
      row({
        name: "ТРИНАДЦЯТИЙ Іван Іванович",
        staffSheetColumns: {
          ...row().staffSheetColumns,
          staff_13: "ТРИНАДЦЯТИЙ",
          staff_15: "ТРИНАДЦЯТИЙ",
        },
      }),
    ]);
    expect(people[0]?.rank).toBe("солдат");
    expect(people[0]?.callsign).toBe("ТРИНАДЦЯТИЙ");
  });

  it("maps shifted staff row number + callsign to correct export columns", () => {
    const source = row({
      name: "КРАВЕЦЬ Олег Іванович",
      rank: "КРОХА",
      staffSheetColumns: {
        ...row().staffSheetColumns,
        staff_12: "солдат",
        staff_14: "43",
        staff_13: "КРОХА",
        staff_15: "",
        staff_19: "3586704134",
        staff_26: "11.06.2026",
        staff_31: "шпиталь",
        staff_33: "",
      },
    });
    expect(resolveRotaGudzCallsign(source)).toBe("КРОХА");
    expect(
      buildRotaGudzPersonRows([source])[0],
    ).toMatchObject({
      rank: "солдат",
      name: "КРАВЕЦЬ Олег Іванович",
      callsign: "КРОХА",
      ipn: "3586704134",
      bzvp: "11.06.2026",
      note: "шпиталь",
    });
  });

  it("strips phantom merges on KROKHA (R54) and BIMA (R35) data rows", async () => {
    const platoonOne = Array.from({ length: 22 }, (_, index) => {
      const number = index + 1;
      return row({
        id: `p1-${number}`,
        name:
          number === 21
            ? "КРОШКА Іван Григорович"
            : `${["БІЛИЦЬКИЙ", "ПЕТРОВ", "КОVAL"][index % 3]!.replace("KOVAL", "КОВАЛЬ")} Іван Іванович`,
        staffSheetColumns: {
          staff_3: "1 піхотний взвод",
          staff_13: "солдат",
          staff_15: number === 21 ? "КРОХА" : `ПОЗ${number}`,
          staff_19: number === 21 ? "3586704134" : String(3_000_000_000 + number),
          staff_26: "11.06.2026",
          staff_31: number === 21 ? "шпиталь" : "вихід",
        },
      });
    });
    const platoonTwo = [
      row({
        id: "p2-1",
        name: "САВЧЕНКО Максим Олександрович",
        staffSheetColumns: {
          staff_3: "2 піхотний взвод",
          staff_5: "Командир взводу",
          staff_7: "командир 2 піхотного взводу",
          staff_13: "молодший лейтенант",
          staff_15: "САВА",
          staff_31: "ТРАНЗИТЕР",
        },
      }),
      row({
        id: "p2-2",
        name: "МАЦЕНКО Дмитро Костянтинович",
        staffSheetColumns: {
          staff_3: "2 піхотний взвод",
          staff_5: "Командир відділення",
          staff_7: "командир відділення",
          staff_13: "молодший сержант",
          staff_15: "БІМА",
          staff_31: "Полігон Д",
        },
      }),
    ];
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/gudz-rota-report-template.xlsx",
    );
    const mergePlan = writeRotaGudzReportWorkbook(wb, {
      unitLabel: "3 РОТА",
      rows: [...platoonOne, ...platoonTwo],
    });
    const sanitized = await sanitizeRotaGudzWorkbookBuffer(
      await wb.outputAsync(),
      mergePlan,
    );
    const zip = await import("jszip").then((m) => m.default.loadAsync(sanitized));
    const xml = await zip.file("xl/worksheets/sheet1.xml")!.async("string");
    expect(xml).not.toMatch(/A54:D54/);
    expect(xml).not.toMatch(/A35:D35/);

    const reopened = await XlsxPopulate.fromDataAsync(sanitized);
    const sheet = reopened.sheet("ОС 3 РОТА");

    let krokhaRow = 0;
    for (let rowNumber = 10; rowNumber <= 60; rowNumber += 1) {
      if (sheet.cell(rowNumber, 5).value() === "КРОХА") {
        krokhaRow = rowNumber;
        break;
      }
    }
    expect(krokhaRow).toBeGreaterThan(0);
    expect(sheet.cell(krokhaRow, 3).value()).toBe("солдат");
    expect(sheet.cell(krokhaRow, 4).value()).toBe("КРОШКА Іван Григорович");
    expect(String(sheet.cell(krokhaRow, 6).value())).toBe("3586704134");

    expect(sheet.cell(35, 3).value()).toBe("Командир відділення");
    expect(sheet.cell(35, 4).value()).toBe("МАЦЕНКО Дмитро Костянтинович");
    expect(sheet.cell(35, 5).value()).toBe("БІМА");
    expect(sheet.cell(35, 8).value()).toBe("Полігон Д");
    expect(sheet.cell(35, 4).style("bold")).toBe(false);
    expect(sheet.cell(krokhaRow, 4).style("bold")).toBe(false);
  });

  it("writes shifted staff rows into correct workbook columns", async () => {
    const wb = await XlsxPopulate.fromFileAsync(
      "public/templates/gudz-rota-report-template.xlsx",
    );
    writeRotaGudzReportWorkbook(wb, {
      unitLabel: "3 РОТА",
      rows: [
        row({
          name: "КРАВЕЦЬ Олег Іванович",
          rank: "КРОХА",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_12: "солдат",
            staff_14: "43",
            staff_13: "КРОХА",
            staff_15: "",
            staff_19: "3586704134",
            staff_26: "11.06.2026",
            staff_31: "шпиталь",
            staff_33: "",
          },
        }),
      ],
    });
    const sheet = wb.sheet("ОС 3 РОТА");
    expect(sheet.cell(9, 1).value()).toBe("КОМАНДИР");
    const dataRow = 10;
    expect(sheet.cell(dataRow, 3).value()).toBe("солдат");
    expect(sheet.cell(dataRow, 4).value()).toBe("КРАВЕЦЬ Олег Іванович");
    expect(sheet.cell(dataRow, 5).value()).toBe("КРОХА");
    expect(String(sheet.cell(dataRow, 6).value())).toBe("3586704134");
    expect(sheet.cell(dataRow, 7).value()).toBe("11.06.2026");
    expect(sheet.cell(dataRow, 8).value()).toBe("шпиталь");
  });

  it("strips embedded callsign from PIB", () => {
    expect(
      resolveRotaGudzPersonName(
        row({
          name: "ГРИЩЕНКО Олександр Сергійович (Кент)",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_15: "Кент",
          },
        }),
      ),
    ).toBe("ГРИЩЕНКО Олександр Сергійович");
  });

  it("rejects location-like text as PIB", () => {
    expect(
      buildRotaGudzPersonRows([
        row({
          name: "Полігон Д",
          staffSheetColumns: {
            ...row().staffSheetColumns,
            staff_15: "БІМА",
            staff_31: "Полігон Д",
          },
        }),
      ]),
    ).toEqual([]);
  });

  it("drops rows without a valid PIB", () => {
    expect(
      buildRotaGudzPersonRows([
        row({
          name: "26",
          staffSheetColumns: {
            staff_14: "26",
            staff_13: "",
          },
        }),
      ]),
    ).toEqual([]);
  });

  it("uses leadership labels instead of rank in platoon rows", () => {
    expect(
      isPlatoonCommanderPosition("Командир 1 піхотного взводу"),
    ).toBe(true);
    expect(isSectionCommanderPosition("Командир відділення")).toBe(true);
    expect(
      resolveRotaGudzRankLabel("Командир відділення", "солдат"),
    ).toBe("Командир відділення");
    expect(
      resolveRotaGudzRankLabel("Командир 2 піхотного взводу", "лейтенант"),
    ).toBe("Командир взводу");

    const platoonCommander = row({
      name: "ІВАНОВ Іван Іванович",
      staffSheetColumns: {
        ...row().staffSheetColumns,
        staff_7: "Командир 1 піхотного взводу",
      },
    });
    const sectionCommander = row({
      name: "ПЕТРОВ Петро Петрович",
      staffSheetColumns: {
        ...row().staffSheetColumns,
        staff_5: "Командир відділення",
      },
    });
    const people = buildRotaGudzPersonRows([
      sectionCommander,
      platoonCommander,
      row({ name: "СІДОРОВ Сідор Сідорович" }),
    ]);
    expect(
      people.find((person) => person.name.startsWith("ІВАНОВ"))?.rank,
    ).toBe("Командир взводу");
    expect(
      people.find((person) => person.name.startsWith("ПЕТРОВ"))?.rank,
    ).toBe("Командир відділення");
    expect(
      people.find((person) => person.name.startsWith("СІДОРОВ"))?.rank,
    ).toBe("солдат");
  });
});
