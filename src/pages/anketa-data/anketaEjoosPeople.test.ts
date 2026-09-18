import { describe, expect, it } from "vitest";
import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  collectEjoosOosAndExcludedPeople,
  filterAnketaRowsByEjoosSource,
  reconcileAnketaRowsWithEjoos,
  type EjoosAnketaCandidate,
} from "./anketaEjoosPeople";
import { createEmptyAnketaRow, type AnketaRow } from "./anketaSheet";

const anketaRow = (
  rowNumber: number,
  values: Partial<AnketaRow>,
): AnketaRow => Object.assign(createEmptyAnketaRow(rowNumber), values);

const candidate = (
  values: Partial<EjoosAnketaCandidate>,
): EjoosAnketaCandidate => ({
  personId: "",
  fullName: "",
  rank: "",
  positionIndex: "",
  source: "oos",
  ...values,
});

const sheet = (
  sheetIndex: number,
  sheetName: string,
  rows: CellValue[][],
): ExcelSheetSnapshot => ({
  sheetIndex,
  sheetName,
  rawRows: rows,
  headerRows: [],
  rows: [],
  columnCount: 32,
  columnIndexes: [],
  dataStartRow: 6,
});

describe("collectEjoosOosAndExcludedPeople", () => {
  it("uses only OOS and Excluded, removes service rows, and deduplicates a person", () => {
    const headerRows = Array.from({ length: 5 }, () => [] as CellValue[]);
    const oos = sheet(0, "2. ООС", [
      ...headerRows,
      ["солдат", "ІВАНЕНКО Іван Іванович", "101", "2103001"],
      ["", "ВИБУВ У РОЗПОРЯДЖЕННЯ КОМАНДИРА", "", ""],
    ]);
    const excluded = sheet(1, "3. Виключені", [
      ...headerRows,
      ["", "ІВАНЕНКО Іван Іванович", "101", ""],
      ["сержант", "ПЕТРЕНКО Петро Петрович", "202", "2103002"],
      ["старший солдат", "", "", "", "2103003", "", "СИДОРЕНКО Сидір Сидорович", "303"],
    ]);
    const workbook = {
      file: new File([], "ejoos.xlsx"),
      fileName: "ejoos.xlsx",
      sheetName: oos.sheetName,
      headerRows: [],
      rows: [],
      columnCount: 32,
      columnIndexes: [],
      dataStartRow: 6,
      sheets: [oos, excluded],
    } satisfies ExcelWorkbookSnapshot;

    expect(
      collectEjoosOosAndExcludedPeople(workbook).map(
        ({ personId, fullName, rank, positionIndex, source }) => ({
          personId,
          fullName,
          rank,
          positionIndex,
          source,
        }),
      ),
    ).toEqual([
      candidate({
        personId: "101",
        fullName: "ІВАНЕНКО Іван Іванович",
        rank: "солдат",
        positionIndex: "2103001",
      }),
      candidate({
        personId: "202",
        fullName: "ПЕТРЕНКО Петро Петрович",
        rank: "сержант",
        positionIndex: "2103002",
        source: "excluded",
      }),
      candidate({
        personId: "303",
        fullName: "СИДОРЕНКО Сидір Сидорович",
        rank: "старший солдат",
        positionIndex: "2103003",
        source: "excluded",
      }),
    ]);
  });

  it("does not merge different OOS people that share the same ID", () => {
    const headerRows = Array.from({ length: 5 }, () => [] as CellValue[]);
    const oos = sheet(0, "2. ООС", [
      ...headerRows,
      ["солдат", "ІВАНЕНКО Іван Іванович", "789", "2103001"],
      ["солдат", "ДМИТРІЄВ Андрій Сергійович", "789", "2103002"],
    ]);
    const workbook = {
      file: new File([], "ejoos.xlsx"),
      fileName: "ejoos.xlsx",
      sheetName: oos.sheetName,
      headerRows: [],
      rows: [],
      columnCount: 32,
      columnIndexes: [],
      dataStartRow: 6,
      sheets: [oos],
    } satisfies ExcelWorkbookSnapshot;

    const people = collectEjoosOosAndExcludedPeople(workbook);

    expect(people).toHaveLength(2);
    expect(people.map((person) => person.fullName)).toEqual([
      "ІВАНЕНКО Іван Іванович",
      "ДМИТРІЄВ Андрій Сергійович",
    ]);
    expect(people.every((person) => person.source === "oos")).toBe(true);
  });
});

describe("reconcileAnketaRowsWithEjoos", () => {
  it("keeps non-empty Anketa fields, fills gaps, adds missing, and removes outsiders", () => {
    const existing = anketaRow(10, {
      externalId: "101",
      fullName: "ІВАНЕНКО Іван Іванович",
      rank: "",
      positionIndex: "анкета-індекс",
      rnokpp: "1234567890",
    });
    const outsider = anketaRow(20, {
      externalId: "999",
      fullName: "СТОРОННІЙ Степан Степанович",
      additionalInfo: "не має залишитися",
    });

    const result = reconcileAnketaRowsWithEjoos(
      [existing, outsider],
      [
        candidate({
          personId: "101",
          fullName: "ІВАНЕНКО Іван Іванович",
          rank: "сержант",
          positionIndex: "2103001",
        }),
        candidate({
          personId: "202",
          fullName: "ПЕТРЕНКО Петро Петрович",
          rank: "солдат",
          positionIndex: "2103002",
          source: "excluded",
        }),
      ],
    );

    expect(result.rows).toHaveLength(2);
    expect(result.rows[0]).toMatchObject({
      __rowNumber: 10,
      externalId: "101",
      rnokpp: "1234567890",
      rank: "сержант",
      positionIndex: "анкета-індекс",
      __ejoosSource: "oos",
    });
    expect(result.rows[1]).toMatchObject({
      __rowNumber: 21,
      externalId: "202",
      fullName: "ПЕТРЕНКО Петро Петрович",
      rank: "солдат",
      positionIndex: "2103002",
      __ejoosSource: "excluded",
    });
    expect(result.report).toMatchObject({
      ejoosPeople: 2,
      merged: 1,
      added: 1,
      removed: 1,
    });
  });

  it("matches by unique name when one side has no ID", () => {
    const current = anketaRow(7, {
      fullName: "КОВАЛЬ Андрій Олегович",
      birthDate: "01.01.1990",
    });
    const result = reconcileAnketaRowsWithEjoos(
      [current],
      [
        candidate({
          personId: "303",
          fullName: "КОВАЛЬ Андрій Олегович",
          rank: "солдат",
        }),
      ],
    );
    expect(result.rows[0]).toMatchObject({
      __rowNumber: 7,
      externalId: "303",
      birthDate: "01.01.1990",
    });
  });

  it("does not merge identical names with different IDs", () => {
    const current = anketaRow(4, {
      externalId: "111",
      fullName: "ТЕСТОВИЙ Іван Іванович",
      rnokpp: "1111111111",
    });
    const result = reconcileAnketaRowsWithEjoos(
      [current],
      [
        candidate({
          personId: "222",
          fullName: "ТЕСТОВИЙ Іван Іванович",
          rank: "солдат",
        }),
      ],
    );
    expect(result.report.conflicts).toHaveLength(1);
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]).toMatchObject({
      __rowNumber: 5,
      externalId: "222",
      rnokpp: "",
    });
    expect(result.report.removed).toBe(1);
  });

  it("reports an ID/name conflict and still adds the OOS person", () => {
    const current = anketaRow(8, {
      externalId: "45867",
      fullName: "ОРЛОВ Дмитро Олександрович",
      rnokpp: "4586700000",
    });
    const result = reconcileAnketaRowsWithEjoos(
      [current],
      [
        candidate({
          personId: "45867",
          fullName: "БОЛЮБАХА Кирило Олександрович",
          rank: "солдат",
        }),
      ],
    );
    expect(result.report.conflicts).toHaveLength(1);
    expect(result.report.added).toBe(1);
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]).toMatchObject({
      fullName: "БОЛЮБАХА Кирило Олександрович",
      externalId: "45867",
      rank: "солдат",
    });
  });

  it("treats ЗАКАЛЮЖНИЙ / ЗАКАЛЮЖНІЙ as the same person", () => {
    const current = anketaRow(12, {
      externalId: "771",
      fullName: "ЗАКАЛЮЖНІЙ Іван Олегович",
      education: "середня",
    });
    const result = reconcileAnketaRowsWithEjoos(
      [current],
      [
        candidate({
          personId: "771",
          fullName: "ЗАКАЛЮЖНИЙ Іван Олегович",
          rank: "солдат",
        }),
      ],
    );
    expect(result.report.conflicts).toHaveLength(0);
    expect(result.report.added).toBe(0);
    expect(result.rows).toHaveLength(1);
    expect(result.rows[0]).toMatchObject({
      __rowNumber: 12,
      fullName: "ЗАКАЛЮЖНІЙ Іван Олегович",
      education: "середня",
      rank: "солдат",
    });
  });

  it("always restores sex from the patronymic during reconcile", () => {
    const result = reconcileAnketaRowsWithEjoos(
      [
        anketaRow(30, {
          externalId: "501",
          fullName: "МИШУК Маргарита Анатоліївна",
          sex: "-",
        }),
        anketaRow(31, {
          externalId: "502",
          fullName: "КІЯНЕНКО Андрій Олександрович",
          sex: "",
        }),
      ],
      [
        candidate({
          personId: "501",
          fullName: "МИШУК Маргарита Анатоліївна",
        }),
        candidate({
          personId: "502",
          fullName: "КІЯНЕНКО Андрій Олександрович",
        }),
      ],
    );

    expect(result.rows.map((row) => row.sex)).toEqual(["Ж", "Ч"]);
  });
});

describe("filterAnketaRowsByEjoosSource", () => {
  it("puts unclassified rows into OOS and keeps excluded separate", () => {
    const oos = anketaRow(2, { fullName: "ООС", __ejoosSource: "oos" });
    const excluded = anketaRow(3, {
      fullName: "ВИКЛ",
      __ejoosSource: "excluded",
    });
    const unknown = anketaRow(4, { fullName: "ЩЕ НЕ ЗВІРЕНО" });

    expect(
      filterAnketaRowsByEjoosSource([oos, excluded, unknown], "oos").map(
        (row) => row.fullName,
      ),
    ).toEqual(["ООС", "ЩЕ НЕ ЗВІРЕНО"]);
    expect(
      filterAnketaRowsByEjoosSource([oos, excluded, unknown], "excluded").map(
        (row) => row.fullName,
      ),
    ).toEqual(["ВИКЛ"]);
  });
});
