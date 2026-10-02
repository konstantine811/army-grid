import { describe, expect, it } from "vitest";
import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  buildFighterStatusAdditions,
  findFighterStatusAddition,
} from "./fighterStatusImport";

const statusSheet = (): ExcelSheetSnapshot => {
  const header = ["Прізвище", "Позивний", "Дата виходу", "Дата повернення", "Днів"];
  const data = [
    ["Іваненко Петро", "Сокіл", "01.01.2026", "10.01.2026", "9"],
    ["Іваненко Петро", "Сокіл", "01.06.2026", "05.06.2026", "4"],
    ["Іваненко Петро", "Сокіл", "01.06.2026", "05.06.2026", "4"],
    ["Іваненко Петро", "Вовк", "01.03.2026", "03.03.2026", "2"],
  ];
  return {
    sheetIndex: 0,
    sheetName: "Статус бійців",
    rawRows: [header, ...data],
    headerRows: [header],
    rows: data.map((values, index) => ({
      id: `row-${index}`,
      excelRowNumber: index + 2,
      values,
      source: "merged" as const,
    })),
    columnCount: header.length,
    columnIndexes: header.map((_, index) => index),
    dataStartRow: 2,
  };
};

describe("buildFighterStatusAdditions", () => {
  it("keeps the latest exit and sums every distinct exit", () => {
    const additions = buildFighterStatusAdditions(statusSheet());
    const person = findFighterStatusAddition(
      { піб: "Іваненко Петро", позивний: "Сокіл" },
      additions,
    );

    expect(person?.fighter_status_exit_date).toBe("01.06.2026");
    expect(person?.fighter_status_return_date).toBe("05.06.2026");
    expect(person?.fighter_status_total_days).toBe("4");
    expect(person?.fighter_status_all_days).toBe("13");
  });

  it("does not add another callsign's exits to the sum", () => {
    const additions = buildFighterStatusAdditions(statusSheet());
    const person = findFighterStatusAddition(
      { піб: "Іваненко Петро", позивний: "Вовк" },
      additions,
    );

    expect(person?.fighter_status_total_days).toBe("2");
    expect(person?.fighter_status_all_days).toBe("2");
  });
});
