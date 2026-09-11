import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  buildOverviewLocationCountCopyText,
  buildOverviewRotaCopyText,
  buildOverviewWhatsAppCopyText,
} from "./overviewCopyText";

const row = (
  id: string,
  name: string,
  statusLabel: string,
  unit: string,
  place = "",
  position = "стрілець",
  fullPosition?: string,
) =>
  ({
    id,
    externalId: id,
    name,
    status: statusLabel === "В строю" ? "ON_DUTY" : "MEDICAL",
    statusLabel,
    unit,
    positionTitle: fullPosition ?? position,
    staffSheetColumns: {
      staff_5: position,
      staff_7: fullPosition,
      staff_21: statusLabel,
      staff_31: place,
    },
  }) as unknown as BackendPersonnelOverviewRow;

describe("buildOverviewWhatsAppCopyText", () => {
  it("copies selected statuses and adds execution people from the selected unit", () => {
    const allRows = [
      row("1", "ГУК Володимир Степанович", "В строю", "3 рота"),
      row("2", "ГАПОН Андрій Вікторович", "Лікування", "3 рота"),
      row("3", "ДРАГОЙ Микола Леонідович", "В строю", "3 рота", "На виконанні"),
      row("4", "ЧУЖИЙ Боєць", "В строю", "2 рота", "На виконанні"),
    ];

    expect(
      buildOverviewWhatsAppCopyText({
        rows: [allRows[0]!],
        allRows,
        columns: [],
        filters: [
          { id: "unit", label: "Підрозділ", values: ["3 рота"] },
          { id: "status", label: "Статус", values: ["В строю"] },
        ],
      }),
    ).toBe(
      [
        "В строю:",
        "1\tГУК Володимир Степанович",
        "",
        "На виконанні:",
        "1\tДРАГОЙ Микола Леонідович",
      ].join("\n"),
    );
  });

  it("copies short position, name and place for filtered rows", () => {
    expect(
      buildOverviewRotaCopyText({
        rows: [
          row(
            "1",
            "ГУК Володимир Степанович",
            "В строю",
            "3 рота",
            "ППД Вишневе",
            "головний сержант",
            "головний сержант 3 роти механізованого батальйону",
          ),
          row(
            "2",
            "ГАПОН Андрій Вікторович",
            "Лікування",
            "3 рота",
            "Шпиталь",
            "стрілець",
            "стрілець стрілець 1 піхотного взводу",
          ),
        ],
        allRows: [],
        columns: [],
        filters: [{ id: "unit", label: "Підрозділ", values: ["3 рота"] }],
      }),
    ).toBe(
      [
        "1 - Головний сержант - ГУК Володимир Степанович - ППД Вишневе",
        "2 - Стрілець - ГАПОН Андрій Вікторович - Шпиталь",
      ].join("\n"),
    );
  });

  it("puts commanders and management before other rows", () => {
    expect(
      buildOverviewRotaCopyText({
        rows: [
          row("1", "СТРІЛЕЦЬ Іван", "В строю", "3 рота", "ППД"),
          row(
            "2",
            "КОМАНДИР Петро",
            "В строю",
            "3 рота",
            "ППД",
            "Командир взводу",
            "Командир 1 піхотного взводу",
          ),
          row(
            "3",
            "СЕРЖАНТ Олег",
            "В строю",
            "3 рота",
            "ППД",
            "головний сержант",
            "головний сержант 3 роти",
          ),
        ],
        allRows: [],
        columns: [],
        filters: [{ id: "unit", label: "Підрозділ", values: ["3 рота"] }],
      }),
    ).toBe(
      [
        "1 - Командир взводу - КОМАНДИР Петро - ППД",
        "2 - Головний сержант - СЕРЖАНТ Олег - ППД",
        "3 - Стрілець - СТРІЛЕЦЬ Іван - ППД",
      ].join("\n"),
    );
  });
});

describe("buildOverviewLocationCountCopyText", () => {
  it("counts every in-service place so the total matches В строю", () => {
    const allRows = [
      row("1", "ГУК Володимир Степанович", "В строю", "3 рота", "ППД Вишневе"),
      row("2", "ГАПОН Андрій Вікторович", "Лікування", "3 рота", "ППД Вишневе"),
      row("3", "ДРАГОЙ Микола Леонідович", "В строю", "3 рота", "ППД Павлоград"),
      row("4", "ПОЛІГОННИЙ Іван", "В строю", "3 рота", "Полігон Д"),
      row("5", "НА ВИКОНАННІ Іван", "В строю", "3 рота", "На виконанні"),
      row("6", "ШПИТАЛЬ Іван", "Лікування", "3 рота", "Шпиталь"),
      row("7", "ВІДПУСТКА Петро", "Відпустка", "3 рота", "Відпустка"),
      row("8", "ЧУЖИЙ Боєць", "В строю", "2 рота", "ППД Вишневе"),
    ];

    expect(
      buildOverviewLocationCountCopyText({
        rows: [allRows[0]!],
        allRows,
        columns: [],
        filters: [{ id: "unit", label: "Підрозділ", values: ["3 рота"] }],
      }),
    ).toBe(
      [
        "В строю: 4",
        "На виконанні: 1",
        "Полігон Д: 1",
        "ППД Вишневе: 1",
        "ППД Павлоград: 1",
        "",
        "Полігон Д: 1",
        "Лікування: 2",
        "Відпустка: 1",
      ].join("\n"),
    );
  });
});
