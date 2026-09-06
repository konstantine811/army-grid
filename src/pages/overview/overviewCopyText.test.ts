import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import { buildOverviewWhatsAppCopyText } from "./overviewCopyText";

const row = (
  id: string,
  name: string,
  statusLabel: string,
  unit: string,
  place = "",
) =>
  ({
    id,
    externalId: id,
    name,
    status: statusLabel === "В строю" ? "ON_DUTY" : "MEDICAL",
    statusLabel,
    unit,
    staffSheetColumns: { staff_31: place },
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
});
