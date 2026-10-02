import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { buildVisibleRosterFieldRows } from "./personnelCardFields";
import { resolveMorningGeneralListColumnLabel } from "./morningGeneralListColumnLabels";
import {
  inferRosterFieldLabel,
  resolvePersonStayPlace,
} from "./personnelUtils";

const row = {
  roster__column_14: "КІЯНЕНКО Андрій Олександрович",
  roster__column_19: "1234567890",
  roster__column_5: "Командир батальйону",
  roster__column_22: "управління",
  roster__column_10: "так",
  roster__column_13: "старший лейтенант",
  roster__column_15: "МАРІК",
} as EjournalPreviewRow;

describe("buildVisibleRosterFieldRows", () => {
  it("hides the staff name, position, and IPN that already sit in the card header", () => {
    const fields = buildVisibleRosterFieldRows({
      row,
      rosterLabels: {
        column_14: "ПІБ",
        column_19: "ІПН",
        column_5: "Посада",
        column_22: "Підрозділ",
        column_10: "Анкета",
        column_13: "Звання",
        column_15: "Позивний",
      },
      personName: "КІЯНЕНКО Андрій Олександрович",
      birthDate: "1992-06-08",
      cardRnokpp: "1234567890",
    });

    expect(fields.map((field) => field.label)).toEqual(["Підрозділ"]);
  });

  it("shows the sum of exit days next to the exit count", () => {
    const fields = buildVisibleRosterFieldRows({
      row: {
        "roster__к-сть_виходів": "2",
        roster__fighter_status_all_days: "24",
        roster__підрозділ: "1 рота",
      } as EjournalPreviewRow,
      rosterLabels: { підрозділ: "Підрозділ" },
      personName: "БЕРЕЖНИЙ Олег Ігорович",
      birthDate: "",
      cardRnokpp: "",
    });

    expect(fields.map((field) => `${field.label}: ${field.value}`)).toEqual([
      "К-сть виходів: 2",
      "Усього днів: 24",
      "Підрозділ: 1 рота",
    ]);
  });

  it("names column 43 as place of stay and keeps restriction columns", () => {
    expect(resolveMorningGeneralListColumnLabel("колонка_43")).toBe(
      "Місце перебування",
    );
    expect(resolveMorningGeneralListColumnLabel("column_44")).toBe("Обмеження");
    expect(resolveMorningGeneralListColumnLabel("колонка_45")).toBe("Статус БГ");
    expect(
      inferRosterFieldLabel("колонка_43", "ЄКЦ Павлоград", {
        колонка_43: "Колонка 43",
      }),
    ).toBe("Місце перебування");

    const fields = buildVisibleRosterFieldRows({
      row: {
        roster__колонка_43: "ЄКЦ Павлоград",
        roster__колонка_44: "Обмеж. придат.",
        roster__колонка_45: "БГ",
        roster__колонка_30: "Полк",
      } as EjournalPreviewRow,
      rosterLabels: {
        колонка_43: "Колонка 43",
        колонка_44: "Колонка 44",
        колонка_45: "Колонка 45",
        колонка_30: "Колонка 30",
      },
      personName: "КІЯНЕНКО Андрій Олександрович",
      birthDate: "",
      cardRnokpp: "",
    });

    expect(fields.map((field) => field.label)).toEqual([
      "Обмеження",
      "Статус БГ",
      "Колонка 30",
    ]);
  });

  it("hides ditto «ж» and keeps a real squad name once", () => {
    const fields = buildVisibleRosterFieldRows({
      row: {
        roster__column_2: "ж",
        roster__підрозділ: "ж",
        roster__column_4: "ж",
        roster__відділення: "ж",
        roster__column_3: "1 взвод",
        roster__взвод: "1 взвод",
      } as EjournalPreviewRow,
      rosterLabels: {
        column_2: "Підрозділ",
        підрозділ: "Підрозділ",
        column_4: "Відділення",
        відділення: "Відділення",
        column_3: "Взвод",
        взвод: "Взвод",
      },
      personName: "КІЯНЕНКО Андрій Олександрович",
      birthDate: "",
      cardRnokpp: "",
    });

    expect(fields.map((field) => `${field.label}: ${field.value}`)).toEqual([
      "Взвод: 1 взвод",
    ]);
    expect(fields.map((field) => field.sourceKey)).toEqual(["взвод"]);
  });
});

describe("resolvePersonStayPlace", () => {
  it("prefers column 43 over the older stay-place column", () => {
    expect(
      resolvePersonStayPlace({
        місце_перебування: "КСП Павлоград",
        колонка_43: "ЄКЦ Павлоград",
      } as EjournalPreviewRow),
    ).toBe("ЄКЦ Павлоград");
    expect(
      resolvePersonStayPlace({
        місце_перебування: "ППД Вишневе",
      } as EjournalPreviewRow),
    ).toBe("ППД Вишневе");
  });
});
