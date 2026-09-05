import { describe, expect, it } from "vitest";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildAnketaSourceFieldUpdates,
  buildPersonnelToAnketaFieldUpdates,
  isReplaceableAnketaMissingValue,
} from "./anketaPersonMerge";
import { createEmptyAnketaRow } from "./anketaSheet";

describe("isReplaceableAnketaMissingValue", () => {
  it.each([
    "",
    "дані відсутні",
    "не має",
    "немає",
    "відсутні",
    "втрачено",
    "уточнюється",
  ])("treats %s as replaceable", (value) => {
    expect(isReplaceableAnketaMissingValue(value)).toBe(true);
  });

  it("preserves real questionnaire data", () => {
    expect(isReplaceableAnketaMissingValue("1234567890")).toBe(false);
  });
});

describe("buildPersonnelToAnketaFieldUpdates", () => {
  it("fills blanks and absence markers without overwriting real Anketa values", () => {
    const anketa = createEmptyAnketaRow(12);
    anketa.fullName = "ТЕСТОВИЙ Іван Іванович";
    anketa.birthDate = "дані відсутні";
    anketa.education = "не має";
    anketa.rnokpp = "9999999999";

    const personnel = {
      дата_народження: "01.02.1990",
      освіта: "вища",
      рнокпп_за_наявності: "1234567890",
    } as EjournalPreviewRow;

    const result = buildPersonnelToAnketaFieldUpdates(personnel, anketa);

    expect(result.fieldUpdates).toMatchObject({
      birthDate: "01.02.1990",
      education: "вища",
    });
    expect(result.fieldUpdates).not.toHaveProperty("rnokpp");
  });

  it("does not copy an absence marker from Personnel", () => {
    const anketa = createEmptyAnketaRow(13);
    const personnel = {
      освіта: "невідомо",
    } as EjournalPreviewRow;

    expect(
      buildPersonnelToAnketaFieldUpdates(personnel, anketa).fieldUpdates,
    ).not.toHaveProperty("education");
  });
});

describe("buildAnketaSourceFieldUpdates", () => {
  it("copies current OOS columns over absence markers only", () => {
    const anketa = createEmptyAnketaRow(14);
    anketa.positionDates = "дані відсутні";
    anketa.positionOrderNumber = "не має";
    anketa.arrivedFrom = "вже заповнено в анкеті";

    const result = buildAnketaSourceFieldUpdates(
      {
        positionDates: "23.08.2026",
        positionOrderNumber: "244",
        arrivedFrom: "2 ШБ",
      },
      anketa,
    );

    expect(result.fieldUpdates).toMatchObject({
      positionDates: "23.08.2026",
      positionOrderNumber: "244",
    });
    expect(result.fieldUpdates).not.toHaveProperty("arrivedFrom");
  });
});
