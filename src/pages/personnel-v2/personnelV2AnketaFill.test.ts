import { describe, expect, it } from "vitest";
import { createEmptyAnketaRow } from "../anketa-data/anketaSheet";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { PERSON_CARD_FIELDS } from "../personnel/personnelUtils";
import {
  indexPersonnelV2AnketaRows,
  matchPersonnelV2AnketaRow,
  mergePersonSummaryWithAnketa,
  personCardFieldValue,
} from "./personnelV2AnketaFill";

const anketa = () => {
  const row = createEmptyAnketaRow(2);
  row.fullName = "Шевченко Тарас Григорович";
  row.externalId = "2103004";
  row.birthDate = "09.03.1814";
  row.birthPlace = "м. Моринці";
  row.rnokpp = "1234567890";
  row.education = "Київський університет";
  return row;
};

describe("personnel v2 anketa fill", () => {
  it("matches by external id and fills empty OOS fields", () => {
    const row = anketa();
    const index = indexPersonnelV2AnketaRows([row]);
    expect(
      matchPersonnelV2AnketaRow(index, "Інше ім'я", "", "2103004"),
    ).toBe(row);

    const personnel = {
      __dbRowId: "roster:1",
      звання: "солдат",
    } as EjournalPreviewRow;
    const birthPlace = PERSON_CARD_FIELDS.find(
      (field) => field.label === "Місце народження",
    )!;
    const rank = PERSON_CARD_FIELDS.find((field) => field.label === "Звання")!;

    expect(personCardFieldValue(personnel, row, birthPlace)).toMatchObject({
      value: "м. Моринці",
      fromAnketa: true,
    });
    expect(personCardFieldValue(personnel, row, rank)?.fromAnketa).toBe(false);
  });

  it("keeps an OOS value and fills the summary from anketa", () => {
    const row = anketa();
    const summary = mergePersonSummaryWithAnketa(
      {
        name: "Шевченко Тарас Григорович",
        rank: "солдат",
        externalId: "2103004",
        positionIndex: "",
        serviceType: "",
        birthDate: "",
        birthPlace: "",
        sex: "",
        rnokpp: "",
        location: "",
        positionTitle: "стрілець",
        arrivedFrom: "",
        education: "",
        relatives: "",
        additionalInfo: "",
        phones: [],
        phonesDisplay: [],
        militaryId: "",
        contractFrom: "",
        contractTo: "",
        callSign: "Кобзар",
      },
      row,
    );
    expect(summary?.rank).toBe("солдат");
    expect(summary?.rnokpp).toBe("1234567890");
    expect(summary?.birthPlace).toBe("м. Моринці");
    expect(summary?.education).toBe("Київський університет");
    expect(summary?.positionTitle).toBe("стрілець");
  });
});
