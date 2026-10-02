import { describe, expect, it } from "vitest";
import { createEmptyAnketaRow } from "../anketa-data/anketaSheet";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { PERSON_CARD_FIELDS } from "../personnel/personnelUtils";
import {
  indexPersonnelV2AnketaRows,
  matchPersonnelV2AnketaRow,
  mergePersonSummaryWithAnketa,
  personCardFieldValue,
  summaryPreferringAnketa,
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

  it("keeps the staff stay place when the anketa location is empty", () => {
    const row = anketa();
    const summary = summaryPreferringAnketa(
      {
        name: "Шевченко Тарас Григорович",
        rank: "",
        externalId: "2103004",
        positionIndex: "",
        serviceType: "",
        birthDate: "",
        birthPlace: "",
        sex: "",
        rnokpp: "",
        location: "ППД Петропавлівка",
        positionTitle: "",
        arrivedFrom: "",
        education: "",
        relatives: "",
        additionalInfo: "",
        phones: [],
        phonesDisplay: [],
        militaryId: "",
        contractFrom: "",
        contractTo: "",
        callSign: "",
      },
      row,
    );
    expect(summary?.location).toBe("ППД Петропавлівка");

    row.location = "Полігон";
    expect(summaryPreferringAnketa(summary, row)?.location).toBe("Полігон");
  });

  it("keeps the staff rank when the anketa still says солдат", () => {
    const row = anketa();
    row.fullName = "ДЕЙНЕГА Юрій Володимирович";
    row.rank = "солдат";
    const summary = summaryPreferringAnketa(
      {
        name: "ДЕЙНЕГА Юрій Володимирович",
        rank: "молодший сержант",
        externalId: "9725",
        positionIndex: "2111778 2103472 2103407",
        serviceType: "мобілізац",
        birthDate: "12.07.1992",
        birthPlace: "",
        sex: "",
        rnokpp: "",
        location: "",
        positionTitle: "",
        arrivedFrom: "",
        education: "",
        relatives: "",
        additionalInfo: "",
        phones: [],
        phonesDisplay: [],
        militaryId: "",
        contractFrom: "",
        contractTo: "",
        callSign: "ДІЄГО",
      },
      row,
    );
    expect(summary?.rank).toBe("молодший сержант");
    expect(summary?.callSign).toBe("ДІЄГО");
  });

  it("does not attach another person's anketa when the external id is shared", () => {
    const sishchuk = createEmptyAnketaRow(70);
    sishchuk.fullName = "СИЩУК Андрій Олександрович";
    sishchuk.externalId = "8422";
    sishchuk.rank = "солдат";
    sishchuk.positionIndex = "2103244 2103397";
    sishchuk.birthDate = "04.05.2000";

    const moroz = createEmptyAnketaRow(438);
    moroz.fullName = "МОРОЗ Андрій Миколайович";
    moroz.externalId = "8422";
    moroz.rank = "молодший сержант";
    moroz.positionIndex = "2103244 2103787";
    moroz.birthDate = "10.12.1986";

    const index = indexPersonnelV2AnketaRows([sishchuk, moroz]);
    const matched = matchPersonnelV2AnketaRow(
      index,
      "МОРОЗ Андрій Миколайович",
      "10.12.1986",
      "8422",
    );

    expect(matched?.fullName).toBe("МОРОЗ Андрій Миколайович");
    expect(matched?.rank).toBe("молодший сержант");
  });
});
