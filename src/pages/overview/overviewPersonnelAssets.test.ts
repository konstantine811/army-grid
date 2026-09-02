import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { applyPersonnelAssetsToOverview } from "./overviewPersonnelAssets";

const overviewRow = (
  name: string,
  extra: Partial<BackendPersonnelOverviewRow> = {},
): BackendPersonnelOverviewRow => ({
  id: extra.id || `overview:${name}`,
  externalId: extra.externalId || `roster:${name}`,
  name,
  rank: "солдат",
  unit: "1 рота",
  status: "ON_DUTY",
  statusLabel: "На службі",
  validFrom: null,
  days: null,
  plannedReturn: null,
  place: "",
  updatedAt: "",
  ...extra,
});

describe("applyPersonnelAssetsToOverview", () => {
  it("copies questionnaire and photo from Особовий склад by ПІБ", () => {
    const personnel = {
      __dbRowId: "p1",
      прізвище: "ПОЛЬОВИЙ Олексій Геннадійович",
      id: "2103001",
    } as EjournalPreviewRow;
    const overview = overviewRow("ПОЛЬОВИЙ Олексій Геннадійович", {
      externalId: "roster:польовий",
    });

    const result = applyPersonnelAssetsToOverview(
      [overview],
      [personnel],
      [{ personExternalId: "2103001", fileName: "Польовий Олексій.pdf" }],
      [{ personExternalId: "2103001", photoData: "data:image/jpeg;base64,abc" }],
    );

    expect(result.questionnairePresence["roster:польовий"]).toBe(true);
    expect(result.questionnaireSourceIds["roster:польовий"]).toBe("2103001");
    expect(result.photos["roster:польовий"]).toBe("data:image/jpeg;base64,abc");
  });

  it("matches Overview ПІБ even when the row has a birth date suffix", () => {
    const personnel = {
      __dbRowId: "p2",
      прізвище: "МОРОЗОВ Віктор Вікторович",
      дата_народження: "27.03.1988",
      id: "2103002",
    } as EjournalPreviewRow;
    const overview = overviewRow("МОРОЗОВ Віктор Вікторович 27.03.1988 р.н.", {
      externalId: "roster:морозов",
    });

    const result = applyPersonnelAssetsToOverview(
      [overview],
      [personnel],
      [{ personExternalId: "2103002", fileName: "Морозов Віктор.pdf" }],
      [{ personExternalId: "2103002", photoData: "data:image/jpeg;base64,xyz" }],
    );

    expect(result.questionnairePresence["roster:морозов"]).toBe(true);
    expect(result.photos["roster:морозов"]).toBe("data:image/jpeg;base64,xyz");
  });

  it("copies documents from Особовий склад by ПІБ", () => {
    const personnel = {
      __dbRowId: "p3",
      прізвище: "ДОРОШЕНКО Дмитро Сергійович",
      id: "2103003",
    } as EjournalPreviewRow;
    const overview = overviewRow("ДОРОШЕНКО Дмитро Сергійович", {
      externalId: "roster:дорошенко",
    });

    const result = applyPersonnelAssetsToOverview(
      [overview],
      [personnel],
      [],
      [],
      [
        {
          id: "d1",
          personExternalId: "2103003",
          type: "form6Report",
          title: "Форма 6",
          createdAt: "",
          updatedAt: "",
        },
      ],
    );

    expect(result.documents["roster:дорошенко"]?.count).toBe(1);
    expect(result.documents["roster:дорошенко"]?.labels).toContain("Форма 6");
  });

  it("matches stored ПІБ+дата without a personnel row", () => {
    const overview = overviewRow("ТРЕГУБ Євген Федорович 12.04.1984 р.н.", {
      externalId: "roster:трегуб",
    });

    const result = applyPersonnelAssetsToOverview(
      [overview],
      [],
      [
        {
          personExternalId: "p:трегуб євген федорович:1984-04-12",
          fileName: "Трегуб Євген.pdf",
        },
      ],
      [
        {
          personExternalId: "p:трегуб євген федорович:1984-04-12",
          photoData: "data:image/jpeg;base64,tregub",
        },
      ],
    );

    expect(result.questionnairePresence["roster:трегуб"]).toBe(true);
    expect(result.photos["roster:трегуб"]).toBe("data:image/jpeg;base64,tregub");
  });

  it("joins a numeric Особовий склад id to Overview by ПІБ + дата", () => {
    const personnel = {
      __dbRowId: "p4",
      прізвище: "ЛАНОВИЙ Володимир Олександрович",
      дата_народження: "03.11.1979",
      id: "11524",
    } as EjournalPreviewRow;
    const overview = overviewRow("ЛАНОВИЙ Володимир Олександрович", {
      externalId: "roster:лановий",
    });

    const result = applyPersonnelAssetsToOverview(
      [overview],
      [personnel],
      [{ personExternalId: "11524", fileName: "questionnaire.pdf" }],
      [{ personExternalId: "11524", photoData: "data:image/jpeg;base64,lan" }],
    );

    expect(result.questionnairePresence["roster:лановий"]).toBe(true);
    expect(result.questionnaireSourceIds["roster:лановий"]).toBe("11524");
    expect(result.photos["roster:лановий"]).toBe("data:image/jpeg;base64,lan");
  });
});
