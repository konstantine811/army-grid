import { describe, expect, it } from "vitest";
import type { BackendPersonnelOverviewRow } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { buildPersonIdentityFingerprint } from "../personnel/personnelUtils";
import { buildOverviewQuestionnairePresence } from "./overviewQuestionnaires";

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

describe("buildOverviewQuestionnairePresence", () => {
  it("marks an overview staff row when the PDF lives under the EJOOS spreadsheet ID", () => {
    const roster = {
      __dbRowId: "r1",
      прізвище: "ПОЛЬОВИЙ Олексій Геннадійович",
      id: "2103001",
      дата_народження: "01.01.1990",
    } as EjournalPreviewRow;

    const result = buildOverviewQuestionnairePresence(
      [
        overviewRow("ПОЛЬОВИЙ Олексій Геннадійович", {
          externalId: "p:польовий олексій",
        }),
      ],
      [{ personExternalId: "2103001", fileName: "Польовий Олексій.pdf" }],
      [roster],
    );

    expect(result.presence["p:польовий олексій"]).toBe(true);
    expect(result.sourceIds["p:польовий олексій"]).toBe("2103001");
  });

  it("marks a row when the PDF is stored under the old name fingerprint", () => {
    const oldKey = buildPersonIdentityFingerprint(
      "МОРОЗОВ Олексій",
      "12.03.1988",
    );
    const result = buildOverviewQuestionnairePresence(
      [
        overviewRow("МОРОЗОВ Олексій Вікторович", {
          externalId: "2103777",
          name: "МОРОЗОВ Олексій Вікторович 12.03.1988",
        }),
      ],
      [{ personExternalId: oldKey, fileName: "Морозов Олексій.pdf" }],
    );

    expect(result.presence["2103777"]).toBe(true);
    expect(result.sourceIds["2103777"]).toBe(oldKey);
  });

  it("marks Польовий by file name when the overview id is only a roster key", () => {
    const result = buildOverviewQuestionnairePresence(
      [
        overviewRow("ПОЛЬОВИЙ Олексій Геннадійович", {
          externalId: "roster:польовий олексій геннадійович",
        }),
      ],
      [
        {
          personExternalId: "orphan-anketa",
          fileName: "Польовий Олексій Геннадійович.pdf",
        },
      ],
    );

    expect(result.presence["roster:польовий олексій геннадійович"]).toBe(true);
    expect(result.sourceIds["roster:польовий олексій геннадійович"]).toBe(
      "orphan-anketa",
    );
  });

  it("does not throw when the questionnaire list is not an array", () => {
    expect(() =>
      buildOverviewQuestionnairePresence(
        [overviewRow("ДОРОШЕНКО Дмитро Сергійович", { externalId: "1" })],
        null as unknown as [],
      ),
    ).not.toThrow();
  });
});
