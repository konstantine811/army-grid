import { describe, expect, it } from "vitest";
import {
  isCallsignPersonId,
  personnelListStatusMessage,
  questionnaireFileRejection,
  questionnairePreviewHeading,
  samePersonNameCount,
} from "./personnelQuestionnaireOpen";

describe("questionnaire open helpers", () => {
  it("builds the preview title and spots a callsign id", () => {
    expect(questionnairePreviewHeading("Кіяненко", "anketa.pdf")).toBe(
      "Анкета · Кіяненко · anketa.pdf",
    );
    expect(isCallsignPersonId("p:кіяненко:c:шеф")).toBe(true);
    expect(isCallsignPersonId("4246")).toBe(false);
  });

  it("rejects a non-pdf and a missing person id", () => {
    expect(
      questionnaireFileRejection(new File(["x"], "a.txt", { type: "text/plain" }), "12"),
    ).toBe("Анкета має бути у форматі PDF.");
    expect(
      questionnaireFileRejection(new File(["x"], "a.pdf", { type: "application/pdf" }), ""),
    ).toBe("Не вдалося додати анкету: у вибраної особи немає ID.");
  });

  it("counts the same display name and describes a cached list", () => {
    expect(samePersonNameCount(["Шевченко", "шевченко", "Кіяненко"], "ШЕВЧЕНКО")).toBe(2);
    expect(
      personnelListStatusMessage({
        focusedName: "",
        fromCache: true,
        rowCount: 662,
        sheetName: "ООС",
      }),
    ).toContain("Кеш: 662");
  });
});
