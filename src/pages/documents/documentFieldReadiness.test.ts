import { describe, expect, it } from "vitest";
import {
  documentBasisFieldHighlightClass,
  documentHasBasisDateMismatch,
  documentHasEmptyInputs,
  documentRequiredFieldIsBlank,
  readDocumentSkippedDueToSzch,
} from "./documentFieldReadiness";

describe("documentFieldReadiness", () => {
  it("marks UBD empty when РНОКПП is short", () => {
    expect(
      documentHasEmptyInputs("ubdReport", {
        fullName: "Іваненко Іван",
        rank: "солдат",
        staffPosition: "стрілець",
        birthDate: "01.01.1990",
        rnokpp: "123",
        taskPeriod: "07.06.2026 — 07.07.2026",
        taskPlace: "схід",
        basisNumber: "4862/ОКП/1471/дск",
        basisDate: "07.06.2026",
      }),
    ).toBe(true);
  });

  it("marks Form 6 yellow when BR date does not match task period start", () => {
    const fields = {
      fullName: "Іваненко Іван",
      rank: "солдат",
      staffPosition: "стрілець",
      birthDate: "01.01.1990",
      idDocument: "Паспорт 002532584",
      rnokpp: "1234567890",
      taskPeriod: "07.06.2026 — 07.07.2026",
      taskPlace: "схід",
      basisNumber: "4862/ОКП/1162/дск",
      basisDate: "09.05.2026",
    };
    expect(documentHasEmptyInputs("form6Report", fields)).toBe(false);
    expect(documentHasBasisDateMismatch("form6Report", fields)).toBe(true);
    expect(
      documentBasisFieldHighlightClass("form6Report", fields, "basisDate"),
    ).toBe("document-field-warning");
  });

  it("marks Form 12 incomplete without a name", () => {
    expect(
      documentHasEmptyInputs("form12Report", {
        commander: "Командиру батальйону",
        fullName: "",
        rank: "солдат",
        staffPosition: "стрілець",
      }),
    ).toBe(true);
    expect(
      documentRequiredFieldIsBlank("form12Report", "fullName", {
        fullName: "",
      }),
    ).toBe(true);
  });

  it("does not yellow documents without a BR picker", () => {
    expect(
      documentHasBasisDateMismatch("form12Report", {
        taskPeriod: "07.06.2026",
        basisDate: "01.01.2020",
      }),
    ).toBe(false);
  });

  it("requires lost-ID movement locations", () => {
    expect(
      documentHasEmptyInputs("lostMilitaryId", {
        fullName: "Іваненко Іван",
        rank: "солдат",
        staffPosition: "стрілець",
        addressee: "командиру",
        lossDate: "01.02.2026",
        circumstanceKind: "movement",
        fromLocation: "",
        toLocation: "пункт",
      }),
    ).toBe(true);
  });

  it("marks Form 6 id document incomplete without a passport number", () => {
    const filled = {
      fullName: "Іваненко Іван",
      rank: "солдат",
      staffPosition: "стрілець",
      birthDate: "01.01.1990",
      rnokpp: "1234567890",
      taskPeriod: "07.06.2026 — 07.07.2026",
      taskPlace: "схід",
      basisNumber: "4862/ОКП/1471/дск",
      basisDate: "07.06.2026",
    };
    expect(
      documentRequiredFieldIsBlank("form6Report", "idDocument", {
        ...filled,
        idDocument: "Паспорт громадянина України",
      }),
    ).toBe(true);
    expect(
      documentRequiredFieldIsBlank("form6Report", "idDocument", {
        ...filled,
        idDocument: "Паспорт громадянина України Серія КВ №001828",
      }),
    ).toBe(false);
    expect(
      documentRequiredFieldIsBlank("form6Report", "idDocument", {
        ...filled,
        idDocument: "Паспорт 002532584",
      }),
    ).toBe(false);
  });

  it("reads the СЗЧ skip flag from document fields", () => {
    expect(readDocumentSkippedDueToSzch({ skippedDueToSzch: true })).toBe(true);
    expect(readDocumentSkippedDueToSzch({ skippedDueToSzch: "true" })).toBe(
      true,
    );
    expect(readDocumentSkippedDueToSzch({ skippedDueToSzch: false })).toBe(
      false,
    );
    expect(readDocumentSkippedDueToSzch({})).toBe(false);
  });
});
