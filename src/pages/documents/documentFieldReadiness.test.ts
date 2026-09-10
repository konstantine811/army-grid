import { describe, expect, it } from "vitest";
import {
  documentBasisFieldHighlightClass,
  documentHasBasisDateMismatch,
  documentHasEmptyInputs,
  documentRequiredFieldIsBlank,
  normalizeUbdReadinessFields,
  readDocumentSkippedDueToSzch,
  readDocumentSkippedDueToStatus200,
  readDocumentSkippedFromWork,
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

  it("clears stale UBD «БР не підходить» when dates match", () => {
    const fields = {
      taskPeriod: "з 04.08.2026-02.09.2026",
      taskPlace: "н.п. Шилівка",
      basisNumber: "4862/ОКП/2264/дск",
      basisDate: "04.08.2026",
      basisNotReady: true,
    };
    expect(documentHasBasisDateMismatch("ubdReport", fields)).toBe(false);
    expect(
      documentBasisFieldHighlightClass("ubdReport", fields, "basisNumber"),
    ).toBeUndefined();
  });

  it("allows manual BR pick when day/month match but years differ", () => {
    const fields = {
      taskPeriod: "з 06.08.2025-13.08.2025",
      taskPlace: "н.п. Садки",
      basisNumber: "4862/ОКП/2292/дск",
      basisDate: "06.08.2026",
      basisNotReady: false,
    };
    expect(documentHasBasisDateMismatch("ubdReport", fields)).toBe(false);
    expect(
      documentBasisFieldHighlightClass("ubdReport", fields, "basisDate"),
    ).toBeUndefined();
  });

  it("clears journal highlight when dates match but stale basisNotReady=true", () => {
    const fields = {
      taskPeriod: "з 04.08.2026-02.09.2026",
      basisNumber: "4862/ОКП/2264/дск",
      basisDate: "04.08.2026",
      basisNotReady: true,
    };
    expect(documentHasBasisDateMismatch("ubdReport", fields)).toBe(false);
    expect(normalizeUbdReadinessFields(fields).basisNotReady).toBe(false);
  });

  it("keeps journal highlight only with explicit basisNotReady on mismatch", () => {
    const fields = {
      taskPeriod: "з 04.08.2026-02.09.2026",
      basisNumber: "4862/ОКП/2292/дск",
      basisDate: "06.08.2026",
      basisNotReady: true,
    };
    expect(documentHasBasisDateMismatch("ubdReport", fields)).toBe(false);
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

  it("accepts lost-ID event with date, place and circumstances only", () => {
    expect(
      documentHasEmptyInputs("lostMilitaryId", {
        fullName: "Іваненко Іван",
        rank: "солдат",
        staffPosition: "стрілець",
        addressee: "командиру",
        lossDate: "12.11.2025",
        circumstanceKind: "custom",
        lossLocation: "с. Гришене",
        customCircumstances:
          "на позицію потрапив каб та FPV дрон і все згоріло",
      }),
    ).toBe(false);
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

  it("reads the status 200 skip flag from document fields", () => {
    expect(
      readDocumentSkippedDueToStatus200({ skippedDueToStatus200: true }),
    ).toBe(true);
    expect(
      readDocumentSkippedDueToStatus200({ skippedDueToStatus200: "true" }),
    ).toBe(true);
    expect(
      readDocumentSkippedDueToStatus200({ skippedDueToStatus200: false }),
    ).toBe(false);
    expect(readDocumentSkippedDueToStatus200({})).toBe(false);
  });

  it("treats either skip flag as skipped from work", () => {
    expect(
      readDocumentSkippedFromWork({
        skippedDueToSzch: true,
        skippedDueToStatus200: false,
      }),
    ).toBe(true);
    expect(
      readDocumentSkippedFromWork({
        skippedDueToSzch: false,
        skippedDueToStatus200: true,
      }),
    ).toBe(true);
    expect(readDocumentSkippedFromWork({})).toBe(false);
  });
});
