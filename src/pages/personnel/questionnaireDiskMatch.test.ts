import { describe, expect, it } from "vitest";
import {
  diskNamePartsClose,
  isPlausibleDiskQuestionnaireMatch,
} from "./questionnaireDiskMatch";

describe("diskNamePartsClose", () => {
  it("treats a one-letter slip or і/и as the same name", () => {
    expect(diskNamePartsClose("максим", "максім")).toBe(true);
    expect(diskNamePartsClose("степанов", "степанів")).toBe(true);
    expect(diskNamePartsClose("вячеславович", "вячеславовіч")).toBe(true);
  });

  it("does not treat a different given name as a typo", () => {
    expect(diskNamePartsClose("евген", "артем")).toBe(false);
    expect(diskNamePartsClose("евген", "назар")).toBe(false);
    expect(diskNamePartsClose("степанов", "лукин")).toBe(false);
  });
});

describe("isPlausibleDiskQuestionnaireMatch", () => {
  it("cuts files whose surname or patronymic belong to someone else", () => {
    const person = "СТЕПАНОВ Максим В'ячеславович";
    expect(
      isPlausibleDiskQuestionnaireMatch(
        person,
        "ЛУКІН Максим Геннадійович (Макс).pdf",
        "макс",
      ),
    ).toBe(false);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        person,
        "Скоркін Максим Олександрович (МАКС).pdf",
        "макс",
      ),
    ).toBe(false);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        person,
        "ЮВЖЕНКО Максим Володимирович (Макс).pdf",
        "макс",
      ),
    ).toBe(false);
  });

  it("cuts the same surname when the given name is a different person", () => {
    const person = "САВЧУК Євген Олександрович";
    expect(
      isPlausibleDiskQuestionnaireMatch(
        person,
        "САВЧУК Артем Андрійович (Сава).pdf",
      ),
    ).toBe(false);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        person,
        "САВЧУК Назар Михайлович (Гуцул).pdf",
      ),
    ).toBe(false);
    expect(
      isPlausibleDiskQuestionnaireMatch(person, "Савчук Назар Михайлович.pdf"),
    ).toBe(false);
  });

  it("keeps a file that only has the surname or only the given name", () => {
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "САВЧУК Євген Олександрович",
        "Савчук.pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "САВЧУК Євген Олександрович",
        "Євген.pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "САВЧУК Євген Олександрович",
        "Артем.pdf",
      ),
    ).toBe(false);
  });

  it("keeps a short file and a small typo for manual review", () => {
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "САВЧУК Євген Олександрович",
        "Савчук Євген.pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "СТЕПАНОВ Максим В'ячеславович",
        "Степанів Максім Вячеславович.pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "СТЕПАНОВ Максим В'ячеславович",
        "СТЕПАНОВ Максим.pdf",
      ),
    ).toBe(true);
  });

  it("keeps a single-token callsign file for manual review", () => {
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "СТЕПАНОВ Максим В'ячеславович",
        "Макс.pdf",
        "макс",
      ),
    ).toBe(true);
  });
});
