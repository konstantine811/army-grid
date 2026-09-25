import { describe, expect, it } from "vitest";
import {
  diskNamePartsClose,
  isPlausibleDiskQuestionnaireMatch,
  isStrictDiskQuestionnaireMatch,
  pickStrictDiskQuestionnaireMatch,
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

  it("keeps a file that only has the surname for manual review", () => {
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
    ).toBe(false);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "САВЧУК Євген Олександрович",
        "Артем.pdf",
      ),
    ).toBe(false);
  });

  it("does not treat a callsign-only file as the person's questionnaire", () => {
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "КЛУБАНЬ Володимир Вікторович",
        "Полтава.pdf",
        "полтава",
      ),
    ).toBe(false);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "СТЕПАНОВ Максим В'ячеславович",
        "Макс.pdf",
        "макс",
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

  it("ignores leading numbers and the word «Анкета» in file names", () => {
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "БЕЗУГЛИЙ Дмитро Юрійович",
        "01 БЕЗУГЛИЙ Дмитро Юрійович (Безик).pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "РЕБІНЧАК Володимир Михайлович",
        "25 Анкета РЕБІНЧАК Володимир Михайлович (Крафт).pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "БІЛИК Сергій Володимирович",
        "26 Білик Сергій Володимирович.pdf",
      ),
    ).toBe(true);
    expect(
      isPlausibleDiskQuestionnaireMatch(
        "НЕЧЕВ Юрій Валерійович",
        "28 Нечев Юрій Валерійович_.pdf",
      ),
    ).toBe(true);
  });
});

describe("isStrictDiskQuestionnaireMatch", () => {
  it("accepts exact FIO and rejects another person with the same surname", () => {
    const person = "КЛУБАНЬ Володимир Вікторович";
    expect(
      isStrictDiskQuestionnaireMatch(
        person,
        "КЛУБАНЬ Володимир Вікторович (Полтава).pdf",
      ),
    ).toBe(true);
    expect(
      isStrictDiskQuestionnaireMatch(person, "КЛУБАНЬ Іван Петрович.pdf"),
    ).toBe(false);
    expect(isStrictDiskQuestionnaireMatch(person, "Полтава.pdf")).toBe(false);
  });

  it("returns null when several different strict files match", () => {
    const picked = pickStrictDiskQuestionnaireMatch(
      "КЛУБАНЬ Володимир Вікторович",
      [
        {
          fileName: "КЛУБАНЬ Володимир Вікторович (A).pdf",
          relativePath: "a/a.pdf",
        },
        {
          fileName: "КЛУБАНЬ Володимир Вікторович (B).pdf",
          relativePath: "b/b.pdf",
        },
      ],
    );
    expect(picked).toBeNull();
  });
});
