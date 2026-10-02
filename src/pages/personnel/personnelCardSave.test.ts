import { describe, expect, it } from "vitest";
import {
  cardSavedMessage,
  mergeSavedPersonRow,
  personCardValuesForSave,
} from "./personnelCardSave";
import {
  phoneDraftProblem,
  phoneRemovedMessage,
  phoneSavedMessage,
} from "./personnelPhoneActions";
import {
  questionnaireOpenError,
  questionnairePreviewPersonId,
} from "./personnelQuestionnaireOpen";
import { questionnaireSavedMessage } from "./personnelQuestionnaireSave";
import { assignPhotoValues, removePhotoKeys } from "./personnelPhotoSave";

describe("personnel card save", () => {
  it("keeps the row id and leaves other people untouched", () => {
    const saved = mergeSavedPersonRow(
      { __dbRowId: "12", column_14: "старий" },
      "12",
      { column_14: "Кіяненко" },
    );
    const other = mergeSavedPersonRow(
      { __dbRowId: "4246", column_14: "Бабченко" },
      "12",
      { column_14: "Кіяненко" },
    );
    expect(saved).toMatchObject({ __dbRowId: "12", column_14: "Кіяненко" });
    expect(other.column_14).toBe("Бабченко");
    expect(cardSavedMessage("Кіяненко")).toBe("Картку оновлено: Кіяненко.");
  });

  it("stores a position index as separate lines", () => {
    const values = personCardValuesForSave(
      [{ key: "індекс", parts: ["індекс", "посади"] }],
      { індекс: "123 · 456" },
    );
    expect(values.індекс).toBe("123\n456");
  });
});

describe("personnel phone messages", () => {
  it("asks for a person before a number", () => {
    expect(phoneDraftProblem(false, "")).toBe("Спочатку виберіть особу зі списку.");
    expect(phoneDraftProblem(true, "")).toContain("063 123 45 67");
  });

  it("says whether the number reached the database", () => {
    expect(phoneSavedMessage("0631234567", true)).toContain("063 123 45 67");
    expect(phoneSavedMessage("0631234567", false)).toContain("локально");
    expect(phoneRemovedMessage("0631234567", false)).toContain("Не вдалося оновити БД");
  });

  it("writes a photo onto every lookup key and drops it again", () => {
    const withPhoto = assignPhotoValues({}, ["4246", "pid_1"], "blob");
    expect(withPhoto).toEqual({ "4246": "blob", pid_1: "blob" });
    expect(removePhotoKeys(withPhoto, ["4246"])).toEqual({ pid_1: "blob" });
  });
});

describe("questionnaire preview id", () => {
  it("prefers the stored questionnaire id over the list id", () => {
    expect(questionnairePreviewPersonId("pid_1", "p:федерко:c:музикант")).toBe("pid_1");
    expect(questionnairePreviewPersonId("  ", "4246")).toBe("4246");
    expect(questionnaireOpenError(new Error("немає файлу"))).toBe(
      "Не вдалося відкрити анкету: немає файлу",
    );
    expect(questionnaireSavedMessage("Бабченко", "anketa.pdf")).toContain("anketa.pdf");
  });
});
