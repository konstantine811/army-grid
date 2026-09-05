import { describe, expect, it } from "vitest";
import { normalizeAnketaNameKey } from "./anketaPersonMatch";
import {
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  applyAbsentQuestionnaireClearsToRows,
  applyAbsentQuestionnaireFillsToRows,
  buildAbsentQuestionnaireAnketaRows,
  collectAbsentQuestionnaireCellClears,
  collectAbsentQuestionnaireCellFills,
  countAnketaBlankFieldPersons,
  findNextAnketaEmptyCell,
  findNextAnketaPersonEmptyCell,
  listAnketaEmptyCells,
  removePresentQuestionnairesFromMissingNameKeys,
} from "./anketaGaps";
import { createEmptyAnketaRow } from "./anketaSheet";

const person = (name: string, extra: Record<string, string> = {}) => {
  const row = createEmptyAnketaRow(2);
  return { ...row, fullName: name, ...extra };
};

describe("collectAbsentQuestionnaireCellFills", () => {
  it("fills only empty selected columns for people without a questionnaire", () => {
    const missing = person("КОВАЛЬ Іван Петрович", {
      rnokpp: "",
      birthDate: "01.01.1990",
      education: "",
    });
    const hasAnketa = person("ШЕВЧЕНКО Тарас Григорович", {
      rnokpp: "",
      education: "",
    });
    const exclude = new Set([normalizeAnketaNameKey(missing.fullName)]);

    const fills = collectAbsentQuestionnaireCellFills(
      [missing, hasAnketa],
      ["rnokpp", "birthDate", "education"],
      exclude,
    );

    expect(fills.map((fill) => fill.columnId).sort()).toEqual([
      "education",
      "rnokpp",
    ]);
    expect(fills.every((fill) => fill.rowId === missing.__rowId)).toBe(true);
  });

  it("fills empty selected cells when the person has no PDF", () => {
    const row = person("НОВИЙ Іван Петрович", { rnokpp: "", education: "" });
    const fills = collectAbsentQuestionnaireCellFills(
      [row],
      ["rnokpp"],
      new Set(),
      () => false,
    );
    expect(fills).toHaveLength(1);
    expect(fills[0]?.columnId).toBe("rnokpp");
  });

  it("does not overwrite people who already have a PDF", () => {
    const row = person("ШЕВЧЕНКО Тарас Григорович", { rnokpp: "" });
    expect(
      collectAbsentQuestionnaireCellFills([row], ["rnokpp"], new Set(), () => true),
    ).toEqual([]);
  });

  it("does not fill a listed-missing person who now has a PDF", () => {
    const row = person("КОВАЛЬ Іван Петрович", { rnokpp: "" });
    const exclude = new Set([normalizeAnketaNameKey(row.fullName)]);
    expect(
      collectAbsentQuestionnaireCellFills([row], ["rnokpp"], exclude, () => true),
    ).toEqual([]);
  });
});

describe("collectAbsentQuestionnaireCellClears", () => {
  it("clears «дані відсутні» only for people who now have a PDF", () => {
    const found = person("ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      birthDate: "27.02.1986",
    });
    const stillMissing = person("КОВАЛЬ Іван Петрович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });

    const clears = collectAbsentQuestionnaireCellClears(
      [found, stillMissing],
      ["rnokpp", "education"],
      (row) => row.fullName === found.fullName,
    );

    expect(clears.map((item) => item.columnId).sort()).toEqual([
      "education",
      "rnokpp",
    ]);
    expect(clears.every((item) => item.rowId === found.__rowId)).toBe(true);
  });

  it("writes empty cells so gap search can fill them", () => {
    const row = person("ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "середня",
    });
    const next = applyAbsentQuestionnaireClearsToRows([row], [
      {
        rowId: row.__rowId,
        rowNumber: row.__rowNumber,
        columnId: "rnokpp",
        fullName: row.fullName,
        externalId: "",
      },
    ]);

    expect(next[0]?.rnokpp).toBe("");
    expect(next[0]?.education).toBe("середня");
    expect(
      listAnketaEmptyCells(next, ["rnokpp", "education"]).map(
        (cell) => cell.columnId,
      ),
    ).toEqual(["rnokpp"]);
  });
});

describe("removePresentQuestionnairesFromMissingNameKeys", () => {
  it("returns a person with a newly found questionnaire to gap search", () => {
    const found = person("ШЕВЦОВ Дмитро Сергійович");
    const stillMissing = person("КОВАЛЬ Іван Петрович");
    const missingKeys = new Set([
      normalizeAnketaNameKey(found.fullName),
      normalizeAnketaNameKey(stillMissing.fullName),
    ]);

    const next = removePresentQuestionnairesFromMissingNameKeys(
      [found, stillMissing],
      missingKeys,
      (row) => row === found,
    );

    expect(next.has(normalizeAnketaNameKey(found.fullName))).toBe(false);
    expect(next.has(normalizeAnketaNameKey(stillMissing.fullName))).toBe(true);
  });
});

describe("buildAbsentQuestionnaireAnketaRows", () => {
  it("adds people from the missing list who are not in the table yet", () => {
    const existing = person("ШЕВЧЕНКО Тарас Григорович");
    const created = buildAbsentQuestionnaireAnketaRows(
      [existing],
      ["ШЕВЧЕНКО Тарас Григорович", "КОВАЛЬ Іван Петрович"],
      ["rnokpp"],
    );
    expect(created).toHaveLength(1);
    expect(created[0]?.fullName).toBe("КОВАЛЬ Іван Петрович");
    expect(created[0]?.rnokpp).toBe(ANKETA_ABSENT_QUESTIONNAIRE_VALUE);
  });
});

describe("listAnketaEmptyCells skips absent questionnaires", () => {
  it("does not return cells marked «дані відсутні»", () => {
    const row = person("КОВАЛЬ Іван Петрович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "",
    });

    expect(
      listAnketaEmptyCells([row], ["rnokpp", "education"]),
    ).toEqual([]);
  });

  it("still finds empty cells for people who have a questionnaire", () => {
    const row = person("ШЕВЧЕНКО Тарас Григорович", { rnokpp: "" });
    const gaps = listAnketaEmptyCells([row], ["rnokpp"]);
    expect(gaps).toHaveLength(1);
    expect(gaps[0]?.columnId).toBe("rnokpp");
  });
});

describe("applyAbsentQuestionnaireFillsToRows", () => {
  it("writes the mark only into collected cells", () => {
    const row = person("КОВАЛЬ Іван Петрович", {
      rnokpp: "",
      education: "середня",
    });
    const next = applyAbsentQuestionnaireFillsToRows([row], [
      {
        rowId: row.__rowId,
        rowNumber: row.__rowNumber,
        columnId: "rnokpp",
        fullName: row.fullName,
        externalId: "",
      },
    ]);

    expect(next[0]?.rnokpp).toBe(ANKETA_ABSENT_QUESTIONNAIRE_VALUE);
    expect(next[0]?.education).toBe("середня");
  });
});

describe("countAnketaBlankFieldPersons", () => {
  it("counts people with completely empty selected fields, not «дані відсутні»", () => {
    const fullyBlank = person("ПЕРШИЙ Іван Іванович");
    const partlyBlank = person("ДРУГИЙ Петро Петрович", {
      rnokpp: "1234567890",
    });
    const marked = person("ТРЕТІЙ Олег Олегович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });

    const stats = countAnketaBlankFieldPersons(
      [fullyBlank, partlyBlank, marked],
      ["rnokpp", "education"],
    );

    expect(stats.personsFullyBlank).toBe(1);
    expect(stats.personsWithBlankFields).toBe(2);
    expect(stats.blankCells).toBe(3);
  });
});

describe("findNextAnketaEmptyCell walks from the current cell", () => {
  it("returns the next empty cell without wrapping when one remains", () => {
    const first = person("ПЕРШИЙ", { rnokpp: "", education: "середня" });
    first.__rowId = "r1";
    first.__rowNumber = 2;
    const second = person("ДРУГИЙ", { rnokpp: "111", education: "" });
    second.__rowId = "r2";
    second.__rowNumber = 3;

    const current = listAnketaEmptyCells([first, second], ["rnokpp", "education"])[0];
    const next = findNextAnketaEmptyCell(
      [first, second],
      current,
      ["rnokpp", "education"],
    );

    expect(next?.rowId).toBe("r2");
    expect(next?.columnId).toBe("education");
  });

  it("jumps to the next person and skips the rest of the current row", () => {
    const first = person("ПЕРШИЙ", { rnokpp: "", education: "" });
    first.__rowId = "r1";
    first.__rowNumber = 2;
    const second = person("ДРУГИЙ", { rnokpp: "", education: "" });
    second.__rowId = "r2";
    second.__rowNumber = 3;

    const current = listAnketaEmptyCells([first, second], ["rnokpp", "education"])[0];
    const next = findNextAnketaPersonEmptyCell(
      [first, second],
      current,
      ["rnokpp", "education"],
    );

    expect(next?.rowId).toBe("r2");
    expect(next?.columnId).toBe("rnokpp");
  });
});
