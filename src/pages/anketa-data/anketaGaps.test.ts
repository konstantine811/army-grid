import { describe, expect, it } from "vitest";
import {
  normalizeAnketaNameKey,
  orderAnketaRowsForGapSearch,
} from "./anketaPersonMatch";
import {
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE,
  ANKETA_RELATIVES_EMPTY_TEMPLATE,
  applyAbsentQuestionnaireClearsToRows,
  initialAnketaCellEditorDraft,
  applyAbsentQuestionnaireFillsToRows,
  buildAbsentQuestionnaireAnketaRows,
  collectAbsentQuestionnaireCellClears,
  collectAbsentQuestionnaireCellFills,
  countAnketaBlankFieldPersons,
  filterAnketaRowsForGapSearch,
  gateAnketaGapSearchRowsInStaffFirst,
  orderAnketaGapRowsBlankFirst,
  orderAnketaGapSearchRows,
  findNextAnketaEmptyCell,
  findNextAnketaPersonEmptyCell,
  isAnketaGapCellEmpty,
  isAnketaGapCellTrulyEmpty,
  isAnketaBlankStructuredTemplate,
  isAnketaEmptyStructuredBlock,
  listAnketaEmptyCells,
  removePresentQuestionnairesFromMissingNameKeys,
} from "./anketaGaps";
import { createEmptyAnketaRow } from "./anketaSheet";

const person = (name: string, extra: Record<string, string> = {}) => {
  const row = createEmptyAnketaRow(2);
  return { ...row, fullName: name, ...extra };
};

describe("initialAnketaCellEditorDraft", () => {
  it("prefills an empty relatives cell with the default template", () => {
    expect(initialAnketaCellEditorDraft("relatives", "", true)).toBe(
      ANKETA_RELATIVES_EMPTY_TEMPLATE,
    );
    expect(initialAnketaCellEditorDraft("relatives", "  ", true)).toBe(
      ANKETA_RELATIVES_EMPTY_TEMPLATE,
    );
  });

  it("prefills an empty additionalInfo cell with the default template", () => {
    expect(initialAnketaCellEditorDraft("additionalInfo", "", true)).toBe(
      ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE,
    );
  });

  it("keeps existing text and other columns unchanged", () => {
    expect(initialAnketaCellEditorDraft("relatives", "Мати: Іванова", false)).toBe(
      "Мати: Іванова",
    );
    expect(initialAnketaCellEditorDraft("additionalInfo", "тел: 050", false)).toBe(
      "тел: 050",
    );
    expect(initialAnketaCellEditorDraft("education", "", true)).toBe("");
  });
});

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

  it("does not clear when only some selected columns are «дані відсутні»", () => {
    const found = person("ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "середня",
    });
    expect(
      collectAbsentQuestionnaireCellClears(
        [found],
        ["rnokpp", "education"],
        () => true,
      ),
    ).toEqual([]);
  });

  it("does not clear when a selected column is still empty", () => {
    const found = person("ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "",
    });
    expect(
      collectAbsentQuestionnaireCellClears(
        [found],
        ["rnokpp", "education"],
        () => true,
      ),
    ).toEqual([]);
  });

  it("clears all selected columns when each one was «дані відсутні»", () => {
    const row = person("ШЕВЦОВ ДМИТРО СЕРГІЙОВИЧ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      birthDate: "27.02.1986",
    });
    const clears = collectAbsentQuestionnaireCellClears(
      [row],
      ["rnokpp", "education"],
      () => true,
    );
    const next = applyAbsentQuestionnaireClearsToRows([row], clears);

    expect(next[0]?.rnokpp).toBe("");
    expect(next[0]?.education).toBe("");
    expect(next[0]?.birthDate).toBe("27.02.1986");
    expect(
      listAnketaEmptyCells(next, ["rnokpp", "education"]).map(
        (cell) => cell.columnId,
      ),
    ).toEqual(["rnokpp", "education"]);
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

describe("isAnketaEmptyStructuredBlock", () => {
  it("treats empty relatives and additionalInfo templates as empty", () => {
    expect(isAnketaEmptyStructuredBlock(ANKETA_RELATIVES_EMPTY_TEMPLATE)).toBe(
      true,
    );
    expect(
      isAnketaEmptyStructuredBlock(ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE),
    ).toBe(true);
    expect(isAnketaBlankStructuredTemplate(ANKETA_RELATIVES_EMPTY_TEMPLATE)).toBe(
      true,
    );
    expect(
      isAnketaBlankStructuredTemplate(ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE),
    ).toBe(true);
  });

  it("treats labeled «дані відсутні» blocks as empty", () => {
    expect(
      isAnketaEmptyStructuredBlock(`Сімейний стан: дані відсутні
Мати: дані відсутні
Батько: дані відсутні
Довірена особа: дані відсутні
Дитина: дані відсутні`),
    ).toBe(true);
    expect(
      isAnketaEmptyStructuredBlock(`тел:
УБД: немає`),
    ).toBe(true);
  });

  it("keeps filled relatives and phones as real data", () => {
    expect(
      isAnketaEmptyStructuredBlock(
        "Сімейний стан: одружений\nМати: Іванова Марія Петрівна, 1955",
      ),
    ).toBe(false);
    expect(
      isAnketaEmptyStructuredBlock("тел: +380501112233\nУБД: немає"),
    ).toBe(false);
  });
});

describe("isAnketaGapCellEmpty", () => {
  it("treats placeholders and «дані відсутні» as gaps for OCR", () => {
    const row = person("КОВАЛЬЧУК Роман Петрович", {
      rnokpp: "забув",
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      birthDate: "01.01.1990",
    });
    expect(isAnketaGapCellEmpty(row, "rnokpp")).toBe(true);
    expect(isAnketaGapCellEmpty(row, "education")).toBe(true);
    expect(isAnketaGapCellEmpty(row, "birthDate")).toBe(false);
  });

  it("includes empty relatives and additionalInfo templates in gap search", () => {
    const row = person("КОВАЛЬЧУК Роман Петрович", {
      relatives: ANKETA_RELATIVES_EMPTY_TEMPLATE,
      additionalInfo: ANKETA_ADDITIONAL_INFO_EMPTY_TEMPLATE,
      education: "середня",
    });
    expect(isAnketaGapCellTrulyEmpty(row, "relatives")).toBe(true);
    expect(isAnketaGapCellTrulyEmpty(row, "additionalInfo")).toBe(true);
    expect(
      listAnketaEmptyCells([row], ["relatives", "additionalInfo", "education"]).map(
        (cell) => cell.columnId,
      ),
    ).toEqual(["relatives", "additionalInfo"]);
  });

  it("treats a lone hyphen as an empty gap cell", () => {
    const row = person("МАКСИМЕНКО Олексій Євгенійович", {
      education: "-",
      birthDate: "—",
      idDocumentNumber: " - ",
    });
    expect(isAnketaGapCellTrulyEmpty(row, "education")).toBe(true);
    expect(isAnketaGapCellTrulyEmpty(row, "birthDate")).toBe(true);
    expect(isAnketaGapCellTrulyEmpty(row, "idDocumentNumber")).toBe(true);
    expect(
      listAnketaEmptyCells([row], ["education", "birthDate", "idDocumentNumber"]).map(
        (cell) => cell.columnId,
      ),
    ).toEqual(["idDocumentNumber", "birthDate", "education"]);
  });

  it("does not treat a hyphen inside a real value as empty", () => {
    const row = person("МАКСИМЮК Іван Дмитрович", {
      idDocumentNumber: "АА-123456",
      education: "середня",
    });
    expect(isAnketaGapCellTrulyEmpty(row, "idDocumentNumber")).toBe(false);
    expect(
      listAnketaEmptyCells([row], ["idDocumentNumber", "education"]),
    ).toEqual([]);
  });

  it("skips any cell that already has text in gap search", () => {
    const row = person("КОВАЛЬЧУК Роман Петрович", {
      rnokpp: "забув",
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      birthDate: "01.01.1990",
    });
    expect(
      listAnketaEmptyCells([row], ["rnokpp", "education", "birthDate"]).map(
        (cell) => cell.columnId,
      ),
    ).toEqual([]);
  });

  it("excludes a person fully marked «дані відсутні» from gap search", () => {
    const row = person("ТУТОВ Сергій Миколайович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });
    expect(
      listAnketaEmptyCells([row], ["rnokpp", "education"]).map(
        (cell) => cell.columnId,
      ),
    ).toEqual([]);
  });

  it("includes empty selected cells even for people on the missing-questionnaire list", () => {
    const row = person("КОВАЛЬ Іван Петрович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "",
    });
    const exclude = new Set([normalizeAnketaNameKey(row.fullName)]);
    expect(
      listAnketaEmptyCells([row], ["rnokpp", "education"], {
        excludeNameKeys: exclude,
      }).map((cell) => cell.columnId),
    ).toEqual(["education"]);
  });

  it("keeps people without a PDF questionnaire in gap search", () => {
    const withPdf = person("КОВАЛЬЧУК Роман Петрович", { rnokpp: "" });
    withPdf.__rowId = "with-pdf";
    withPdf.externalId = "123";
    const withoutPdf = person("ТУТОВ Сергій Миколайович", { rnokpp: "" });
    withoutPdf.__rowId = "without-pdf";
    withoutPdf.externalId = "456";
    const filtered = filterAnketaRowsForGapSearch([withPdf, withoutPdf], {
      questionnaireMeta: [{ personExternalId: "123", fileName: "anketa.pdf" }],
    });
    expect(filtered.map((row) => row.__rowId)).toEqual([
      withPdf.__rowId,
      withoutPdf.__rowId,
    ]);
    expect(
      listAnketaEmptyCells(filtered, ["rnokpp"]).map((cell) => cell.rowId),
    ).toEqual([withPdf.__rowId, withoutPdf.__rowId]);
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

describe("orderAnketaGapRowsBlankFirst", () => {
  it("puts fully blank rows before rows that already have text", () => {
    const blank = person("ПОРОЖНІЙ", { rnokpp: "", education: "" });
    blank.__rowId = "blank";
    const partial = person("ЧАСТКОВИЙ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "",
    });
    partial.__rowId = "partial";
    const ordered = orderAnketaGapRowsBlankFirst(
      [partial, blank],
      ["rnokpp", "education"],
    );
    expect(ordered.map((row) => row.__rowId)).toEqual(["blank", "partial"]);
  });
});

describe("orderAnketaGapSearchRows", () => {
  it("orders in-staff before rest and blank before partial in each group", () => {
    const inStaffPartial = person("ШТАТ ЧАСТКОВО", {
      rnokpp: "забув",
      education: "",
    });
    inStaffPartial.__rowId = "s-partial";
    const inStaffBlank = person("ШТАТ ПОРОЖНІЙ");
    inStaffBlank.__rowId = "s-blank";
    const restBlank = person("РЕШТА ПОРОЖНЯ");
    restBlank.__rowId = "r-blank";
    const ordered = orderAnketaGapSearchRows(
      [inStaffPartial, restBlank, inStaffBlank],
      new Set(["s-partial", "s-blank"]),
      ["rnokpp", "education"],
    );
    expect(ordered.map((row) => row.__rowId)).toEqual([
      "s-blank",
      "s-partial",
      "r-blank",
    ]);
  });

  it("puts people with a PDF before people without one inside the same staff group", () => {
    const staffNoPdf = person("ШТАТ БЕЗ PDF", { education: "" });
    staffNoPdf.__rowId = "s-no-pdf";
    const staffPdf = person("ШТАТ З PDF", { education: "" });
    staffPdf.__rowId = "s-pdf";
    const ordered = orderAnketaGapSearchRows(
      [staffNoPdf, staffPdf],
      new Set(["s-no-pdf", "s-pdf"]),
      ["education"],
      new Set(["s-pdf"]),
    );
    expect(ordered.map((row) => row.__rowId)).toEqual(["s-pdf", "s-no-pdf"]);
  });
});

describe("gateAnketaGapSearchRowsInStaffFirst", () => {
  it("keeps rest people with empty selected columns in the queue while staff still has gaps", () => {
    const inStaffBlank = person("У ШТАТІ ПОРОЖНІЙ", { rnokpp: "" });
    inStaffBlank.__rowId = "staff-blank";
    const inStaffPartial = person("У ШТАТІ З ТЕКСТОМ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });
    inStaffPartial.__rowId = "staff-partial";
    const rest = person("АРХІВ", { rnokpp: "" });
    rest.__rowId = "rest";
    const rows = [inStaffPartial, inStaffBlank, rest];
    const gated = gateAnketaGapSearchRowsInStaffFirst(
      rows,
      new Set(["staff-blank", "staff-partial"]),
      ["rnokpp"],
    );
    expect(gated.map((row) => row.__rowId)).toEqual(["staff-blank", "rest"]);
  });

  it("includes in-staff rows that still have empty selected columns, not only fully blank", () => {
    const inStaffBlank = person("У ШТАТІ ПОРОЖНІЙ", {
      rnokpp: "",
      education: "",
    });
    inStaffBlank.__rowId = "staff-blank";
    const inStaffPartial = person("У ШТАТІ З ТЕКСТОМ", {
      rnokpp: "1234567890",
      education: "",
    });
    inStaffPartial.__rowId = "staff-partial";
    const rest = person("АРХІВ", { rnokpp: "", education: "" });
    rest.__rowId = "rest";
    const gated = gateAnketaGapSearchRowsInStaffFirst(
      [inStaffPartial, inStaffBlank, rest],
      new Set(["staff-blank", "staff-partial"]),
      ["rnokpp", "education"],
    );
    expect(gated.map((row) => row.__rowId)).toEqual([
      "staff-blank",
      "staff-partial",
      "rest",
    ]);
  });

  it("includes in-staff rows with at least one truly empty selected field", () => {
    const inStaffPartial = person("У ШТАТІ З ТЕКСТОМ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "",
    });
    inStaffPartial.__rowId = "staff-partial";
    const gated = gateAnketaGapSearchRowsInStaffFirst(
      [inStaffPartial],
      new Set(["staff-partial"]),
      ["rnokpp", "education"],
    );
    expect(gated.map((row) => row.__rowId)).toEqual(["staff-partial"]);
  });

  it("excludes rows where every selected field already has text", () => {
    const inStaffPartial = person("У ШТАТІ З ТЕКСТОМ", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: "забув",
    });
    inStaffPartial.__rowId = "staff-partial";
    const gated = gateAnketaGapSearchRowsInStaffFirst(
      [inStaffPartial],
      new Set(["staff-partial"]),
      ["rnokpp", "education"],
    );
    expect(gated).toEqual([]);
  });

  it("includes the rest once all in-staff gaps are filled", () => {
    const inStaff = person("У ШТАТІ", { rnokpp: "1234567890" });
    inStaff.__rowId = "staff";
    const rest = person("АРХІВ", { rnokpp: "" });
    rest.__rowId = "rest";
    const rows = [inStaff, rest];
    const gated = gateAnketaGapSearchRowsInStaffFirst(
      rows,
      new Set(["staff"]),
      ["rnokpp"],
    );
    expect(gated.map((row) => row.__rowId)).toEqual(["rest"]);
  });
});

describe("findNextAnketaEmptyCell walks from the current cell", () => {
  it("returns the next empty cell without wrapping when one remains", () => {
    const first = person("ПЕРШИЙ", { rnokpp: "", education: "середня" });
    first.__rowId = "r1";
    first.__rowNumber = 2;
    const second = person("ДРУГИЙ", { rnokpp: "1234567890", education: "" });
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

  it("starts with in-staff rows when the list is reordered for gap search", () => {
    const archive = person("АРХІВНИЙ", { rnokpp: "" });
    archive.__rowId = "archive";
    archive.__rowNumber = 2;
    const inStaff = person("У ШТАТІ", { rnokpp: "" });
    inStaff.__rowId = "staff";
    inStaff.__rowNumber = 3;
    const ordered = orderAnketaRowsForGapSearch(
      [archive, inStaff],
      new Set(["staff"]),
    );
    const first = findNextAnketaEmptyCell(ordered, null, ["rnokpp"]);
    expect(first?.rowId).toBe("staff");
  });
});
