import { describe, expect, it } from "vitest";
import {
  buildAnketaGapReviewEntry,
  buildAnketaGapReviewFingerprint,
  filterSkippedAnketaGapRows,
  isAnketaGapPersonReviewed,
  mergeAnketaGapReviewEntry,
  shouldSkipAnketaGapPerson,
} from "./anketaGapProgress";
import { ANKETA_ABSENT_QUESTIONNAIRE_VALUE } from "./anketaGaps";
import { createEmptyAnketaRow } from "./anketaSheet";

const person = (name: string, extra: Record<string, string> = {}) => {
  const row = createEmptyAnketaRow(2);
  row.__rowId = `row-${name}`;
  row.externalId = "100";
  return { ...row, fullName: name, ...extra };
};

describe("anketaGapProgress", () => {
  it("remembers a person after review even when cells contain «дані відсутні»", () => {
    const row = person("ТУТОВ Сергій Миколайович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });
    const gapColumns = ["rnokpp", "education"] as const;
    const meta = [{ personExternalId: "100", fileName: "tutov.pdf" }];
    const entry = buildAnketaGapReviewEntry(row, gapColumns, meta);
    const progress = mergeAnketaGapReviewEntry({}, entry);
    expect(
      isAnketaGapPersonReviewed(row, gapColumns, meta, progress),
    ).toBe(true);
    expect(
      filterSkippedAnketaGapRows([row], gapColumns, meta, progress),
    ).toEqual([]);
  });

  it("keeps a reviewed person in the queue while selected columns are still empty", () => {
    const row = person("НАУМОВ Дмитро Сергійович", {
      rnokpp: "",
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });
    const gapColumns = ["rnokpp", "education"] as const;
    const entry = buildAnketaGapReviewEntry(row, gapColumns, []);
    const progress = mergeAnketaGapReviewEntry({}, entry);
    expect(
      isAnketaGapPersonReviewed(row, gapColumns, [], progress),
    ).toBe(true);
    expect(
      filterSkippedAnketaGapRows([row], gapColumns, [], progress).map(
        (item) => item.__rowId,
      ),
    ).toEqual([row.__rowId]);
    expect(shouldSkipAnketaGapPerson(row, gapColumns, [], progress)).toBe(false);
  });

  it("skips a person without truly empty selected fields even without review journal", () => {
    const row = person("ТУТОВ Сергій Миколайович", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
      education: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });
    const gapColumns = ["rnokpp", "education"] as const;
    expect(
      shouldSkipAnketaGapPerson(row, gapColumns, [], {}),
    ).toBe(true);
    expect(
      filterSkippedAnketaGapRows([row], gapColumns, [], {}),
    ).toEqual([]);
  });

  it("still matches review when questionnaire meta loads after mark", () => {
    const row = person("КОВАЛЬ Іван", {
      rnokpp: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
    });
    const gapColumns = ["rnokpp"] as const;
    const entry = buildAnketaGapReviewEntry(row, gapColumns, []);
    const progress = mergeAnketaGapReviewEntry({}, entry);
    const meta = [{ personExternalId: "100", fileName: "new.pdf" }];
    expect(
      isAnketaGapPersonReviewed(row, gapColumns, meta, progress),
    ).toBe(true);
  });

  it("returns the person to the queue when a gap column value changes", () => {
    const row = person("КОВАЛЬ Іван", { rnokpp: "" });
    const gapColumns = ["rnokpp"] as const;
    const entry = buildAnketaGapReviewEntry(row, gapColumns, []);
    const progress = mergeAnketaGapReviewEntry({}, entry);
    const updated = person("КОВАЛЬ Іван", { rnokpp: "1234567890" });
    updated.__rowId = row.__rowId;
    updated.externalId = row.externalId;
    expect(
      isAnketaGapPersonReviewed(updated, gapColumns, [], progress),
    ).toBe(false);
  });

  it("returns the person to the queue when the questionnaire stamp changes", () => {
    const row = person("КОВАЛЬ Іван", { rnokpp: "1234567890" });
    const gapColumns = ["rnokpp"] as const;
    const oldMeta = [{ personExternalId: "100", fileName: "old.pdf" }];
    const entry = buildAnketaGapReviewEntry(row, gapColumns, oldMeta);
    const progress = mergeAnketaGapReviewEntry({}, entry);
    const newMeta = [{ personExternalId: "100", fileName: "new.pdf" }];
    expect(
      isAnketaGapPersonReviewed(row, gapColumns, newMeta, progress),
    ).toBe(false);
  });

  it("builds stable fingerprints for the same row state", () => {
    const row = person("ТЕСТ", { rnokpp: "1111111111" });
    const fp1 = buildAnketaGapReviewFingerprint(row, ["rnokpp"], "100|a.pdf");
    const fp2 = buildAnketaGapReviewFingerprint(row, ["rnokpp"], "100|a.pdf");
    expect(fp1).toBe(fp2);
  });
});
