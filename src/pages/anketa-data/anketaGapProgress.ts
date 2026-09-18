import type { BackendPersonQuestionnaireMeta } from "../../api";
import { readDataCache, writeDataCache } from "../../data/idbDataCache";
import { questionnaireFileMatchesPerson } from "../personnel/personAttachments";
import { isAnketaGapPersonWithoutTrulyEmptyCells } from "./anketaGaps";
import {
  normalizeAnketaExternalIdKey,
  normalizeAnketaNameKey,
} from "./anketaPersonMatch";
import type { AnketaColumnKey, AnketaRow } from "./anketaSheet";

export const ANKETA_GAP_PROGRESS_CACHE_KEY = "anketa:gap-review-progress:v1";

export type AnketaGapReviewEntry = {
  personKey: string;
  fullName: string;
  reviewedAt: string;
  gapColumns: AnketaColumnKey[];
  fingerprint: string;
};

export type AnketaGapReviewMap = Record<string, AnketaGapReviewEntry>;

/** Стабільний ключ особи: ID з таблиці або нормалізоване ПІБ. */
export const buildAnketaGapReviewPersonKey = (
  row: Pick<AnketaRow, "fullName" | "externalId">,
) => {
  const externalId = normalizeAnketaExternalIdKey(row.externalId);
  if (externalId) return `id:${externalId}`;
  const nameKey = normalizeAnketaNameKey(row.fullName);
  return nameKey ? `name:${nameKey}` : "";
};

export const buildAnketaGapQuestionnaireStamp = (
  row: AnketaRow,
  items: BackendPersonQuestionnaireMeta[] | null | undefined,
) => {
  if (!items?.length) return "";
  const spreadsheetId = normalizeAnketaExternalIdKey(row.externalId);
  for (const meta of items) {
    const metaId = String(meta.personExternalId ?? "").trim();
    if (!metaId) continue;
    const normalizedMetaId = normalizeAnketaExternalIdKey(metaId);
    if (spreadsheetId && normalizedMetaId === spreadsheetId) {
      return `${normalizedMetaId}|${String(meta.fileName ?? "").trim()}`;
    }
    if (questionnaireFileMatchesPerson(meta.fileName, [row.fullName])) {
      return `${normalizedMetaId}|${String(meta.fileName ?? "").trim()}`;
    }
  }
  return "";
};

export const buildAnketaGapReviewColumnFingerprint = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
) => {
  const cols = [...gapColumnKeys].sort();
  const values = cols
    .map((key) => `${key}=${String(row[key] ?? "").trim()}`)
    .join("\n");
  return `${cols.join(",")}\n${values}`;
};

export const buildAnketaGapReviewFingerprint = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
  questionnaireStamp: string,
) =>
  `${buildAnketaGapReviewColumnFingerprint(row, gapColumnKeys)}\n${questionnaireStamp}`;

const parseAnketaGapReviewFingerprint = (fingerprint: string) => {
  const lastNewline = fingerprint.lastIndexOf("\n");
  if (lastNewline < 0) {
    return { columnFingerprint: fingerprint, questionnaireStamp: "" };
  }
  return {
    columnFingerprint: fingerprint.slice(0, lastNewline),
    questionnaireStamp: fingerprint.slice(lastNewline + 1),
  };
};

export const buildAnketaGapReviewEntry = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
  questionnaireMeta: BackendPersonQuestionnaireMeta[] | null | undefined,
): AnketaGapReviewEntry | null => {
  const personKey = buildAnketaGapReviewPersonKey(row);
  if (!personKey) return null;
  const stamp = buildAnketaGapQuestionnaireStamp(row, questionnaireMeta);
  return {
    personKey,
    fullName: String(row.fullName ?? "").trim(),
    reviewedAt: new Date().toISOString(),
    gapColumns: [...gapColumnKeys].sort(),
    fingerprint: buildAnketaGapReviewFingerprint(row, gapColumnKeys, stamp),
  };
};

/** Особа вже переглянута і дані/анкета не змінились з моменту перегляду. */
export const isAnketaGapPersonReviewed = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
  questionnaireMeta: BackendPersonQuestionnaireMeta[] | null | undefined,
  progress: AnketaGapReviewMap,
) => {
  const personKey = buildAnketaGapReviewPersonKey(row);
  if (!personKey) return false;
  const entry = progress[personKey];
  if (!entry) return false;
  const currentCols = [...gapColumnKeys].sort().join(",");
  const entryCols = [...entry.gapColumns].sort().join(",");
  if (currentCols !== entryCols) return false;
  const columnFingerprint = buildAnketaGapReviewColumnFingerprint(
    row,
    gapColumnKeys,
  );
  const entryParts = parseAnketaGapReviewFingerprint(entry.fingerprint);
  if (entryParts.columnFingerprint !== columnFingerprint) return false;

  const currentStamp = buildAnketaGapQuestionnaireStamp(row, questionnaireMeta);
  if (!entryParts.questionnaireStamp || !currentStamp) return true;
  return entryParts.questionnaireStamp === currentStamp;
};

/** Пропустити в пошуку лише тих, у кого немає зовсім порожніх вибраних полів. */
export const shouldSkipAnketaGapPerson = (
  row: AnketaRow,
  gapColumnKeys: readonly AnketaColumnKey[],
  _questionnaireMeta?: BackendPersonQuestionnaireMeta[] | null,
  _progress?: AnketaGapReviewMap,
) => isAnketaGapPersonWithoutTrulyEmptyCells(row, gapColumnKeys);

export const filterSkippedAnketaGapRows = (
  rows: AnketaRow[],
  gapColumnKeys: readonly AnketaColumnKey[],
  questionnaireMeta: BackendPersonQuestionnaireMeta[] | null | undefined,
  progress: AnketaGapReviewMap,
) =>
  rows.filter(
    (row) =>
      !shouldSkipAnketaGapPerson(
        row,
        gapColumnKeys,
        questionnaireMeta,
        progress,
      ),
  );

/** @deprecated Використовуйте filterSkippedAnketaGapRows */
export const filterReviewedAnketaGapRows = filterSkippedAnketaGapRows;

export const countActiveAnketaGapReviews = (
  rows: AnketaRow[],
  gapColumnKeys: readonly AnketaColumnKey[],
  questionnaireMeta: BackendPersonQuestionnaireMeta[] | null | undefined,
  progress: AnketaGapReviewMap,
) =>
  rows.filter((row) =>
    isAnketaGapPersonReviewed(row, gapColumnKeys, questionnaireMeta, progress),
  ).length;

export const loadAnketaGapReviewProgress = async (): Promise<AnketaGapReviewMap> => {
  const stored = await readDataCache<AnketaGapReviewMap>(
    ANKETA_GAP_PROGRESS_CACHE_KEY,
  );
  if (!stored || typeof stored !== "object") return {};
  return stored;
};

export const saveAnketaGapReviewProgress = async (progress: AnketaGapReviewMap) => {
  await writeDataCache(ANKETA_GAP_PROGRESS_CACHE_KEY, progress);
};

export const mergeAnketaGapReviewEntry = (
  progress: AnketaGapReviewMap,
  entry: AnketaGapReviewEntry | null,
): AnketaGapReviewMap => {
  if (!entry) return progress;
  return { ...progress, [entry.personKey]: entry };
};
