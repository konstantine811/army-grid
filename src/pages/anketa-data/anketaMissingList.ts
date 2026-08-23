import { normalizeAnketaNameKey } from "./anketaPersonMatch";
import type { AnketaRow } from "./anketaSheet";

/** Google Sheet «Відсутні анкети» — ПІБ без анкетних даних. */
export const ANKETA_MISSING_SHEET_ID =
  "1EW1nZT_f9QlH8CvFAKQh4vo3M7WrOaCBGcc0WcGzPtE";
export const ANKETA_MISSING_SHEET_GID = "0";

export const anketaMissingSheetEditUrl = () =>
  `https://docs.google.com/spreadsheets/d/${ANKETA_MISSING_SHEET_ID}/edit?gid=${ANKETA_MISSING_SHEET_GID}#gid=${ANKETA_MISSING_SHEET_GID}`;

const anketaMissingSheetCsvUrls = () => {
  const query = `format=csv&gid=${ANKETA_MISSING_SHEET_GID}`;
  return [
    `/google-sheets/spreadsheets/d/${ANKETA_MISSING_SHEET_ID}/export?${query}`,
    `https://docs.google.com/spreadsheets/d/${ANKETA_MISSING_SHEET_ID}/export?${query}`,
  ];
};

const looksLikeMissingListName = (value: string) => {
  const trimmed = value.trim();
  if (!trimmed) return false;
  const parts = trimmed.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return false;
  return /^[\p{L}\s'ʼ’-]+$/u.test(trimmed);
};

export const parseAnketaMissingNameKeys = (csvText: string): Set<string> => {
  const keys = new Set<string>();
  for (const line of csvText.split(/\r?\n/)) {
    const raw = line.split(",")[0]?.trim() ?? "";
    if (!looksLikeMissingListName(raw)) continue;
    const key = normalizeAnketaNameKey(raw);
    if (key) keys.add(key);
  }
  return keys;
};

const fetchMissingListCsv = async () => {
  let lastError: Error | null = null;
  for (const url of anketaMissingSheetCsvUrls()) {
    try {
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) {
        lastError = new Error(`HTTP ${response.status} · ${url}`);
        continue;
      }
      const text = await response.text();
      if (!text.trim() || text.trimStart().startsWith("<!DOCTYPE")) {
        lastError = new Error("Отримано HTML замість CSV (доступ обмежено).");
        continue;
      }
      return text;
    } catch (error) {
      lastError =
        error instanceof Error ? error : new Error("Невідома помилка мережі");
    }
  }
  throw lastError ?? new Error("Не вдалося завантажити список відсутніх анкет.");
};

export const loadAnketaMissingNameKeys = async (): Promise<Set<string>> => {
  const csvText = await fetchMissingListCsv();
  return parseAnketaMissingNameKeys(csvText);
};

export const isAnketaRowMissingQuestionnaire = (
  row: Pick<AnketaRow, "fullName">,
  excludeNameKeys: Set<string> | null | undefined,
) => {
  if (!excludeNameKeys?.size) return false;
  const key = normalizeAnketaNameKey(row.fullName);
  return Boolean(key && excludeNameKeys.has(key));
};
