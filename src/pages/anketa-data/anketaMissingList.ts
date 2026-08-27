import { normalizeAnketaNameKey } from "./anketaPersonMatch";
import type { AnketaRow } from "./anketaSheet";

/** Google Sheet «Відсутні анкети» — ПІБ без анкетних даних (сторінка анкетних даних). */
export const ANKETA_MISSING_SHEET_ID =
  "1EW1nZT_f9QlH8CvFAKQh4vo3M7WrOaCBGcc0WcGzPtE";
export const ANKETA_MISSING_SHEET_LEGACY_ID = ANKETA_MISSING_SHEET_ID;
export const ANKETA_MISSING_SHEET_GID = "0";

export type AnketaMissingListSource = "primary" | "legacy";

const anketaMissingSheetFetchUrls = (sheetId: string) => {
  const gvizJsonQuery = `tqx=out:json&gid=${ANKETA_MISSING_SHEET_GID}`;
  const gvizCsvQuery = `tqx=out:csv&gid=${ANKETA_MISSING_SHEET_GID}`;
  const exportQuery = `format=csv&gid=${ANKETA_MISSING_SHEET_GID}`;
  return [
    `/google-sheets/spreadsheets/d/${sheetId}/gviz/tq?${gvizJsonQuery}`,
    `/google-sheets/spreadsheets/d/${sheetId}/gviz/tq?${gvizCsvQuery}`,
    `/google-sheets/spreadsheets/d/${sheetId}/export?${exportQuery}`,
  ];
};

const formatGvizCell = (value: unknown): string => {
  if (value == null) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  return String(value).trim();
};

const parseCsvFirstCell = (line: string): string => {
  const trimmed = line.trim();
  if (!trimmed) return "";
  if (trimmed.startsWith('"')) {
    const match = trimmed.match(/^"((?:[^"]|"")*)"/);
    if (match) return match[1].replace(/""/g, '"').trim();
  }
  return trimmed.split(",")[0]?.trim() ?? "";
};

const looksLikeMissingListName = (value: string) => {
  const trimmed = value.trim();
  if (!trimmed) return false;
  const lower = trimmed.toLocaleLowerCase("uk-UA");
  if (
    lower === "піб" ||
    lower.startsWith("прізвище") ||
    lower.includes("відсутн") ||
    lower.includes("анкет") ||
    lower.includes("відустн")
  ) {
    return false;
  }
  const parts = trimmed.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return false;
  return /^[\p{L}\s'ʼ’-]+$/u.test(trimmed);
};

const collectMissingNames = (rawNames: string[]): string[] => {
  const names: string[] = [];
  const seen = new Set<string>();
  for (const raw of rawNames) {
    if (!looksLikeMissingListName(raw)) continue;
    const key = normalizeAnketaNameKey(raw);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    names.push(raw);
  }
  return names;
};

export const parseAnketaMissingNames = (csvText: string): string[] =>
  collectMissingNames(
    csvText.split(/\r?\n/).map((line) => parseCsvFirstCell(line)),
  );

export const parseAnketaMissingNameKeys = (csvText: string): Set<string> =>
  new Set(parseAnketaMissingNames(csvText).map((name) => normalizeAnketaNameKey(name)));

const parseGvizMissingNames = (text: string): string[] => {
  const trimmed = text.trim();
  if (!trimmed || trimmed.startsWith("<!DOCTYPE")) {
    throw new Error("Отримано HTML замість даних Google Sheet.");
  }
  const match = trimmed.match(
    /google\.visualization\.Query\.setResponse\(([\s\S]*)\)\s*;?\s*$/,
  );
  const raw = match?.[1] ?? trimmed;
  const payload = JSON.parse(raw) as {
    status?: string;
    table?: {
      rows?: Array<{ c?: Array<{ v?: unknown } | null> }>;
    };
  };
  if (payload.status && payload.status !== "ok") {
    throw new Error("Google Sheet gviz повернув помилку.");
  }
  const rows = payload.table?.rows ?? [];
  return collectMissingNames(
    rows.map((row) => formatGvizCell(row.c?.[0]?.v)),
  );
};

const fetchMissingNamesFromSheet = async (
  sheetId: string,
): Promise<string[]> => {
  let lastError: Error | null = null;
  for (const url of anketaMissingSheetFetchUrls(sheetId)) {
    try {
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) {
        lastError = new Error(`HTTP ${response.status}`);
        continue;
      }
      const text = await response.text();
      if (!text.trim()) {
        lastError = new Error("Порожня відповідь Google Sheet.");
        continue;
      }
      if (text.trimStart().startsWith("<!DOCTYPE")) {
        lastError = new Error("Доступ до Google Sheet обмежено.");
        continue;
      }
      if (text.includes("setResponse(")) {
        const names = parseGvizMissingNames(text);
        if (names.length) return names;
        lastError = new Error("Google Sheet не містить ПІБ.");
        continue;
      }
      const names = parseAnketaMissingNames(text);
      if (names.length) return names;
      lastError = new Error("Google Sheet не містить ПІБ.");
    } catch (error) {
      lastError =
        error instanceof Error ? error : new Error("Невідома помилка мережі");
    }
  }
  throw lastError ?? new Error("Не вдалося завантажити список відсутніх анкет.");
};

export const loadAnketaMissingNames = async (): Promise<{
  names: string[];
  source: AnketaMissingListSource;
}> => {
  const names = await fetchMissingNamesFromSheet(ANKETA_MISSING_SHEET_ID);
  return { names, source: "primary" };
};

export const loadAnketaMissingNameKeys = async (): Promise<{
  keys: Set<string>;
  source: AnketaMissingListSource;
  count: number;
}> => {
  const { names, source } = await loadAnketaMissingNames();
  return {
    keys: new Set(names.map((name) => normalizeAnketaNameKey(name))),
    source,
    count: names.length,
  };
};

export const isAnketaRowMissingQuestionnaire = (
  row: Pick<AnketaRow, "fullName">,
  excludeNameKeys: Set<string> | null | undefined,
) => {
  if (!excludeNameKeys?.size) return false;
  const key = normalizeAnketaNameKey(row.fullName);
  return Boolean(key && excludeNameKeys.has(key));
};
