import type { BackendPersonQuestionnaireMeta } from "../../api";
import { questionnaireFileMatchesPerson } from "../personnel/personAttachments";
import {
  buildPersonIdentityFingerprint,
  cleanPersonDisplayName,
  extractPersonCallSign,
} from "../personnel/personnelUtils";
import { isAnketaRowMissingQuestionnaire } from "./anketaMissingList";
import {
  anketaNameKeyVariants,
  expandAnketaNameKeySet,
  normalizeAnketaExternalIdKey,
} from "./anketaPersonMatch";
import {
  ANKETA_COLUMNS,
  ANKETA_SHEET_GID,
  ANKETA_SHEET_ID,
  createEmptyAnketaRow,
  isAnketaColumnReadonly,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";

export type AnketaGapSearchOptions = {
  skipKeys?: Iterable<string> | null;
  excludeNameKeys?: Set<string> | null;
};

export type AnketaEmptyCell = {
  rowId: string;
  columnId: AnketaColumnKey;
  rowNumber: number;
  columnIndex: number; // 0-based sheet column
  header: string;
  a1: string;
};

/** Позначка: в особи немає анкети — порожні вибрані колонки заповнюємо цим текстом. */
export const ANKETA_ABSENT_QUESTIONNAIRE_VALUE = "дані відсутні";

/** Константи для комірок без реальних даних — комірка вважається заповненою. */
export const ANKETA_MISSING_VALUE_PRESETS = [
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  "не вказав",
  "забув",
  "загубив",
  "відмовився вказувати",
  "не має",
  "не памʼятає",
  "невідомо",
  "втрачено",
  "не застосовується",
  "уточнюється",
  "в процесі виготовлення",
] as const;

export type AnketaMissingValuePreset =
  (typeof ANKETA_MISSING_VALUE_PRESETS)[number];

const APPS_SCRIPT_STORAGE_KEY = "army-grid:anketa-apps-script-url";
const GAP_COLUMNS_STORAGE_KEY = "army-grid:anketa-gap-columns.v1";

export const DEFAULT_ANKETA_GAP_COLUMNS: AnketaColumnKey[] = [
  "rnokpp",
  "birthDate",
  "birthPlace",
  "sex",
  "idDocumentNumber",
  "arrivedFrom",
  "serviceType",
  "contractFrom",
  "contractTo",
  "additionalInfo",
  "location",
  "militaryId",
  "conscriptedWhen",
  "conscriptedBy",
  "education",
];

const sanitizeGapColumns = (value: unknown): AnketaColumnKey[] | null => {
  if (!Array.isArray(value)) return null;
  const allowed = new Set(ANKETA_COLUMNS.map((column) => column.key));
  return value
    .map(String)
    .filter(
      (key): key is AnketaColumnKey =>
        allowed.has(key as AnketaColumnKey) &&
        !isAnketaColumnReadonly(key as AnketaColumnKey),
    );
};

export const readAnketaGapColumns = (): AnketaColumnKey[] => {
  try {
    const raw = window.localStorage.getItem(GAP_COLUMNS_STORAGE_KEY);
    if (!raw) return [...DEFAULT_ANKETA_GAP_COLUMNS];
    const parsed = JSON.parse(raw) as unknown;
    // v1 payload: string[] or { keys: string[] }
    const keys = sanitizeGapColumns(
      parsed &&
        typeof parsed === "object" &&
        !Array.isArray(parsed) &&
        "keys" in parsed
        ? (parsed as { keys: unknown }).keys
        : parsed,
    );
    if (!keys) return [...DEFAULT_ANKETA_GAP_COLUMNS];
    return keys;
  } catch {
    return [...DEFAULT_ANKETA_GAP_COLUMNS];
  }
};

export const writeAnketaGapColumns = (keys: AnketaColumnKey[]) => {
  try {
    const sanitized = sanitizeGapColumns(keys) ?? [];
    window.localStorage.setItem(
      GAP_COLUMNS_STORAGE_KEY,
      JSON.stringify({ keys: sanitized }),
    );
  } catch {
    /* ignore quota / private mode */
  }
};

const columnLetter = (index: number) => {
  let n = index + 1;
  let label = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    n = Math.floor((n - 1) / 26);
  }
  return label;
};

export const anketaCellA1 = (rowNumber: number, columnIndex: number) =>
  `${columnLetter(columnIndex)}${rowNumber}`;

export const anketaGoogleCellUrl = (rowNumber: number, columnIndex: number) =>
  `https://docs.google.com/spreadsheets/d/${ANKETA_SHEET_ID}/edit#gid=${ANKETA_SHEET_GID}&range=${anketaCellA1(rowNumber, columnIndex)}`;

export const isAnketaAbsentQuestionnaireValue = (value: unknown) =>
  String(value ?? "")
    .trim()
    .toLocaleLowerCase("uk-UA") === ANKETA_ABSENT_QUESTIONNAIRE_VALUE;

const writableGapColumnKeys = (
  columnKeys?: Iterable<AnketaColumnKey> | null,
) =>
  [...(columnKeys ?? [])].filter(
    (key) =>
      ANKETA_COLUMNS.some((column) => column.key === key) &&
      !isAnketaColumnReadonly(key),
  );

/** Рядок уже позначений як «без анкети» у вибраних колонках. */
export const isAnketaRowMarkedAbsentQuestionnaire = (
  row: AnketaRow,
  columnKeys?: Iterable<AnketaColumnKey> | null,
) => {
  const requested = columnKeys ? [...columnKeys] : [];
  const keys = writableGapColumnKeys(
    requested.length ? requested : ANKETA_COLUMNS.map((column) => column.key),
  );
  return keys.some((key) => isAnketaAbsentQuestionnaireValue(row[key]));
};

export type AnketaAbsentQuestionnaireFill = {
  rowId: string;
  rowNumber: number;
  columnId: AnketaColumnKey;
  fullName: string;
  externalId: string;
};

export const anketaRowHasQuestionnairePdf = (
  row: AnketaRow,
  items: BackendPersonQuestionnaireMeta[] | null | undefined,
) => {
  if (!items?.length) return false;
  const spreadsheetId = normalizeAnketaExternalIdKey(row.externalId);
  const fingerprint = buildPersonIdentityFingerprint(
    cleanPersonDisplayName(row.fullName).replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim(),
    String(row.birthDate ?? "").trim(),
    extractPersonCallSign(row.fullName),
  );
  for (const meta of items) {
    const metaId = String(meta.personExternalId ?? "").trim();
    if (!metaId) continue;
    if (spreadsheetId && normalizeAnketaExternalIdKey(metaId) === spreadsheetId) {
      return true;
    }
    if (fingerprint && metaId === fingerprint) return true;
    if (questionnaireFileMatchesPerson(meta.fileName, [row.fullName])) {
      return true;
    }
  }
  return false;
};

const shouldFillAbsentQuestionnaireRow = (
  row: AnketaRow,
  excludeNameKeys: Set<string> | null | undefined,
  hasQuestionnaire?: (row: AnketaRow) => boolean,
) => {
  if (hasQuestionnaire?.(row)) return false;
  const listedMissing = isAnketaRowMissingQuestionnaire(row, excludeNameKeys);
  if (listedMissing) return true;
  if (hasQuestionnaire) return true;
  return !excludeNameKeys?.size;
};

/** Порожні комірки вибраних колонок у осіб без анкети (список / немає PDF). */
export const collectAbsentQuestionnaireCellFills = (
  rows: AnketaRow[],
  columnKeys: Iterable<AnketaColumnKey> | null | undefined,
  excludeNameKeys?: Set<string> | null,
  hasQuestionnaire?: (row: AnketaRow) => boolean,
): AnketaAbsentQuestionnaireFill[] => {
  const keys = writableGapColumnKeys(columnKeys);
  if (!keys.length) return [];

  const fills: AnketaAbsentQuestionnaireFill[] = [];
  for (const row of rows) {
    if (!shouldFillAbsentQuestionnaireRow(row, excludeNameKeys, hasQuestionnaire)) {
      continue;
    }
    for (const columnId of keys) {
      const value = String(row[columnId] ?? "").trim();
      if (value) continue;
      fills.push({
        rowId: row.__rowId,
        rowNumber: row.__rowNumber,
        columnId,
        fullName: row.fullName,
        externalId: row.externalId,
      });
    }
  }
  return fills;
};

/** «дані відсутні» у вибраних колонках, якщо анкета вже є — стерти, щоб пошук міг заповнити. */
export const collectAbsentQuestionnaireCellClears = (
  rows: AnketaRow[],
  columnKeys: Iterable<AnketaColumnKey> | null | undefined,
  hasQuestionnaire: (row: AnketaRow) => boolean,
): AnketaAbsentQuestionnaireFill[] => {
  const keys = writableGapColumnKeys(columnKeys);
  if (!keys.length) return [];

  const clears: AnketaAbsentQuestionnaireFill[] = [];
  for (const row of rows) {
    if (!hasQuestionnaire(row)) continue;
    for (const columnId of keys) {
      if (!isAnketaAbsentQuestionnaireValue(row[columnId])) continue;
      clears.push({
        rowId: row.__rowId,
        rowNumber: row.__rowNumber,
        columnId,
        fullName: row.fullName,
        externalId: row.externalId,
      });
    }
  }
  return clears;
};

const nextAnketaSheetRowNumber = (rows: AnketaRow[]) =>
  rows.reduce((max, row) => Math.max(max, Number(row.__rowNumber) || 0), 1) + 1;

/** Імена зі списку «без анкет», яких ще немає в таблиці — нові рядки. */
export const buildAbsentQuestionnaireAnketaRows = (
  existingRows: AnketaRow[],
  missingNames: string[],
  columnKeys: Iterable<AnketaColumnKey> | null | undefined,
): AnketaRow[] => {
  const keys = writableGapColumnKeys(columnKeys);
  const present = expandAnketaNameKeySet(
    existingRows.map((row) => row.fullName).filter(Boolean),
  );
  const created: AnketaRow[] = [];
  let rowNumber = nextAnketaSheetRowNumber(existingRows);
  const seen = new Set<string>();

  for (const rawName of missingNames) {
    const name = String(rawName ?? "").replace(/\s+/g, " ").trim();
    const variants = anketaNameKeyVariants(name);
    if (![...variants].some(Boolean)) continue;
    if ([...variants].some((key) => present.has(key) || seen.has(key))) continue;
    for (const key of variants) seen.add(key);

    const row = createEmptyAnketaRow(rowNumber);
    row.fullName = name;
    for (const columnId of keys) {
      row[columnId] = ANKETA_ABSENT_QUESTIONNAIRE_VALUE;
    }
    created.push(row);
    rowNumber += 1;
  }
  return created;
};

export const applyAbsentQuestionnaireFillsToRows = (
  rows: AnketaRow[],
  fills: AnketaAbsentQuestionnaireFill[],
) => {
  if (!fills.length) return rows;
  const byRow = new Map<string, AnketaColumnKey[]>();
  for (const fill of fills) {
    const list = byRow.get(fill.rowId) ?? [];
    list.push(fill.columnId);
    byRow.set(fill.rowId, list);
  }
  return rows.map((row) => {
    const columns = byRow.get(row.__rowId);
    if (!columns?.length) return row;
    let next = row;
    for (const columnId of columns) {
      next = { ...next, [columnId]: ANKETA_ABSENT_QUESTIONNAIRE_VALUE };
    }
    return next;
  });
};

export const applyAbsentQuestionnaireClearsToRows = (
  rows: AnketaRow[],
  clears: AnketaAbsentQuestionnaireFill[],
) => {
  if (!clears.length) return rows;
  const byRow = new Map<string, AnketaColumnKey[]>();
  for (const clear of clears) {
    const list = byRow.get(clear.rowId) ?? [];
    list.push(clear.columnId);
    byRow.set(clear.rowId, list);
  }
  return rows.map((row) => {
    const columns = byRow.get(row.__rowId);
    if (!columns?.length) return row;
    let next = row;
    for (const columnId of columns) {
      next = { ...next, [columnId]: "" };
    }
    return next;
  });
};

export const listAnketaEmptyCells = (
  rows: AnketaRow[],
  columnKeys?: Iterable<AnketaColumnKey> | null,
  options?: AnketaGapSearchOptions,
): AnketaEmptyCell[] => {
  const allowed = columnKeys
    ? new Set(
        [...columnKeys].filter((key) =>
          ANKETA_COLUMNS.some((column) => column.key === key),
        ),
      )
    : null;
  if (allowed && allowed.size === 0) return [];
  const skip = options?.skipKeys ? new Set([...options.skipKeys]) : null;
  const excludeNameKeys = options?.excludeNameKeys ?? null;

  const gaps: AnketaEmptyCell[] = [];
  for (const row of rows) {
    if (isAnketaRowMissingQuestionnaire(row, excludeNameKeys)) continue;
    if (isAnketaRowMarkedAbsentQuestionnaire(row, allowed)) continue;
    ANKETA_COLUMNS.forEach((column, columnIndex) => {
      if (isAnketaColumnReadonly(column.key)) return;
      if (allowed && !allowed.has(column.key)) return;
      const value = String(row[column.key] ?? "").trim();
      if (value) return;
      const key = `${row.__rowId}:${column.key}`;
      if (skip?.has(key)) return;
      gaps.push({
        rowId: row.__rowId,
        columnId: column.key,
        rowNumber: row.__rowNumber,
        columnIndex,
        header: column.header,
        a1: anketaCellA1(row.__rowNumber, columnIndex),
      });
    });
  }
  return gaps;
};

export type AnketaGapStats = {
  emptyCells: number;
  personsWithGaps: number;
  /** Особи, у яких хоча б одне вибране поле ще зовсім порожнє (не «дані відсутні»). */
  personsWithBlankFields: number;
  /** Особи, у яких усі вибрані поля ще зовсім порожні. */
  personsFullyBlank: number;
  blankCells: number;
  totalRows: number;
  columns: number;
};

const selectedWritableKeys = (
  columnKeys?: Iterable<AnketaColumnKey> | null,
) =>
  writableGapColumnKeys(
    columnKeys && [...columnKeys].length
      ? columnKeys
      : ANKETA_COLUMNS.map((column) => column.key),
  );

/** Скільки службовців ще мають зовсім порожні вибрані поля. */
export const countAnketaBlankFieldPersons = (
  rows: AnketaRow[],
  columnKeys?: Iterable<AnketaColumnKey> | null,
) => {
  const keys = selectedWritableKeys(columnKeys);
  if (!keys.length) {
    return { personsWithBlankFields: 0, personsFullyBlank: 0, blankCells: 0 };
  }

  let personsWithBlankFields = 0;
  let personsFullyBlank = 0;
  let blankCells = 0;

  for (const row of rows) {
    let blanks = 0;
    for (const key of keys) {
      if (String(row[key] ?? "").trim()) continue;
      blanks += 1;
    }
    if (!blanks) continue;
    blankCells += blanks;
    personsWithBlankFields += 1;
    if (blanks === keys.length) personsFullyBlank += 1;
  }

  return { personsWithBlankFields, personsFullyBlank, blankCells };
};

/** Статистика пропусків лише по вибраних колонках (без «відкладених»). */
export const summarizeAnketaGaps = (
  rows: AnketaRow[],
  columnKeys?: Iterable<AnketaColumnKey> | null,
  options?: Pick<AnketaGapSearchOptions, "excludeNameKeys">,
): AnketaGapStats => {
  const keys = columnKeys ? [...new Set(columnKeys)] : [];
  const gaps = listAnketaEmptyCells(rows, keys, options);
  const personIds = new Set(gaps.map((gap) => gap.rowId));
  const blanks = countAnketaBlankFieldPersons(rows, keys);
  return {
    emptyCells: gaps.length,
    personsWithGaps: personIds.size,
    personsWithBlankFields: blanks.personsWithBlankFields,
    personsFullyBlank: blanks.personsFullyBlank,
    blankCells: blanks.blankCells,
    totalRows: rows.length,
    columns: keys.length,
  };
};

export const anketaGapSkipKey = (gap: Pick<AnketaEmptyCell, "rowId" | "columnId">) =>
  `${gap.rowId}:${gap.columnId}`;

export const findNextAnketaEmptyCell = (
  rows: AnketaRow[],
  current: AnketaEmptyCell | null,
  columnKeys?: Iterable<AnketaColumnKey> | null,
  options?: AnketaGapSearchOptions,
): AnketaEmptyCell | null => {
  const gaps = listAnketaEmptyCells(rows, columnKeys, options);
  if (!gaps.length) return null;
  if (!current) return gaps[0] ?? null;

  const currentIndex = gaps.findIndex(
    (gap) =>
      gap.rowId === current.rowId && gap.columnId === current.columnId,
  );
  if (currentIndex < 0) {
    // Current was skipped/filled — take first gap after current position in sheet order.
    const after = gaps.find(
      (gap) =>
        gap.rowNumber > current.rowNumber ||
        (gap.rowNumber === current.rowNumber &&
          gap.columnIndex > current.columnIndex),
    );
    return after ?? gaps[0] ?? null;
  }
  return gaps[currentIndex + 1] ?? gaps[0] ?? null;
};

/** First empty cell of the next person (skips remaining gaps of the current row). */
export const findNextAnketaPersonEmptyCell = (
  rows: AnketaRow[],
  current: AnketaEmptyCell | null,
  columnKeys?: Iterable<AnketaColumnKey> | null,
  options?: AnketaGapSearchOptions,
): AnketaEmptyCell | null => {
  const gaps = listAnketaEmptyCells(rows, columnKeys, options);
  if (!gaps.length) return null;
  if (!current) return gaps[0] ?? null;

  const currentIndex = gaps.findIndex(
    (gap) =>
      gap.rowId === current.rowId && gap.columnId === current.columnId,
  );
  const startFrom = currentIndex >= 0 ? currentIndex + 1 : 0;

  for (let index = startFrom; index < gaps.length; index += 1) {
    const gap = gaps[index];
    if (gap && gap.rowId !== current.rowId) return gap;
  }
  for (let index = 0; index < startFrom; index += 1) {
    const gap = gaps[index];
    if (gap && gap.rowId !== current.rowId) return gap;
  }
  return null;
};

/** All empty-cell skip keys for one person (to defer the whole row). */
export const listAnketaPersonGapSkipKeys = (
  rows: AnketaRow[],
  rowId: string,
  columnKeys?: Iterable<AnketaColumnKey> | null,
  options?: AnketaGapSearchOptions,
) => {
  const personRows = rows.filter((row) => row.__rowId === rowId);
  return listAnketaEmptyCells(personRows, columnKeys, options).map(
    anketaGapSkipKey,
  );
};

export const updateAnketaRowCell = (
  rows: AnketaRow[],
  rowId: string,
  columnId: AnketaColumnKey,
  value: string,
): AnketaRow[] => {
  if (isAnketaColumnReadonly(columnId)) return rows;
  return rows.map((row) =>
    row.__rowId === rowId ? { ...row, [columnId]: value } : row,
  );
};

export const readAnketaAppsScriptUrl = () => {
  try {
    return window.localStorage.getItem(APPS_SCRIPT_STORAGE_KEY)?.trim() || "";
  } catch {
    return "";
  }
};

export const writeAnketaAppsScriptUrl = (url: string) => {
  try {
    const next = url.trim();
    if (next) window.localStorage.setItem(APPS_SCRIPT_STORAGE_KEY, next);
    else window.localStorage.removeItem(APPS_SCRIPT_STORAGE_KEY);
  } catch {
    /* ignore */
  }
};

/**
 * Optional write-back via a Google Apps Script Web App.
 * Deploy script as web app (execute as you, anyone with link) with doPost handler.
 */
export const pushAnketaCellToGoogle = async (payload: {
  a1: string;
  value: string;
  sheetId?: string;
  gid?: string;
}) => {
  const endpoint = readAnketaAppsScriptUrl();
  if (!endpoint) {
    throw new Error(
      "Немає URL Apps Script. Вставте посилання вебзастосунку в налаштуваннях синхронізації.",
    );
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      sheetId: payload.sheetId ?? ANKETA_SHEET_ID,
      gid: payload.gid ?? ANKETA_SHEET_GID,
      a1: payload.a1,
      value: payload.value,
    }),
  });

  if (!response.ok) {
    throw new Error(`Google sync HTTP ${response.status}`);
  }

  const text = await response.text();
  try {
    const json = JSON.parse(text) as { ok?: boolean; error?: string };
    if (json.error) throw new Error(json.error);
    if (json.ok === false) throw new Error("Apps Script повернув помилку.");
  } catch (error) {
    if (error instanceof SyntaxError) return;
    throw error;
  }
};

export const ANKETA_APPS_SCRIPT_TEMPLATE = `/**
 * Google Apps Script — вставте в Extensions → Apps Script таблиці «Анкети».
 * Deploy → New deployment → Web app
 * Execute as: Me
 * Who has access: Anyone
 * Скопіюйте URL вебзастосунку в army-grid (поле синхронізації).
 */
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var sheetId = body.sheetId || SpreadsheetApp.getActiveSpreadsheet().getId();
    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = ss.getSheets()[0];
    if (!body.a1) throw new Error('a1 required');
    sheet.getRange(String(body.a1)).setValue(body.value == null ? '' : body.value);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, a1: body.a1 }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
`;
