import {
  ANKETA_COLUMNS,
  ANKETA_SHEET_GID,
  ANKETA_SHEET_ID,
  isAnketaColumnReadonly,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";

export type AnketaEmptyCell = {
  rowId: string;
  columnId: AnketaColumnKey;
  rowNumber: number;
  columnIndex: number; // 0-based sheet column
  header: string;
  a1: string;
};

/** Константи для комірок без реальних даних — комірка вважається заповненою. */
export const ANKETA_MISSING_VALUE_PRESETS = [
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

export const listAnketaEmptyCells = (
  rows: AnketaRow[],
  columnKeys?: Iterable<AnketaColumnKey> | null,
  options?: { skipKeys?: Iterable<string> | null },
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

  const gaps: AnketaEmptyCell[] = [];
  for (const row of rows) {
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
  totalRows: number;
  columns: number;
};

/** Статистика пропусків лише по вибраних колонках (без «відкладених»). */
export const summarizeAnketaGaps = (
  rows: AnketaRow[],
  columnKeys?: Iterable<AnketaColumnKey> | null,
): AnketaGapStats => {
  const keys = columnKeys ? [...new Set(columnKeys)] : [];
  const gaps = listAnketaEmptyCells(rows, keys);
  const personIds = new Set(gaps.map((gap) => gap.rowId));
  return {
    emptyCells: gaps.length,
    personsWithGaps: personIds.size,
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
  options?: { skipKeys?: Iterable<string> | null },
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
  options?: { skipKeys?: Iterable<string> | null },
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
  options?: { skipKeys?: Iterable<string> | null },
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
