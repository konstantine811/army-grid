import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { MORNING_GENERAL_LIST_COLUMN_LABELS } from "../personnel/personnelUtils";
import {
  FIGHTER_STATUS_FIELDS,
  buildFighterStatusAdditions,
  findFighterStatusAddition,
} from "../personnel/fighterStatusImport";
import { normalizeDatasetKey } from "../../shared/format";
import { parseCsv } from "../anketa-data/anketaSheet";
import type { CellValue, ExcelRow, ExcelSheetSnapshot } from "../../excelRoundTrip";
import { readRosterColumnValue } from "./rosterSourceSnapshot";

/** Google Sheet «Штатка». */
export const STAFF_SHEET_ID = "1dqorfOj7TTZt6sBBV1mc_dupxRkUSyeRjxJjlV07A38";
/** Перший аркуш «1.ОС Загальний список». */
export const STAFF_SHEET_GID = "0";
/** Аркуш «Статус бійців» (дати виходу / повернення). */
export const STAFF_SHEET_FIGHTER_STATUS_GID = "390853017";
export const STAFF_SHEET_ROSTER_GID = "46231430";

export const staffSheetEditUrl = () =>
  `https://docs.google.com/spreadsheets/d/${STAFF_SHEET_ID}/edit?usp=sharing`;

const APPS_SCRIPT_STORAGE_KEY = "army-grid:staff-sheet-apps-script-url";
const LAST_SYNC_STORAGE_KEY = "army-grid:staff-sheet-roster-sync-at";

/** Kyiv local hours for automatic morning / lunch / evening roster sync. */
export const STAFF_SHEET_SYNC_MORNING_HOUR = 7;
export const STAFF_SHEET_SYNC_LUNCH_HOUR = 13;
export const STAFF_SHEET_SYNC_EVENING_HOUR = 18;

export const STAFF_SHEET_SYNC_HOURS_KYIV = [
  STAFF_SHEET_SYNC_MORNING_HOUR,
  STAFF_SHEET_SYNC_LUNCH_HOUR,
  STAFF_SHEET_SYNC_EVENING_HOUR,
] as const;

export const staffSheetGvizUrls = (gid = STAFF_SHEET_GID) => {
  const query = `tqx=out:json&gid=${gid}`;
  return [
    `/google-sheets/spreadsheets/d/${STAFF_SHEET_ID}/gviz/tq?${query}`,
    `https://docs.google.com/spreadsheets/d/${STAFF_SHEET_ID}/gviz/tq?${query}`,
  ];
};

/** Roster column numbers → sequential columns A… in «Штатка» (row 1 headers). */
export const STAFF_SHEET_ROSTER_COLUMNS = [
  1, 2, 3, 4, 5, 8, 10, 11, 12, 13, 14, 15, 21, 22, 23, 24, 25, 26, 27, 28, 29,
  31, 32, 33, 34,
] as const;

export const STAFF_SHEET_HEADERS = [
  "№",
  "Підрозділ",
  "Взвод",
  "Відділення",
  "Посада",
  "ШПК факт",
  "Анкета",
  "Військовий квиток",
  "Мобілізація/контракт",
  "Звання",
  "ПІБ",
  "Позивний",
  "СТАТУС",
  "Тип В\\С",
  "Статус БГ",
  "БЗВП/БРЕЗ",
  "Наявність БЗВП",
  "Курс БЗВП",
  "Відрядження (БРЕЗ)",
  "Обмеження",
  "В якому підрозділі",
  "Місце перебування",
  "Примітки",
  "Напрямок",
  "Примітка 3",
] as const;

const hasPersonName = (row: EjournalPreviewRow) => {
  const name = readRosterColumnValue(row, 14);
  return name.trim().length >= 3;
};

export const rosterRowsToStaffSheetValues = (
  rows: EjournalPreviewRow[],
): string[][] =>
  rows
    .filter(hasPersonName)
    .map((row, index) =>
      STAFF_SHEET_ROSTER_COLUMNS.map((columnNumber, columnIndex) => {
        if (columnIndex === 0) return String(index + 1);
        return readRosterColumnValue(row, columnNumber);
      }),
    );

/** 1-based Excel column numbers for «Штатка» push (same order as values rows). */
export const STAFF_SHEET_PUSH_COLUMN_NUMBERS = [
  ...STAFF_SHEET_ROSTER_COLUMNS,
] as number[];

/** Лише колонки, які доповнюємо в імпортованій «Штатці». */
export const STAFF_SHEET_ANKETA_VK_COLUMNS = [10, 11] as const;

/** Колонки для доповнення з анкет / ООС / ВК ТПВ: Анкета, ВК, дата, рік, вік, ІПН. */
export const STAFF_SHEET_ENRICHMENT_COLUMNS = [10, 11, 16, 17, 18, 19] as const;
export const STAFF_SHEET_ENRICHMENT_VALUE_INDEX = {
  anketa: 0,
  militaryId: 1,
  birthDate: 2,
  birthYear: 3,
  fullYears: 4,
  inn: 5,
} as const;
export const STAFF_SHEET_ENRICHMENT_START_ROW = 2;
export const staffSheetBirthDateColumn = 16;

export const columnNumberToSheetLetter = (columnNumber: number): string => {
  let column = columnNumber;
  let letters = "";
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    column = Math.floor((column - 1) / 26);
  }
  return letters;
};

export const staffSheetBirthYearFormula = (
  excelRowNumber: number,
  birthDateColumn = staffSheetBirthDateColumn,
) => {
  const letter = columnNumberToSheetLetter(birthDateColumn);
  return `=IF(${letter}${excelRowNumber}="";"";YEAR(${letter}${excelRowNumber}))`;
};

export const staffSheetFullYearsFormula = (
  excelRowNumber: number,
  birthDateColumn = staffSheetBirthDateColumn,
) => {
  const letter = columnNumberToSheetLetter(birthDateColumn);
  return `=IF(${letter}${excelRowNumber}="";"";DATEDIF(${letter}${excelRowNumber};TODAY();"Y"))`;
};

/** Формули для Excel (.xlsx) — коми замість крапок з комою. */
export const staffSheetBirthYearExcelFormula = (
  excelRowNumber: number,
  birthDateColumn = staffSheetBirthDateColumn,
) => {
  const letter = columnNumberToSheetLetter(birthDateColumn);
  return `IF(${letter}${excelRowNumber}="","",YEAR(${letter}${excelRowNumber}))`;
};

export const staffSheetFullYearsExcelFormula = (
  excelRowNumber: number,
  birthDateColumn = staffSheetBirthDateColumn,
) => {
  const letter = columnNumberToSheetLetter(birthDateColumn);
  return `IF(${letter}${excelRowNumber}="","",DATEDIF(${letter}${excelRowNumber},TODAY(),"Y"))`;
};

export const STAFF_SHEET_EXPORT_COLUMN_NUMBERS = Object.keys(
  MORNING_GENERAL_LIST_COLUMN_LABELS,
)
  .map((key) => Number(key))
  .filter((column) => Number.isFinite(column) && column > 0)
  .sort((left, right) => left - right);

export const pushStaffSheetEnrichmentToGoogle = async (
  entries: Array<{ excelRowNumber: number; values: string[] }>,
  options?: {
    sheetId?: string;
    gid?: string;
    columns?: readonly number[];
  },
) => {
  const endpoint = readStaffAppsScriptUrl();
  if (!endpoint) {
    throw new Error(
      "Немає URL Apps Script для «Штатка». Налаштуйте його на сторінці «Заповнення Excel».",
    );
  }
  if (!entries.length) {
    throw new Error("Немає рядків для оновлення Google Sheet.");
  }

  const columns = options?.columns ?? STAFF_SHEET_ENRICHMENT_COLUMNS;
  const numberColumns = columns.filter((column) => [17, 18, 19].includes(column));
  const textColumns = columns.filter((column) => [10, 11, 16].includes(column));

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      sheetId: options?.sheetId ?? STAFF_SHEET_ID,
      gid: options?.gid ?? STAFF_SHEET_ROSTER_GID,
      mode: "bulkColumns",
      startRow: STAFF_SHEET_ENRICHMENT_START_ROW,
      rowNumbers: entries.map((entry) => entry.excelRowNumber),
      columns: [...columns],
      values: entries.map((entry) => entry.values),
      numberColumns,
      textColumns,
      clearBackground: columns.some((column) => [16, 17, 18, 19].includes(column)),
      clearAfterRow: false,
    }),
  });

  if (!response.ok) {
    throw new Error(`Google sync HTTP ${response.status}`);
  }

  const text = await response.text();
  try {
    const json = JSON.parse(text) as { ok?: boolean; error?: string; rows?: number };
    if (json.error) throw new Error(json.error);
    if (json.ok === false) throw new Error("Apps Script повернув помилку.");
    return json.rows ?? entries.length;
  } catch (error) {
    if (error instanceof SyntaxError) return entries.length;
    throw error;
  }
};

export const readStaffAppsScriptUrl = () => {
  try {
    return window.localStorage.getItem(APPS_SCRIPT_STORAGE_KEY)?.trim() || "";
  } catch {
    return "";
  }
};

export const writeStaffAppsScriptUrl = (url: string) => {
  try {
    const next = url.trim();
    if (next) window.localStorage.setItem(APPS_SCRIPT_STORAGE_KEY, next);
    else window.localStorage.removeItem(APPS_SCRIPT_STORAGE_KEY);
  } catch {
    /* ignore */
  }
};

export const pushStaffSheetToGoogle = async (
  values: string[][],
  options?: { sheetId?: string; gid?: string },
) => {
  const endpoint = readStaffAppsScriptUrl();
  if (!endpoint) {
    throw new Error(
      "Немає URL Apps Script для «Штатка». Вставте посилання вебзастосунку нижче.",
    );
  }
  if (!values.length) {
    throw new Error("Немає рядків для оновлення Google Sheet.");
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      sheetId: options?.sheetId ?? STAFF_SHEET_ID,
      // Пишемо в аркуш «1.ОС Загальний список», не в «Статус бійців».
      gid: options?.gid ?? STAFF_SHEET_ROSTER_GID,
      // bulkColumns: значення йдуть у рідні колонки Excel (ПІБ=14, Статус=21…),
      // а не зсувом у A..Y — інакше ламається ранковий звіт після ручного .xlsx.
      mode: "bulkColumns",
      startRow: 2,
      columns: STAFF_SHEET_PUSH_COLUMN_NUMBERS,
      values,
      clearAfterRow: true,
    }),
  });

  if (!response.ok) {
    throw new Error(`Google sync HTTP ${response.status}`);
  }

  const text = await response.text();
  try {
    const json = JSON.parse(text) as { ok?: boolean; error?: string; rows?: number };
    if (json.error) throw new Error(json.error);
    if (json.ok === false) throw new Error("Apps Script повернув помилку.");
    return json.rows ?? values.length;
  } catch (error) {
    if (error instanceof SyntaxError) return values.length;
    throw error;
  }
};

const cellText = (value: unknown) => String(value ?? "").trim();

const normalizeStaffSheetHeaderKey = (value: string) =>
  value
    .trim()
    .toLocaleLowerCase("uk-UA")
    .replace(/\s+/g, " ")
    .replace(/[\\]/g, "");

/** Прибирає назву колонки з комірки («ПІБ ІВАНОВ…» → «ІВАНОВ…», «Анкета» → «»). */
export const stripStaffSheetColumnHeaderPrefix = (
  value: string,
  headerLabel: string,
) => {
  const text = value.trim();
  if (!text) return "";
  const headerNorm = normalizeStaffSheetHeaderKey(headerLabel);
  const textNorm = normalizeStaffSheetHeaderKey(text);
  if (textNorm === headerNorm) return "";
  const prefix = `${headerNorm} `;
  if (textNorm.startsWith(prefix)) {
    return text.slice(text.toLocaleLowerCase("uk-UA").indexOf(headerNorm) + headerNorm.length).trim();
  }
  const headerWords = headerLabel.trim().split(/\s+/);
  const textWords = text.split(/\s+/);
  if (
    headerWords.length <= 3 &&
    normalizeStaffSheetHeaderKey(textWords.slice(0, headerWords.length).join(" ")) ===
      headerNorm
  ) {
    return textWords.slice(headerWords.length).join(" ").trim();
  }
  return text;
};

export const sanitizeStaffSheetCellValue = (
  value: string,
  columnNumber?: number,
) => {
  const label =
    columnNumber != null
      ? MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber] ?? ""
      : "";
  const stripped = label
    ? stripStaffSheetColumnHeaderPrefix(value, label)
    : value.trim();
  return stripped.trim();
};

const formatSheetCell = (value: unknown): string => {
  if (value == null) return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(value);
  }
  const text = String(value).trim();
  const gvizDate = text.match(/^Date\((\d+),(\d+),(\d+)(?:,(\d+),(\d+)(?:,(\d+))?)?\)$/);
  if (gvizDate) {
    const year = gvizDate[1];
    const month = String(Number(gvizDate[2]) + 1).padStart(2, "0");
    const day = String(Number(gvizDate[3])).padStart(2, "0");
    return `${day}.${month}.${year}`;
  }
  return text;
};

const looksLikePersonName = (value: string) => {
  const text = value
    .trim()
    .replace(
      /\s*\(\s*\d{1,2}\.\d{1,2}\.\d{4}\s*(?:р\.?\s*н\.?)?\s*\)\s*$/iu,
      "",
    )
    .trim();
  if (text.length < 5 || text.length > 80) return false;
  if (/\d/.test(text)) return false;
  const parts = text.split(/\s+/).filter(Boolean);
  if (parts.length < 2 || parts.length > 4) return false;
  if (
    /^(посада|підрозділ|взвод|відділення|звання|статус|напрям|прізвищ|ім[ʼ'’]?я|дата|позивн|збро|приміт|засоби|повна)\b/i.test(
      text,
    )
  ) {
    return false;
  }
  // Посади / довгі описи, не ПІБ.
  if (
    /посада|командир|заступник|згідно\s+штат|прізвищ|власне\s+ім|психолог|офіцер\s+гру|радіотелеф|механік|водій|стрілець|оператор|навідник|санітар|батальйон|відділенн|взвод\b/i.test(
      text,
    )
  ) {
    return false;
  }
  // Типовий ПІБ: 2–4 слова з великої літери (або весь CAPS).
  return parts.every((part) =>
    /^[A-ZА-ЯІЇЄҐ][A-Za-zА-Яа-яІіЇїЄєҐґʼ'’-]*$/u.test(part),
  );
};

const SECOND_JOB_TITLE_SPLIT =
  /\s+(?=(?:заступник|помічник|начальник|головний)(?:\s|$))/iu;

const splitStackedGvizCellValues = (
  rest: string,
  expectedCount: number,
  headerLabel: string,
): string[] => {
  const text = rest.trim();
  if (!text) return Array.from({ length: expectedCount }, () => "");
  if (expectedCount <= 1) return [text];

  if (headerLabel === "№") {
    const words = text.split(/\s+/).filter(Boolean);
    if (words.length >= expectedCount) return words.slice(0, expectedCount);
    if (words.length > 1) {
      return [
        ...words,
        ...Array.from({ length: expectedCount - words.length }, () => ""),
      ].slice(0, expectedCount);
    }
  }

  if (
    headerLabel === "ШПК факт" ||
    headerLabel === "Звання" ||
    headerLabel === "Категорія складу"
  ) {
    const words = text.split(/\s+/).filter(Boolean);
    if (
      words.length === 2 &&
      /^(?:лейтенант|сержант|капітан)$/iu.test(words[1] ?? "")
    ) {
      return [text, ...Array.from({ length: expectedCount - 1 }, () => "")];
    }
    if (
      words.length === expectedCount &&
      words.every(
        (word) =>
          word.length <= 24 &&
          !/батальйон|командир|заступник|піхотного/i.test(word),
      )
    ) {
      return words;
    }
  }

  if (
    headerLabel === "Посада" ||
    headerLabel === "Повна посада" ||
    headerLabel.includes("осада")
  ) {
    const parts = text.split(SECOND_JOB_TITLE_SPLIT).map((part) => part.trim());
    if (parts.length >= expectedCount) {
      return parts.slice(0, expectedCount);
    }
    if (parts.length === 2 && expectedCount === 2) return parts;
  }

  if (looksLikePersonName(text)) {
    return [text, ...Array.from({ length: expectedCount - 1 }, () => "")];
  }

  if (text.split(/\s+/).length === expectedCount) {
    return text.split(/\s+/).filter(Boolean);
  }

  return [text, ...Array.from({ length: expectedCount - 1 }, () => "")];
};

const parseGvizStackedColumnLabel = (
  label: string,
  headerLabel: string,
  expectedCount: number,
) => {
  const text = label.trim();
  if (!text) {
    return {
      header: headerLabel,
      values: Array.from({ length: expectedCount }, () => ""),
    };
  }
  const stripped = stripStaffSheetColumnHeaderPrefix(text, headerLabel);
  if (!stripped) {
    return {
      header: headerLabel,
      values: Array.from({ length: expectedCount }, () => ""),
    };
  }
  return {
    header: headerLabel,
    values: splitStackedGvizCellValues(stripped, expectedCount, headerLabel),
  };
};

const countGvizStackedLabelRows = (cols: Array<{ label?: string }>) => {
  const numberLabel = formatSheetCell(cols[0]?.label ?? "");
  const numberTokens = stripStaffSheetColumnHeaderPrefix(numberLabel, "№")
    .split(/\s+/)
    .filter(Boolean);
  if (numberTokens.length > 1) return numberTokens.length;

  const posadaLabel = formatSheetCell(cols[4]?.label ?? "");
  const posadaParts = stripStaffSheetColumnHeaderPrefix(posadaLabel, "Посада")
    .split(SECOND_JOB_TITLE_SPLIT)
    .map((part) => part.trim())
    .filter(Boolean);
  if (posadaParts.length > 1) return posadaParts.length;

  return 1;
};

const detectGvizPibColumnIndex = (cols: Array<{ label?: string }>) => {
  for (let index = 0; index < cols.length; index += 1) {
    const header =
      MORNING_GENERAL_LIST_COLUMN_LABELS[index + 1] ??
      `Колонка ${index + 1}`;
    const stripped = stripStaffSheetColumnHeaderPrefix(
      formatSheetCell(cols[index]?.label ?? ""),
      header === "ПІБ" ? "ПІБ" : header,
    );
    if (looksLikePersonName(stripped)) return index;
  }
  return 13;
};

const isGvizStackedLabelsMode = (
  cols: Array<{ label?: string }>,
  body: string[][],
) => {
  const pibIndex = detectGvizPibColumnIndex(cols);
  const labelPib = stripStaffSheetColumnHeaderPrefix(
    formatSheetCell(cols[pibIndex]?.label ?? ""),
    "ПІБ",
  );
  if (!looksLikePersonName(labelPib)) return false;
  if (!body.length) return countGvizStackedLabelRows(cols) > 1;
  const bodyPib = formatSheetCell(body[0]?.[pibIndex] ?? "");
  if (!looksLikePersonName(bodyPib)) return countGvizStackedLabelRows(cols) > 1;
  const normalizeName = (value: string) =>
    value.trim().toLocaleUpperCase("uk-UA").replace(/\s+/g, " ");
  return normalizeName(labelPib) !== normalizeName(bodyPib);
};

const expandGvizStackedLabels = (
  cols: Array<{ label?: string }>,
  body: string[][],
): string[][] => {
  const stackedCount = countGvizStackedLabelRows(cols);
  const columnCount = cols.length;
  const syntheticHeader = Array.from({ length: columnCount }, (_, index) => {
    const fallback =
      MORNING_GENERAL_LIST_COLUMN_LABELS[index + 1] ?? `Колонка ${index + 1}`;
    const label = formatSheetCell(cols[index]?.label ?? "");
    if (!label.trim()) return fallback;
    const headerNorm = normalizeStaffSheetHeaderKey(fallback);
    if (normalizeStaffSheetHeaderKey(label).startsWith(headerNorm)) {
      return fallback;
    }
    return fallback;
  });

  const stackedRows = Array.from({ length: stackedCount }, (_, rowIndex) =>
    Array.from({ length: columnCount }, (_, colIndex) => {
      const header =
        MORNING_GENERAL_LIST_COLUMN_LABELS[colIndex + 1] ??
        `Колонка ${colIndex + 1}`;
      const label = formatSheetCell(cols[colIndex]?.label ?? "");
      const { values } = parseGvizStackedColumnLabel(
        label,
        header,
        stackedCount,
      );
      return sanitizeStaffSheetCellValue(values[rowIndex] ?? "", colIndex + 1);
    }),
  );

  return [syntheticHeader, ...stackedRows, ...body];
};

const isStaffSheetLabelOnlyRow = (values: Record<string, string>) => {
  const texts = Object.entries(values)
    .filter(([key]) => /^column_\d+$/i.test(key))
    .map(([, value]) => value.trim())
    .filter(Boolean);
  if (!texts.length) return false;
  const headerLabels = new Set(
    Object.values(MORNING_GENERAL_LIST_COLUMN_LABELS).map((label) =>
      normalizeStaffSheetHeaderKey(label),
    ),
  );
  return texts.every((text) =>
    headerLabels.has(normalizeStaffSheetHeaderKey(text)),
  );
};

const looksLikeHeaderRow = (row: string[]) => {
  const normalized = row.map((cell) =>
    cell.trim().toLocaleLowerCase("uk-UA").replace(/\s+/g, " "),
  );
  const hasName = normalized.some(
    (cell) =>
      cell === "піб" ||
      cell.startsWith("піб ") ||
      cell.includes("прізви"),
  );
  const hasPosada = normalized.some((cell) => cell.includes("посада"));
  const hasExitDate = normalized.some(
    (cell) => cell.includes("дата") && cell.includes("вих"),
  );
  const hasStatus = normalized.some((cell) => cell.includes("статус"));
  const hasCallsign = normalized.some((cell) => cell.includes("позивн"));
  return (
    (hasName && hasPosada) ||
    (hasName && hasExitDate) ||
    (hasName && hasStatus) ||
    (hasName && hasCallsign)
  );
};

export const FIGHTER_STATUS_FALLBACK_HEADERS = [
  "Підрозділ",
  "Напрямок",
  "Посада згідно штату (коротка)",
  "Військове звання",
  "Прізвище, власне ім’я",
  "Позивний",
  "Дата виходу",
  "Дата повернення або 200/300/500",
  "Днів",
  "Статус",
  "Зброя",
  "Засоби зв'язку",
  "ПРИМІТКА",
] as const;

export const tableToExcelSheetSnapshot = (
  sheetName: string,
  table: string[][],
  sheetIndex = 0,
): ExcelSheetSnapshot => {
  if (!table.length) {
    return {
      sheetIndex,
      sheetName,
      rawRows: [],
      headerRows: [],
      rows: [],
      columnCount: 0,
      columnIndexes: [],
      dataStartRow: 2,
    };
  }

  const first = table[0]?.map((cell) => formatSheetCell(cell)) ?? [];
  const headerIsReal = looksLikeHeaderRow(first);
  const headerValues: CellValue[] = headerIsReal
    ? first
    : [...FIGHTER_STATUS_FALLBACK_HEADERS];
  const dataRows = headerIsReal ? table.slice(1) : table;
  const columnCount = Math.max(
    headerValues.length,
    ...dataRows.map((row) => row.length),
    1,
  );
  const paddedHeader = Array.from(
    { length: columnCount },
    (_, index) => headerValues[index] ?? "",
  );

  const rows: ExcelRow[] = dataRows.map((cells, index) => ({
    id: `${sheetName}-${index + 1}`,
    excelRowNumber: index + (headerIsReal ? 2 : 1),
    values: Array.from({ length: columnCount }, (_, columnIndex) =>
      formatSheetCell(cells[columnIndex]),
    ),
    source: "merged",
  }));

  return {
    sheetIndex,
    sheetName,
    rawRows: [paddedHeader, ...rows.map((row) => row.values)],
    headerRows: [paddedHeader],
    rows,
    columnCount,
    columnIndexes: Array.from({ length: columnCount }, (_, index) => index),
    dataStartRow: 2,
  };
};

const resolveHeaderRosterColumn = (
  header: string,
  columnIndex: number,
): number => {
  const trimmed = header.trim();
  if (!trimmed) return columnIndex + 1;

  const exactStaff = STAFF_SHEET_HEADERS.findIndex(
    (label) =>
      label.toLocaleLowerCase("uk-UA") === trimmed.toLocaleLowerCase("uk-UA"),
  );
  if (exactStaff >= 0) return STAFF_SHEET_ROSTER_COLUMNS[exactStaff];

  const wanted = trimmed.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ");
  for (const [columnNumber, label] of Object.entries(
    MORNING_GENERAL_LIST_COLUMN_LABELS,
  )) {
    if (label.toLocaleLowerCase("uk-UA") === wanted) {
      return Number(columnNumber);
    }
  }

  for (const [columnNumber, label] of Object.entries(
    MORNING_GENERAL_LIST_COLUMN_LABELS,
  )) {
    const known = label.toLocaleLowerCase("uk-UA");
    if (wanted.startsWith(known) || known.startsWith(wanted)) {
      return Number(columnNumber);
    }
  }

  return columnIndex + 1;
};

export type StaffSheetRosterImportPayload = {
  name: string;
  sourceFileName: string;
  notes: string;
  sheets: Array<{
    name: string;
    sheetIndex: number;
    columns: Array<{
      key: string;
      label: string;
      order: number;
      originalIndex: number;
      letter: string;
    }>;
    rows: Array<{ excelRowNumber: number; values: Record<string, string> }>;
  }>;
  personCount: number;
  fighterStatusCount: number;
  fighterStatusMatched: number;
  source: "apps-script" | "gviz";
};

export const buildStaffSheetRosterImportPayload = (
  table: string[][],
  meta: {
    source: "apps-script" | "gviz";
    sourceLabel: string;
    fighterStatusTable?: string[][] | null;
    /** Експорт: зберегти рядки підрозділів / порожні позиції, не лише ПІБ. */
    includeAllRows?: boolean;
  },
): StaffSheetRosterImportPayload => {
  if (!table.length) {
    throw new Error("Таблиця «Штатка» порожня.");
  }

  const firstRow = (table[0] ?? []).map((cell) => cellText(cell));
  const strippedFirstRow = firstRow.map((cell, index) =>
    sanitizeStaffSheetCellValue(cell, index + 1),
  );
  const firstRowLooksLikeHeader = looksLikeHeaderRow(firstRow);
  const treatFirstRowAsHeader =
    firstRowLooksLikeHeader &&
    !strippedFirstRow.some((cell) => looksLikePersonName(cell));
  const headerIsReal = treatFirstRowAsHeader;
  const headerRow = treatFirstRowAsHeader ? firstRow : [];
  const dataRows = treatFirstRowAsHeader ? table.slice(1) : table;

  const columnCount = Math.max(
    headerRow.length,
    ...dataRows.map((row) => row.length),
    STAFF_SHEET_HEADERS.length,
  );

  // Навіть з заголовками gviz може зсунути колонки — ПІБ шукаємо по вмісту.
  const detectedPibIndex = detectPersonNameColumnIndex(dataRows);
  const headerPibIndex = headerIsReal
    ? headerRow.findIndex((cell) => {
        const lower = cell.toLocaleLowerCase("uk-UA");
        return lower === "піб" || lower.startsWith("піб ");
      })
    : -1;

  const columnNumberOffset =
    !headerIsReal && detectedPibIndex >= 0
      ? 14 - (detectedPibIndex + 1)
      : 0;

  const rosterColumns = Array.from({ length: columnCount }, (_, index) => {
    const header = headerRow[index] ?? "";
    let rosterColumn = header
      ? resolveHeaderRosterColumn(header, index)
      : Math.max(1, index + 1 + columnNumberOffset);

    // Якщо заголовок «ПІБ» стоїть не там, де реальні імена — вирівнюємо.
    if (
      headerIsReal &&
      detectedPibIndex >= 0 &&
      headerPibIndex >= 0 &&
      headerPibIndex !== detectedPibIndex &&
      index === detectedPibIndex
    ) {
      rosterColumn = 14;
    }
    if (
      headerIsReal &&
      detectedPibIndex >= 0 &&
      headerPibIndex === index &&
      headerPibIndex !== detectedPibIndex
    ) {
      // Хибний «ПІБ»-заголовок не чіпаємо як 14 — лишаємо інший номер.
      rosterColumn =
        MORNING_GENERAL_LIST_COLUMN_LABELS[index + 1] && index + 1 !== 14
          ? index + 1
          : Math.max(1, index + 1);
    }

    const label =
      header ||
      MORNING_GENERAL_LIST_COLUMN_LABELS[rosterColumn] ||
      `Колонка ${rosterColumn}`;
    const baseKey =
      normalizeDatasetKey(label) || `column_${rosterColumn}`;
    return {
      rosterColumn,
      label,
      baseKey,
      order: index,
      originalIndex: rosterColumn - 1,
    };
  });

  const usedKeys = new Map<string, number>();
  const columns = rosterColumns.map((column) => {
    const count = usedKeys.get(column.baseKey) ?? 0;
    usedKeys.set(column.baseKey, count + 1);
    const key =
      count === 0 ? column.baseKey : `${column.baseKey}_${count + 1}`;
    return {
      key,
      label: column.label,
      order: column.order,
      originalIndex: column.originalIndex,
      letter: "",
      rosterColumn: column.rosterColumn,
    };
  });

  const fighterStatusSheet = meta.fighterStatusTable?.length
    ? tableToExcelSheetSnapshot(
        "Статус бійців",
        meta.fighterStatusTable,
        1,
      )
    : null;
  const fighterStatusAdditions = fighterStatusSheet
    ? buildFighterStatusAdditions(fighterStatusSheet)
    : new Map<string, Record<string, unknown>>();

  const statusColumns = FIGHTER_STATUS_FIELDS.map((field, index) => ({
    key: field.key,
    label: field.label,
    order: columns.length + index,
    originalIndex: columns.length + index,
    letter: "",
  }));

  let pibColumnIndex = columns.findIndex(
    (column) =>
      column.rosterColumn === 14 ||
      /^піб$/i.test(column.label.trim()) ||
      column.key === "піб",
  );
  if (detectedPibIndex >= 0) {
    pibColumnIndex = detectedPibIndex;
  }
  const mappedStatusColumnIndex = rosterColumns.findIndex(
    (column) => column.rosterColumn === 21,
  );
  const statusColumnIndex =
    mappedStatusColumnIndex >= 0
      ? mappedStatusColumnIndex
      : detectStatusColumnIndex(dataRows);

  // У «Штатці» рядок 1 — заголовок; дані з рядка 2 навіть якщо gviz
  // «з’їв» заголовки в labels.
  const excelRowOffset = 2;

  let fighterStatusMatched = 0;
  const rows = dataRows
    .map((cells, index) => {
      const values: Record<string, string> = {};
      columns.forEach((column, columnIndex) => {
        const text = sanitizeStaffSheetCellValue(
          formatSheetCell(cells[columnIndex]),
          column.rosterColumn,
        );
        if (!text) return;
        if (columnIndex === pibColumnIndex) return;
        if (statusColumnIndex >= 0 && columnIndex === statusColumnIndex) return;
        values[column.key] = text;
        values[`column_${column.rosterColumn}`] = text;
      });
      const fullName =
        (pibColumnIndex >= 0
          ? formatSheetCell(cells[pibColumnIndex])
          : values.column_14 || values.піб || "") || "";
      const hasFullName = looksLikePersonName(fullName);
      const isStaffPosition = [5, 8, 13].some((columnNumber) =>
        Boolean(values[`column_${columnNumber}`]?.trim()),
      );
      const hasAnyContent = cells.some((cell) =>
        sanitizeStaffSheetCellValue(formatSheetCell(cell)).trim(),
      );
      if (!meta.includeAllRows && !hasFullName && !isStaffPosition) {
        return null;
      }
      if (meta.includeAllRows && !hasAnyContent) {
        return {
          excelRowNumber: index + excelRowOffset,
          values: {},
        };
      }
      if (hasFullName) {
        values.column_14 = fullName;
        values.піб = fullName;
        values.ПІБ = fullName;
      }

      const formattedCells = cells.map((cell) => formatSheetCell(cell));
      const status =
        mappedStatusColumnIndex >= 0
          ? formattedCells[mappedStatusColumnIndex] ?? ""
          : pickStatusFromRow(
              formattedCells,
              statusColumnIndex,
              pibColumnIndex,
            );
      if (status) {
        values.column_21 = status;
        values.статус = status;
        values.Статус = status;
      }

      const callsignIndex =
        pibColumnIndex >= 0 && pibColumnIndex + 1 < cells.length
          ? pibColumnIndex + 1
          : -1;
      if (callsignIndex >= 0 && callsignIndex !== statusColumnIndex) {
        const callsign = formatSheetCell(cells[callsignIndex]);
        if (
          callsign &&
          !looksLikePersonName(callsign) &&
          callsign.length <= 32 &&
          !/^\d{2}\.\d{2}\.\d{4}$/.test(callsign)
        ) {
          values.column_15 = callsign;
          values.позивний = callsign;
          values.Позивний = callsign;
        }
      }

      const statusAddition = hasFullName
        ? findFighterStatusAddition(values, fighterStatusAdditions)
        : null;
      if (statusAddition) {
        fighterStatusMatched += 1;
        Object.entries(statusAddition).forEach(([key, value]) => {
          const text = formatSheetCell(value);
          if (text) values[key] = text;
        });
      }
      return {
        excelRowNumber: index + excelRowOffset,
        values,
      };
    })
    .filter(
      (
        row,
      ): row is { excelRowNumber: number; values: Record<string, string> } => {
        if (!row) return false;
        if (isStaffSheetLabelOnlyRow(row.values)) return false;
        return true;
      },
    );

  if (!rows.length) {
    throw new Error("У «Штатці» не знайдено рядків з ПІБ.");
  }

  const stamp = new Date().toISOString().slice(0, 16).replace("T", "_");
  return {
    name: `Штатка_${stamp}`,
    sourceFileName: meta.sourceLabel,
    notes: `Синхронізація особового складу з Google Sheet «Штатка» (${meta.source})${
      fighterStatusAdditions.size
        ? ` · Статус бійців: ${fighterStatusAdditions.size} записів, привʼязано ${fighterStatusMatched}`
        : ""
    }.`,
    sheets: [
      {
        name: "1.ОС Загальний список",
        sheetIndex: 0,
        columns: [
          ...columns.map(({ key, label, order, originalIndex, letter }) => ({
            key,
            label,
            order,
            originalIndex,
            letter,
          })),
          ...statusColumns,
        ],
        rows,
      },
    ],
    personCount: rows.filter((row) => Boolean(row.values.column_14)).length,
    fighterStatusCount: fighterStatusAdditions.size,
    fighterStatusMatched,
    source: meta.source,
  };
};

type PulledStaffTables = {
  rosterTable: string[][];
  fighterStatusTable: string[][] | null;
  source: "apps-script" | "gviz";
};

const pullTableViaAppsScript = async (
  gid: string,
): Promise<string[][]> => {
  const endpoint = readStaffAppsScriptUrl();
  if (!endpoint) {
    throw new Error("NO_APPS_SCRIPT");
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      sheetId: STAFF_SHEET_ID,
      gid,
      mode: "pull",
    }),
  });
  if (!response.ok) {
    throw new Error(`Apps Script HTTP ${response.status}`);
  }
  const json = (await response.json()) as {
    ok?: boolean;
    error?: string;
    values?: unknown[][];
    sheets?: Array<{ gid?: string; name?: string; values?: unknown[][] }>;
  };
  if (json.error) throw new Error(json.error);
  if (json.sheets?.length) {
    const match =
      json.sheets.find((sheet) => String(sheet.gid) === String(gid)) ??
      json.sheets[0];
    if (!match?.values?.length) {
      throw new Error("Apps Script повернув порожню таблицю.");
    }
    return match.values.map((row) =>
      (Array.isArray(row) ? row : []).map((cell) => formatSheetCell(cell)),
    );
  }
  if (!json.values?.length) {
    throw new Error("Apps Script повернув порожню таблицю.");
  }
  return json.values.map((row) =>
    (Array.isArray(row) ? row : []).map((cell) => formatSheetCell(cell)),
  );
};

const pullAllTablesViaAppsScript = async (): Promise<PulledStaffTables> => {
  const endpoint = readStaffAppsScriptUrl();
  if (!endpoint) {
    throw new Error("NO_APPS_SCRIPT");
  }

  const response = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      sheetId: STAFF_SHEET_ID,
      mode: "pull",
      gids: [STAFF_SHEET_ROSTER_GID, STAFF_SHEET_FIGHTER_STATUS_GID],
    }),
  });
  if (!response.ok) {
    throw new Error(`Apps Script HTTP ${response.status}`);
  }
  const json = (await response.json()) as {
    ok?: boolean;
    error?: string;
    values?: unknown[][];
    sheets?: Array<{ gid?: string; name?: string; values?: unknown[][] }>;
  };
  if (json.error) throw new Error(json.error);

  const toTable = (values: unknown[][] | undefined) =>
    (values ?? []).map((row) =>
      (Array.isArray(row) ? row : []).map((cell) => formatSheetCell(cell)),
    );

  if (json.sheets?.length) {
    const byGid = new Map(
      json.sheets.map((sheet) => [String(sheet.gid ?? ""), sheet]),
    );
    const byName = json.sheets.find((sheet) =>
      /статус\s*бійц/i.test(String(sheet.name ?? "")),
    );
    const roster =
      byGid.get(STAFF_SHEET_ROSTER_GID) ??
      byGid.get(STAFF_SHEET_GID) ??
      json.sheets[0];
    const fighter =
      byGid.get(STAFF_SHEET_FIGHTER_STATUS_GID) ?? byName ?? null;
    const rosterTable = toTable(roster?.values);
    if (!rosterTable.length) {
      throw new Error("Apps Script: порожній «Загальний список».");
    }
    return {
      rosterTable,
      fighterStatusTable: fighter?.values?.length ? toTable(fighter.values) : null,
      source: "apps-script",
    };
  }

  // Old script without multi-sheet pull — roster only, then fetch status separately.
  const rosterTable = toTable(json.values);
  if (!rosterTable.length) {
    throw new Error("Apps Script повернув порожню таблицю.");
  }
  let fighterStatusTable: string[][] | null = null;
  try {
    fighterStatusTable = await pullTableViaAppsScript(
      STAFF_SHEET_FIGHTER_STATUS_GID,
    );
  } catch {
    fighterStatusTable = null;
  }
  return { rosterTable, fighterStatusTable, source: "apps-script" };
};

const parseGvizJson = (text: string): string[][] => {
  const trimmed = text.trim();
  const match = trimmed.match(
    /google\.visualization\.Query\.setResponse\(([\s\S]*)\)\s*;?\s*$/,
  );
  const raw = match?.[1] ?? trimmed;
  const payload = JSON.parse(raw) as {
    table?: {
      cols?: Array<{ label?: string }>;
      rows?: Array<{ c?: Array<{ v?: unknown } | null> }>;
    };
    status?: string;
  };
  if (payload.status && payload.status !== "ok") {
    throw new Error("Google Sheet gviz повернув помилку.");
  }
  const cols = payload.table?.cols ?? [];
  const rows = payload.table?.rows ?? [];
  const header = cols.map((col) => formatSheetCell(col.label ?? ""));
  const body = rows.map((row) =>
    (row.c ?? []).map((cell) => formatSheetCell(cell?.v)),
  );
  const headerHasPerson = header.some((cell) => looksLikePersonName(cell));
  const headerOk =
    looksLikeHeaderRow(header) && !headerHasPerson;

  if (isGvizStackedLabelsMode(cols, body)) {
    return expandGvizStackedLabels(cols, body);
  }

  if (headerOk) {
    return [header, ...body];
  }

  // gviz кладе перший рядок даних у labels — повертаємо його як data.
  // Не плутати зі справжніми заголовками без «Посада» (ПІБ+Статус).
  if (headerHasPerson) {
    const syntheticHeader = cols.map((_, index) => {
      const fallback =
        MORNING_GENERAL_LIST_COLUMN_LABELS[index + 1] ??
        `Колонка ${index + 1}`;
      return fallback;
    });
    const firstDataRow = header.map((cell, index) =>
      sanitizeStaffSheetCellValue(cell, index + 1),
    );
    return [syntheticHeader, firstDataRow, ...body];
  }

  if (looksLikeHeaderRow(header)) {
    return [header, ...body];
  }

  // Слабкі labels (лише «ПІБ» / «Колонка N») — не як data-рядок.
  if (header.some((cell) => cell.trim()) && !headerHasPerson) {
    const weakHeader = header.map((cell, index) =>
      cell.trim() || `Колонка ${index + 1}`,
    );
    return [weakHeader, ...body];
  }

  return body;
};

export const parseStaffSheetGvizResponse = parseGvizJson;

const isLikelyRosterStatusValue = (value: string) => {
  const text = value.trim().toLocaleLowerCase("uk-UA");
  if (!text || text.length > 80) return false;
  return (
    text === "в строю" ||
    text.includes("новоприбулий") ||
    text.includes("відком") ||
    text.includes("відрядж") ||
    text.includes("сзч") ||
    text.includes("ліку") ||
    text.includes("лікарн") ||
    text.includes("шпит") ||
    text.includes("загиб") ||
    text.includes("помер") ||
    text.includes("відпустк") ||
    text.includes("звільнен") ||
    text.includes("полон") ||
    /(?:^|\D)(?:200|300|500)(?:\D|$)/.test(text) ||
    text.includes("тцк") ||
    text.includes("без віс")
  );
};

/** Індекс колонки з найбільшою кількістю ПІБ у рядках. */
const detectPersonNameColumnIndex = (dataRows: string[][]) => {
  const columnCount = Math.max(0, ...dataRows.map((row) => row.length));
  let bestIndex = -1;
  let bestCount = 0;
  for (let index = 0; index < columnCount; index += 1) {
    let count = 0;
    for (const row of dataRows) {
      if (looksLikePersonName(formatSheetCell(row[index]))) count += 1;
    }
    if (count > bestCount) {
      bestCount = count;
      bestIndex = index;
    }
  }
  return bestIndex;
};

/** Індекс колонки з типовими статусами «В строю» тощо. */
const detectStatusColumnIndex = (dataRows: string[][]) => {
  const columnCount = Math.max(0, ...dataRows.map((row) => row.length));
  let bestIndex = -1;
  let bestCount = 0;
  for (let index = 0; index < columnCount; index += 1) {
    let count = 0;
    for (const row of dataRows) {
      if (isLikelyRosterStatusValue(formatSheetCell(row[index]))) count += 1;
    }
    if (count > bestCount) {
      bestCount = count;
      bestIndex = index;
    }
  }
  return bestCount >= 3 ? bestIndex : -1;
};

const pickStatusFromRow = (
  cells: string[],
  preferredIndex: number,
  pibIndex: number,
) => {
  if (preferredIndex >= 0) {
    const preferred = formatSheetCell(cells[preferredIndex]);
    if (isLikelyRosterStatusValue(preferred)) return preferred;
  }
  // Перший рядок з gviz-labels часто коротший — шукаємо статус по всьому рядку.
  for (let index = 0; index < cells.length; index += 1) {
    if (index === pibIndex) continue;
    const text = formatSheetCell(cells[index]);
    if (isLikelyRosterStatusValue(text)) return text;
  }
  return preferredIndex >= 0 ? formatSheetCell(cells[preferredIndex]) : "";
};

const pullTableViaGviz = async (gid: string): Promise<string[][]> => {
  let lastError: Error | null = null;
  for (const url of staffSheetGvizUrls(gid)) {
    try {
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) {
        lastError = new Error(`HTTP ${response.status}`);
        continue;
      }
      const text = await response.text();
      if (!text.trim() || text.trimStart().startsWith("<!DOCTYPE")) {
        lastError = new Error("Отримано HTML замість даних Google Sheet.");
        continue;
      }
      if (text.includes("setResponse(") || text.trimStart().startsWith("{")) {
        return parseGvizJson(text);
      }
      const csvTable = parseCsv(text);
      if (csvTable.length) return csvTable;
      lastError = new Error("Порожня відповідь gviz.");
    } catch (error) {
      lastError =
        error instanceof Error ? error : new Error("Помилка мережі gviz");
    }
  }
  throw lastError ?? new Error("Не вдалося прочитати «Штатку» з Google.");
};

export const pullStaffSheetGvizRosterTable = async (): Promise<string[][]> => {
  try {
    return await pullTableViaGviz(STAFF_SHEET_ROSTER_GID);
  } catch {
    return pullTableViaGviz(STAFF_SHEET_GID);
  }
};

const pullAllTablesViaGviz = async (): Promise<PulledStaffTables> => {
  const rosterTable = await pullStaffSheetGvizRosterTable();
  let fighterStatusTable: string[][] | null = null;
  try {
    fighterStatusTable = await pullTableViaGviz(
      STAFF_SHEET_FIGHTER_STATUS_GID,
    );
  } catch (error) {
    console.warn("Не вдалося прочитати аркуш «Статус бійців»", error);
  }
  return { rosterTable, fighterStatusTable, source: "gviz" };
};

export const pullStaffSheetRosterImportPayload =
  async (options?: {
    source?: "auto" | "gviz";
  }): Promise<StaffSheetRosterImportPayload> => {
    let tables: PulledStaffTables;
    if (options?.source === "gviz") {
      tables = await pullAllTablesViaGviz();
    } else {
      try {
        tables = await pullAllTablesViaAppsScript();
      } catch (error) {
        if (error instanceof Error && error.message !== "NO_APPS_SCRIPT") {
          console.warn("Staff sheet Apps Script pull failed, trying gviz", error);
        }
        tables = await pullAllTablesViaGviz();
      }
    }

    return buildStaffSheetRosterImportPayload(tables.rosterTable, {
      source: tables.source,
      sourceLabel: `Штатка (${tables.source}) · ${STAFF_SHEET_ID}`,
      fighterStatusTable: tables.fighterStatusTable,
    });
  };

const base64ToArrayBuffer = (base64: string) => {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes.buffer;
};

/** Справжній .xlsx «Штатки» (з кольорами/ширинами), не зібраний з комірок. */
export const downloadStaffSheetXlsxBytes = async (options?: {
  sheetId?: string;
}): Promise<{ fileName: string; fileData: ArrayBuffer }> => {
  const sheetId = options?.sheetId ?? STAFF_SHEET_ID;
  const endpoint = readStaffAppsScriptUrl();

  if (endpoint) {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        sheetId,
        mode: "exportXlsx",
      }),
    });
    if (!response.ok) {
      throw new Error(`exportXlsx HTTP ${response.status}`);
    }
    const json = (await response.json()) as {
      ok?: boolean;
      error?: string;
      base64?: string;
      fileName?: string;
    };
    if (json.error) throw new Error(json.error);
    if (!json.base64) throw new Error("Apps Script не повернув файл .xlsx.");
    return {
      fileName: json.fileName || "Штатка.xlsx",
      fileData: base64ToArrayBuffer(json.base64),
    };
  }

  const urls = [
    `/google-sheets/spreadsheets/d/${sheetId}/export?format=xlsx`,
    `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`,
  ];
  let lastError: Error | null = null;
  for (const url of urls) {
    try {
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) {
        lastError = new Error(`HTTP ${response.status}`);
        continue;
      }
      const fileData = await response.arrayBuffer();
      if (fileData.byteLength < 1000) {
        lastError = new Error("Порожня або HTML-відповідь замість .xlsx.");
        continue;
      }
      return { fileName: "Штатка.xlsx", fileData };
    } catch (error) {
      lastError =
        error instanceof Error ? error : new Error("Не вдалося скачати .xlsx");
    }
  }
  throw (
    lastError ??
    new Error(
      "Не вдалося скачати .xlsx «Штатки». Налаштуйте Apps Script (режим exportXlsx) або зробіть таблицю доступною.",
    )
  );
};

const kyivParts = (date: Date) => {
  const parts = new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/Kyiv",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).formatToParts(date);
  const get = (type: Intl.DateTimeFormatPartTypes) =>
    parts.find((part) => part.type === type)?.value ?? "0";
  return {
    year: Number(get("year")),
    month: Number(get("month")),
    day: Number(get("day")),
    hour: Number(get("hour")),
    minute: Number(get("minute")),
  };
};

/** Instant of a Kyiv wall-clock time as UTC Date. */
const kyivWallTimeToUtc = (
  year: number,
  month: number,
  day: number,
  hour: number,
  minute = 0,
) => {
  const utcGuess = new Date(Date.UTC(year, month - 1, day, hour, minute, 0));
  const asKyiv = kyivParts(utcGuess);
  const desiredAsMinutes = ((hour * 60 + minute) | 0);
  const actualAsMinutes = asKyiv.hour * 60 + asKyiv.minute;
  const deltaMinutes = desiredAsMinutes - actualAsMinutes;
  return new Date(utcGuess.getTime() + deltaMinutes * 60_000);
};

export const getLatestStaffSheetSyncSlot = (now = new Date()): Date => {
  const parts = kyivParts(now);
  const passedToday = STAFF_SHEET_SYNC_HOURS_KYIV.map((hour) =>
    kyivWallTimeToUtc(parts.year, parts.month, parts.day, hour),
  ).filter((slot) => slot.getTime() <= now.getTime());

  if (passedToday.length) {
    return passedToday[passedToday.length - 1]!;
  }

  const yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  const y = kyivParts(yesterday);
  return kyivWallTimeToUtc(
    y.year,
    y.month,
    y.day,
    STAFF_SHEET_SYNC_EVENING_HOUR,
  );
};

export const readStaffSheetLastSyncAt = (): string | null => {
  try {
    return window.localStorage.getItem(LAST_SYNC_STORAGE_KEY);
  } catch {
    return null;
  }
};

export const writeStaffSheetLastSyncAt = (iso = new Date().toISOString()) => {
  try {
    window.localStorage.setItem(LAST_SYNC_STORAGE_KEY, iso);
  } catch {
    /* ignore */
  }
};

export const shouldAutoSyncStaffSheetRoster = (now = new Date()) => {
  const last = readStaffSheetLastSyncAt();
  const due = getLatestStaffSheetSyncSlot(now);
  if (!last) return true;
  const lastMs = Date.parse(last);
  if (!Number.isFinite(lastMs)) return true;
  return lastMs < due.getTime();
};

export const STAFF_APPS_SCRIPT_TEMPLATE = `/**
 * Google Apps Script — таблиця «Штатка».
 * Deploy → Web app · Execute as: Me · Who has access: Anyone
 * Підтримує mode=bulk / bulkColumns / pull / exportXlsx.
 */
function findSheetByGid_(ss, gid) {
  if (gid == null || gid === '') return ss.getSheets()[0];
  var list = ss.getSheets();
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].getSheetId()) === String(gid)) return list[i];
  }
  return ss.getSheets()[0];
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var ss = SpreadsheetApp.openById(body.sheetId);

    if (body.mode === 'exportXlsx') {
      var exportUrl = 'https://docs.google.com/spreadsheets/d/' + body.sheetId + '/export?format=xlsx';
      var token = ScriptApp.getOAuthToken();
      var resp = UrlFetchApp.fetch(exportUrl, {
        headers: { Authorization: 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() !== 200) {
        throw new Error('exportXlsx HTTP ' + resp.getResponseCode());
      }
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          fileName: (ss.getName() || 'Staff') + '.xlsx',
          base64: Utilities.base64Encode(resp.getContent())
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.mode === 'pull' || body.mode === 'export') {
      var gids = body.gids;
      if (gids && gids.length) {
        var sheets = [];
        for (var i = 0; i < gids.length; i++) {
          var sheet = findSheetByGid_(ss, gids[i]);
          sheets.push({
            gid: String(sheet.getSheetId()),
            name: sheet.getName(),
            values: sheet.getDataRange().getDisplayValues()
          });
        }
        return ContentService
          .createTextOutput(JSON.stringify({ ok: true, sheets: sheets }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      var one = findSheetByGid_(ss, body.gid);
      var values = one.getDataRange().getDisplayValues();
      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          values: values,
          rows: values.length,
          gid: String(one.getSheetId()),
          name: one.getName()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = findSheetByGid_(ss, body.gid);

    if (body.mode === 'bulkColumns' && body.columns && body.values && body.values.length) {
      var startRowCols = Number(body.startRow) || 2;
      var rowNumbers = body.rowNumbers;
      var columns = body.columns;
      var valuesCols = body.values;
      var rowCount = valuesCols.length;
      var numberColumns = {};
      var textColumns = {};
      (body.numberColumns || []).forEach(function (n) { numberColumns[Number(n)] = true; });
      (body.textColumns || []).forEach(function (n) { textColumns[Number(n)] = true; });
      var clearBackground = body.clearBackground === true;
      for (var c = 0; c < columns.length; c++) {
        var col = Number(columns[c]);
        if (!col) continue;
        for (var r = 0; r < rowCount; r++) {
          var cell = valuesCols[r] && valuesCols[r][c];
          var text = cell == null ? '' : String(cell);
          var targetRow = rowNumbers && rowNumbers[r] != null && rowNumbers[r] !== ''
            ? Number(rowNumbers[r])
            : (startRowCols + r);
          if (!targetRow) continue;
          var cellRange = sheet.getRange(targetRow, col);
          if (clearBackground) {
            cellRange.setBackground(null);
            cellRange.setFontColor(null);
          }
          if (textColumns[col]) {
            cellRange.setNumberFormat('@');
          } else if (numberColumns[col]) {
            cellRange.setNumberFormat('0');
          }
          if (text.charAt(0) === '=') {
            cellRange.setFormula(text);
          } else if (text === '') {
            cellRange.setValue('');
          } else if (numberColumns[col] && /^-?\\d+(\\.\\d+)?$/.test(text)) {
            cellRange.setValue(Number(text));
          } else {
            cellRange.setValue(text);
          }
        }
      }
      if (body.clearAfterRow) {
        var lastCols = sheet.getLastRow();
        var clearFromCols = startRowCols + rowCount;
        if (lastCols >= clearFromCols) {
          for (var c2 = 0; c2 < columns.length; c2++) {
            var col2 = Number(columns[c2]);
            if (!col2) continue;
            sheet.getRange(clearFromCols, col2, lastCols - clearFromCols + 1, 1).clearContent();
          }
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, rows: rowCount, mode: 'bulkColumns' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.mode === 'bulk' && body.values && body.values.length) {
      var startRow = Number(body.startRow) || 2;
      var startCol = Number(body.startCol) || 1;
      var rows = body.values.length;
      var cols = body.values[0].length;
      sheet.getRange(startRow, startCol, rows, cols).setValues(body.values);

      if (body.clearAfterRow) {
        var last = sheet.getLastRow();
        var clearFrom = startRow + rows;
        if (last >= clearFrom) {
          sheet.getRange(clearFrom, startCol, last - clearFrom + 1, cols).clearContent();
        }
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, rows: rows }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (body.a1) {
      sheet.getRange(String(body.a1)).setValue(body.value == null ? '' : body.value);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, a1: body.a1 }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error('Unknown payload');
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
`;
