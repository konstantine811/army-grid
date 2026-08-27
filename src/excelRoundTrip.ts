import JSZip from "jszip";
import { formatUkDate, tryParseExcelSerialDate } from "./shared/format";

export type CellValue = string | number | boolean | Date | null | undefined;

type RawCellValue =
  | CellValue
  | {
      text?: () => string;
      value?: () => unknown;
      children?: unknown[];
      name?: string;
      constructor?: { name?: string };
    }
  | unknown[];

export type ExcelRow = {
  id: string;
  excelRowNumber: number;
  values: CellValue[];
  source: "template" | "merged";
};

export type ExcelSheetSnapshot = {
  sheetIndex: number;
  sheetName: string;
  rawRows: CellValue[][];
  headerRows: CellValue[][];
  rows: ExcelRow[];
  columnCount: number;
  columnIndexes: number[];
  dataStartRow: number;
  /** Excel row number → ARGB fill of column N (ПІБ), when read from workbook styles. */
  pibFillByExcelRow?: Record<number, string>;
};

export type ExcelWorkbookSnapshot = {
  file: File;
  fileName: string;
  sheetName: string;
  headerRows: CellValue[][];
  rows: ExcelRow[];
  columnCount: number;
  columnIndexes: number[];
  dataStartRow: number;
  sheets: ExcelSheetSnapshot[];
};

type DebugCellValue = string | number | boolean | null;

const DEFAULT_DATA_START_ROW = 5;
const DEFAULT_HEADER_ROW_INDEX = 3;
const KEY_COLUMN_INDEX = 1;
const FALLBACK_COLUMN_COUNT = 30;
const HEADER_SCAN_LIMIT = 12;
const HEADER_KEYWORDS = [
  "аркуш",
  "колонка",
  "№ колонки",
  "назва стовпця",
  "обовязковість",
  "обовʼязковість",
  "обов’язковість",
  "правило",
  "коментар",
  "піб",
  "прізвище",
  "посада",
  "статус",
  "підрозділ",
  "звання",
  "дата",
];

const loadXlsxPopulate = async () => {
  const module =
    await import("xlsx-populate/browser/xlsx-populate-no-encryption");
  return module.default;
};

const BCHS_ROSTER_PIB_COLUMN = 14;
/** Office theme accents — theme 9 is the standard green (70AD47). */
const OFFICE_THEME_RGB = [
  "000000",
  "FFFFFF",
  "44546A",
  "E7E6E6",
  "5B9BD5",
  "ED7D31",
  "A5A5A5",
  "FFC000",
  "4472C4",
  "70AD47",
];

const applyExcelTint = (hex: string, tint: number) => {
  const raw = hex.replace(/^FF/i, "").replace(/^#/, "");
  if (raw.length !== 6 || !Number.isFinite(tint) || tint === 0) {
    return raw.toUpperCase();
  }
  const apply = (channel: number) => {
    const next =
      tint < 0 ? channel * (1 + tint) : channel * (1 - tint) + 255 * tint;
    return Math.max(0, Math.min(255, Math.round(next)));
  };
  const r = apply(Number.parseInt(raw.slice(0, 2), 16));
  const g = apply(Number.parseInt(raw.slice(2, 4), 16));
  const b = apply(Number.parseInt(raw.slice(4, 6), 16));
  return [r, g, b]
    .map((value) => value.toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
};

const EMPTY_FILL_RGB = new Set(["FFFFFFFF", "FFFFFF", "00000000", "000000"]);

const columnNumberToLetter = (columnNumber: number) => {
  let label = "";
  let value = columnNumber;
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
};

const parseXmlThemeRgbByIndex = (themeXml: string): string[] => {
  const scheme = themeXml.match(/<a:clrScheme\b[^>]*>([\s\S]*?)<\/a:clrScheme>/i);
  if (!scheme) return [...OFFICE_THEME_RGB];
  const values = [...scheme[1].matchAll(/val="([0-9A-Fa-f]{6})"/g)].map((match) =>
    match[1].toUpperCase(),
  );
  return values.length >= 10 ? values.slice(0, 10) : [...OFFICE_THEME_RGB];
};

const parseFgColorRgb = (attributes: string, themeRgb: string[]): string | null => {
  const rgb = attributes.match(/\brgb="([0-9A-Fa-f]{6,8})"/i)?.[1];
  if (rgb)
    return rgb.replace(/^FF/i, "").toUpperCase().padStart(6, "0").slice(-6);
  const theme = attributes.match(/\btheme="(\d+)"/i)?.[1];
  if (theme == null) return null;
  const base = themeRgb[Number(theme)];
  if (!base) return null;
  const tintRaw = attributes.match(/\btint="(-?[\d.]+)"/i)?.[1];
  const tint = tintRaw == null ? 0 : Number(tintRaw);
  return applyExcelTint(base, tint);
};

const parseStylesFillRgbs = (stylesXml: string, themeRgb: string[]) => {
  const fillsBlock =
    stylesXml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/i)?.[1] ?? "";
  const fills = [...fillsBlock.matchAll(/<fill\b[\s\S]*?<\/fill>/gi)].map(
    (match) => {
      const fg = match[0].match(/<fgColor\b([^/>]*)/i)?.[1];
      return fg ? parseFgColorRgb(fg, themeRgb) : null;
    },
  );
  const xfsBlock =
    stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
  const xfFillIds = [...xfsBlock.matchAll(/<xf\b([^>]*)>/gi)].map((match) => {
    const fillId = match[1].match(/\bfillId="(\d+)"/i)?.[1];
    return fillId == null ? 0 : Number(fillId);
  });
  return { fills, xfFillIds };
};

const parseWorksheetColumnFills = (
  sheetXml: string,
  columnLetter: string,
  fills: Array<string | null>,
  xfFillIds: number[],
) => {
  const result: Record<number, string> = {};
  for (const match of sheetXml.matchAll(/<c\b([^>]*)>/gi)) {
    const attrs = match[1];
    const ref = attrs.match(/\br="([A-Z]+)(\d+)"/i);
    if (!ref || ref[1].toUpperCase() !== columnLetter) continue;
    const styleId = attrs.match(/\bs="(\d+)"/i)?.[1];
    if (styleId == null) continue;
    const fillId = xfFillIds[Number(styleId)];
    if (fillId == null) continue;
    const rgb = fills[fillId];
    if (!rgb) continue;
    const argb = rgb.length === 6 ? `FF${rgb}` : rgb.toUpperCase();
    if (EMPTY_FILL_RGB.has(argb) || EMPTY_FILL_RGB.has(rgb.toUpperCase())) {
      continue;
    }
    result[Number(ref[2])] = argb;
  }
  return result;
};

const readXlsxSheetColumnFills = async (
  file: File,
  columnNumber: number,
): Promise<{
  byName: Map<string, Record<number, string>>;
  byIndex: Array<Record<number, string>>;
}> => {
  const byName = new Map<string, Record<number, string>>();
  const byIndex: Array<Record<number, string>> = [];
  if (columnNumber < 1) return { byName, byIndex };
  try {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
    const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
    const stylesXml = await zip.file("xl/styles.xml")?.async("string");
    if (!workbookXml || !relsXml || !stylesXml) return { byName, byIndex };

    const themeXml =
      (await zip.file("xl/theme/theme1.xml")?.async("string")) ?? "";
    const themeRgb = parseXmlThemeRgbByIndex(themeXml);
    const { fills, xfFillIds } = parseStylesFillRgbs(stylesXml, themeRgb);
    const columnLetter = columnNumberToLetter(columnNumber);

    const ridToTarget = new Map<string, string>();
    for (const rel of relsXml.matchAll(
      /\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g,
    )) {
      ridToTarget.set(rel[1], rel[2]);
    }
    for (const rel of relsXml.matchAll(
      /\bTarget="([^"]+)"[^>]*\bId="([^"]+)"/g,
    )) {
      if (!ridToTarget.has(rel[2])) ridToTarget.set(rel[2], rel[1]);
    }

    for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*)>/gi)) {
      const name = sheet[1].match(/\bname="([^"]+)"/)?.[1];
      const rid =
        sheet[1].match(/\br:id="([^"]+)"/)?.[1] ??
        sheet[1].match(/\bid="([^"]+)"/)?.[1];
      if (!name || !rid) {
        byIndex.push({});
        continue;
      }
      const target = ridToTarget.get(rid);
      if (!target) {
        byIndex.push({});
        continue;
      }
      const path = target.startsWith("/")
        ? target.slice(1)
        : target.startsWith("xl/")
          ? target
          : `xl/${target.replace(/^\.\//, "")}`;
      const sheetXml = await zip.file(path)?.async("string");
      if (!sheetXml) {
        byIndex.push({});
        continue;
      }
      const fillsByRow = parseWorksheetColumnFills(
        sheetXml,
        columnLetter,
        fills,
        xfFillIds,
      );
      byName.set(name, fillsByRow);
      byIndex.push(fillsByRow);
    }
  } catch {
    return { byName, byIndex };
  }
  return { byName, byIndex };
};

const colorObjectToRgb = (color: unknown): string | null => {
  if (!color) return null;
  if (typeof color === "string") {
    const hex = color.replace(/^#/, "").toUpperCase();
    return hex || null;
  }
  if (typeof color !== "object") return null;
  const rec = color as {
    rgb?: string;
    theme?: number | string;
    tint?: number | string;
    indexed?: number | string;
  };
  if (rec.rgb) return rec.rgb.replace(/^#/, "").toUpperCase();
  if (rec.theme != null && rec.theme !== "") {
    const themeIndex = Number(rec.theme);
    const base = OFFICE_THEME_RGB[themeIndex];
    if (!base) return null;
    const tint = rec.tint == null || rec.tint === "" ? 0 : Number(rec.tint);
    return applyExcelTint(base, tint);
  }
  return null;
};

export const extractCellFillRgb = (fill: unknown): string | null => {
  if (!fill) return null;
  if (typeof fill === "string") {
    const normalized = fill.replace(/^#/, "").toUpperCase();
    return normalized || null;
  }
  if (typeof fill !== "object") return null;
  const rec = fill as {
    type?: string;
    color?: unknown;
    rgb?: string;
    foreground?: unknown;
    fgColor?: unknown;
    background?: unknown;
    bgColor?: unknown;
    theme?: number | string;
    tint?: number | string;
    stops?: Array<{ color?: unknown }>;
  };
  return (
    colorObjectToRgb(rec.color) ||
    colorObjectToRgb(rec.foreground) ||
    colorObjectToRgb(rec.fgColor) ||
    colorObjectToRgb(rec.background) ||
    colorObjectToRgb(rec.bgColor) ||
    colorObjectToRgb(rec.stops?.[0]?.color) ||
    colorObjectToRgb(rec)
  );
};

const findBchsRosterPibColumnNumber = (rows: CellValue[][]) => {
  const scanLimit = Math.min(5, rows.length);
  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      const header = valueToDisplay(row[columnIndex]).toLowerCase();
      if (header.includes("піб") || header.includes("п.і.б")) {
        return columnIndex + 1;
      }
    }
  }
  return 0;
};

const readBchsRosterPibFills = (
  sheet: {
    cell: (row: number, column: number) => { style: (name: string) => unknown };
  },
  rowCount: number,
  pibColumnNumber = BCHS_ROSTER_PIB_COLUMN,
): Record<number, string> => {
  const fills: Record<number, string> = {};

  for (let rowNumber = 2; rowNumber <= rowCount; rowNumber += 1) {
    try {
      const rgb = extractCellFillRgb(
        sheet.cell(rowNumber, pibColumnNumber).style("fill"),
      );
      if (
        !rgb ||
        rgb === "FFFFFFFF" ||
        rgb === "FFFFFF" ||
        rgb === "00000000" ||
        rgb === "000000"
      ) {
        continue;
      }
      fills[rowNumber] = rgb;
    } catch {
      // Skip cells whose style cannot be read.
    }
  }

  return fills;
};

const isEmptyValue = (value: CellValue) =>
  value === null || value === undefined || String(value).trim() === "";

const normalizeKey = (value: CellValue) =>
  String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();

export const hasRowData = (row: CellValue[]) =>
  row.some((value) => !isEmptyValue(value));

const pickColumns = (row: CellValue[], columnIndexes: number[]) =>
  columnIndexes.map((columnIndex) => row[columnIndex] ?? null);

const normalizeHeaderText = (value: CellValue) =>
  valueToDisplay(value)
    .replace(/[ʼ’']/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

const scoreHeaderRow = (row: CellValue[]) => {
  const values = row.filter((value) => !isEmptyValue(value));
  const joined = values.map(normalizeHeaderText).join(" ");
  const keywordScore = HEADER_KEYWORDS.reduce((score, keyword) => {
    const normalizedKeyword = keyword.replace(/[ʼ’']/g, "").toLowerCase();

    return joined.includes(normalizedKeyword) ? score + 8 : score;
  }, 0);
  const textScore = values.filter(
    (value) => typeof value === "string" && String(value).trim().length <= 90,
  ).length;
  const numericPenalty =
    values.filter((value) => typeof value === "number").length * 2;

  return values.length + textScore + keywordScore - numericPenalty;
};

const detectHeaderRowIndex = (rows: CellValue[][]) => {
  const candidates = rows.slice(0, HEADER_SCAN_LIMIT).map((row, index) => ({
    index,
    score: scoreHeaderRow(row),
  }));
  const bestCandidate = candidates.reduce(
    (best, candidate) => (candidate.score > best.score ? candidate : best),
    { index: DEFAULT_HEADER_ROW_INDEX, score: Number.NEGATIVE_INFINITY },
  );
  const defaultCandidate = candidates.find(
    (candidate) => candidate.index === DEFAULT_HEADER_ROW_INDEX,
  );

  if (
    bestCandidate.score >= 18 ||
    bestCandidate.score >= (defaultCandidate?.score ?? 0) + 8
  ) {
    return bestCandidate.index;
  }

  return rows[DEFAULT_HEADER_ROW_INDEX]
    ? DEFAULT_HEADER_ROW_INDEX
    : Math.max(0, bestCandidate.index);
};

const extractXmlText = (value: unknown): string => {
  if (value === null || value === undefined) return "";
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  )
    return String(value);
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(extractXmlText).join("");

  if (typeof value === "object") {
    const node = value as { children?: unknown[]; value?: () => unknown };
    if (typeof node.value === "function") return extractXmlText(node.value());
    if (Array.isArray(node.children))
      return node.children.map(extractXmlText).join("");
  }

  return String(value);
};

const sanitizeCellValue = (value: RawCellValue): CellValue => {
  if (value === undefined || value === null) return null;
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  )
    return value;
  if (value instanceof Date) return value;

  if (Array.isArray(value)) {
    const text = value.map(extractXmlText).join("").trim();
    
    return text || null;
  }

  if (typeof value === "object") {
    const maybeRichText = value as {
      text?: () => string;
      value?: () => unknown;
      constructor?: { name?: string };
    };

    if (typeof maybeRichText.text === "function") return maybeRichText.text();
    if (typeof maybeRichText.value === "function")
      return extractXmlText(maybeRichText.value());

    return extractXmlText(value);
  }

  return String(value);
};

const normalizeDateDisplay = (value: CellValue) => {
  if (value instanceof Date) return formatUkDate(value);

  const parsedSerial = tryParseExcelSerialDate(value);
  if (parsedSerial) return formatUkDate(parsedSerial);

  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}T/.test(value)) {
    const parsedIso = new Date(value);
    if (!Number.isNaN(parsedIso.getTime())) return formatUkDate(parsedIso);
  }

  return value ?? "";
};

export const valueToDisplay = (value: CellValue) =>
  String(normalizeDateDisplay(value));

const valueToDebug = (value: CellValue): DebugCellValue => {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
};

export function createWorkbookDebugPayload(snapshot: ExcelWorkbookSnapshot) {
  const createSheetPayload = (sheet: ExcelSheetSnapshot) => {
    const columns = Array.from({ length: sheet.columnCount }, (_, index) => ({
      index,
      originalIndex: sheet.columnIndexes[index],
      letter: getColumnLabel(sheet.columnIndexes[index] ?? index),
      header: getColumnHeader(sheet, index),
    }));

    const parsedRows = sheet.rows
      .filter((row) => hasRowData(row.values))
      .map((row) => ({
        excelRowNumber: row.excelRowNumber,
        values: row.values.map(valueToDebug),
        byColumn: Object.fromEntries(
          columns.map((column, index) => [
            column.header || `column_${index + 1}`,
            valueToDebug(row.values[index]),
          ]),
        ),
      }));

    return {
      sheetIndex: sheet.sheetIndex,
      sheetName: sheet.sheetName,
      dataStartRow: sheet.dataStartRow,
      columnCount: sheet.columnCount,
      rawRowCount: sheet.rawRows.length,
      parsedRowCount: parsedRows.length,
      rawRows: sheet.rawRows.map((row, index) => ({
        excelRowNumber: index + 1,
        values: row.map(valueToDebug),
      })),
      headerRows: sheet.headerRows.map((row) => row.map(valueToDebug)),
      columns,
      rows: parsedRows,
    };
  };

  return {
    fileName: snapshot.fileName,
    sheetCount: snapshot.sheets.length,
    activeSheetName: snapshot.sheetName,
    sheets: snapshot.sheets.map(createSheetPayload),
  };
}

export type ReadWorkbookOptions = {
  /** Keep only matching sheets (by name). */
  sheetFilter?: (sheetName: string) => boolean;
  /** Cap columns per sheet (ЕЖООС ООС/статистика мають 1000+ колонок). */
  maxColumns?: number;
  /** Skip PIB cell-fill scanning (very expensive on large workbooks). */
  skipStyleFills?: boolean;
  /** Skip building `rows[]` objects — rawRows only (less memory). */
  skipRowObjects?: boolean;
};

/** Fast path for ЕЖООС↔1ПБ sync: few sheets, few columns, no style scans. */
export const EJOOS_SYNC_READ_OPTIONS: ReadWorkbookOptions = {
  sheetFilter: (name) =>
    /^(sh|рух|archive)$/i.test(name.trim()) ||
    /^1\.\s*шпо/i.test(name) ||
    /^2\.\s*оос/i.test(name) ||
    /^3\.\s*виключ/i.test(name) ||
    /^5\.\s*тимчасов/i.test(name) ||
    /^6\.\s*табель/i.test(name) ||
    /^10\.\s*історія/i.test(name),
  maxColumns: 48,
  skipStyleFills: true,
  skipRowObjects: true,
};

export async function readWorkbookSnapshot(
  file: File,
  options: ReadWorkbookOptions = {},
): Promise<ExcelWorkbookSnapshot> {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(file);
  const zipFillsByColumn = new Map<
    number,
    Awaited<ReturnType<typeof readXlsxSheetColumnFills>>
  >();
  const loadZipFills = async (columnNumber: number) => {
    const cached = zipFillsByColumn.get(columnNumber);
    if (cached) return cached;
    const loaded = await readXlsxSheetColumnFills(file, columnNumber);
    zipFillsByColumn.set(columnNumber, loaded);
    return loaded;
  };

  const workbookSheets = workbook.sheets();
  const sheets: ExcelWorkbookSnapshot["sheets"] = [];
  for (let sheetIndex = 0; sheetIndex < workbookSheets.length; sheetIndex += 1) {
    const sheet = workbookSheets[sheetIndex];
    const sheetName = sheet.name();
    if (options.sheetFilter && !options.sheetFilter(sheetName)) continue;

    const values = (sheet.usedRange()?.value() ?? []) as RawCellValue[][];
    const detectedColumnCount = Math.max(
      FALLBACK_COLUMN_COUNT,
      0,
      ...values.map((row) => row.length),
    );
    const rawColumnCount = options.maxColumns
      ? Math.min(options.maxColumns, detectedColumnCount)
      : detectedColumnCount;
    const fullRows = values.map((row) =>
      Array.from({ length: rawColumnCount }, (_, index) =>
        sanitizeCellValue(row[index] ?? null),
      ),
    );
    const headerRowIndex = detectHeaderRowIndex(fullRows);
    const dataStartRow = headerRowIndex + 2;
    const fullDataRows = fullRows.slice(dataStartRow - 1);
    const usedColumnIndexes = Array.from(
      { length: rawColumnCount },
      (_, index) => index,
    ).filter((columnIndex) => {
      const hasHeader = !isEmptyValue(fullRows[headerRowIndex]?.[columnIndex]);
      const hasData = fullDataRows.some(
        (row) => !isEmptyValue(row[columnIndex]),
      );

      return hasHeader || hasData;
    });
    const firstUsedColumnIndex = usedColumnIndexes[0] ?? 0;
    let lastUsedColumnIndex =
      usedColumnIndexes.at(-1) ?? Math.max(0, rawColumnCount - 1);
    if (options.maxColumns) {
      lastUsedColumnIndex = Math.min(
        lastUsedColumnIndex,
        firstUsedColumnIndex + options.maxColumns - 1,
      );
    }
    const columnIndexes =
      usedColumnIndexes.length > 0
        ? Array.from(
            { length: lastUsedColumnIndex - firstUsedColumnIndex + 1 },
            (_, index) => firstUsedColumnIndex + index,
          )
        : [];
    const normalizedRows = fullRows.map((row) =>
      pickColumns(row, columnIndexes),
    );
    const columnCount = columnIndexes.length;

    let pibFillByExcelRow: Record<number, string> | undefined;
    if (!options.skipStyleFills) {
      const pibColumnNumber = findBchsRosterPibColumnNumber(fullRows);
      const fromPopulate =
        pibColumnNumber > 0
          ? readBchsRosterPibFills(sheet, fullRows.length, pibColumnNumber)
          : {};
      const zipFills =
        pibColumnNumber > 0 ? await loadZipFills(pibColumnNumber) : null;
      const fromZip =
        zipFills?.byName.get(sheetName) ?? zipFills?.byIndex[sheetIndex] ?? {};
      const merged = { ...fromPopulate, ...fromZip };
      if (Object.keys(merged).length > 0) pibFillByExcelRow = merged;
    }

    sheets.push({
      sheetIndex,
      sheetName,
      rawRows: normalizedRows,
      headerRows: normalizedRows.slice(0, headerRowIndex + 1),
      rows: options.skipRowObjects
        ? []
        : normalizedRows.slice(dataStartRow - 1).map((rowValues, index) => ({
            id: `${file.name}-${sheetIndex}-${dataStartRow + index}`,
            excelRowNumber: dataStartRow + index,
            values: rowValues,
            source: "template" as const,
          })),
      columnCount,
      columnIndexes,
      dataStartRow,
      pibFillByExcelRow,
    });
  }

  const activeSheet = sheets[0] ?? {
    sheetIndex: 0,
    sheetName: "Аркуш1",
    rawRows: [],
    headerRows: [],
    rows: [],
    columnCount: 0,
    columnIndexes: [],
    dataStartRow: DEFAULT_DATA_START_ROW,
  };

  return {
    file,
    fileName: file.name,
    sheetName: activeSheet.sheetName,
    headerRows: activeSheet.headerRows,
    rows: activeSheet.rows,
    columnCount: activeSheet.columnCount,
    columnIndexes: activeSheet.columnIndexes,
    dataStartRow: activeSheet.dataStartRow,
    sheets,
  };
}

export function getColumnLabel(index: number) {
  let label = "";
  let value = index + 1;

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

export function getColumnHeader(
  snapshot: Pick<ExcelWorkbookSnapshot, "headerRows">,
  index: number,
) {
  const rawHeader = snapshot.headerRows.at(-1)?.[index];
  const header = valueToDisplay(rawHeader).replace(/\s+/g, " ").trim();

  return header;
}

export function updateCell(
  rows: ExcelRow[],
  rowId: string,
  columnIndex: number,
  value: string,
) {
  return rows.map((row) =>
    row.id === rowId
      ? {
          ...row,
          values: row.values.map((cellValue, index) =>
            index === columnIndex ? value : cellValue,
          ),
        }
      : row,
  );
}

export function mergeRows(baseRows: ExcelRow[], incomingRows: ExcelRow[]) {
  const rows = baseRows.map((row) => ({ ...row, values: [...row.values] }));
  const indexByKey = new Map<string, number>();

  rows.forEach((row, index) => {
    const key = normalizeKey(row.values[KEY_COLUMN_INDEX]);
    if (key) indexByKey.set(key, index);
  });

  const stats = {
    updated: 0,
    created: 0,
    skipped: 0,
  };

  incomingRows
    .filter((row) => hasRowData(row.values))
    .forEach((incomingRow) => {
      const key = normalizeKey(incomingRow.values[KEY_COLUMN_INDEX]);

      if (!key) {
        stats.skipped += 1;
        return;
      }

      const existingIndex = indexByKey.get(key);
      if (existingIndex !== undefined) {
        rows[existingIndex] = {
          ...rows[existingIndex],
          values: rows[existingIndex].values.map((value, index) =>
            isEmptyValue(incomingRow.values[index])
              ? value
              : incomingRow.values[index],
          ),
        };
        stats.updated += 1;
        return;
      }

      const emptyIndex = rows.findIndex((row) => !hasRowData(row.values));
      const nextRowNumber =
        emptyIndex >= 0
          ? rows[emptyIndex].excelRowNumber
          : Math.max(...rows.map((row) => row.excelRowNumber)) + 1;
      const nextRow = {
        id: `merged-${key}-${nextRowNumber}`,
        excelRowNumber: nextRowNumber,
        values: incomingRow.values,
        source: "merged" as const,
      };

      if (emptyIndex >= 0) {
        rows[emptyIndex] = nextRow;
      } else {
        rows.push(nextRow);
      }

      indexByKey.set(
        key,
        rows.findIndex((row) => row.id === nextRow.id),
      );
      stats.created += 1;
    });

  return { rows, stats };
}

export async function exportStyledWorkbook(
  snapshot: ExcelWorkbookSnapshot,
  rows: ExcelRow[],
  sheetIndex = 0,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(snapshot.file);
  const sheet = workbook.sheet(sheetIndex);
  const targetSnapshot =
    snapshot.sheets.find((item) => item.sheetIndex === sheetIndex) ?? snapshot;

  rows.forEach((row) => {
    row.values.forEach((value, columnIndex) => {
      const originalColumnIndex =
        targetSnapshot.columnIndexes[columnIndex] ?? columnIndex;
      sheet
        .cell(row.excelRowNumber, originalColumnIndex + 1)
        .value(isEmptyValue(value) ? null : value);
    });
  });

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const baseName = snapshot.fileName.replace(/\.xlsx$/i, "");
  link.href = url;
  link.download = `${baseName}_edited.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportWorkbookWithMutations(
  snapshot: ExcelWorkbookSnapshot,
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(snapshot.file);

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportWorkbookFileWithMutations(
  file: File,
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(file);

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportBlankWorkbookWithMutations(
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromBlankAsync();

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportBlankWorkbookSheetWithMutations(
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  sheetName: string,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook: any = await XlsxPopulate.fromBlankAsync();

  await mutateWorkbook(workbook);

  const targetSheet = workbook.sheet(sheetName);
  if (!targetSheet) {
    throw new Error(`Не знайдено аркуш "${sheetName}" для експорту.`);
  }

  [...workbook.sheets()].reverse().forEach((sheet: any) => {
    if (sheet.name() !== sheetName) workbook.deleteSheet(sheet);
  });

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportTemplateWorkbookWithMutations(
  templateUrl: string,
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const response = await fetch(templateUrl);
  if (!response.ok) {
    throw new Error(`Не вдалося завантажити Excel-шаблон: ${templateUrl}`);
  }
  const workbook: any = await XlsxPopulate.fromDataAsync(
    await response.arrayBuffer(),
  );

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
