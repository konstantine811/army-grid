import JSZip from "jszip";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { canonicalName, normId } from "./ejoosIdentity";
import { extractSheetFormulasByCell } from "./ejoosTimesheetLayout";
import { resolveSheetPath } from "./ejoosZipCellWrites";

const LEFTOVER_NOTE = /вибув|переведен|згідно|розпорядж/i;

const decodeXmlText = (xml: string) =>
  xml
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#10;/g, "\n")
    .replace(/&#xA;/gi, "\n")
    .replace(/_x000A_/gi, "\n")
    .replace(/_x000D_/gi, "\n")
    .replace(/&amp;/g, "&");

const cellText = (value: unknown) =>
  value instanceof Date
    ? ""
    : typeof value === "number" && Number.isFinite(value)
      ? String(Math.trunc(value))
      : String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();

const parseSharedStrings = (sstXml: string) =>
  [...sstXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)].map((match) =>
    [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
      .map((part) => decodeXmlText(part[1]))
      .join(""),
  );

const columnLetterToNumber = (letter: string) => {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index;
};

const decodeSheetCell = (
  cellXml: string,
  attrs: string,
  strings: string[],
): unknown => {
  const type = attrs.match(/\bt="([^"]*)"/)?.[1] || "";
  if (type === "s") {
    const index = Number(cellXml.match(/<v\b[^>]*>(\d+)<\/v>/i)?.[1]);
    return Number.isInteger(index) ? strings[index] ?? "" : "";
  }
  if (type === "inlineStr" || /<is\b/i.test(cellXml)) {
    return [...cellXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
      .map((part) => decodeXmlText(part[1]))
      .join("");
  }
  const raw = cellXml.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1];
  if (raw == null) return "";
  const text = decodeXmlText(raw);
  if (type !== "str" && /^-?\d+(\.\d+)?$/.test(text)) return Number(text);
  return text;
};

/**
 * Рядок Табеля належить особі, якщо збігається ID або ПІБ
 * у будь-якій з перших колонок (не лише G/H).
 */
export const timesheetRowMatchesPerson = (
  cells: unknown[] | undefined,
  personId: string,
  fullName: string,
) => {
  const row = Array.from(cells ?? []);
  const wantId = normId(personId);
  const wantName = canonicalName(fullName);
  const wantSurname = wantName.split(" ")[0] || "";
  const texts = row.map((cell) => canonicalName(cellText(cell)));
  const ids = row.map((cell) => normId(cell));
  if (wantId && ids.some((id) => id === wantId)) return true;
  if (wantName && texts.some((text) => text === wantName)) return true;
  if (wantName && texts.some((text) => text.includes(wantName))) return true;
  if (wantSurname.length < 4) return false;
  const named = texts.find(
    (text) =>
      text === wantSurname ||
      text.startsWith(`${wantSurname} `) ||
      text.includes(` ${wantSurname} `),
  );
  if (!named) return false;
  if (named.split(" ").length >= 2) return true;
  return LEFTOVER_NOTE.test(texts.join(" "));
};

const MAX_TIMESHEET_PERSON_ROW = 4000;

export const findTimesheetPersonRowsInGrid = (
  rawRows: Array<unknown[] | undefined>,
  personId: string,
  fullName: string,
  startExcelRow = 6,
) => {
  const rows: number[] = [];
  const last = Math.min(rawRows.length, MAX_TIMESHEET_PERSON_ROW);
  for (let index = startExcelRow - 1; index < last; index += 1) {
    if (timesheetRowMatchesPerson(rawRows[index], personId, fullName)) {
      rows.push(index + 1);
    }
  }
  return rows;
};

/** Повний XML аркуша, включно з рядками поза usedRange / dimension. */
export const buildSheetGridFromXml = (sheetXml: string, sstXml = "") => {
  const strings = parseSharedStrings(sstXml);
  const rows: unknown[][] = [];
  const cellRe =
    /<c\b([^<>]*\br="([A-Z]{1,3})(\d+)"(?![0-9A-Za-z])[^<>/]*)(\/\s*>|>[\s\S]*?<\/c>)/gi;
  for (const cell of sheetXml.matchAll(cellRe)) {
    const row = Number(cell[3]);
    const column = columnLetterToNumber(cell[2]);
    if (!row || column < 1 || column > 48) continue;
    const gridRow = rows[row - 1] ?? (rows[row - 1] = []);
    gridRow[column - 1] = decodeSheetCell(cell[0], cell[1], strings);
  }
  return rows;
};

/** Єдине джерело для placement-логіки: snapshot + XML разом. */
export const placementSheetFromMergedGrid = (
  sheet: ExcelSheetSnapshot,
  merged: Array<unknown[] | undefined>,
): ExcelSheetSnapshot => {
  const lastRow = Math.max(sheet.rawRows.length, merged.length);
  const rawRows: string[][] = [];
  for (let row = 0; row < lastRow; row += 1) {
    const mergedRow = merged[row];
    const snapshotRow = sheet.rawRows[row];
    rawRows.push(
      Array.from({ length: 40 }, (_, column) => {
        const mergedValue = mergedRow?.[column];
        if (mergedValue != null && String(mergedValue).trim() !== "") {
          return String(mergedValue).trim();
        }
        return String(snapshotRow?.[column] ?? "").trim();
      }),
    );
  }
  return { ...sheet, rawRows };
};

export const mergeTimesheetGrids = (
  ...grids: Array<Array<unknown[] | undefined>>
) => {
  const merged: unknown[][] = [];
  for (const grid of grids) {
    for (let index = 0; index < grid.length; index += 1) {
      const row = grid[index];
      if (!row) continue;
      const target = merged[index] ?? (merged[index] = []);
      for (let column = 0; column < row.length; column += 1) {
        if (row[column] != null && row[column] !== "") {
          target[column] = row[column];
        }
      }
    }
  }
  return merged;
};

export const pickTimesheetKeepRow = (
  rows: number[],
  prefer: (row: number) => boolean,
  planned = 0,
) => {
  const preferred = rows.find((row) => prefer(row));
  if (preferred) return preferred;
  if (planned && rows.includes(planned)) return planned;
  return rows[0] || 0;
};

export const uniqueExcelRows = (rows: number[]) =>
  [...new Set(rows.filter((row) => row > 0))].sort((left, right) => left - right);

export const loadTimesheetSheetArtifacts = async (
  file: Blob | File | undefined,
  sheetName: string,
) => {
  const empty = {
    grid: [] as unknown[][],
    formulas: new Map<string, string>(),
  };
  if (!file || !sheetName) return empty;
  try {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const path = await resolveSheetPath(zip, sheetName);
    if (!path) return empty;
    const sheetXml = (await zip.file(path)?.async("string")) || "";
    const sstXml =
      (await zip.file("xl/sharedStrings.xml")?.async("string")) || "";
    return {
      grid: buildSheetGridFromXml(sheetXml, sstXml),
      formulas: extractSheetFormulasByCell(sheetXml),
    };
  } catch {
    return empty;
  }
};

export const loadTimesheetGridFromFile = async (
  file: Blob | File | undefined,
  sheetName: string,
) => (await loadTimesheetSheetArtifacts(file, sheetName)).grid;

/** Рядки Табеля поза usedRange (дописана історія внизу) підмішуємо в snapshot. */
export const hydrateEjoosTimesheetSnapshot = async (
  snapshot: ExcelWorkbookSnapshot,
): Promise<ExcelWorkbookSnapshot> => {
  const timesheet = snapshot.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
  if (!timesheet) return snapshot;
  const xmlGrid = await loadTimesheetGridFromFile(
    snapshot.file,
    timesheet.sheetName,
  );
  if (!xmlGrid.length) return snapshot;
  timesheet.rawRows = mergeTimesheetGrids(timesheet.rawRows, xmlGrid);
  return snapshot;
};

export const readEjoosWorkbookSnapshot = async (file: File) =>
  hydrateEjoosTimesheetSnapshot(
    await readWorkbookSnapshot(file, EJOOS_SYNC_READ_OPTIONS),
  );
