import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  formatUkDateForExcelWrite,
  tryParseExcelSerialDate,
} from "../../shared/format";
import { findEjoosSheet } from "./ejoosSyncPlan";
import { readEjoosWorkbookSnapshot } from "./ejoosTimesheetPersonRows";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

const ARRIVAL_DATE_HEADER_PATTERNS = [
  /дата\s*прибут/,
  /фактичн.*(?:поверн|прибут)/,
  /дата\s*вибут(?!.*наказ)/,
  /дата\s*наказ/,
  /планова\s*дата/,
];

const ABSENT_DATE_HEADER_PATTERNS = [
  /дата\s*вибут(?!.*наказ)/,
  /дата\s*наказ(?!\s*про\s*поверн)/,
  /планова\s*дата/,
  /фактичн.*(?:дата\s*)?поверн/,
  /дата\s*(?:наказу\s*)?про\s*поверн/,
  /дата\s*поверн/,
];

const DEFAULT_ABSENT_DATE_COLUMNS = [7, 8, 12, 13, 15];

const findDateColumnsByPatterns = (
  sheet: ExcelSheetSnapshot,
  patterns: RegExp[],
  defaultColumns: number[],
): number[] => {
  const columns = new Set<number>();
  for (let row = 1; row <= Math.min(6, sheet.rawRows.length); row += 1) {
    for (let column = 1; column <= 40; column += 1) {
      const text = String(sheet.rawRows[row - 1]?.[column - 1] ?? "")
        .toLocaleLowerCase("uk-UA")
        .replace(/\s+/g, " ")
        .trim();
      if (!text) continue;
      if (patterns.some((pattern) => pattern.test(text))) {
        columns.add(column);
      }
    }
  }
  if (columns.size === 0) {
    for (const column of defaultColumns) columns.add(column);
  }
  return [...columns].sort((left, right) => left - right);
};

export const findArrivalDateColumns = (
  sheet: ExcelSheetSnapshot,
): number[] =>
  findDateColumnsByPatterns(sheet, ARRIVAL_DATE_HEADER_PATTERNS, [8]);

export const findAbsentDateColumns = (
  sheet: ExcelSheetSnapshot,
): number[] =>
  findDateColumnsByPatterns(
    sheet,
    ABSENT_DATE_HEADER_PATTERNS,
    DEFAULT_ABSENT_DATE_COLUMNS,
  );

export const ejoosDateCellNeedsStringWrite = (value: unknown) =>
  tryParseExcelSerialDate(value) !== null;

export const buildEjoosDateCellWrites = (
  sheet: ExcelSheetSnapshot,
  options?: {
    rows?: Iterable<number>;
    minRow?: number;
    dateColumns?: number[];
  },
): ZipCellWrite[] => {
  const minRow = options?.minRow ?? 6;
  const dateColumns = options?.dateColumns ?? findArrivalDateColumns(sheet);
  const rowSet =
    options?.rows != null
      ? new Set([...options.rows].filter((row) => row >= minRow))
      : null;
  const writes: ZipCellWrite[] = [];

  for (let row = minRow; row <= sheet.rawRows.length; row += 1) {
    if (rowSet && !rowSet.has(row)) continue;
    for (const column of dateColumns) {
      const raw = sheet.rawRows[row - 1]?.[column - 1];
      if (!ejoosDateCellNeedsStringWrite(raw)) continue;
      const formatted = formatUkDateForExcelWrite(raw);
      if (!formatted) continue;
      writes.push({ row, column, value: formatted });
    }
  }

  return writes;
};

const fixSheetExcelSerialDates = async (
  file: Blob,
  sheetPattern: RegExp,
  dateColumns: (sheet: ExcelSheetSnapshot) => number[],
): Promise<Blob> => {
  const workbook = await readEjoosWorkbookSnapshot(
    new File([file], "ejoos-date-fix.xlsx", {
      type:
        file.type ||
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
  );
  const sheet = findEjoosSheet(workbook, sheetPattern);
  if (!sheet) return file;
  const writes = buildEjoosDateCellWrites(sheet, {
    dateColumns: dateColumns(sheet),
  });
  if (!writes.length) return file;
  return applyInlineStringWritesToWorkbook(file, sheet.sheetName, writes);
};

/** Після xlsx-populate: «Дата прибуття» часто лишається 46193 без формату дати. */
export async function fixArrivalSheetExcelSerialDates(file: Blob): Promise<Blob> {
  return fixSheetExcelSerialDates(
    file,
    /тимчасов.*прибул/i,
    findArrivalDateColumns,
  );
}

/** Те саме для «5. Тимчасово відсутні» (G/H/L/M/O). */
export async function fixAbsentSheetExcelSerialDates(file: Blob): Promise<Blob> {
  return fixSheetExcelSerialDates(
    file,
    /тимчасов.*відсут/i,
    findAbsentDateColumns,
  );
}

export async function fixEjoosSheetExcelSerialDates(file: Blob): Promise<Blob> {
  let blob = file;
  blob = await fixArrivalSheetExcelSerialDates(blob);
  blob = await fixAbsentSheetExcelSerialDates(blob);
  return blob;
}

export { fixOosPersonIdCells } from "./ejoosOosPersonIdRepair";
