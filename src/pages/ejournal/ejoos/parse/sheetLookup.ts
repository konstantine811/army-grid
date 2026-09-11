import type {
  CellValue,
  ExcelWorkbookSnapshot,
} from "../../../../excelRoundTrip";
import { isJournalPersonId, norm, normId, normKey } from "./cellText";

export const findEjoosSheet = (
  workbook: ExcelWorkbookSnapshot,
  matcher: RegExp,
) => workbook.sheets.find((sheet) => matcher.test(sheet.sheetName));

export const headerMap = (row: CellValue[]) => {
  const map = new Map<string, number>();
  row.forEach((cell, index) => {
    const key = normKey(norm(cell));
    if (key && !map.has(key)) map.set(key, index);
  });
  return map;
};

export const findCol = (map: Map<string, number>, ...needles: RegExp[]) => {
  for (const [key, index] of map.entries()) {
    if (needles.some((re) => re.test(key))) return index;
  }
  return -1;
};

export const cell = (row: CellValue[] | undefined, index: number) =>
  index >= 0 ? norm(row?.[index]) : "";

export const idCell = (row: CellValue[] | undefined, index: number) =>
  index >= 0 ? normId(row?.[index]) : "";

export const journalIdCell = (row: CellValue[] | undefined, index: number) => {
  const value = idCell(row, index);
  return isJournalPersonId(value) ? value : "";
};
