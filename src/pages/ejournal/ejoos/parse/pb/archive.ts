import type { ExcelWorkbookSnapshot } from "../../../../../excelRoundTrip";
import { norm } from "../cellText";
import type { PbArchivePeriod } from "../../types/pb";
import {
  cell,
  findCol,
  findEjoosSheet,
  headerMap,
  idCell,
} from "../sheetLookup";

export const parsePbArchive = (
  workbook: ExcelWorkbookSnapshot,
): PbArchivePeriod[] => {
  const sheet = findEjoosSheet(workbook, /^archive$/i);
  if (!sheet) return [];
  const headerRowIndex = sheet.rawRows.findIndex((row) =>
    /вид\s*вибут|прізвище/i.test(row.map(norm).join(" ")),
  );
  if (headerRowIndex < 0) return [];
  const headers = headerMap(sheet.rawRows[headerRowIndex] ?? []);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /прізвище|піб/);
  const rankCol = findCol(headers, /зван/);
  const typeCol = findCol(headers, /вид\s*вибут/);
  const dateCol = findCol(headers, /з якої дати|дата вибут/);
  const placeCol = findCol(headers, /куди виб/);
  const orderNumCol = findCol(headers, /номер наказу вибут/);
  const orderDateCol = findCol(headers, /дата наказу вибут/);
  const plannedCol = findCol(headers, /планова дата/);
  const returnCol = findCol(
    headers,
    /^дата прибуття$/,
    /фактичн.*(?:поверн|прибут)/,
    /^дата поверн/,
  );
  const returnOrderNumCol = findCol(
    headers,
    /^номер наказу$/,
    /номер наказу\s*(?:прибут|поверн)/,
    /наказ\s*(?:на\s+)?(?:прибут|поверн)/,
  );
  const returnOrderDateCol = findCol(
    headers,
    /^дата наказу$/,
    /дата наказу\s*(?:прибут|поверн)/,
  );
  const posCol = findCol(headers, /займана посад|посада/);
  const numCol = findCol(headers, /№\s*з\/п|^№$/);
  const periods: PbArchivePeriod[] = [];

  sheet.rawRows.slice(headerRowIndex + 1).forEach((row, offset) => {
    const fullName = cell(row, nameCol);
    if (!fullName) return;
    periods.push({
      excelRow: headerRowIndex + offset + 2,
      periodNumber: cell(row, numCol),
      personId: idCell(row, idCol),
      fullName,
      rank: cell(row, rankCol),
      positionTitle: cell(row, posCol),
      absenceType: cell(row, typeCol),
      departDate: cell(row, dateCol),
      place: cell(row, placeCol),
      orderNumber: cell(row, orderNumCol),
      orderDate: cell(row, orderDateCol),
      plannedReturn: cell(row, plannedCol),
      returnDate: cell(row, returnCol),
      returnOrderNumber: cell(row, returnOrderNumCol),
      returnOrderDate: cell(row, returnOrderDateCol),
    });
  });
  return periods;
};
