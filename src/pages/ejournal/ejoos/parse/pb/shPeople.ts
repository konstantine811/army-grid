import type { ExcelWorkbookSnapshot } from "../../../../../excelRoundTrip";
import type { PbShPerson } from "../../types/pb";
import {
  cell,
  findCol,
  findEjoosSheet,
  headerMap,
  idCell,
} from "../sheetLookup";

export const parsePbShPeople = (
  workbook: ExcelWorkbookSnapshot,
): PbShPerson[] => {
  const sheet = findEjoosSheet(workbook, /^sh$/i);
  if (!sheet) return [];
  const headers = headerMap(sheet.rawRows[0] ?? []);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /^піб$/, /прізвище/);
  const rankCol = findCol(headers, /зван/);
  const indexCol = findCol(headers, /індекс\s*посад/);
  const posCol = findCol(headers, /^посада$/);
  const statusCol = findCol(headers, /^статус$/, /^перебуван/);
  const fromCol = findCol(headers, /звідки.*прибув|^звідки$/);
  const people: PbShPerson[] = [];

  sheet.rawRows.slice(1).forEach((row, offset) => {
    const fullName = cell(row, nameCol);
    const positionIndex = cell(row, indexCol);
    const status = cell(row, statusCol);
    const personId = idCell(row, idCol);
    if (!fullName && !personId && !positionIndex) return;
    if (!fullName && (!status || status === "0")) return;
    people.push({
      excelRow: offset + 2,
      personId: personId && personId !== "0" ? personId : "",
      fullName,
      rank: cell(row, rankCol),
      positionIndex,
      positionTitle: cell(row, posCol),
      status,
      arrivedFrom: cell(row, fromCol),
    });
  });
  return people;
};
