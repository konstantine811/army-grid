import type { ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { tryParseExcelSerialDate } from "../../shared/format";
import { canonicalName, isJournalPersonId, normId } from "./ejoosIdentity";
import { norm } from "./ejoos/parse/cellText";
import {
  cellValueToOosText,
  isOosSectionHeaderText,
  readOosPersonIdCell,
} from "./ejoosOosText";
import {
  parseEjoosAbsents,
  parseEjoosExcluded,
  parseEjoosShpo,
  parseEjoosTimesheetPeople,
  personIdFromShpo,
} from "./ejoosSyncPlan";
import { findEjoosSheet } from "./ejoos/parse/sheetLookup";
import { readEjoosWorkbookSnapshot } from "./ejoosTimesheetPersonRows";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

const OOS_ID_COLUMN = 3;
const OOS_NAME_COLUMN = 2;
const OOS_INDEX_COLUMN = 4;
const OOS_DATA_START_ROW = 6;

type PersonIdRow = {
  personId: string;
  fullName: string;
  positionIndex?: string;
};

const usableJournalId = (value: unknown) => {
  const id = normId(value);
  return isJournalPersonId(id) ? id : "";
};

const oosRowPositionIndex = (row: ExcelSheetSnapshot["rawRows"][number]) => {
  const raw = norm(row?.[OOS_INDEX_COLUMN - 1]);
  return raw.match(/\d{5,}/)?.[0] || raw;
};

/** Колонка C показує дату замість штатного ID. */
export const oosPersonIdCellLooksCorrupted = (value: unknown) => {
  if (readOosPersonIdCell(value)) return false;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  const text = cellValueToOosText(value);
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text)) return true;
  return tryParseExcelSerialDate(value) !== null;
};

const buildPersonIdLookups = (rows: PersonIdRow[]) => {
  const byName = new Map<string, string>();
  const byIndex = new Map<string, string>();
  for (const row of rows) {
    const id = usableJournalId(row.personId);
    if (!id) continue;
    const name = canonicalName(row.fullName);
    if (name && !byName.has(name)) byName.set(name, id);
    const index = String(row.positionIndex || "").trim();
    if (index && !byIndex.has(index)) byIndex.set(index, id);
  }
  return { byName, byIndex };
};

const resolveOosPersonId = (input: {
  fullName: string;
  positionIndex: string;
  shpoRows: ReturnType<typeof parseEjoosShpo>;
  timesheetLookups: ReturnType<typeof buildPersonIdLookups>;
  excludedLookups: ReturnType<typeof buildPersonIdLookups>;
  absentLookups: ReturnType<typeof buildPersonIdLookups>;
}) => {
  const fromShpo = personIdFromShpo(input.shpoRows, {
    fullName: input.fullName,
    positionIndex: input.positionIndex,
  });
  if (isJournalPersonId(fromShpo)) return fromShpo;

  const name = canonicalName(input.fullName);
  if (name) {
    const fromTimesheet = input.timesheetLookups.byName.get(name);
    if (fromTimesheet) return fromTimesheet;
    const fromExcluded = input.excludedLookups.byName.get(name);
    if (fromExcluded) return fromExcluded;
    const fromAbsent = input.absentLookups.byName.get(name);
    if (fromAbsent) return fromAbsent;
  }

  if (input.positionIndex) {
    const fromTimesheet = input.timesheetLookups.byIndex.get(input.positionIndex);
    if (fromTimesheet) return fromTimesheet;
    const fromExcluded = input.excludedLookups.byIndex.get(input.positionIndex);
    if (fromExcluded) return fromExcluded;
    const fromAbsent = input.absentLookups.byIndex.get(input.positionIndex);
    if (fromAbsent) return fromAbsent;
  }

  return "";
};

export const buildOosPersonIdRepairWrites = (
  workbook: ExcelWorkbookSnapshot,
): ZipCellWrite[] => {
  const oosSheet = findEjoosSheet(workbook, /(^|[.\s])оос($|[\s])/i);
  if (!oosSheet) return [];

  const shpoSheet = findEjoosSheet(workbook, /шпо|штатно.?посад/i);
  const timesheetSheet = findEjoosSheet(workbook, /табель/i);
  const excludedSheet = findEjoosSheet(workbook, /виключен/i);
  const absentSheet = findEjoosSheet(workbook, /тимчасов.*відсут/i);

  const shpoRows = parseEjoosShpo(shpoSheet);
  const timesheetLookups = buildPersonIdLookups(parseEjoosTimesheetPeople(timesheetSheet));
  const excludedLookups = buildPersonIdLookups(parseEjoosExcluded(excludedSheet));
  const absentLookups = buildPersonIdLookups(parseEjoosAbsents(absentSheet));

  const writes: ZipCellWrite[] = [];

  for (let row = OOS_DATA_START_ROW; row <= oosSheet.rawRows.length; row += 1) {
    const cells = oosSheet.rawRows[row - 1];
    const fullName = norm(cells?.[OOS_NAME_COLUMN - 1]);
    const rawIdCell = cells?.[OOS_ID_COLUMN - 1];
    if (!fullName && rawIdCell == null) continue;
    if (isOosSectionHeaderText(fullName)) continue;

    const positionIndex = oosRowPositionIndex(cells);
    const resolvedId = resolveOosPersonId({
      fullName,
      positionIndex,
      shpoRows,
      timesheetLookups,
      excludedLookups,
      absentLookups,
    });
    if (!resolvedId) continue;

    const currentId = readOosPersonIdCell(rawIdCell);
    const corrupted = oosPersonIdCellLooksCorrupted(rawIdCell);
    if (!corrupted && currentId === resolvedId) continue;

    writes.push({ row, column: OOS_ID_COLUMN, value: resolvedId });
  }

  return writes;
};

export async function fixOosPersonIdCells(file: Blob): Promise<Blob> {
  let workbook: ExcelWorkbookSnapshot;
  try {
    workbook = await readEjoosWorkbookSnapshot(
      new File([file], "ejoos-oos-id-repair.xlsx", {
        type:
          file.type ||
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
    );
  } catch {
    return file;
  }
  const oosSheet = findEjoosSheet(workbook, /(^|[.\s])оос($|[\s])/i);
  if (!oosSheet) return file;
  const writes = buildOosPersonIdRepairWrites(workbook);
  if (!writes.length) return file;
  try {
    return await applyInlineStringWritesToWorkbook(file, oosSheet.sheetName, writes);
  } catch {
    return file;
  }
}
