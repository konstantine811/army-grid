import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  EJOOS_PERSON_DATA_START_ROW,
  collectExcludedPositionDateWrites,
  formatExcludedDestination,
  formatExcludedListBasis,
  formatExcludedPositionDates,
  formatTimesheetTransferMark,
  OOS_TO_EXCLUDED_BASE,
} from "./ejoosExcludedColumns";
import {
  parseEjoosShpo,
  personIdFromShpo,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";
import { stripTimesheetDivisionLabel } from "./ejoosTimesheetUnitSections";
import {
  findTimesheetPersonRowsInGrid,
  loadTimesheetGridFromFile,
  mergeTimesheetGrids,
  pickTimesheetKeepRow,
  uniqueExcelRows,
} from "./ejoosTimesheetPersonRows";
import {
  dayFromOrderLabel,
  isTimesheetDepartureMark,
  parseTimesheetAbsenceSpans,
  timesheetTransferMarkForDay,
} from "./ejoosTimesheetText";
import { excludeWritePlan } from "./ejoosExcludePolicy";

const valueOf = (value: CellValue | unknown): string | number | null => {
  if (value === undefined || value === null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Math.floor(
      (Date.UTC(
        value.getFullYear(),
        value.getMonth(),
        value.getDate(),
      ) -
        Date.UTC(1899, 11, 30)) /
        86400000,
    );
  }
  if (typeof value === "number" || typeof value === "string") return value;
  const rich = value as { text?: () => string; value?: () => unknown };
  if (typeof rich.text === "function") return rich.text() || null;
  if (typeof rich.value === "function") return valueOf(rich.value());
  return null;
};

const findSheet = (book: ExcelWorkbookSnapshot, pattern: RegExp) =>
  book.sheets.find((sheet) => pattern.test(sheet.sheetName));

const cellText = (sheet: ExcelSheetSnapshot, row: number, column: number) => {
  const text = String(valueOf(sheet.rawRows[row - 1]?.[column - 1]) ?? "").trim();
  if (column === 1 && text) return stripTimesheetDivisionLabel(text);
  return text;
};

const nameKey = (value: string) =>
  value.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ").trim();

const findPersonRow = (
  sheet: ExcelSheetSnapshot,
  input: { personId: string; fullName: string; nameColumn: number; idColumn: number },
) => {
  const targetName = nameKey(input.fullName);
  for (let row = 6; row <= sheet.rawRows.length; row += 1) {
    const rowId = cellText(sheet, row, input.idColumn);
    const rowName = nameKey(cellText(sheet, row, input.nameColumn));
    if (input.personId && rowId === input.personId) return row;
    if (targetName && rowName === targetName) return row;
  }
  return 0;
};

const timesheetGridHasDeparture = (
  grid: Array<unknown[] | undefined>,
  row: number,
) => {
  const cells = grid[row - 1] ?? [];
  for (let day = 1; day <= 31; day += 1) {
    if (isTimesheetDepartureMark(cells[7 + day])) return true;
  }
  return false;
};

const nextExcludedRow = (sheet: ExcelSheetSnapshot) => {
  let last = 5;
  for (let row = 6; row <= sheet.rawRows.length; row += 1) {
    if ([1, 2, 28, 29, 30, 31, 32].some((column) => cellText(sheet, row, column))) {
      last = row;
    }
  }
  return last + 1;
};

const hasExcludedData = (sheet: ExcelSheetSnapshot, row: number) =>
  [1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 28, 29, 30, 31, 32].some((column) =>
    cellText(sheet, row, column),
  );

const findExcludedStyleSourceRow = (
  sheet: ExcelSheetSnapshot,
  beforeRow: number,
) => {
  for (let row = beforeRow - 1; row >= EJOOS_PERSON_DATA_START_ROW; row -= 1) {
    if (hasExcludedData(sheet, row)) return row;
  }
  for (
    let row = EJOOS_PERSON_DATA_START_ROW;
    row <= sheet.rawRows.length;
    row += 1
  ) {
    if (hasExcludedData(sheet, row)) return row;
  }
  return EJOOS_PERSON_DATA_START_ROW;
};

const timesheetRowHasDeparture = (
  sheet: ExcelSheetSnapshot,
  row: number,
) => {
  for (let day = 1; day <= 31; day += 1) {
    if (isTimesheetDepartureMark(cellText(sheet, row, 8 + day))) return true;
  }
  return false;
};

const findTimesheetStyleRow = (sheet: ExcelSheetSnapshot) => {
  for (let row = 7; row <= sheet.rawRows.length; row += 1) {
    if (cellText(sheet, row, 2) || cellText(sheet, row, 7)) return row;
  }
  return 7;
};

export const nextTimesheetHistoryRow = (
  grid: Array<unknown[] | undefined>,
  sourceRow: number,
  reserved: Set<number>,
) => {
  const start = sourceRow >= 7 ? sourceRow + 1 : 7;
  for (let row = start; row <= grid.length + 30; row += 1) {
    if (reserved.has(row)) continue;
    const gridRow = grid[row - 1];
    if (
      !String(gridRow?.[0] ?? "").trim() &&
      !String(gridRow?.[1] ?? "").trim() &&
      !String(gridRow?.[6] ?? "").trim() &&
      !String(gridRow?.[7] ?? "").trim()
    ) {
      reserved.add(row);
      return row;
    }
  }
  const row = grid.length + reserved.size + 1;
  reserved.add(row);
  return row;
};

const excludedNumberValue = (
  column: number,
  value: string | number | null,
): string | number | null => {
  if (value == null || value === "") return null;
  if (typeof value === "number") return value;
  const trimmed = String(value).trim();
  if (
    (column === 3 ||
      column === 4 ||
      column === 9 ||
      column === 10 ||
      column === 12 ||
      column === 30) &&
    /^\d+$/.test(trimmed)
  ) {
    return Number(trimmed);
  }
  return value;
};

const excludedStyledWrite = (
  row: number,
  column: number,
  value: string | number | null,
  styleSourceRow: number,
): ZipCellWrite => {
  const next =
    column === 5 && value != null && value !== ""
      ? formatExcludedPositionDates(value) || value
      : value;
  return {
    row,
    column,
    value: excludedNumberValue(column, next),
    styleSourceRow,
    styleSourceColumn: column,
    // Не шукаємо «канонічний» стиль по аркушу: у шаблоні Виключених
    // нижні рядки часто жовті, і тоді нова підстава фарбується жовтим.
    copyNeighborStyle: false,
    heightSourceRow: styleSourceRow,
    wrapText:
      column === 5 ||
      (typeof next === "string" && next.includes("\n")),
  };
};

const timesheetHistoryDayMarks = (
  op: EjoosSyncOp,
  plan: EjoosSyncPlan,
  departMark: string,
) => {
  const departDay = dayFromOrderLabel(op.payload.excludeDate);
  const activeFromDay = dayFromOrderLabel(op.payload.timesheetActiveFrom);
  const absenceSpans = parseTimesheetAbsenceSpans(
    op.payload.timesheetAbsenceSpans || "",
  );
  const lastDay = Math.min(31, Math.max(departDay, plan.timesheetDay));
  const days: Array<{ day: number; value: string | null }> = [];
  let presentDays = 0;
  for (let day = 1; day <= 31; day += 1) {
    const value = timesheetTransferMarkForDay({
      day,
      departDay,
      lastDay,
      activeFromDay,
      absenceSpans,
      departMark,
    });
    days.push({ day, value });
    if (value === "+") presentDays += 1;
  }
  return { days, presentDays };
};

const timesheetStyledWrite = (
  row: number,
  column: number,
  value: string | number | null,
  sourceRow: number,
  sourceColumn = column,
): ZipCellWrite => {
  const wrapText =
    typeof value === "string" &&
    (value.includes("\n") || /вибув/iu.test(value));
  return {
    row,
    column,
    value,
    styleSourceRow: sourceRow,
    styleSourceColumn: sourceColumn,
    // Як у Виключених: не шукати «канонічний» стиль по аркушу.
    // Порожні рядки Табеля — жовтий шаблон вакансії без синьої ПІБ і товстої сітки.
    copyNeighborStyle: false,
    keepNeighborStyle: !wrapText,
    heightSourceRow: sourceRow,
    wrapText,
  };
};

const TIMESHEET_STYLE_LAST_COLUMN = 40;

/** Існуючі дати прийняття посади у Виключених: кожна з нового рядка, рано→пізно. */
export async function applyExcludedPositionDatesPresentation(input: {
  file: Blob;
  ejoos: ExcelWorkbookSnapshot;
}): Promise<Blob> {
  const excluded = findSheet(input.ejoos, /виключен/i);
  if (!excluded) return input.file;
  const writes = collectExcludedPositionDateWrites(excluded.rawRows);
  if (!writes.length) return input.file;
  return applyInlineStringWritesToWorkbook(
    input.file,
    excluded.sheetName,
    writes,
  );
}

/** Після xlsx-populate: перенести `s=` зайнятого рядка на новий історичний. */
export async function copyTimesheetRowStylesWithZip(
  file: Blob,
  jobs: Array<{ sourceRow: number; targetRow: number }>,
): Promise<Blob> {
  const unique = new Map<string, { sourceRow: number; targetRow: number }>();
  for (const job of jobs) {
    if (
      job.sourceRow < 7 ||
      job.targetRow < 7 ||
      job.sourceRow === job.targetRow
    ) {
      continue;
    }
    unique.set(`${job.sourceRow}:${job.targetRow}`, job);
  }
  if (!unique.size) return file;
  const writes: ZipCellWrite[] = [];
  for (const { sourceRow, targetRow } of unique.values()) {
    for (let column = 1; column <= TIMESHEET_STYLE_LAST_COLUMN; column += 1) {
      writes.push({
        row: targetRow,
        column,
        value: null,
        styleOnly: true,
        styleSourceRow: sourceRow,
        styleSourceColumn: column,
        copyNeighborStyle: false,
        keepNeighborStyle: true,
        heightSourceRow: sourceRow,
      });
    }
  }
  return applyInlineStringWritesToWorkbook(file, /табель/i, writes);
}

const findArrivalCloseColumns = (sheet: ExcelSheetSnapshot) => {
  let departDateCol = 0;
  let orderDateCol = 0;
  let orderNumberCol = 0;
  let orderCombinedCol = 0;
  for (let row = 1; row <= 6; row += 1) {
    for (let column = 1; column <= 40; column += 1) {
      const key = cellText(sheet, row, column)
        .toLocaleLowerCase("uk-UA")
        .replace(/\s+/g, " ")
        .trim();
      if (!key) continue;
      if (/дата вибут/.test(key) && !/наказ/.test(key)) {
        departDateCol = column;
      } else if (
        /дата.{0,12}наказ.{0,12}вибут|наказ.{0,12}вибут.{0,12}дата/.test(key)
      ) {
        orderDateCol = column;
      } else if (
        /номер.{0,12}наказ.{0,12}вибут|наказ.{0,12}вибут.{0,12}номер/.test(key)
      ) {
        orderNumberCol = column;
      } else if (/наказ.{0,16}вибут/.test(key) && !orderCombinedCol) {
        orderCombinedCol = column;
      }
    }
  }
  return { departDateCol, orderDateCol, orderNumberCol, orderCombinedCol };
};

const excludedDestinationValue = (op: EjoosSyncOp) =>
  formatExcludedDestination(op.payload.documentsDest || op.payload.changeText) ||
  null;

const excludedPatchValues = (
  op: EjoosSyncOp,
  personId: string,
  positionIndex: string,
): Array<[number, string | number | null]> => [
  [1, op.payload.fromRank || op.rank || null],
  [2, op.payload.fromName || op.fullName || null],
  [3, personId || null],
  [4, positionIndex || null],
  [6, op.payload.arrivedFrom || null],
  [7, op.payload.enlistDate || null],
  [8, op.payload.enlistOrderDate || null],
  [9, op.payload.enlistOrderNumber || null],
  [10, op.payload.appointmentOrderNumber || null],
  [11, op.payload.appointmentOrderDate || null],
  [12, op.payload.lastRankOrderDate || null],
  [13, op.payload.lastRankOrderNumber || null],
  [28, op.payload.excludeDate || null],
  [29, op.payload.orderDate || null],
  [30, op.payload.orderNumber || null],
  [31, excludedDestinationValue(op)],
  [32, formatExcludedListBasis(op.payload)],
];

export async function applyExcludeTransfersWithZip(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
}) {
  const { ejoos, plan, ops } = input;
  const excluded = findSheet(ejoos, /виключен/i);
  const oos = findSheet(ejoos, /(^|[.\s])оос($|[\s])/i);
  const shpo = findSheet(ejoos, /шпо|штатно.?посад/i);
  const timesheet = findSheet(ejoos, /табель/i);
  const arrivals = findSheet(ejoos, /тимчасов.*прибул/i);
  if (!excluded || !oos || !shpo || !timesheet) {
    throw new Error("У ЕЖООС не знайдено всі аркуші для переведення");
  }

  const arrivalCloseColumns = arrivals
    ? findArrivalCloseColumns(arrivals)
    : null;
  const reservedTimesheetRows = new Set<number>();
  const writtenTimesheetByPerson = new Map<string, number[]>();
  const timesheetXmlGrid = await loadTimesheetGridFromFile(
    ejoos.file,
    timesheet.sheetName,
  );
  const timesheetGrid = mergeTimesheetGrids(
    timesheet.rawRows,
    timesheetXmlGrid,
  );
  let excludedRow = nextExcludedRow(excluded);
  const excludedStyleSourceRow = findExcludedStyleSourceRow(
    excluded,
    excludedRow,
  );
  const shpoRows = parseEjoosShpo(shpo);

  // Atomic preflight: an external transfer must have a safe Timesheet action
  // before we even start accumulating writes for Excluded/OOS/SHPO.
  for (const op of ops) {
    if (op.kind !== "exclude_transfer") continue;
    const personId = personIdFromShpo(shpoRows, {
      fullName: op.payload.fromName || op.fullName,
      positionIndex:
        op.payload.occupiedPositionIndex ||
        op.payload.fromPositionIndex ||
        op.payload.previousIndex ||
        op.positionIndex,
      personId: op.payload.fromPersonId || op.personId,
    });
    const rows = findTimesheetPersonRowsInGrid(
      timesheetGrid,
      personId,
      op.payload.fromName || op.fullName,
    );
    const planned = Number(op.payload.timesheetExcelRow || 0);
    const sourceRow = planned || rows[0] || 0;
    const writePlan = excludeWritePlan(op.payload);
    const timesheetAction =
      sourceRow > 0
        ? writePlan.replaceInPlace
          ? "PATCH_HISTORY"
          : "MOVE_TO_HISTORY"
        : writePlan.createTimesheetHistory
          ? "CREATE_HISTORY_IN_SOURCE_SECTION"
          : "";
    if (!timesheetAction) {
      throw new Error(
        `TIMESHEET_SOURCE_UNRESOLVED: ${op.fullName || op.personId}`,
      );
    }
  }

  const excludedWrites: ZipCellWrite[] = [];
  const oosWrites: ZipCellWrite[] = [];
  const shpoWrites: ZipCellWrite[] = [];
  const timesheetWrites: ZipCellWrite[] = [];
  const arrivalWrites: ZipCellWrite[] = [];

  for (const op of ops) {
    const personId = personIdFromShpo(shpoRows, {
      fullName: op.payload.fromName || op.fullName,
      positionIndex:
        op.payload.occupiedPositionIndex ||
        op.payload.fromPositionIndex ||
        op.payload.previousIndex ||
        op.positionIndex,
      personId: op.payload.fromPersonId || op.personId,
    });
    const existingExcludedRow = Number(op.payload.excludedExcelRow || 0);
    let oosRow = Number(op.payload.oosExcelRow || 0);
    const positionIndex =
      op.payload.fromPositionIndex ||
      op.payload.previousIndex ||
      op.positionIndex;

    if (existingExcludedRow) {
      for (const [column, value] of excludedPatchValues(
        op,
        personId,
        positionIndex,
      )) {
        if (value == null || value === "") continue;
        excludedWrites.push(
          excludedStyledWrite(
            existingExcludedRow,
            column,
            value,
            existingExcludedRow,
          ),
        );
      }
      const currentDates =
        oosRow > 0
          ? oos.rawRows[oosRow - 1]?.[4]
          : excluded.rawRows[existingExcludedRow - 1]?.[4];
      const formattedDates = formatExcludedPositionDates(currentDates);
      if (formattedDates) {
        excludedWrites.push(
          excludedStyledWrite(
            existingExcludedRow,
            5,
            formattedDates,
            existingExcludedRow,
          ),
        );
      }
    } else {
      const targetRow = excludedRow++;
      if (oosRow > 0) {
        for (const [fromColumn, toColumn] of OOS_TO_EXCLUDED_BASE) {
          const raw = valueOf(oos.rawRows[oosRow - 1]?.[fromColumn - 1]);
          excludedWrites.push(
            excludedStyledWrite(
              targetRow,
              toColumn,
              toColumn === 5 ? formatExcludedPositionDates(raw) || raw : raw,
              excludedStyleSourceRow,
            ),
          );
        }
      }
      for (const [column, value] of excludedPatchValues(
        op,
        personId,
        positionIndex,
      )) {
        if (value == null || value === "") continue;
        excludedWrites.push(
          excludedStyledWrite(targetRow, column, value, excludedStyleSourceRow),
        );
      }
    }

    const timesheetName = op.payload.fromName || op.fullName;
    const personTimesheetKey = `${personId || ""}|${nameKey(timesheetName)}`;
    const existingTimesheetRows = uniqueExcelRows([
      ...findTimesheetPersonRowsInGrid(
        timesheetGrid,
        personId,
        timesheetName,
      ),
      ...(writtenTimesheetByPerson.get(personTimesheetKey) ?? []),
    ]);
    const plannedKeep = Number(op.payload.timesheetExcelRow || 0);
    const isClosedHistoryRow = (row: number) =>
      timesheetRowHasDeparture(timesheet, row) ||
      timesheetGridHasDeparture(timesheetGrid, row);
    const openTimesheetRows = existingTimesheetRows.filter(
      (row) => !isClosedHistoryRow(row),
    );
    let timesheetKeepRow = pickTimesheetKeepRow(
      openTimesheetRows,
      () => true,
      plannedKeep && openTimesheetRows.includes(plannedKeep) ? plannedKeep : 0,
    );
    if (!timesheetKeepRow && !openTimesheetRows.length) {
      timesheetKeepRow = pickTimesheetKeepRow(
        existingTimesheetRows,
        isClosedHistoryRow,
        plannedKeep,
      );
    }
    const writePlan = excludeWritePlan(op.payload);
    if (!timesheetKeepRow && writePlan.replaceInPlace && plannedKeep > 0) {
      timesheetKeepRow = plannedKeep;
    }
    const { days, presentDays } = timesheetHistoryDayMarks(
      op,
      plan,
      formatTimesheetTransferMark(op.payload),
    );
    const timesheetCellValue = (row: number, column: number) => {
      const fromRaw = valueOf(timesheet.rawRows[row - 1]?.[column - 1]);
      if (fromRaw != null && String(fromRaw).trim() !== "") return fromRaw;
      return valueOf(timesheetGrid[row - 1]?.[column - 1]);
    };
    const copyTimesheetRow = (sourceRow: number, targetRow: number) => {
      // Колонка A — повтор роти/батальйону на рядку даних (не заголовок наступної секції).
      timesheetWrites.push({
        row: targetRow,
        column: 1,
        value:
          timesheetCellValue(sourceRow, 1) ||
          timesheetCellValue(Math.max(targetRow - 1, 7), 1) ||
          null,
        styleSourceRow: sourceRow,
        styleSourceColumn: 1,
        copyNeighborStyle: false,
        keepNeighborStyle: true,
      });
      for (let column = 2; column <= TIMESHEET_STYLE_LAST_COLUMN; column += 1) {
        timesheetWrites.push(
          timesheetStyledWrite(
            targetRow,
            column,
            timesheetCellValue(sourceRow, column),
            sourceRow,
            column,
          ),
        );
      }
    };
    const clearTimesheetOccupant = (row: number) => {
      for (const column of [1, 6, 7, 8, 40]) {
        timesheetWrites.push({
          row,
          column,
          value: null,
          styleSourceRow: row,
          styleSourceColumn: column,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        });
      }
      for (let day = 1; day <= 31; day += 1) {
        timesheetWrites.push({
          row,
          column: 8 + day,
          value: null,
          styleSourceRow: row,
          styleSourceColumn: 9,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        });
      }
    };
    const writeHistoryOnRow = (targetRow: number, styleRow: number) => {
      const styleDayColumn = (() => {
        for (let day = 1; day <= 31; day += 1) {
          const mark = cellText(timesheet, styleRow, 8 + day);
          if (mark === "+" || /вибув/iu.test(mark)) return 8 + day;
        }
        return 9;
      })();
      if (op.payload.fromRank || op.rank) {
        timesheetWrites.push(
          timesheetStyledWrite(
            targetRow,
            6,
            op.payload.fromRank || op.rank,
            styleRow,
            6,
          ),
        );
      }
      timesheetWrites.push(
        timesheetStyledWrite(
          targetRow,
          7,
          op.payload.fromName || op.fullName || null,
          styleRow,
          7,
        ),
        timesheetStyledWrite(targetRow, 8, personId || null, styleRow, 8),
      );
      if (positionIndex && !cellText(timesheet, targetRow, 2)) {
        timesheetWrites.push(
          timesheetStyledWrite(targetRow, 2, positionIndex, styleRow, 2),
        );
      }
      for (const { day, value } of days) {
        timesheetWrites.push(
          timesheetStyledWrite(targetRow, 8 + day, value, styleRow, styleDayColumn),
        );
      }
      timesheetWrites.push(
        timesheetStyledWrite(targetRow, 40, presentDays, styleRow, 40),
      );
    };
    const sourceRow = timesheetKeepRow || plannedKeep;
    if (writePlan.replaceInPlace && sourceRow > 0) {
      writeHistoryOnRow(sourceRow, sourceRow);
      timesheetKeepRow = sourceRow;
    } else if (sourceRow > 0) {
      timesheetKeepRow = nextTimesheetHistoryRow(
        timesheetXmlGrid,
        sourceRow,
        reservedTimesheetRows,
      );
      copyTimesheetRow(sourceRow, timesheetKeepRow);
      writeHistoryOnRow(timesheetKeepRow, sourceRow);
      clearTimesheetOccupant(sourceRow);
    } else if (writePlan.createTimesheetHistory) {
      const styleRow = findTimesheetStyleRow(timesheet);
      timesheetKeepRow = nextTimesheetHistoryRow(
        timesheetXmlGrid,
        styleRow,
        reservedTimesheetRows,
      );
      writeHistoryOnRow(timesheetKeepRow, styleRow);
    }
    if (timesheetKeepRow > 0) {
      writtenTimesheetByPerson.set(personTimesheetKey, [timesheetKeepRow]);
      reservedTimesheetRows.add(timesheetKeepRow);
    }
    for (const row of existingTimesheetRows) {
      if (row === timesheetKeepRow) continue;
      if (isClosedHistoryRow(row) && writePlan.replaceInPlace) {
        continue;
      }
      for (const column of [1, 6, 7, 8, 40]) {
        timesheetWrites.push({
          row,
          column,
          value: null,
          styleSourceRow: row,
          styleSourceColumn: column,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        });
      }
      for (let day = 1; day <= 31; day += 1) {
        timesheetWrites.push({
          row,
          column: 8 + day,
          value: null,
          styleSourceRow: row,
          styleSourceColumn: 9,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        });
      }
    }

    const shpoRow = Number(op.payload.shpoExcelRow || 0);
    if (shpoRow > 0) {
      const rowPersonId = cellText(shpo, shpoRow, 8);
      const rowName = cellText(shpo, shpoRow, 7);
      const fromName = op.payload.fromName || op.fullName || "";
      const clearingSamePerson =
        (!rowPersonId && !rowName) ||
        (personId && rowPersonId === personId) ||
        (fromName &&
          rowName &&
          nameKey(rowName) === nameKey(fromName));
      if (clearingSamePerson) {
        for (const column of [6, 7, 8, 18]) {
          shpoWrites.push({ row: shpoRow, column, value: null });
        }
      }
    }
    if (oosRow <= 0) {
      oosRow = findPersonRow(oos, {
        personId,
        fullName: op.payload.fromName || op.fullName,
        nameColumn: 2,
        idColumn: 3,
      });
    }
    if (oosRow > 0) {
      for (let column = 1; column <= 40; column += 1) {
        oosWrites.push({ row: oosRow, column, value: null });
      }
    }

    const arrivalRow = Number(op.payload.arrivalExcelRow || 0);
    if (arrivals && arrivalCloseColumns && arrivalRow > 0) {
      const departDate =
        op.payload.arrivalDepartDate || op.payload.appointmentOrderDate || "";
      const orderNumber =
        op.payload.arrivalDepartOrderNumber ||
        op.payload.appointmentOrderNumber ||
        op.payload.orderNumber ||
        "";
      const orderDate =
        op.payload.arrivalDepartOrderDate ||
        op.payload.appointmentOrderDate ||
        "";
      if (arrivalCloseColumns.departDateCol) {
        arrivalWrites.push({
          row: arrivalRow,
          column: arrivalCloseColumns.departDateCol,
          value: departDate || null,
        });
      }
      if (arrivalCloseColumns.orderDateCol) {
        arrivalWrites.push({
          row: arrivalRow,
          column: arrivalCloseColumns.orderDateCol,
          value: orderDate || null,
        });
      }
      if (arrivalCloseColumns.orderNumberCol) {
        arrivalWrites.push({
          row: arrivalRow,
          column: arrivalCloseColumns.orderNumberCol,
          value: orderNumber ? `№${orderNumber}` : null,
        });
      }
      if (arrivalCloseColumns.orderCombinedCol) {
        arrivalWrites.push({
          row: arrivalRow,
          column: arrivalCloseColumns.orderCombinedCol,
          value: [orderNumber ? `№${orderNumber}` : "", orderDate ? `від ${orderDate}` : ""]
            .filter(Boolean)
            .join(" "),
        });
      }
    }
  }

  let blob: Blob | File = ejoos.file;
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    excluded.sheetName,
    excludedWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(blob, oos.sheetName, oosWrites);
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    shpo.sheetName,
    shpoWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    timesheet.sheetName,
    timesheetWrites,
  );
  if (arrivals && arrivalWrites.length) {
    blob = await applyInlineStringWritesToWorkbook(
      blob,
      arrivals.sheetName,
      arrivalWrites,
    );
  }
  return blob;
}
