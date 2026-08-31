import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  EJOOS_PERSON_DATA_START_ROW,
  formatExcludedDestinationText,
  formatExcludedListBasis,
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
import {
  dayFromOrderLabel,
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

const cellText = (sheet: ExcelSheetSnapshot, row: number, column: number) =>
  String(valueOf(sheet.rawRows[row - 1]?.[column - 1]) ?? "").trim();

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

const findTimesheetStyleRow = (sheet: ExcelSheetSnapshot) => {
  for (let row = 7; row <= sheet.rawRows.length; row += 1) {
    if (cellText(sheet, row, 2) || cellText(sheet, row, 7)) return row;
  }
  return 7;
};

const nextTimesheetHistoryRow = (
  sheet: ExcelSheetSnapshot,
  sourceRow: number,
  reserved: Set<number>,
) => {
  for (
    let row = sourceRow + 1;
    row <= sheet.rawRows.length + 30;
    row += 1
  ) {
    if (reserved.has(row)) continue;
    if (
      !cellText(sheet, row, 2) &&
      !cellText(sheet, row, 7) &&
      !cellText(sheet, row, 8)
    ) {
      reserved.add(row);
      return row;
    }
  }
  const row = sheet.rawRows.length + reserved.size + 1;
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
): ZipCellWrite => ({
  row,
  column,
  value: excludedNumberValue(column, value),
  styleSourceRow,
  styleSourceColumn: column,
  // Не шукаємо «канонічний» стиль по аркушу: у шаблоні Виключених
  // нижні рядки часто жовті, і тоді нова підстава фарбується жовтим.
  copyNeighborStyle: false,
  heightSourceRow: styleSourceRow,
  wrapText: typeof value === "string" && value.includes("\n"),
});

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
  formatExcludedDestinationText(
    op.payload.documentsDest || op.payload.destination || "",
  ) || null;

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

  const excludedWrites: ZipCellWrite[] = [];
  const oosWrites: ZipCellWrite[] = [];
  const shpoWrites: ZipCellWrite[] = [];
  const timesheetWrites: ZipCellWrite[] = [];
  const arrivalWrites: ZipCellWrite[] = [];
  const arrivalCloseColumns = arrivals
    ? findArrivalCloseColumns(arrivals)
    : null;
  const reservedTimesheetRows = new Set<number>();
  let excludedRow = nextExcludedRow(excluded);
  const excludedStyleSourceRow = findExcludedStyleSourceRow(
    excluded,
    excludedRow,
  );
  const shpoRows = parseEjoosShpo(shpo);

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
    } else {
      const targetRow = excludedRow++;
      if (oosRow > 0) {
        for (const [fromColumn, toColumn] of OOS_TO_EXCLUDED_BASE) {
          excludedWrites.push(
            excludedStyledWrite(
              targetRow,
              toColumn,
              valueOf(oos.rawRows[oosRow - 1]?.[fromColumn - 1]),
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

    const sourceTimesheetRow = Number(op.payload.timesheetExcelRow || 0);
    if (sourceTimesheetRow > 0) {
      const targetRow = nextTimesheetHistoryRow(
        timesheet,
        sourceTimesheetRow,
        reservedTimesheetRows,
      );
      const styleDayColumn = (() => {
        for (let day = 1; day <= 31; day += 1) {
          const mark = cellText(timesheet, sourceTimesheetRow, 8 + day);
          if (mark === "+" || /вибув/iu.test(mark)) return 8 + day;
        }
        return 9;
      })();
      for (let column = 1; column <= 40; column += 1) {
        timesheetWrites.push(
          timesheetStyledWrite(
            targetRow,
            column,
            valueOf(timesheet.rawRows[sourceTimesheetRow - 1]?.[column - 1]),
            sourceTimesheetRow,
            column <= 8 ? column : styleDayColumn,
          ),
        );
      }
      const { days, presentDays } = timesheetHistoryDayMarks(
        op,
        plan,
        formatTimesheetTransferMark(op.payload),
      );
      for (const { day, value } of days) {
        timesheetWrites.push(
          timesheetStyledWrite(
            targetRow,
            8 + day,
            value,
            sourceTimesheetRow,
            styleDayColumn,
          ),
        );
        timesheetWrites.push({
          row: sourceTimesheetRow,
          column: 8 + day,
          value: null,
          styleSourceRow: sourceTimesheetRow,
          styleSourceColumn: styleDayColumn,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        });
      }
      timesheetWrites.push(
        timesheetStyledWrite(
          targetRow,
          40,
          presentDays,
          sourceTimesheetRow,
          40,
        ),
      );
      if (op.payload.fromRank || op.rank) {
        timesheetWrites.push(
          timesheetStyledWrite(
            targetRow,
            6,
            op.payload.fromRank || op.rank,
            sourceTimesheetRow,
            6,
          ),
        );
      }
      if (op.payload.fromName || op.fullName) {
        timesheetWrites.push(
          timesheetStyledWrite(
            targetRow,
            7,
            op.payload.fromName || op.fullName,
            sourceTimesheetRow,
            7,
          ),
        );
      }
      if (personId && !cellText(timesheet, sourceTimesheetRow, 8)) {
        timesheetWrites.push(
          timesheetStyledWrite(targetRow, 8, personId, sourceTimesheetRow, 8),
        );
      }
      for (const column of [6, 7, 8]) {
        timesheetWrites.push({
          row: sourceTimesheetRow,
          column,
          value: null,
          styleSourceRow: sourceTimesheetRow,
          styleSourceColumn: column,
          copyNeighborStyle: false,
          keepNeighborStyle: true,
        });
      }
    } else if (excludeWritePlan(op.payload).createTimesheetHistory) {
      const styleRow = findTimesheetStyleRow(timesheet);
      const targetRow = nextTimesheetHistoryRow(
        timesheet,
        styleRow,
        reservedTimesheetRows,
      );
      const styleDayColumn = 9;
      timesheetWrites.push(
        timesheetStyledWrite(targetRow, 2, positionIndex || null, styleRow, 2),
        timesheetStyledWrite(
          targetRow,
          6,
          op.payload.fromRank || op.rank || null,
          styleRow,
          6,
        ),
        timesheetStyledWrite(
          targetRow,
          7,
          op.payload.fromName || op.fullName || null,
          styleRow,
          7,
        ),
        timesheetStyledWrite(targetRow, 8, personId || null, styleRow, 8),
      );
      const { days, presentDays } = timesheetHistoryDayMarks(
        op,
        plan,
        formatTimesheetTransferMark(op.payload),
      );
      for (const { day, value } of days) {
        timesheetWrites.push(
          timesheetStyledWrite(targetRow, 8 + day, value, styleRow, styleDayColumn),
        );
      }
      timesheetWrites.push(
        timesheetStyledWrite(targetRow, 40, presentDays, styleRow, 40),
      );
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
