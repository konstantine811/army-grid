import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  parseEjoosShpo,
  personIdFromShpo,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import { formatTimesheetDeparture } from "./ejoosTimesheetText";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

const valueOf = (value: CellValue | unknown): string | number | null => {
  if (value === undefined || value === null || value === "") return null;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Math.floor(
      (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) -
        Date.UTC(1899, 11, 30)) /
        86400000,
    );
  }
  const rich = value as { text?: () => string; value?: () => unknown };
  if (typeof rich.text === "function") return rich.text() || null;
  if (typeof rich.value === "function") return valueOf(rich.value());
  return null;
};

const findSheet = (book: ExcelWorkbookSnapshot, pattern: RegExp) =>
  book.sheets.find((sheet) => pattern.test(sheet.sheetName));

const textAt = (sheet: ExcelSheetSnapshot, row: number, column: number) =>
  String(valueOf(sheet.rawRows[row - 1]?.[column - 1]) ?? "").trim();

const dispositionTarget = (
  sheet: ExcelSheetSnapshot,
  reserved: Set<number>,
) => {
  let styleRow = 0;
  let lastRow = 0;
  for (let row = 1; row <= sheet.rawRows.length; row += 1) {
    const location = textAt(sheet, row, 2);
    if (!/РОЗПОРЯДЖ/iu.test(location)) continue;
    lastRow = row;
    if (textAt(sheet, row, 7) || textAt(sheet, row, 8)) styleRow = row;
    if (
      !reserved.has(row) &&
      !textAt(sheet, row, 7) &&
      !textAt(sheet, row, 8)
    ) {
      reserved.add(row);
      return { row, styleRow: styleRow || row };
    }
  }
  let row = lastRow + 1;
  while (reserved.has(row)) row += 1;
  reserved.add(row);
  return { row, styleRow: styleRow || Math.max(1, lastRow) };
};

const nextAbsentRow = (sheet: ExcelSheetSnapshot, reserved: Set<number>) => {
  for (let row = 6; row <= sheet.rawRows.length + 5; row += 1) {
    if (reserved.has(row) || textAt(sheet, row, 2)) continue;
    reserved.add(row);
    return row;
  }
  const row = sheet.rawRows.length + reserved.size + 1;
  reserved.add(row);
  return row;
};

const nextTimesheetRow = (
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
      !textAt(sheet, row, 2) &&
      !textAt(sheet, row, 7) &&
      !textAt(sheet, row, 8)
    ) {
      reserved.add(row);
      return row;
    }
  }
  const row = sheet.rawRows.length + reserved.size + 1;
  reserved.add(row);
  return row;
};

/**
 * У РУХ призначення приходить у різних формах: «у розпорядження командира…»,
 * «, який знаходиться у розпорядженні командира…». Зводимо до знахідного
 * відмінка без прийменника, щоб далі підставити потрібну форму.
 */
const dispositionPlace = (destination: string) =>
  String(destination || "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^[,;]\s*/, "")
    .replace(/^який\s+знаходиться\s+/iu, "")
    .replace(/^(?:у|в)\s+розпорядженн\S*/iu, "розпорядження")
    .replace(/^розпорядженн\S*/iu, "розпорядження")
    .replace(/\s+/g, " ")
    .trim();

const dispositionLocation = (place: string) => {
  if (!place) return ", який знаходиться у розпорядженні";
  const tail = place.replace(/^розпорядження\s*/iu, "").trim();
  return tail === place
    ? `, який знаходиться ${place}`
    : `, який знаходиться у розпорядженні ${tail}`.trimEnd();
};

/** У ШПО блок розпорядження тримає короткі коди: ЗБ, СЗЧ, лік — не «БЕЗВІСТИ». */
const shpoDispositionMark = (op: EjoosSyncOp) => {
  const code = String(op.payload.absenceCode || "").trim();
  if (code && code.length <= 8) return code;
  const raw = String(op.payload.absenceType || "").trim();
  if (/безвіст/iu.test(raw)) return "ЗБ";
  if (/сзч|самовіл/iu.test(raw)) return "СЗЧ";
  if (/полон/iu.test(raw)) return "пол";
  return raw || "РОЗПОРЯДЖЕННЯ";
};

type DispositionContext = {
  plan: EjoosSyncPlan;
  shpo: ExcelSheetSnapshot;
  absent: ExcelSheetSnapshot;
  timesheet: ExcelSheetSnapshot;
  shpoWrites: ZipCellWrite[];
  absentWrites: ZipCellWrite[];
  timesheetWrites: ZipCellWrite[];
  reservedShpoRows: Set<number>;
  reservedAbsentRows: Set<number>;
  reservedTimesheetRows: Set<number>;
};

const collectWrites = (op: EjoosSyncOp, ctx: DispositionContext) => {
  const { plan, shpo, absent, timesheet } = ctx;
  const shpoWrites = ctx.shpoWrites;
  const absentWrites = ctx.absentWrites;
  const timesheetWrites = ctx.timesheetWrites;
  const oldShpoRow = Number(op.payload.shpoExcelRow || 0);
  if (oldShpoRow > 0) {
    for (const column of [6, 7, 8, 18]) {
      shpoWrites.push({ row: oldShpoRow, column, value: null });
    }
  }

  const absenceLabel = op.payload.absenceType || "РОЗПОРЯДЖЕННЯ";
  const statusMark = shpoDispositionMark(op);
  const place = dispositionPlace(op.payload.destination);
  // Якщо особа вже стоїть у блоці «у розпорядженні», другий раз не додаємо.
  if (op.payload.skipShpoDisposition !== "1") {
    const target = dispositionTarget(shpo, ctx.reservedShpoRows);
    for (const [column, value] of [
      [2, dispositionLocation(place)],
      [3, statusMark],
      [6, op.rank],
      [7, op.fullName],
      [8, op.personId],
    ] as Array<[number, string]>) {
      shpoWrites.push({
        row: target.row,
        column,
        value: value || null,
        styleSourceRow: target.styleRow,
      });
    }
  }

  // Новий запис відсутності додаємо лише коли чинного немає, і лише з даними
  // архіву — порожній рядок тільки засмічує аркуш.
  if (op.payload.needsAbsenceRecord === "1" && op.payload.absenceDate) {
    const absentRow = nextAbsentRow(absent, ctx.reservedAbsentRows);
    const absentValues: Array<[number, string | null]> = [
      [1, op.rank || null],
      [2, op.fullName || null],
      [
        3,
        personIdFromShpo(parseEjoosShpo(shpo), {
          fullName: op.fullName,
          positionIndex: op.positionIndex,
          personId: op.personId,
        }) || null,
      ],
      [4, op.positionIndex || null],
      [5, absenceLabel || null],
      [6, op.payload.absencePlace || null],
      [7, op.payload.absenceDate || null],
      [
        9,
        [op.payload.absenceOrderDate, op.payload.absenceOrderNumber]
          .filter(Boolean)
          .join(" ") || null,
      ],
      [12, op.payload.plannedReturn || null],
    ];
    for (const [column, value] of absentValues) {
      absentWrites.push({
        row: absentRow,
        column,
        value,
        styleSourceRow: 6,
      });
    }
  }

  const sourceTimesheetRow = Number(op.payload.timesheetExcelRow || 0);
  if (sourceTimesheetRow > 0) {
    const historyRow = nextTimesheetRow(
      timesheet,
      sourceTimesheetRow,
      ctx.reservedTimesheetRows,
    );
    for (let column = 1; column <= 40; column += 1) {
      timesheetWrites.push({
        row: historyRow,
        column,
        value: valueOf(
          timesheet.rawRows[sourceTimesheetRow - 1]?.[column - 1],
        ),
        styleSourceRow: sourceTimesheetRow,
      });
    }
    const orderDay = Number(
      String(op.payload.orderDate || "").match(/^(\d{1,2})/)?.[1] || 0,
    );
    const absenceCode = op.payload.absenceCode || statusMark;
    for (let day = 1; day <= Math.min(31, plan.timesheetDay); day += 1) {
      const value =
        day === orderDay
          ? [
              formatTimesheetDeparture(place),
              op.payload.orderNumber
                ? `наказ №${op.payload.orderNumber}`
                : "",
              op.payload.orderDate ? `від ${op.payload.orderDate}` : "",
            ]
              .filter(Boolean)
              .join(" ")
          : day > orderDay && orderDay > 0
            ? absenceCode
            : valueOf(
                timesheet.rawRows[sourceTimesheetRow - 1]?.[8 + day - 1],
              );
      timesheetWrites.push({
        row: historyRow,
        column: 8 + day,
        value,
        styleSourceRow: sourceTimesheetRow,
      });
      timesheetWrites.push({
        row: sourceTimesheetRow,
        column: 8 + day,
        value: null,
      });
    }
    const personId = personIdFromShpo(parseEjoosShpo(shpo), {
      fullName: op.fullName,
      positionIndex: op.positionIndex || op.payload.previousIndex,
      personId: op.personId,
    });
    if (personId && !textAt(timesheet, sourceTimesheetRow, 8)) {
      timesheetWrites.push({
        row: historyRow,
        column: 8,
        value: personId,
        styleSourceRow: sourceTimesheetRow,
      });
    }
    for (const column of [6, 7, 8]) {
      timesheetWrites.push({
        row: sourceTimesheetRow,
        column,
        value: null,
      });
    }
  }
};

/**
 * Виведення в розпорядження: звільняємо штатну позицію в ШПО й додаємо особу
 * до блоку «у розпорядженні», за потреби відкриваємо запис відсутності та
 * закриваємо штатний рядок Табеля. ООС і Виключені не змінюються.
 */
export async function applyDispositionWithZip(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
}) {
  const { ejoos, plan, ops } = input;
  const shpo = findSheet(ejoos, /шпо|штатно.?посад/i);
  const absent = findSheet(ejoos, /тимчасов.*відсут/i);
  const timesheet = findSheet(ejoos, /табель/i);
  if (!shpo || !absent || !timesheet) {
    throw new Error("Не знайдено ШПО, Тимчасово відсутні або Табель");
  }

  const ctx: DispositionContext = {
    plan,
    shpo,
    absent,
    timesheet,
    shpoWrites: [],
    absentWrites: [],
    timesheetWrites: [],
    reservedShpoRows: new Set<number>(),
    reservedAbsentRows: new Set<number>(),
    reservedTimesheetRows: new Set<number>(),
  };
  for (const op of ops) collectWrites(op, ctx);
  const { shpoWrites, absentWrites, timesheetWrites } = ctx;

  let blob: Blob | File = ejoos.file;
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    shpo.sheetName,
    shpoWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    absent.sheetName,
    absentWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    timesheet.sheetName,
    timesheetWrites,
  );
  return blob;
}
