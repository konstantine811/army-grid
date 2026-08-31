import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  EJOOS_PERSON_DATA_START_ROW,
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
  cellValueToOosText,
  findOosStyleSourceRow,
  mergeOosHistoryValue,
} from "./ejoosOosText";
import {
  clipAbsenceSpansToActiveEpisode,
  dayFromOrderLabel,
  historyAbsenceSpansForClosedEpisode,
  isTimesheetDepartureMark,
  parseTimesheetAbsenceSpans,
  timesheetCodeOnDay,
  timesheetMarkBeforeDeparture,
  timesheetMarkFromArchive,
} from "./ejoosTimesheetText";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";

const valueOf = (value: CellValue | unknown): string | number | null => {
  if (value === undefined || value === null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Math.floor(
      (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) -
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

const textAt = (sheet: ExcelSheetSnapshot, row: number, column: number) =>
  String(valueOf(sheet.rawRows[row - 1]?.[column - 1]) ?? "").trim();

const nameKey = (value: string) =>
  value.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ").trim();

const nextExcludedRow = (sheet: ExcelSheetSnapshot) => {
  let last = 5;
  for (let row = 6; row <= sheet.rawRows.length; row += 1) {
    if (
      [1, 2, 28, 29, 30, 31, 32].some((column) => textAt(sheet, row, column))
    ) {
      last = row;
    }
  }
  return last + 1;
};

const nextTimesheetFreeRow = (
  sheet: ExcelSheetSnapshot,
  sourceRow: number,
  reserved: Set<number>,
) => {
  for (let row = sourceRow + 1; row <= sheet.rawRows.length + 30; row += 1) {
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

/** Історію в ООС ведемо в одній клітинці: найновіше значення — першим рядком. */
const historyValue = (
  sheet: ExcelSheetSnapshot,
  row: number,
  column: number,
  next: string,
) =>
  mergeOosHistoryValue(
    sheet.rawRows[row - 1]?.[column - 1],
    next,
  ) || null;

type PositionChangeContext = {
  plan: EjoosSyncPlan;
  oos: ExcelSheetSnapshot;
  timesheet: ExcelSheetSnapshot;
  shpo: ExcelSheetSnapshot;
  excludedWrites: ZipCellWrite[];
  shpoWrites: ZipCellWrite[];
  oosWrites: ZipCellWrite[];
  timesheetWrites: ZipCellWrite[];
  reservedTimesheetRows: Set<number>;
  takeExcludedRow: () => number;
  excludedStyleSourceRow: number;
};

const collectWrites = (op: EjoosSyncOp, ctx: PositionChangeContext) => {
  const { plan, oos, timesheet, shpo } = ctx;
  const rank = op.payload.nextRank || op.rank;
  const fullName = op.payload.nextName || op.fullName;
  const personId = personIdFromShpo(parseEjoosShpo(shpo), {
    fullName,
    positionIndex: op.payload.nextIndex || op.positionIndex || op.payload.previousIndex,
    personId: op.payload.nextPersonId || op.payload.fromPersonId || op.personId,
  });
  const nextIndex = op.payload.nextIndex || op.positionIndex;
  const oosHistoryIndexes =
    op.payload.oosHistoryIndexes || nextIndex || "";
  const previousIndex =
    op.payload.fromPositionIndex || op.payload.previousIndex;
  const appointmentDate = op.payload.orderDate || plan.timesheetDayLabel;
  const departDay = dayFromOrderLabel(op.payload.excludeDate || appointmentDate);
  const lastDay = Math.min(31, Math.max(departDay, plan.timesheetDay));
  const preserveHistory = op.payload.timesheetPreserveHistory === "1";
  const returningToStaffIndex = op.payload.returningToStaffIndex === "1";
  const activeFromDay =
    dayFromOrderLabel(op.payload.timesheetActiveFrom || op.payload.transferCancelDate) ||
    departDay;

  // 1) Виключені — історія закриття старої штатної посади.
  const oosRow = Number(op.payload.oosExcelRow || 0);
  const excludedRow = ctx.takeExcludedRow();
  if (oosRow > 0) {
    for (const [fromColumn, toColumn] of OOS_TO_EXCLUDED_BASE) {
      ctx.excludedWrites.push({
        row: excludedRow,
        column: toColumn,
        value: valueOf(oos.rawRows[oosRow - 1]?.[fromColumn - 1]),
        styleSourceRow: ctx.excludedStyleSourceRow,
        wrapText: true,
      });
    }
  }
  const excludedValues: Array<[number, string | number | null]> = [
    [1, op.payload.fromRank || rank || null],
    [2, op.payload.fromName || fullName || null],
    [3, op.payload.fromPersonId || personId || null],
    [4, previousIndex || null],
    [28, op.payload.excludeDate || appointmentDate || null],
    [29, op.payload.orderDate || null],
    [
      30,
      op.payload.orderNumber && /^\d+$/.test(op.payload.orderNumber.trim())
        ? Number(op.payload.orderNumber.trim())
        : op.payload.orderNumber || null,
    ],
    [
      31,
      String(op.payload.documentsDest || "")
        .replace(/\s+/g, " ")
        .trim()
        .toLocaleLowerCase("uk-UA") || null,
    ],
    [
      32,
      formatExcludedListBasis({
        ...op.payload,
        exclusionReason: op.payload.exclusionReason || "ПЕРЕВЕДЕННЯ 1 ПБ",
      }),
    ],
  ];
  for (const [column, value] of excludedValues) {
    ctx.excludedWrites.push({
      row: excludedRow,
      column,
      value,
      styleSourceRow: ctx.excludedStyleSourceRow,
      wrapText: true,
    });
  }

  // 2) Табель — стару позицію закриваємо копією рядка з датою вибуття.
  const oldTimesheetRow = Number(
    op.payload.previousIndexTimesheetExcelRow || 0,
  );
  if (oldTimesheetRow > 0) {
    const historyRow = nextTimesheetFreeRow(
      timesheet,
      oldTimesheetRow,
      ctx.reservedTimesheetRows,
    );
    for (let column = 1; column <= 40; column += 1) {
      ctx.timesheetWrites.push({
        row: historyRow,
        column,
        value: valueOf(timesheet.rawRows[oldTimesheetRow - 1]?.[column - 1]),
        styleSourceRow: oldTimesheetRow,
      });
    }
    let presentDays = 0;
    const historySpans = historyAbsenceSpansForClosedEpisode(
      op.payload,
      departDay,
    );
    for (let day = 1; day <= 31; day += 1) {
      let value: string | null = null;
      if (departDay > 0 && day < departDay) {
        value = timesheetMarkBeforeDeparture(day, departDay, historySpans);
        if (value === "+") presentDays += 1;
      } else if (day === departDay) {
        value = formatTimesheetTransferMark(op.payload);
      }
      ctx.timesheetWrites.push({
        row: historyRow,
        column: 8 + day,
        value,
        styleSourceRow: oldTimesheetRow,
        wrapText: typeof value === "string" && /вибув/iu.test(value),
      });
      ctx.timesheetWrites.push({
        row: oldTimesheetRow,
        column: 8 + day,
        value: null,
      });
    }
    ctx.timesheetWrites.push({
      row: historyRow,
      column: 40,
      value: presentDays,
      styleSourceRow: oldTimesheetRow,
    });
    if (personId && !textAt(timesheet, oldTimesheetRow, 8)) {
      ctx.timesheetWrites.push({
        row: historyRow,
        column: 8,
        value: personId,
        styleSourceRow: oldTimesheetRow,
      });
    }
    for (const column of [6, 7, 8]) {
      ctx.timesheetWrites.push({ row: oldTimesheetRow, column, value: null });
    }
  }

  // 3) Табель — нова штатна позиція з «+» від дати наказу.
  const newTimesheetRow = Number(op.payload.timesheetExcelRow || 0);
  if (newTimesheetRow > 0) {
    const occupiedBy = textAt(timesheet, newTimesheetRow, 7);
    const occupiedId = textAt(timesheet, newTimesheetRow, 8);
    const samePerson =
      Boolean(occupiedBy || occupiedId) &&
      (nameKey(occupiedBy) === nameKey(fullName) ||
        (Boolean(personId) && occupiedId === personId));
    const otherPerson = Boolean(occupiedBy || occupiedId) && !samePerson;
    const existingHistoryRow = Number(op.payload.previousTimesheetExcelRow || 0);
    const presenceFrom =
      activeFromDay ||
      dayFromOrderLabel(op.payload.timesheetActiveFrom || appointmentDate) ||
      departDay;
    const absenceSpans = clipAbsenceSpansToActiveEpisode(
      parseTimesheetAbsenceSpans(op.payload.timesheetAbsenceSpans || ""),
      presenceFrom > 1 ? presenceFrom : 1,
    );

    if (returningToStaffIndex && existingHistoryRow > 0) {
      const activeRow = nextTimesheetFreeRow(
        timesheet,
        Math.max(newTimesheetRow, existingHistoryRow),
        ctx.reservedTimesheetRows,
      );
      const indexValue = valueOf(timesheet.rawRows[newTimesheetRow - 1]?.[1]);
      if (indexValue != null) {
        ctx.timesheetWrites.push({
          row: activeRow,
          column: 2,
          value: indexValue,
          styleSourceRow: newTimesheetRow,
        });
      }
      for (const [column, value] of [
        [6, rank],
        [7, fullName],
        [8, personId],
      ] as Array<[number, string]>) {
        ctx.timesheetWrites.push({
          row: activeRow,
          column,
          value: value || null,
          styleSourceRow: newTimesheetRow,
        });
      }
      for (let day = 1; day <= 31; day += 1) {
        const value = timesheetMarkFromArchive(day, {
          activeFromDay: presenceFrom,
          lastDay,
          spans: absenceSpans,
          fillBeforeActive: presenceFrom > 1,
        });
        if (value == null) continue;
        ctx.timesheetWrites.push({
          row: activeRow,
          column: 8 + day,
          value,
          styleSourceRow: newTimesheetRow,
        });
      }
      if (otherPerson) {
        for (const column of [6, 7, 8]) {
          ctx.timesheetWrites.push({
            row: newTimesheetRow,
            column,
            value: null,
            styleSourceRow: newTimesheetRow,
          });
        }
        for (let day = 1; day <= 31; day += 1) {
          ctx.timesheetWrites.push({
            row: newTimesheetRow,
            column: 8 + day,
            value: null,
            styleSourceRow: newTimesheetRow,
          });
        }
        ctx.timesheetWrites.push({
          row: newTimesheetRow,
          column: 40,
          value: null,
          styleSourceRow: newTimesheetRow,
        });
      }
    } else if (otherPerson || (preserveHistory && existingHistoryRow <= 0)) {
      const historyRow = nextTimesheetFreeRow(
        timesheet,
        newTimesheetRow,
        ctx.reservedTimesheetRows,
      );
      for (let column = 1; column <= 40; column += 1) {
        ctx.timesheetWrites.push({
          row: historyRow,
          column,
          value: valueOf(timesheet.rawRows[newTimesheetRow - 1]?.[column - 1]),
          styleSourceRow: newTimesheetRow,
        });
      }
      if (otherPerson) {
        for (let day = Math.max(1, departDay); day <= 31; day += 1) {
          ctx.timesheetWrites.push({
            row: historyRow,
            column: 8 + day,
            value: null,
            styleSourceRow: newTimesheetRow,
          });
        }
      } else if (preserveHistory && !samePerson) {
        ctx.timesheetWrites.push(
          { row: historyRow, column: 6, value: rank || null },
          { row: historyRow, column: 7, value: fullName || null },
          { row: historyRow, column: 8, value: personId || null },
        );
      }
      if (preserveHistory && samePerson) {
        const staffHasDeparture = Array.from({ length: 31 }, (_, i) => i + 1).some(
          (day) =>
            isTimesheetDepartureMark(
              timesheet.rawRows[newTimesheetRow - 1]?.[7 + day],
            ),
        );
        if (!staffHasDeparture) {
          const cancelDepartDay = dayFromOrderLabel(op.payload.cancelledTransferDate);
          if (cancelDepartDay > 0) {
            let presentDays = 0;
            for (let day = 1; day <= 31; day += 1) {
              let value: string | null = null;
              if (day < cancelDepartDay) {
                value = "+";
                presentDays += 1;
              } else if (day === cancelDepartDay) {
                value = formatTimesheetTransferMark({
                  destination: op.payload.cancelledTransferDest,
                  timesheetDestination: op.payload.cancelledTransferDest,
                });
              } else if (day <= lastDay) {
                value = "-";
              }
              ctx.timesheetWrites.push({
                row: historyRow,
                column: 8 + day,
                value,
                styleSourceRow: newTimesheetRow,
                wrapText: typeof value === "string" && /вибув/iu.test(value),
              });
            }
            ctx.timesheetWrites.push({
              row: historyRow,
              column: 40,
              value: presentDays,
              styleSourceRow: newTimesheetRow,
            });
          }
        }
      }
    }
    if (!returningToStaffIndex || existingHistoryRow <= 0) {
    for (const [column, value] of [
      [6, rank],
      [7, fullName],
      [8, personId],
    ] as Array<[number, string]>) {
      ctx.timesheetWrites.push({
        row: newTimesheetRow,
        column,
        value: value || null,
      });
    }
    for (let day = 1; day <= 31; day += 1) {
        const value = timesheetMarkFromArchive(day, {
          activeFromDay: presenceFrom,
          lastDay,
          spans: absenceSpans,
          fillBeforeActive: presenceFrom > 1,
        });
      if (value == null) continue;
      ctx.timesheetWrites.push({
        row: newTimesheetRow,
        column: 8 + day,
        value,
      });
    }
    }
  }

  const historyAbsenceRow = Number(
    op.payload.historyTimesheetExcelRow ||
      op.payload.previousTimesheetExcelRow ||
      0,
  );
  const historyAbsenceSpans = parseTimesheetAbsenceSpans(
    op.payload.historyTimesheetAbsenceSpans || "",
  );
  if (
    historyAbsenceRow > 0 &&
    historyAbsenceSpans.length &&
    historyAbsenceRow !== newTimesheetRow
  ) {
    for (let day = 1; day <= 31; day += 1) {
      const code = timesheetCodeOnDay(day, historyAbsenceSpans);
      if (!code) continue;
      const current = textAt(timesheet, historyAbsenceRow, 8 + day);
      if (isTimesheetDepartureMark(current)) continue;
      ctx.timesheetWrites.push({
        row: historyAbsenceRow,
        column: 8 + day,
        value: code,
        styleSourceRow: historyAbsenceRow,
      });
    }
  }

  // 4) ШПО — стара посада стає вакантною, особа переходить на новий індекс.
  const oldShpoRow = Number(op.payload.previousShpoExcelRow || 0);
  if (oldShpoRow > 0) {
    for (const column of [6, 7, 8, 18]) {
      ctx.shpoWrites.push({ row: oldShpoRow, column, value: null });
    }
  }
  const dispositionShpoRow = Number(op.payload.dispositionShpoExcelRow || 0);
  if (dispositionShpoRow > 0) {
    for (const column of [2, 3, 6, 7, 8]) {
      ctx.shpoWrites.push({ row: dispositionShpoRow, column, value: null });
    }
  }
  const dispositionTimesheetRow = Number(
    op.payload.dispositionTimesheetExcelRow || 0,
  );
  if (dispositionTimesheetRow > 0) {
    for (const column of [6, 7, 8]) {
      ctx.timesheetWrites.push({
        row: dispositionTimesheetRow,
        column,
        value: null,
      });
    }
    for (let day = 1; day <= 31; day += 1) {
      ctx.timesheetWrites.push({
        row: dispositionTimesheetRow,
        column: 8 + day,
        value: null,
      });
    }
    ctx.timesheetWrites.push({
      row: dispositionTimesheetRow,
      column: 40,
      value: null,
    });
  }
  const staleTimesheetRow = Number(op.payload.clearTimesheetExcelRow || 0);
  if (
    staleTimesheetRow > 0 &&
    staleTimesheetRow !== Number(op.payload.timesheetExcelRow || 0)
  ) {
    for (const column of [6, 7, 8]) {
      ctx.timesheetWrites.push({
        row: staleTimesheetRow,
        column,
        value: null,
      });
    }
    for (let day = 1; day <= 31; day += 1) {
      ctx.timesheetWrites.push({
        row: staleTimesheetRow,
        column: 8 + day,
        value: null,
      });
    }
    ctx.timesheetWrites.push({ row: staleTimesheetRow, column: 40, value: null });
  }
  const newShpoRow = Number(op.payload.shpoExcelRow || 0);
  if (newShpoRow > 0) {
    const orderText = [
      "ПОСАДА",
      op.payload.orderNumber ? `№${op.payload.orderNumber}` : "",
      appointmentDate ? `від ${appointmentDate}` : "",
    ]
      .filter(Boolean)
      .join(" ");
    for (const [column, value] of [
      [6, rank],
      [7, fullName],
      [8, personId],
      [18, orderText],
    ] as Array<[number, string]>) {
      ctx.shpoWrites.push({ row: newShpoRow, column, value: value || null });
    }
  }

  // 5) ООС — запис лишається активним, дописуємо історію посад.
  if (oosRow > 0) {
    const oosStyleRow = findOosStyleSourceRow(
      (row, column) => oos.rawRows[row - 1]?.[column - 1],
      { skipRows: [oosRow], lastRow: oos.rawRows.length },
    );
    const oosId =
      personId || cellValueToOosText(oos.rawRows[oosRow - 1]?.[2]) || null;
    const oosIdentity: Array<[number, string | number | null]> = [
      [1, rank || cellValueToOosText(oos.rawRows[oosRow - 1]?.[0]) || null],
      [2, fullName || cellValueToOosText(oos.rawRows[oosRow - 1]?.[1]) || null],
      [3, oosId && /^\d+$/.test(oosId) ? Number(oosId) : oosId],
    ];
    for (const [column, value] of oosIdentity) {
      ctx.oosWrites.push({
        row: oosRow,
        column,
        value,
        copyNeighborStyle: true,
        styleSourceRow: oosStyleRow,
        wrapText: true,
      });
    }
    const historyDates =
      op.payload.oosHistoryDates ||
      oosHistoryIndexes
        .split("\n")
        .filter(Boolean)
        .map(() => appointmentDate)
        .join("\n");
    const oosHistory: Array<[number, string]> = [
      [4, oosHistoryIndexes || nextIndex],
      [5, historyDates || appointmentDate],
      [11, op.payload.orderNumber],
      [12, appointmentDate],
    ];
    for (const [column, value] of oosHistory) {
      const next = historyValue(oos, oosRow, column, value || "");
      if (next === null) continue;
      ctx.oosWrites.push({
        row: oosRow,
        column,
        value: next,
        copyNeighborStyle: true,
        styleSourceRow: oosStyleRow,
        wrapText: true,
      });
    }
  }
};

/**
 * Зміна штатної посади всередині 1ПБ: стару посаду закриваємо як історію
 * (Виключені + Табель + звільнення індексу в ШПО), особу ставимо на новий
 * індекс, а активний запис в ООС зберігаємо — людина лишається в частині.
 */
export async function applyPositionChangeWithZip(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
}) {
  const { ejoos, plan, ops } = input;
  const excluded = findSheet(ejoos, /виключен/i);
  const shpo = findSheet(ejoos, /шпо|штатно.?посад/i);
  const oos = findSheet(ejoos, /(^|[.\s])оос($|[\s])/i);
  const timesheet = findSheet(ejoos, /табель/i);
  if (!excluded || !shpo || !oos || !timesheet) {
    throw new Error("У ЕЖООС не знайдено всі аркуші для зміни посади");
  }

  let excludedRow = nextExcludedRow(excluded);
  const ctx: PositionChangeContext = {
    plan,
    oos,
    timesheet,
    shpo,
    excludedWrites: [],
    shpoWrites: [],
    oosWrites: [],
    timesheetWrites: [],
    reservedTimesheetRows: new Set<number>(),
    takeExcludedRow: () => excludedRow++,
    excludedStyleSourceRow: Math.max(
      EJOOS_PERSON_DATA_START_ROW,
      excludedRow - 1,
    ),
  };
  for (const op of ops) {
    collectWrites(op, ctx);
    const duplicateRow = Number(op.payload.duplicateTimesheetExcelRow || 0);
    if (duplicateRow > 0) {
      if (op.payload.clearTimesheetIndex === "1") {
        ctx.timesheetWrites.push({
          row: duplicateRow,
          column: 2,
          value: null,
          styleSourceRow: duplicateRow,
        });
      }
      for (const column of [6, 7, 8, 40]) {
        ctx.timesheetWrites.push({
          row: duplicateRow,
          column,
          value: null,
          styleSourceRow: duplicateRow,
        });
      }
      for (let day = 1; day <= 31; day += 1) {
        ctx.timesheetWrites.push({
          row: duplicateRow,
          column: 8 + day,
          value: null,
          styleSourceRow: duplicateRow,
        });
      }
    }
  }

  let blob: Blob | File = ejoos.file;
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    excluded.sheetName,
    ctx.excludedWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    timesheet.sheetName,
    ctx.timesheetWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    shpo.sheetName,
    ctx.shpoWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    oos.sheetName,
    ctx.oosWrites,
  );
  return blob;
}
