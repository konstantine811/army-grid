import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import { canonicalName, normId, normKey } from "./ejoosIdentity";
import { formatTimesheetTransferMark } from "./ejoosExcludedColumns";
import { isTimesheetDepartureMark, stripTimesheetDirectionPrefix } from "./ejoosTimesheetText";
import type { TimesheetLayout } from "./ejoosTimesheetLayout";

export type TimesheetTransferActionKind =
  | "MOVE_TO_HISTORY"
  | "PATCH_HISTORY"
  | "CREATE_HISTORY_IN_SOURCE_SECTION";

export type TimesheetTransferAction = {
  kind: TimesheetTransferActionKind;
  sourceIndex: string;
  sourceRow: number;
  sourceSection: string;
  targetHistoryRow: number;
  createTimesheetHistory: boolean;
  replaceInPlace: boolean;
};

export type TimesheetTransferErrorCode =
  | "TIMESHEET_SOURCE_UNRESOLVED"
  | "EXTERNAL_TRANSFER_WITHOUT_TIMESHEET_ACTION";

export class TimesheetTransferError extends Error {
  code: TimesheetTransferErrorCode;

  constructor(code: TimesheetTransferErrorCode, message: string) {
    super(message);
    this.name = "TimesheetTransferError";
    this.code = code;
  }
}

const DATA_START_ROW = 7;

const cellText = (
  sheet: ExcelSheetSnapshot,
  row: number,
  column: number,
  grid?: Array<unknown[] | undefined>,
) => {
  const fromGrid = String(grid?.[row - 1]?.[column - 1] ?? "").trim();
  if (fromGrid) return fromGrid;
  return String(sheet.rawRows[row - 1]?.[column - 1] ?? "").trim();
};

const lastDataRow = (
  sheet: ExcelSheetSnapshot,
  grid?: Array<unknown[] | undefined>,
) => Math.max(sheet.rawRows.length, grid?.length ?? 0);

export const belongedToUnitThisMonth = (payload: Record<string, string>) =>
  payload.transitSameMonth === "1" ||
  Boolean(payload.timesheetActiveFrom?.trim()) ||
  Boolean(payload.priorPlacementDate?.trim()) ||
  Boolean(payload.appointmentOrderDate?.trim()) ||
  Boolean(payload.shpoExcelRow?.trim()) ||
  Boolean(payload.oosExcelRow?.trim()) ||
  Boolean(payload.arrivalExcelRow?.trim());

export const historyMatchesCurrentTransfer = (input: {
  departureText: string;
  payload: Record<string, string>;
}) => {
  const departure = normKey(input.departureText);
  if (!departure) return false;
  const mark = normKey(formatTimesheetTransferMark(input.payload));
  if (mark && (departure.includes(mark) || mark.includes(departure))) {
    return true;
  }
  const order = normKey(input.payload.orderNumber || "");
  if (order && departure.includes(order)) return true;
  const dest = normKey(
    stripTimesheetDirectionPrefix(
      input.payload.timesheetDestination ||
        input.payload.destination ||
        input.payload.documentsDest ||
        input.payload.changeText ||
        "",
    ),
  );
  if (dest.length >= 6 && departure.includes(dest.slice(0, 24))) return true;
  const unit = dest.match(/[аa]\s*(\d{4})/i);
  if (unit && departure.includes(unit[1])) return true;
  return false;
};

type ScannedTimesheetRow = {
  row: number;
  personId: string;
  fullName: string;
  positionIndex: string;
  open: boolean;
  departureText: string;
};

const scanTimesheetPersonRows = (
  sheet: ExcelSheetSnapshot,
  grid?: Array<unknown[] | undefined>,
): ScannedTimesheetRow[] => {
  const rows: ScannedTimesheetRow[] = [];
  const last = lastDataRow(sheet, grid);
  for (let row = DATA_START_ROW; row <= last; row += 1) {
    const personId = normId(cellText(sheet, row, 8, grid));
    const fullName = canonicalName(cellText(sheet, row, 7, grid));
    const positionIndex = cellText(sheet, row, 2, grid);
    if (!personId && !fullName) continue;
    const marks: string[] = [];
    for (let day = 1; day <= 31; day += 1) {
      const value = cellText(sheet, row, 8 + day, grid);
      if (isTimesheetDepartureMark(value)) marks.push(value);
    }
    rows.push({
      row,
      personId,
      fullName,
      positionIndex: /^\d{5,}$/.test(positionIndex) ? positionIndex : "",
      open: marks.length === 0,
      departureText: marks[0] || "",
    });
  }
  return rows;
};

const actionOf = (
  kind: TimesheetTransferActionKind,
  extra: Partial<TimesheetTransferAction>,
): TimesheetTransferAction => ({
  kind,
  sourceIndex: extra.sourceIndex || "",
  sourceRow: extra.sourceRow || 0,
  sourceSection: extra.sourceSection || "",
  targetHistoryRow: extra.targetHistoryRow || 0,
  createTimesheetHistory: kind === "CREATE_HISTORY_IN_SOURCE_SECTION",
  replaceInPlace: kind === "PATCH_HISTORY",
});

export const assertTimesheetTransferAction = (
  action: TimesheetTransferAction | null | undefined,
) => {
  if (!action) {
    throw new TimesheetTransferError(
      "EXTERNAL_TRANSFER_WITHOUT_TIMESHEET_ACTION",
      "EXTERNAL_TRANSFER_WITHOUT_TIMESHEET_ACTION",
    );
  }
  if (
    action.sourceRow === 0 &&
    !action.createTimesheetHistory &&
    !action.replaceInPlace
  ) {
    throw new TimesheetTransferError(
      "TIMESHEET_SOURCE_UNRESOLVED",
      "Не визначено рядок Табеля для зовнішнього переведення. Apply заблоковано — Виключені не пишемо.",
    );
  }
  return action;
};

export const timesheetActionPreview = (payload: Record<string, string>) => {
  const inferred =
    payload.timesheetAction ||
    (payload.timesheetReplaceInPlace === "1"
      ? "PATCH_HISTORY"
      : payload.timesheetCreateHistory === "1"
        ? "CREATE_HISTORY_IN_SOURCE_SECTION"
        : payload.timesheetExcelRow
          ? "MOVE_TO_HISTORY"
          : "");
  return {
    sourceIndex:
      payload.timesheetSourceIndex ||
      payload.fromPositionIndex ||
      payload.previousIndex ||
      "",
    sourceRow: Number(payload.timesheetSourceRow || payload.timesheetExcelRow || 0),
    sourceSection: payload.timesheetSourceSection || "",
    action: inferred || "UNRESOLVED",
    targetHistoryRow: Number(payload.timesheetTargetHistoryRow || 0),
  };
};

export const resolveTimesheetTransferAction = (input: {
  personId: string;
  fullName: string;
  fromPositionIndex: string;
  payload: Record<string, string>;
  layout: TimesheetLayout;
  sheet: ExcelSheetSnapshot;
  grid?: Array<unknown[] | undefined>;
}): TimesheetTransferAction => {
  const personId = normId(input.personId);
  const fullName = canonicalName(input.fullName);
  const fromIndex = String(input.fromPositionIndex || "").trim();
  const scanned = scanTimesheetPersonRows(input.sheet, input.grid);
  const openById = personId
    ? scanned.filter((row) => row.open && row.personId === personId)
    : [];
  const openByName = fullName
    ? scanned.filter((row) => row.open && row.fullName === fullName)
    : [];
  const slot = fromIndex ? input.layout.byIndex[fromIndex] ?? null : null;
  const slotScan = slot
    ? scanned.find((row) => row.row === slot.row)
    : undefined;
  const slotOccupiedByPerson =
    Boolean(slot) &&
    slotScan?.open &&
    ((personId && normId(slot.occupantId) === personId) ||
      (fullName && canonicalName(slot.occupantName) === fullName));

  const moveFrom = (row: number, index = "", section = "") =>
    actionOf("MOVE_TO_HISTORY", {
      sourceRow: row,
      sourceIndex:
        index ||
        fromIndex ||
        scanned.find((item) => item.row === row)?.positionIndex ||
        "",
      sourceSection:
        section ||
        input.layout.sectionByRow.get(row)?.label ||
        slot?.section ||
        "",
    });

  if (openById[0]) {
    return moveFrom(openById[0].row);
  }
  if (openByName[0]) {
    return moveFrom(openByName[0].row);
  }
  if (slot && slotOccupiedByPerson) {
    return moveFrom(slot.row, slot.index, slot.section);
  }

  const expectedUnit = String(
    input.payload.positionTitle || input.payload.timesheetSourceSection || "",
  )
    .toLocaleLowerCase("uk-UA")
    .replace(/\s+/g, " ")
    .trim();
  const sectionMatchesExpected = (row: number) => {
    if (!expectedUnit) return true;
    const label = (input.layout.sectionByRow.get(row)?.label || "")
      .toLocaleLowerCase("uk-UA")
      .replace(/\s+/g, " ");
    if (!label) return true;
    const unitCore = expectedUnit
      .replace(/^.*?(?=\d+\s*піхотн)/i, "")
      .trim();
    return (
      label.includes(expectedUnit) ||
      expectedUnit.includes(label) ||
      (unitCore.length > 4 && label.includes(unitCore.slice(0, 18)))
    );
  };

  const closed = scanned.filter((row) => {
    if (row.open) return false;
    const samePerson =
      (personId && row.personId === personId) ||
      (fullName && row.fullName === fullName);
    return samePerson && historyMatchesCurrentTransfer({
      departureText: row.departureText,
      payload: input.payload,
    });
  });
  if (closed[0] && sectionMatchesExpected(closed[0].row)) {
    return actionOf("PATCH_HISTORY", {
      sourceRow: closed[0].row,
      sourceIndex: closed[0].positionIndex || fromIndex,
      sourceSection:
        input.layout.sectionByRow.get(closed[0].row)?.label || slot?.section || "",
    });
  }
  if (closed[0]) {
    return moveFrom(
      closed[0].row,
      closed[0].positionIndex || fromIndex,
      input.layout.sectionByRow.get(closed[0].row)?.label || "",
    );
  }

  if (slot && belongedToUnitThisMonth(input.payload)) {
    return actionOf("CREATE_HISTORY_IN_SOURCE_SECTION", {
      sourceRow: slot.row,
      sourceIndex: slot.index,
      sourceSection: slot.section,
    });
  }

  throw new TimesheetTransferError(
    "TIMESHEET_SOURCE_UNRESOLVED",
    "Не визначено секцію Табеля для зовнішнього переведення. Apply заблоковано — Виключені не пишемо.",
  );
};
