import type { EjoosSyncOp } from "./ejoosSyncPlan";

const PLACEHOLDER_TIMESHEET_CODE = "(оберіть код)";

export function timesheetOpUsesDerivedPayload(op: EjoosSyncOp) {
  if (op.kind !== "timesheet_day") return false;
  return Boolean(
    op.payload.timesheetAbsenceSpans?.trim() ||
      op.payload.timesheetActiveFrom?.trim() ||
      op.payload.timesheetPreserveHistory === "1" ||
      op.payload.timesheetSkipHistory === "1" ||
      op.payload.type === "PAINT_ARCHIVE",
  );
}

export function timesheetOpNeedsManualCode(op: EjoosSyncOp) {
  if (op.kind !== "timesheet_day") return false;
  if (op.payload.clearStalePerson === "1") return false;
  if (timesheetOpUsesDerivedPayload(op)) return false;

  const code = op.payload.timesheetCode?.trim() || "";
  return !code || code === PLACEHOLDER_TIMESHEET_CODE;
}

export function timesheetOpBlocksApply(op: EjoosSyncOp) {
  if (op.kind !== "timesheet_day") return false;
  if (op.payload.clearStalePerson === "1") return !op.payload.excelRow;
  if (timesheetOpNeedsManualCode(op)) return true;
  return !op.payload.excelRow;
}
