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

export function excludeTransferOpBlocksApply(op: EjoosSyncOp) {
  if (op.kind !== "exclude_transfer") return false;
  return !(
    op.payload.destination?.trim() &&
    op.payload.excludeDate?.trim() &&
    op.payload.orderNumber?.trim() &&
    op.payload.orderDate?.trim()
  );
}

export function ambiguousTransferOpBlocksApply(op: EjoosSyncOp) {
  return (
    op.payload.transferScope === "unclear" ||
    (op.kind === "other_manual" && op.payload.type === "TRANSFER_SCOPE_UNCLEAR")
  );
}

/** Відкритий СЗЧ / розпорядження суперечить новій штатній постановці. */
export function contradictoryStatusOpsBlockApply(ops: EjoosSyncOp[]) {
  const placement = ops.find(
    (op) =>
      op.kind === "position_change" && Boolean(op.payload.openAbsenceExcelRow),
  );
  if (!placement) return false;
  return !ops.some(
    (op) =>
      (op.kind === "absent_close" &&
        op.payload.excelRow === placement.payload.openAbsenceExcelRow) ||
      (op.kind === "absent_upsert" &&
        Boolean(op.payload.returnDate) &&
        op.payload.existingExcelRow === placement.payload.openAbsenceExcelRow),
  );
}

export function personOpsBlockApply(ops: EjoosSyncOp[]) {
  if (ops.some((op) => op.class === "conflict")) return true;
  if (ops.some(excludeTransferOpBlocksApply)) return true;
  if (ops.some(ambiguousTransferOpBlocksApply)) return true;
  if (ops.some(timesheetOpBlocksApply)) return true;
  if (contradictoryStatusOpsBlockApply(ops)) return true;
  if (
    ops.some(
      (op) => op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT",
    )
  ) {
    return true;
  }
  return false;
}
