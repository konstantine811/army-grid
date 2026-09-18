import {
  excludeTransferDestination,
  excludedRowsToClear,
} from "./ejoosExcludePolicy";
import type { EjoosSyncOp } from "./ejoosSyncPlan";

const PLACEHOLDER_TIMESHEET_CODE = "(оберіть код)";

/** Лише перегляд: не пишемо в книгу і не блокуємо ПЕРЕВ / ПОСАДА. */
export const isReviewOnlyMismatchOp = (op: EjoosSyncOp) =>
  op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT" ||
  op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING";

const applyCandidateOps = (ops: EjoosSyncOp[]) =>
  ops.filter(
    (op) =>
      op.class !== "conflict" &&
      !isReviewOnlyMismatchOp(op) &&
      op.kind !== "data_mismatch",
  );

const personKey = (op: EjoosSyncOp) =>
  String(op.personId || op.fullName || "")
    .trim()
    .toLocaleLowerCase("uk-UA");

const isIntegratedPersonMovement = (op: EjoosSyncOp) =>
  op.kind === "exclude_transfer" || op.kind === "move_to_disposition";

/**
 * Як у apply: комплексний ПЕРЕВ / РОЗПОРЯДЖ уже пише всі аркуші особи.
 * Archive / Табель тієї ж людини не мають блокувати кнопку.
 */
export function opsConsideredForWorkbookApply(ops: EjoosSyncOp[]) {
  const integratedPersons = new Set(
    ops.filter(isIntegratedPersonMovement).map(personKey).filter(Boolean),
  );
  return ops.filter(
    (op) =>
      isIntegratedPersonMovement(op) ||
      op.kind === "rank_change" ||
      isReturnThenDispositionPlacement(op, ops) ||
      !integratedPersons.has(personKey(op)),
  );
}

/** ПОСАДА з розпорядження + РОЗПОРЯДЖ у тій же картці (СЗЧ лишається відкритим). */
export const isReturnThenDispositionPlacement = (
  op: EjoosSyncOp,
  ops: EjoosSyncOp[],
) => {
  if (
    op.kind !== "position_change" ||
    op.payload.returningFromDisposition !== "1"
  ) {
    return false;
  }
  const key = personKey(op);
  return (
    Boolean(key) &&
    ops.some(
      (other) =>
        other.kind === "move_to_disposition" && personKey(other) === key,
    )
  );
};

export const hasReturnThenDispositionChain = (ops: EjoosSyncOp[]) =>
  ops.some((op) => isReturnThenDispositionPlacement(op, ops));

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

const timesheetOpClearsExcludedOnly = (op: EjoosSyncOp) =>
  excludedRowsToClear(op.payload).length > 0 &&
  !op.payload.timesheetCode?.trim() &&
  !timesheetOpUsesDerivedPayload(op);

export function timesheetOpNeedsManualCode(op: EjoosSyncOp) {
  if (op.kind !== "timesheet_day") return false;
  if (op.payload.clearStalePerson === "1") return false;
  if (timesheetOpClearsExcludedOnly(op)) return false;
  if (timesheetOpUsesDerivedPayload(op)) return false;

  const code = op.payload.timesheetCode?.trim() || "";
  return !code || code === PLACEHOLDER_TIMESHEET_CODE;
}

export function timesheetOpBlocksApply(op: EjoosSyncOp) {
  if (op.kind !== "timesheet_day") return false;
  if (op.payload.clearStalePerson === "1") return !op.payload.excelRow;
  if (timesheetOpClearsExcludedOnly(op)) return false;
  if (timesheetOpNeedsManualCode(op)) return true;
  return !op.payload.excelRow;
}

export function excludeTransferOpBlocksApply(op: EjoosSyncOp) {
  if (op.kind !== "exclude_transfer") return false;
  return !(
    excludeTransferDestination(op.payload) &&
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
  // Повернення з СЗЧ/розпорядження: відсутність може лишатися відкритою.
  if (placement.payload.returningFromDisposition === "1") {
    return false;
  }
  if (isReturnThenDispositionPlacement(placement, ops)) {
    return false;
  }
  return !ops.some(
    (op) =>
      (op.kind === "absent_close" &&
        op.payload.excelRow === placement.payload.openAbsenceExcelRow) ||
      (op.kind === "absent_upsert" &&
        Boolean(op.payload.returnDate) &&
        op.payload.existingExcelRow === placement.payload.openAbsenceExcelRow),
  );
}

export function personApplyBlockReason(ops: EjoosSyncOp[]): string | null {
  const considered = opsConsideredForWorkbookApply(ops);
  if (considered.some((op) => op.class === "conflict")) {
    return "Конфлікт — спочатку розберіть вручну.";
  }
  const candidates = applyCandidateOps(considered);
  if (candidates.some(excludeTransferOpBlocksApply)) {
    const blocked = candidates.find(excludeTransferOpBlocksApply);
    if (!excludeTransferDestination(blocked?.payload ?? {})) {
      return "Заповніть «куди вибув» перед застосуванням переведення.";
    }
    return "Вкажіть дату виключення, номер і дату стройового наказу.";
  }
  if (candidates.some(ambiguousTransferOpBlocksApply)) {
    return "Не визначено, внутрішнє чи зовнішнє переведення — уточніть «куди вибув».";
  }
  if (candidates.some(timesheetOpBlocksApply)) {
    return "Оберіть код для Табеля або вкажіть рядок.";
  }
  if (contradictoryStatusOpsBlockApply(candidates)) {
    return "Спочатку закрийте відкритий СЗЧ / тимчасову відсутність, потім ставте на штат.";
  }
  return null;
}

export function personOpsBlockApply(ops: EjoosSyncOp[]) {
  return personApplyBlockReason(ops) != null;
}
