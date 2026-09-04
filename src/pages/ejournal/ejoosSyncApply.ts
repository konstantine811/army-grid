import JSZip from "jszip";
import { type ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { readEjoosWorkbookSnapshot } from "./ejoosTimesheetPersonRows";
import type { EjoosSyncOp, EjoosSyncPlan } from "./ejoosSyncPlan";
import {
  buildProtocolText,
  findEjoosSheet,
  planBlocksWorkbookApply,
  workbookApplyBlockMessage,
} from "./ejoosSyncPlan";
import {
  formatExcludedDestination,
  formatExcludedListBasis,
  formatExcludedPositionDates,
  formatTimesheetTransferMark,
  OOS_TO_EXCLUDED_BASE,
} from "./ejoosExcludedColumns";
import {
  applyExcludeTransfersWithZip,
  applyExcludedPositionDatesPresentation,
  copyTimesheetRowStylesWithZip,
} from "./ejoosExcludeTransferZip";
import { applyDispositionWithZip } from "./ejoosDispositionZip";
import { applyPositionChangeWithZip } from "./ejoosPositionChangeZip";
import {
  applyOosHistoryPresentation,
  applyRankLabelsWithZip,
} from "./ejoosOosZip";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";
import {
  repairBrokenCellOpenTags,
  stripInvalidNumericCellValues,
} from "./ejoosWorkbookSanitize";
import {
  clipAbsenceSpansToActiveEpisode,
  dayFromOrderLabel,
  findTimesheetMonthHeaderCell,
  formatTimesheetMonthHeader,
  historyAbsenceSpansForClosedEpisode,
  isTimesheetAbsenceCode,
  isTimesheetDepartureMark,
  parseTimesheetAbsenceSpans,
  replaceTimesheetMonthHeaderText,
  timesheetCodeOnDay,
  timesheetHorizonFillDays,
  timesheetMarkFromArchive,
  timesheetTransferMarkForDay,
  type TimesheetAbsenceSpan,
} from "./ejoosTimesheetText";
import {
  excludeWritePlan,
  excludedRowsToClear,
  positionCloseWritesExcluded,
} from "./ejoosExcludePolicy";
import { findTimesheetPersonRowsInGrid } from "./ejoosTimesheetPersonRows";
import { personChangesFromOps, isWorkbookApplyOp } from "./ejoosPersonDiff";
import {
  excludeTransferOpBlocksApply,
  personOpsBlockApply,
} from "./ejoosOpRequirements";
import {
  serializeAppliedPerson,
  serializeSyncOp,
} from "./ejoosAppliedHistory";
import { mergeOosHistoryValue } from "./ejoosOosText";

const col = (letter: string) => {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index; // 1-based
};

const isSinglePersonMovement = (op: EjoosSyncOp) =>
  op.kind === "exclude_transfer" || op.kind === "move_to_disposition";

const isRankBeforeExclude = (op: EjoosSyncOp) => op.kind === "rank_change";

const applyKindOrder = (op: EjoosSyncOp) => {
  if (op.kind === "rank_change") return 0;
  if (op.kind === "absent_close") return 1;
  if (op.kind === "absent_upsert") return 2;
  if (op.kind === "move_to_disposition") return 3;
  if (op.kind === "exclude_transfer") return 4;
  if (op.kind === "position_change") return 5;
  if (
    op.kind === "other_manual" &&
    op.payload.type === "TRANSFER_CANCELLED"
  ) {
    return 7;
  }
  return 6;
};

const transferCancelUndoBlocksTimesheet = (
  ops: EjoosSyncOp[],
  op: EjoosSyncOp,
) =>
  ops.some(
    (other) =>
      other.payload.type === "TRANSFER_CANCELLED" &&
      other.payload.restoreTimesheet === "1" &&
      other.payload.reviewReason !== "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH" &&
      (other.personId === op.personId ||
        other.fullName.toLocaleLowerCase("uk-UA") ===
          op.fullName.toLocaleLowerCase("uk-UA")),
  );

const closesOldPosition = (op: EjoosSyncOp) =>
  op.kind === "position_change" && op.payload.closeOldPosition === "1";

const personKeyOf = (op: EjoosSyncOp) =>
  String(op.personId || op.fullName || "")
    .trim()
    .toLocaleLowerCase("uk-UA");

const XLSX_MIME =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

/** Після пачки треба читати вже записаний .xlsx, а не лишати старі rawRows. */
const rereadWorkbookAfterBatch = async (
  blob: Blob,
  fileName: string,
): Promise<ExcelWorkbookSnapshot> =>
  readEjoosWorkbookSnapshot(
    new File([blob], fileName, {
      type: blob.type || XLSX_MIME,
    }),
  );

/**
 * Змішана черга (виключення + посади інших людей) не йде одним rewrite.
 * Ділимо на безпечні ZIP-проходи: спочатку розпорядження (звільнити штат),
 * потім виключення, потім прості постановки. Інакше новий occupant
 * запишеться на індекс, а РОЗПОРЯДЖ потім зробить clear().
 */
const partitionSafeApplyBatches = (ops: EjoosSyncOp[]): EjoosSyncOp[][] => {
  const excludePeople = new Set(
    ops
      .filter((op) => op.kind === "exclude_transfer")
      .map(personKeyOf)
      .filter(Boolean),
  );
  const dispositionPeople = new Set(
    ops
      .filter((op) => op.kind === "move_to_disposition")
      .map(personKeyOf)
      .filter(Boolean),
  );
  const simple: EjoosSyncOp[] = [];
  const excludes: EjoosSyncOp[] = [];
  const dispositions: EjoosSyncOp[] = [];
  for (const op of ops) {
    const key = personKeyOf(op);
    if (key && excludePeople.has(key)) {
      if (op.kind === "exclude_transfer" || op.kind === "rank_change") {
        excludes.push(op);
      }
      continue;
    }
    if (key && dispositionPeople.has(key)) {
      if (
        op.kind === "move_to_disposition" ||
        op.kind === "rank_change" ||
        op.kind === "absent_upsert" ||
        op.kind === "absent_close"
      ) {
        dispositions.push(op);
      }
      continue;
    }
    simple.push(op);
  }
  return [dispositions, excludes, simple].filter((batch) => batch.length > 0);
};

async function applyUniformBatchToBlob(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
}): Promise<{ blob: Blob; directXml: boolean }> {
  const { ejoos, plan, ops: appliedOps } = input;
  const personKey = personKeyOf;
  const integratedPersons = new Set(
    appliedOps.filter(isSinglePersonMovement).map(personKey).filter(Boolean),
  );
  const onlyExcludeAndRank = appliedOps.every(
    (op) => op.kind === "exclude_transfer" || op.kind === "rank_change",
  );
  const onlyDispositionFamily = appliedOps.every(
    (op) =>
      op.kind === "move_to_disposition" ||
      op.kind === "rank_change" ||
      op.kind === "absent_upsert" ||
      op.kind === "absent_close",
  );
  const onePerson =
    new Set(appliedOps.map(personKey).filter(Boolean)).size <= 1;
  const uniformIntegratedKind =
    appliedOps.every((op) => op.kind === "exclude_transfer") ||
    appliedOps.every((op) => op.kind === "move_to_disposition") ||
    (integratedPersons.size > 0 && onlyExcludeAndRank) ||
    onlyDispositionFamily;
  if (integratedPersons.size > 0 && !uniformIntegratedKind && !onePerson) {
    throw new Error(
      "Комплексний рух (виключення або розпорядження) застосовуйте окремо від інших змін. Це захищає ЕЖООС від повного перезапису та пошкодження аркушів.",
    );
  }

  const excludeOps = appliedOps.filter((op) => op.kind === "exclude_transfer");
  const dispositionOps = appliedOps.filter(
    (op) => op.kind === "move_to_disposition",
  );
  const positionOps = appliedOps.filter(closesOldPosition);
  const useExcludeZip =
    excludeOps.length > 0 &&
    (onePerson || onlyExcludeAndRank || excludeOps.length === appliedOps.length);
  const useDispositionZip =
    !useExcludeZip &&
    dispositionOps.length > 0 &&
    (onePerson ||
      onlyDispositionFamily ||
      dispositionOps.length === appliedOps.length);
  const usePositionZip =
    !useExcludeZip &&
    !useDispositionZip &&
    positionOps.length > 0 &&
    appliedOps.every((op) => closesOldPosition(op) || op.kind === "rank_change");
  const directXml = useExcludeZip || useDispositionZip || usePositionZip;
  const rawBlob = useExcludeZip
    ? await applyExcludeTransfersWithZip({ ejoos, plan, ops: excludeOps })
    : useDispositionZip
      ? await applyDispositionWithZip({ ejoos, plan, ops: dispositionOps })
      : usePositionZip
        ? await applyPositionChangeWithZip({ ejoos, plan, ops: positionOps })
        : await mutateToBlob(ejoos, appliedOps, plan);
  const blob = await applyTimesheetMonthHeader({
    file: await applyExcludedClearsWithZip({
      file: await applyRankLabelsWithZip({
        file: await applyExcludedPositionDatesPresentation({
          file: await applyOosHistoryPresentation({
            file: rawBlob,
            ejoos,
            ops: appliedOps,
            plan,
          }),
          ejoos,
        }),
        ops: appliedOps,
      }),
      ejoos,
      ops: appliedOps,
    }),
    ejoos,
    plan,
  });
  return { blob, directXml };
}

/**
 * Apply only confirmed ops to a copy of the live ЕЖООС workbook.
 * Does not write anything until the caller uploads the returned blob as a new DB version.
 */
export async function applyConfirmedEjoosOps(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
  actor?: string;
}): Promise<{
  blob: Blob;
  fileName: string;
  protocolText: string;
  changeProtocol: Record<string, unknown>;
  /** Результат зібрано точковим записом XML і вже не потребує санітизації книги. */
  directXml: boolean;
}> {
  const { ejoos, actor } = input;
  const plan = input.plan;
  if (planBlocksWorkbookApply(plan)) {
    throw new Error(workbookApplyBlockMessage(plan));
  }
  // Позначка помилки даних — інформаційна: у книгу нічого не пишемо.
  const ops = input.ops.filter((op) => op.kind !== "data_mismatch");
  if (!input.ops.length) {
    throw new Error("Немає підтверджених змін для застосування");
  }
  if (!ops.length) {
    throw new Error(
      "Обрано лише позначки ПІБ / ID / звання. Їх виправляють у джерелах — у ЕЖООС нічого не пишемо.",
    );
  }
  const integratedPersons = new Set(
    ops.filter(isSinglePersonMovement).map(personKeyOf).filter(Boolean),
  );
  // Комплексний рух уже включає всі потрібні аркуші. Окремі операції
  // цієї ж особи не повинні запускати повторний повний rewrite книги.
  // Звання перед ПЕРЕВ лишаємо: інакше у «Виключені» піде старе звання.
  const consideredOps = ops
    .filter(
      (op) =>
        isSinglePersonMovement(op) ||
        isRankBeforeExclude(op) ||
        !integratedPersons.has(personKeyOf(op)),
    )
    .sort((left, right) => applyKindOrder(left) - applyKindOrder(right));
  const appliedOps = consideredOps.filter(isWorkbookApplyOp);
  if (!appliedOps.length) {
    const unsupported = consideredOps[0] || ops[0];
    throw new Error(
      `Операція «${unsupported.kind}» для ${unsupported.fullName || unsupported.personId} не має apply — підтвердження не записує ЕЖООС.`,
    );
  }

  const incompleteTransfer = consideredOps.find(excludeTransferOpBlocksApply);
  if (incompleteTransfer) {
    throw new Error(
      `Для ${incompleteTransfer.fullName || incompleteTransfer.personId} вкажіть куди вибув, дату виключення, номер і дату стройового наказу.`,
    );
  }
  const unclearTransfer = consideredOps.find(
    (op) =>
      op.payload.transferScope === "unclear" ||
      op.payload.type === "TRANSFER_SCOPE_UNCLEAR",
  );
  if (unclearTransfer) {
    throw new Error(
      `Для ${unclearTransfer.fullName || unclearTransfer.personId} не визначено, внутрішнє чи зовнішнє переведення. Уточніть «куди вибув» або в/ч.`,
    );
  }
  if (personOpsBlockApply(consideredOps)) {
    const contradiction = appliedOps.find(
      (op) =>
        op.kind === "position_change" && op.payload.openAbsenceExcelRow,
    );
    if (
      contradiction &&
      !appliedOps.some(
        (op) =>
          op.kind === "absent_close" ||
          (op.kind === "absent_upsert" && op.payload.returnDate),
      )
    ) {
      throw new Error(
        `Для ${contradiction.fullName || contradiction.personId} спочатку закрийте відкритий СЗЧ / тимчасову відсутність, потім ставте на штат.`,
      );
    }
    throw new Error(
      "Одна або кілька операцій не мають усіх даних для безпечного застосування",
    );
  }

  const fileName = `ЄЖООС_станом_на_${plan.timesheetDayLabel.replaceAll(".", "-")}.xlsx`;
  const batches = partitionSafeApplyBatches(appliedOps);
  let working = ejoos;
  let blob: Blob = ejoos.file;
  let directXml = true;
  for (const batch of batches) {
    const applied = await applyUniformBatchToBlob({
      ejoos: working,
      plan,
      ops: batch,
    });
    blob = applied.blob;
    directXml = directXml && applied.directXml;
    working = await rereadWorkbookAfterBatch(blob, ejoos.fileName);
  }
  const protocolText = buildProtocolText(plan, appliedOps, {
    actor,
    at: new Date().toLocaleString("uk-UA"),
  });
  const people = personChangesFromOps(appliedOps, plan.timesheetDay, {
    decision: "accepted",
  }).map(serializeAppliedPerson);
  const changeProtocol = {
    kind: "apply",
    pbFileName: plan.pbName,
    timesheetDay: plan.timesheetDay,
    timesheetDayLabel: plan.timesheetDayLabel,
    appliedCount: appliedOps.length,
    ops: appliedOps.map(serializeSyncOp),
    people,
    protocolText,
  };
  return { blob, fileName, protocolText, changeProtocol, directXml };
}

async function mutateToBlob(
  ejoos: ExcelWorkbookSnapshot,
  ops: EjoosSyncOp[],
  plan: EjoosSyncPlan,
) {
  const module = await import("xlsx-populate/browser/xlsx-populate-no-encryption");
  const XlsxPopulate = module.default;
  const workbook = (await XlsxPopulate.fromDataAsync(
    ejoos.file,
  )) as unknown as WorkbookLike;

  const timesheet = workbook.sheet("6. Табель");
  const arrivals = workbook.sheet("4. Тимчасово прибулі");
  const absent = workbook.sheet("5. Тимчасово відсутні");
  const excluded = workbook.sheet("3. Виключені");
  const oos = workbook.sheet("2. ООС");
  const shpo = workbook.sheet("1. ШПО");

  const sortedOps = [...ops].sort(
    (left, right) => applyKindOrder(left) - applyKindOrder(right),
  );
  const timesheetStyleCopies: Array<{ sourceRow: number; targetRow: number }> =
    [];

  sortedOps.forEach((op) => {
    if (op.kind === "shpo_occupant") {
      const shpoRow = Number(op.payload.shpoExcelRow || 0);
      let tsRow = Number(op.payload.timesheetExcelRow || 0);
      const name = op.payload.nextName || op.fullName;
      const rank = op.payload.nextRank || op.rank;
      const personId = op.payload.nextPersonId || op.personId;
      if (shpo && shpoRow > 0) {
        if (rank) shpo.cell(shpoRow, col("F")).value(rank);
        if (name) shpo.cell(shpoRow, col("G")).value(name);
        if (personId) shpo.cell(shpoRow, col("H")).value(personId);
      }
      if (timesheet && tsRow <= 0 && op.positionIndex) {
        tsRow = findPersonRowFlexible(timesheet, {
          personId: "",
          fullName: "",
          positionIndex: op.positionIndex,
          startRow: 7,
          nameCol: col("G"),
          idCol: col("H"),
          indexCol: col("B"),
        });
      }
      if (timesheet && tsRow > 0) {
        if (rank) timesheet.cell(tsRow, col("F")).value(rank);
        if (name) timesheet.cell(tsRow, col("G")).value(name);
        if (personId) timesheet.cell(tsRow, col("H")).value(personId);
        if (
          !transferCancelUndoBlocksTimesheet(sortedOps, op) &&
          (op.payload.timesheetAbsenceSpans || op.payload.timesheetActiveFrom)
        ) {
          paintTimesheetArchiveDays(timesheet, tsRow, plan, op.payload);
        }
      }
      clearStaleExcludedFromPayload(excluded, op.payload);
      return;
    }
    if (op.kind === "timesheet_day" && timesheet) {
      const rowNumber = Number(op.payload.excelRow || 0);
      if (op.payload.clearStalePerson === "1" && rowNumber > 0) {
        if (op.payload.clearTimesheetIndex === "1") {
          timesheet.cell(rowNumber, col("B")).value(null);
        }
        timesheet.cell(rowNumber, col("F")).value(null);
        timesheet.cell(rowNumber, col("G")).value(null);
        timesheet.cell(rowNumber, col("H")).value(null);
        for (let day = 1; day <= 31; day += 1) {
          timesheet.cell(rowNumber, col("I") + day - 1).value(null);
        }
        timesheet.cell(rowNumber, col("AN")).value(null);
        clearStaleExcludedFromPayload(excluded, op.payload);
        return;
      }
      const day = Number(op.payload.day || plan.timesheetDay);
      const code = op.payload.timesheetCode || String(op.after || "").trim();
      if (op.payload.restorePerson === "1" && rowNumber > 0) {
        const name = op.payload.nextName || op.fullName;
        const personId = op.payload.nextPersonId || op.personId;
        if (name) timesheet.cell(rowNumber, col("G")).value(name);
        if (personId) timesheet.cell(rowNumber, col("H")).value(personId);
        if (op.payload.nextRank) {
          timesheet.cell(rowNumber, col("F")).value(op.payload.nextRank);
        }
      }
      if (op.payload.type === "PAINT_ARCHIVE") {
        if (!transferCancelUndoBlocksTimesheet(sortedOps, op)) {
          if (rowNumber > 0 && op.rank) {
            timesheet.cell(rowNumber, col("F")).value(op.rank);
          }
          paintTimesheetArchiveDays(timesheet, rowNumber, plan, op.payload);
          repairHistoryTimesheetRow(timesheet, op, plan.timesheetDay);
        }
        clearStaleExcludedFromPayload(excluded, op.payload);
        return;
      }
      if (
        rowNumber > 0 &&
        day >= 1 &&
        day <= 31 &&
        code &&
        code !== "(оберіть код)"
      ) {
        const existing: string[] = [];
        for (let scanDay = 1; scanDay <= 31; scanDay += 1) {
          existing[scanDay] = String(
            timesheet.cell(rowNumber, col("I") + scanDay - 1).value() ?? "",
          ).trim();
        }
        const fills = timesheetHorizonFillDays({
          dayCodes: existing,
          horizon: plan.timesheetDay || day,
          reportCode: code,
          confirmedReturn: op.payload.confirmedReturn === "1",
        });
        if (fills.length) {
          for (const fill of fills) {
            timesheet.cell(rowNumber, col("I") + fill.day - 1).value(fill.mark);
          }
        } else {
          const existingMark = existing[day] || "";
          const skipPlusOverAbsence =
            isTimesheetAbsenceCode(existingMark) &&
            code === "+" &&
            op.payload.confirmedReturn !== "1";
          if (!skipPlusOverAbsence) {
            timesheet.cell(rowNumber, col("I") + day - 1).value(code);
          }
        }
      }
      clearStaleExcludedFromPayload(excluded, op.payload);
      return;
    }
    if (op.kind === "absent_close" && absent) {
      const rowNumber = Number(op.payload.excelRow || 0);
      if (rowNumber > 0) {
        writeAbsenceReturn(absent, rowNumber, op.payload, plan.timesheetDayLabel);
      }
      const timesheetRow = Number(op.payload.timesheetExcelRow || 0);
      if (timesheet && timesheetRow > 0) {
        if (!transferCancelUndoBlocksTimesheet(sortedOps, op)) {
          paintTimesheetArchiveDays(timesheet, timesheetRow, plan, {
            ...op.payload,
            timesheetActiveFrom: op.payload.timesheetActiveFrom || "",
            timesheetSkipHistory: op.payload.timesheetSkipHistory || "1",
          });
        }
      }
      return;
    }
    if (op.kind === "absent_upsert" && absent) {
      const existing = Number(op.payload.existingExcelRow || 0);
      const targetRow = existing > 0 ? existing : nextEmptyAbsentRow(absent);
      const absenceRank = op.payload.nextRank || op.rank;
      if (absenceRank) absent.cell(targetRow, col("A")).value(absenceRank);
      if (op.fullName) absent.cell(targetRow, col("B")).value(op.fullName);
      if (op.personId) absent.cell(targetRow, col("C")).value(op.personId);
      if (op.positionIndex) {
        absent.cell(targetRow, col("D")).value(op.positionIndex);
      }
      if (op.payload.absenceType) {
        absent.cell(targetRow, col("E")).value(op.payload.absenceType);
      }
      if (op.payload.place) absent.cell(targetRow, col("F")).value(op.payload.place);
      if (op.payload.departDate) {
        absent.cell(targetRow, col("G")).value(op.payload.departDate);
      }
      if (op.payload.orderDate) {
        absent.cell(targetRow, col("H")).value(op.payload.orderDate);
      }
      if (op.payload.orderNumber) {
        absent.cell(targetRow, col("I")).value(op.payload.orderNumber);
      }
      if (op.payload.duration) {
        absent.cell(targetRow, col("J")).value(op.payload.duration);
      }
      if (op.payload.plannedReturn && op.payload.plannedReturn !== "?") {
        absent.cell(targetRow, col("L")).value(op.payload.plannedReturn);
      }
      if (op.payload.returnDate) {
        writeAbsenceReturn(absent, targetRow, op.payload, "");
      }
      absent
        .cell(targetRow, col("R"))
        .value(`archive №${op.payload.periodNumber || "—"}`);
      const timesheetRow = Number(op.payload.timesheetExcelRow || 0);
      if (timesheet && timesheetRow > 0) {
        if (!transferCancelUndoBlocksTimesheet(sortedOps, op)) {
          paintTimesheetArchiveDays(timesheet, timesheetRow, plan, op.payload);
          repairHistoryTimesheetRow(timesheet, op, plan.timesheetDay);
        }
      }
      return;
    }
    if (op.kind === "rank_change") {
      return;
    }
    if (op.kind === "exclude_transfer") {
      applyExcludeTransfer({
        op,
        plan,
        excluded,
        timesheet,
        shpo,
        oos,
        arrivals,
        timesheetStyleCopies,
      });
      return;
    }
    if (op.kind === "position_change") {
      applyPositionChange({
        op,
        plan,
        shpo,
        oos,
        timesheet,
        arrivals,
        excluded,
        timesheetStyleCopies,
      });
      return;
    }
    if (
      op.kind === "other_manual" &&
      op.payload.type === "TRANSFER_CANCELLED"
    ) {
      applyTransferCancelled({ op, excluded, shpo, oos, timesheet, plan });
      return;
    }
  });
  if (timesheet) {
    ops.forEach((op) => clearDuplicateTimesheetRow(timesheet, op));
  }

  if (shpo) removeShpoNameOnlyRows(shpo);

  const populatedBlob = (await workbook.outputAsync("blob")) as Blob;
  const finalized = await finalizeEjoosWorkbookBlob(
    populatedBlob,
    ejoos.file,
    touchedSheetNamesForOps(ops),
  );
  return copyTimesheetRowStylesWithZip(finalized, timesheetStyleCopies);
}

/** Скасоване переведення: rollback попереднього ПЕРЕВ (не новий рух). */
function applyTransferCancelled(input: {
  op: EjoosSyncOp;
  plan: EjoosSyncPlan;
  excluded: SheetLike | undefined;
  shpo: SheetLike | undefined;
  oos: SheetLike | undefined;
  timesheet: SheetLike | undefined;
}) {
  const { op, plan, excluded, shpo, timesheet } = input;
  if (op.payload.reviewReason === "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH") {
    return;
  }

  const excludedRowNumber = Number(op.payload.excludedExcelRow || 0);
  const personId = op.personId;
  const fullName = op.fullName;
  const rank = op.rank;
  const positionIndex = op.payload.previousIndex || op.positionIndex;

  if (excluded && excludedRowNumber > 0) {
    for (let column = col("A"); column <= col("AF"); column += 1) {
      excluded.cell(excludedRowNumber, column).value(null);
    }
    excluded.cell(excludedRowNumber, col("Z")).value(null);
  }

  const shpoRowNumber = Number(op.payload.shpoExcelRow || 0);
  const cancelNote = [
    op.payload.cancelledTransferOrder
      ? `скасовано ПЕРЕВ №${op.payload.cancelledTransferOrder}`
      : "",
    op.payload.transferCancelOrder
      ? `наказ №${op.payload.transferCancelOrder}`
      : "",
    op.payload.transferCancelDate ? `від ${op.payload.transferCancelDate}` : "",
  ]
    .filter(Boolean)
    .join(" ");
  if (op.payload.restoreShpo === "1" && shpo && shpoRowNumber > 0) {
    if (rank) shpo.cell(shpoRowNumber, col("F")).value(rank);
    if (fullName) shpo.cell(shpoRowNumber, col("G")).value(fullName);
    if (personId) shpo.cell(shpoRowNumber, col("H")).value(personId);
  }
  if (shpo && shpoRowNumber > 0 && cancelNote) {
    appendCellHistory(shpo, shpoRowNumber, col("R"), cancelNote);
  }

  const timesheetRowNumber = Number(op.payload.timesheetExcelRow || 0);
  if (
    op.payload.restoreTimesheet === "1" &&
    timesheet &&
    timesheetRowNumber > 0
  ) {
    undoCancelledTransferOnTimesheet(timesheet, timesheetRowNumber, plan, {
      rank,
      fullName,
      personId,
      positionIndex,
      payload: op.payload,
    });
  }

  const historyRowNumber = Number(op.payload.historyTimesheetExcelRow || 0);
  if (
    timesheet &&
    historyRowNumber > 0 &&
    timesheetRowNumber > 0 &&
    historyRowNumber !== timesheetRowNumber
  ) {
    const restoredName = String(
      timesheet.cell(timesheetRowNumber, col("G")).value() ?? "",
    ).trim();
    const restoredId = String(
      timesheet.cell(timesheetRowNumber, col("H")).value() ?? "",
    ).trim();
    const historyName = String(
      timesheet.cell(historyRowNumber, col("G")).value() ?? "",
    ).trim();
    const historyId = String(
      timesheet.cell(historyRowNumber, col("H")).value() ?? "",
    ).trim();
    const restoredHasPerson = Boolean(restoredName || restoredId);
    const historyIsSamePerson =
      (personId && historyId === personId) ||
      (fullName &&
        historyName.toLocaleLowerCase("uk-UA") ===
          fullName.toLocaleLowerCase("uk-UA"));
    if (restoredHasPerson && historyIsSamePerson) {
      for (let column = col("B"); column <= col("AN"); column += 1) {
        timesheet.cell(historyRowNumber, column).value(null);
      }
    }
  }

  clearDuplicateTimesheetRow(timesheet, op);
}

function undoCancelledTransferOnTimesheet(
  timesheet: SheetLike,
  rowNumber: number,
  plan: EjoosSyncPlan,
  input: {
    rank: string;
    fullName: string;
    personId: string;
    positionIndex: string;
    payload: Record<string, string>;
  },
) {
  const { rank, fullName, personId, payload } = input;
  if (rank) timesheet.cell(rowNumber, col("F")).value(rank);
  if (fullName) timesheet.cell(rowNumber, col("G")).value(fullName);
  if (personId) timesheet.cell(rowNumber, col("H")).value(personId);
  if (input.positionIndex) {
    timesheet.cell(rowNumber, col("B")).value(input.positionIndex);
  }

  const departDay = dayFromOrderLabel(payload.cancelledTransferDate);
  const cancelDay =
    dayFromOrderLabel(payload.transferCancelDate) || plan.timesheetDay;
  const activeFromDay = dayFromOrderLabel(payload.timesheetActiveFrom) || 1;
  const keepGap = payload.timesheetCancelKeepGap === "1";
  const spans = parseTimesheetAbsenceSpans(payload.timesheetAbsenceSpans || "");

  let presentDays = 0;
  for (let day = 1; day <= 31; day += 1) {
    const cell = timesheet.cell(rowNumber, col("I") + day - 1);
    const value = cell.value();
    const text = String(value ?? "").trim();
    const isDepart = isTimesheetDepartureMark(value);
    const inDepartGap =
      keepGap &&
      departDay > 0 &&
      day >= departDay &&
      day < cancelDay;
    const archiveCode = timesheetCodeOnDay(day, spans);

    if (day > plan.timesheetDay) {
      if (
        isDepart ||
        (departDay > 0 && day >= departDay && (text === "-" || isDepart))
      ) {
        cell.value(null);
      }
      continue;
    }

    if (day < activeFromDay) {
      cell.value("-");
      continue;
    }

    if (inDepartGap) {
      cell.value("-");
      continue;
    }

    if (archiveCode) {
      cell.value(archiveCode);
      if (archiveCode === "+") presentDays += 1;
      continue;
    }

    const staleDepart =
      isDepart ||
      (departDay > 0 && day >= departDay && (text === "-" || isDepart));
    const staleCancelNote = /скасован/i.test(text);
    if (staleDepart || staleCancelNote || !text || text === "-") {
      cell.value("+");
      presentDays += 1;
      continue;
    }

    if (text === "+") {
      presentDays += 1;
    }
  }
  timesheet.cell(rowNumber, col("AN")).value(presentDays);
}

function findArrivalCloseColumns(arrivals: SheetLike) {
  let departDateCol = 0;
  let orderDateCol = 0;
  let orderNumberCol = 0;
  let orderCombinedCol = 0;
  for (let row = 1; row <= 6; row += 1) {
    for (let column = 1; column <= 40; column += 1) {
      const key = String(arrivals.cell(row, column).value() ?? "")
        .toLocaleLowerCase("uk-UA")
        .replace(/\s+/g, " ")
        .trim();
      if (!key) continue;
      if (/дата вибут/.test(key) && !/наказ/.test(key)) {
        departDateCol = column;
      } else if (/дата.{0,12}наказ.{0,12}вибут|наказ.{0,12}вибут.{0,12}дата/.test(key)) {
        orderDateCol = column;
      } else if (/номер.{0,12}наказ.{0,12}вибут|наказ.{0,12}вибут.{0,12}номер/.test(key)) {
        orderNumberCol = column;
      } else if (/наказ.{0,16}вибут/.test(key) && !orderCombinedCol) {
        orderCombinedCol = column;
      }
    }
  }
  return { departDateCol, orderDateCol, orderNumberCol, orderCombinedCol };
}

/** ПОСАДА: постановка на штатну позицію; для тимчасово прибулого — перехід у штат. */
function applyPositionChange(input: {
  op: EjoosSyncOp;
  plan: EjoosSyncPlan;
  shpo: SheetLike | undefined;
  oos: SheetLike | undefined;
  timesheet: SheetLike | undefined;
  arrivals: SheetLike | undefined;
  excluded: SheetLike | undefined;
  timesheetStyleCopies: TimesheetStyleCopy[];
}) {
  const {
    op,
    plan,
    shpo,
    oos,
    timesheet,
    arrivals,
    excluded,
    timesheetStyleCopies,
  } = input;
  if (op.payload.closeOldPosition === "1") {
    closeOldPositionRows({
      op,
      plan,
      shpo,
      oos,
      timesheet,
      excluded,
      timesheetStyleCopies,
    });
  }
  const personId = op.payload.nextPersonId || op.personId;
  const fullName = op.payload.nextName || op.fullName;
  const rank = op.payload.nextRank || op.rank;
  const nextIndex = op.payload.nextIndex || op.positionIndex;
  const appointmentDate = op.payload.orderDate || plan.timesheetDayLabel;
  const appointmentDay = dayFromOrderLabel(appointmentDate) || plan.timesheetDay;
  const preserveHistory = op.payload.timesheetPreserveHistory === "1";
  const activeFromDay =
    dayFromOrderLabel(op.payload.timesheetActiveFrom || op.payload.transferCancelDate) ||
    appointmentDay;

  let shpoRow = Number(op.payload.shpoExcelRow || 0);
  if (shpo && shpoRow <= 0) {
    shpoRow = findPersonRowFlexible(shpo, {
      personId: "",
      fullName: "",
      positionIndex: nextIndex,
      startRow: 7,
      nameCol: col("G"),
      idCol: col("H"),
      indexCol: col("A"),
    });
  }
  const dispositionShpoRow = Number(op.payload.dispositionShpoExcelRow || 0);
  if (shpo && dispositionShpoRow > 0) {
    shpo.cell(dispositionShpoRow, col("B")).value(null);
    shpo.cell(dispositionShpoRow, col("C")).value(null);
    shpo.cell(dispositionShpoRow, col("F")).value(null);
    shpo.cell(dispositionShpoRow, col("G")).value(null);
    shpo.cell(dispositionShpoRow, col("H")).value(null);
  }
  const dispositionTimesheetRow = Number(
    op.payload.dispositionTimesheetExcelRow || 0,
  );
  if (timesheet && dispositionTimesheetRow > 0) {
    clearTimesheetOccupantRow(timesheet, dispositionTimesheetRow);
  }
  if (shpo && shpoRow > 0) {
    shpo.cell(shpoRow, col("F")).value(rank || null);
    shpo.cell(shpoRow, col("G")).value(fullName || null);
    shpo.cell(shpoRow, col("H")).value(personId || null);
    if (op.payload.orderNumber || appointmentDate) {
      shpo
        .cell(shpoRow, col("R"))
        .value(
          [
            "ПОСАДА",
            op.payload.orderNumber ? `№${op.payload.orderNumber}` : "",
            appointmentDate ? `від ${appointmentDate}` : "",
          ]
            .filter(Boolean)
            .join(" "),
        );
    }
  }

  let timesheetRow = Number(op.payload.timesheetExcelRow || 0);
  if (timesheet && timesheetRow <= 0) {
    timesheetRow = findPersonRowFlexible(timesheet, {
      personId: "",
      fullName: "",
      positionIndex: nextIndex,
      startRow: 7,
      nameCol: col("G"),
      idCol: col("H"),
      indexCol: col("B"),
    });
  }
  if (timesheet && timesheetRow > 0 && op.payload.timesheetHandledByCancel !== "1") {
    const previousName = String(
      timesheet.cell(timesheetRow, col("G")).value() ?? "",
    ).trim();
    const previousId = String(
      timesheet.cell(timesheetRow, col("H")).value() ?? "",
    ).trim();
    const samePersonOnStaff =
      Boolean(previousName || previousId) &&
      (previousName.toLocaleLowerCase("uk-UA") ===
        fullName.toLocaleLowerCase("uk-UA") ||
        (Boolean(personId) && previousId === personId));
    const otherPerson =
      Boolean(previousName || previousId) && !samePersonOnStaff;
    const existingHistoryRow = Number(
      op.payload.historyTimesheetExcelRow ||
        op.payload.previousTimesheetExcelRow ||
        0,
    );
    const historyAlreadySeparate =
      existingHistoryRow > 0 && existingHistoryRow !== timesheetRow;
    if (
      samePersonOnStaff &&
      timesheetRowHasDeparture(timesheet, timesheetRow) &&
      !historyAlreadySeparate &&
      !op.payload.transferCancelOrder
    ) {
      const historyRow = findTimesheetAppendNear(timesheet, timesheetRow);
      copyTimesheetHistoryValues(
        timesheet,
        timesheetRow,
        historyRow,
        timesheetStyleCopies,
      );
    } else if (otherPerson) {
      const historyRow = findTimesheetAppendNear(timesheet, timesheetRow);
      copyTimesheetHistoryValues(
        timesheet,
        timesheetRow,
        historyRow,
        timesheetStyleCopies,
      );
      for (let day = appointmentDay; day <= 31; day += 1) {
        timesheet.cell(historyRow, col("I") + day - 1).value(null);
      }
    }
    if (!previousName && !previousId) {
      recordTimesheetStyleCopy(
        timesheetStyleCopies,
        occupiedTimesheetStyleRow(timesheet, timesheetRow),
        timesheetRow,
      );
    }
    timesheet.cell(timesheetRow, col("F")).value(rank || null);
    timesheet.cell(timesheetRow, col("G")).value(fullName || null);
    timesheet.cell(timesheetRow, col("H")).value(personId || null);
    const bindIndex = op.payload.timesheetBindStaffIndex || nextIndex;
    if (bindIndex && op.payload.isTempArrivalPlacement === "1") {
      timesheet.cell(timesheetRow, col("B")).value(bindIndex);
    }
    paintTimesheetArchiveDays(timesheet, timesheetRow, plan, {
      ...op.payload,
      timesheetActiveFrom:
        op.payload.timesheetActiveFrom ||
        (preserveHistory ? String(activeFromDay) : String(appointmentDay)),
      timesheetPreserveHistory: preserveHistory ? "1" : op.payload.timesheetPreserveHistory,
    });
  }

  // ООС не пишемо через xlsx-populate — лише applyOosHistoryPresentation.

  // Історичний рядок не видаляємо: позначаємо завершення тимчасового статусу.
  const arrivalRow = Number(op.payload.arrivalExcelRow || 0);
  if (arrivals && arrivalRow > 0) {
    const departDate =
      op.payload.arrivalDepartDate || appointmentDate;
    const orderNumber = op.payload.arrivalDepartOrderNumber || op.payload.orderNumber;
    const orderDate = op.payload.arrivalDepartOrderDate || appointmentDate;
    const closeColumns = findArrivalCloseColumns(arrivals);
    if (closeColumns.departDateCol) {
      arrivals.cell(arrivalRow, closeColumns.departDateCol).value(departDate || null);
    }
    if (closeColumns.orderDateCol) {
      arrivals.cell(arrivalRow, closeColumns.orderDateCol).value(orderDate || null);
    }
    if (closeColumns.orderNumberCol) {
      arrivals
        .cell(arrivalRow, closeColumns.orderNumberCol)
        .value(orderNumber ? `№${orderNumber}` : null);
    }
    if (closeColumns.orderCombinedCol) {
      arrivals.cell(arrivalRow, closeColumns.orderCombinedCol).value(
        [orderNumber ? `№${orderNumber}` : "", orderDate ? `від ${orderDate}` : ""]
          .filter(Boolean)
          .join(" "),
      );
    }
    appendCellHistory(
      arrivals,
      arrivalRow,
      col("R"),
      [
        `ЗАРАХОВАНО ДО ШТАТУ ${nextIndex}`,
        appointmentDate,
        op.payload.orderNumber ? `наказ №${op.payload.orderNumber}` : "",
      ]
        .filter(Boolean)
        .join(" · "),
    );
  }

  clearStaleExcludedFromPayload(excluded, op.payload);
  const staleTimesheetRow = Number(op.payload.clearTimesheetExcelRow || 0);
  if (
    timesheet &&
    staleTimesheetRow > 0 &&
    staleTimesheetRow !== Number(op.payload.timesheetExcelRow || 0)
  ) {
    clearTimesheetOccupantRow(timesheet, staleTimesheetRow);
  }
}

const EXCLUDED_CLEAR_LAST_COLUMN = 32;

const clearTimesheetOccupantRow = (
  timesheet: SheetLike,
  rowNumber: number,
) => {
  if (rowNumber <= 0) return;
  timesheet.cell(rowNumber, col("F")).value(null);
  timesheet.cell(rowNumber, col("G")).value(null);
  timesheet.cell(rowNumber, col("H")).value(null);
  for (let day = 1; day <= 31; day += 1) {
    timesheet.cell(rowNumber, col("I") + day - 1).value(null);
  }
  timesheet.cell(rowNumber, col("AN")).value(null);
};

const clearStaleExcludedRow = (
  excluded: SheetLike | undefined,
  rowRaw: string | undefined,
) => {
  const rowNumber = Number(rowRaw || 0);
  if (!excluded || rowNumber <= 0) return;
  for (let column = col("A"); column <= EXCLUDED_CLEAR_LAST_COLUMN; column += 1) {
    excluded.cell(rowNumber, column).value(null);
  }
  excluded.cell(rowNumber, col("Z")).value(null);
};

const clearStaleExcludedFromPayload = (
  excluded: SheetLike | undefined,
  payload: { clearExcludedExcelRow?: string; clearExcludedExcelRows?: string },
) => {
  for (const row of excludedRowsToClear(payload)) {
    clearStaleExcludedRow(excluded, String(row));
  }
};

async function applyTimesheetMonthHeader(input: {
  file: Blob;
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
}): Promise<Blob> {
  const timesheet = findEjoosSheet(input.ejoos, /табель/i);
  if (!timesheet) return input.file;
  if (input.plan.monthRolloverRequired) return input.file;
  const expected = formatTimesheetMonthHeader(input.plan.timesheetDayLabel);
  if (!expected) return input.file;
  const found = findTimesheetMonthHeaderCell(timesheet.rawRows);
  if (
    found &&
    found.matched.replace(/\s+/g, " ").trim().toLocaleLowerCase("uk-UA") ===
      expected.toLocaleLowerCase("uk-UA")
  ) {
    return input.file;
  }
  const row = found?.row || 2;
  const column = found?.column || 9;
  const current = found
    ? String(timesheet.rawRows[found.row - 1]?.[found.column - 1] ?? "")
    : "";
  const value = replaceTimesheetMonthHeaderText(current, expected);
  return applyInlineStringWritesToWorkbook(input.file, timesheet.sheetName, [
    { row, column, value },
  ]);
}

async function applyExcludedClearsWithZip(input: {
  file: Blob;
  ejoos: ExcelWorkbookSnapshot;
  ops: EjoosSyncOp[];
}): Promise<Blob> {
  const excluded = findEjoosSheet(input.ejoos, /виключен/i);
  if (!excluded) return input.file;
  const rows = new Set<number>();
  for (const op of input.ops) {
    for (const row of excludedRowsToClear(op.payload)) rows.add(row);
  }
  if (!rows.size) return input.file;
  const writes: ZipCellWrite[] = [];
  for (const row of rows) {
    for (let column = 1; column <= EXCLUDED_CLEAR_LAST_COLUMN; column += 1) {
      writes.push({
        row,
        column,
        value: null,
        styleSourceRow: 6,
        copyNeighborStyle: true,
      });
    }
  }
  return applyInlineStringWritesToWorkbook(
    input.file,
    excluded.sheetName,
    writes,
  );
}

/**
 * Зміна посади всередині 1ПБ закриває стару штатну посаду в ШПО/Табелі.
 * Рядок у «3. Виключені» — лише якщо це не внутрішній стрибок індексу.
 * ООС не чіпаємо — особа лишається в частині.
 */
function closeOldPositionRows(input: {
  op: EjoosSyncOp;
  plan: EjoosSyncPlan;
  shpo: SheetLike | undefined;
  oos: SheetLike | undefined;
  timesheet: SheetLike | undefined;
  excluded: SheetLike | undefined;
  timesheetStyleCopies: TimesheetStyleCopy[];
}) {
  const { op, shpo, oos, timesheet, excluded, timesheetStyleCopies } = input;
  const personId = op.payload.fromPersonId || op.personId;
  const fullName = op.payload.fromName || op.fullName;
  const rank = op.payload.fromRank || op.rank;
  const previousIndex =
    op.payload.fromPositionIndex || op.payload.previousIndex;
  const excludeDateLabel = op.payload.excludeDate || op.payload.orderDate;
  const departDay = dayFromOrderLabel(excludeDateLabel);

  if (excluded && positionCloseWritesExcluded(op.payload)) {
    const targetRow = nextExcludedManualRow(excluded, 6);
    copyRowHeight(excluded, 6, targetRow);
    const oosRow =
      oos && personId ? findPersonIdRow(oos, personId, 6, col("C")) : 0;
    if (oos && oosRow > 0) {
      fillExcludedBaseValuesFromOos(oos, oosRow, excluded, targetRow);
    }
    if (rank) excluded.cell(targetRow, col("A")).value(rank);
    if (fullName) excluded.cell(targetRow, col("B")).value(fullName);
    if (personId) excluded.cell(targetRow, col("C")).value(personId);
    if (previousIndex) excluded.cell(targetRow, col("D")).value(previousIndex);
    excluded.cell(targetRow, col("AB")).value(excludeDateLabel || null);
    excluded.cell(targetRow, col("AC")).value(op.payload.orderDate || null);
    excluded.cell(targetRow, col("AD")).value(op.payload.orderNumber || null);
    excluded
      .cell(targetRow, col("AE"))
      .value(
        formatExcludedDestination(
          op.payload.documentsDest || op.payload.changeText || "",
        ) || null,
      );
    excluded.cell(targetRow, col("AF")).value(
      formatExcludedListBasis({
        ...op.payload,
        exclusionReason: op.payload.exclusionReason || "ПЕРЕВЕДЕННЯ 1 ПБ",
      }),
    );
    excluded.cell(targetRow, col("AF")).style?.("wrapText", true);
    excluded.cell(targetRow, col("Z")).style?.("wrapText", true);
    excluded.cell(targetRow, col("E")).style?.("wrapText", true);
  }

  const oldTimesheetRow = Number(
    op.payload.previousIndexTimesheetExcelRow || 0,
  );
  if (timesheet && oldTimesheetRow > 0) {
    const preserveHistory = op.payload.timesheetPreserveHistory === "1";
    if (preserveHistory) {
      const historyRow = findTimesheetAppendNear(timesheet, oldTimesheetRow);
      copyTimesheetHistoryValues(
        timesheet,
        oldTimesheetRow,
        historyRow,
        timesheetStyleCopies,
      );
      if (rank) timesheet.cell(historyRow, col("F")).value(rank);
      writeTimesheetTransferHistory(timesheet, historyRow, {
        departDay,
        destination: op.payload.timesheetDestination,
        payload: op.payload,
        orderNumber: op.payload.orderNumber,
        orderDate: op.payload.orderDate,
        // Особа не відсутня в частині, а переходить на іншу посаду:
        // після дня вибуття старий рядок лишається порожнім.
        lastDay: departDay,
        absenceSpans: historyAbsenceSpansForClosedEpisode(op.payload, departDay),
      });
    }
    timesheet.cell(oldTimesheetRow, col("F")).value(null);
    timesheet.cell(oldTimesheetRow, col("G")).value(null);
    timesheet.cell(oldTimesheetRow, col("H")).value(null);
    for (let day = 1; day <= 31; day += 1) {
      timesheet.cell(oldTimesheetRow, col("I") + day - 1).value(null);
    }
  }

  const oldShpoRow = Number(op.payload.previousShpoExcelRow || 0);
  if (shpo && oldShpoRow > 0) {
    shpo.cell(oldShpoRow, col("F")).value(null);
    shpo.cell(oldShpoRow, col("G")).value(null);
    shpo.cell(oldShpoRow, col("H")).value(null);
    shpo.cell(oldShpoRow, col("R")).value(null);
  }
}

/** Прибирає службовий хвіст ШПО, де в рядку є лише ПІБ без штатної посади/ID. */
function removeShpoNameOnlyRows(sheet: SheetLike) {
  const endRow = sheet.usedRange()?.endCell().rowNumber() ?? 0;
  for (let row = 7; row <= endRow; row += 1) {
    const nameCell = sheet.cell(row, col("G"));
    if (!isMeaningfulCellValue(nameCell.value())) continue;

    let hasOtherData = false;
    for (let column = col("A"); column <= col("R"); column += 1) {
      if (column === col("G")) continue;
      if (isMeaningfulCellValue(sheet.cell(row, column).value())) {
        hasOtherData = true;
        break;
      }
    }
    if (hasOtherData) continue;

    nameCell.value(null);
  }
}

/**
 * Після зміни формул/рядків старий calcChain містить неактуальні посилання.
 * Excel відновлює такий файл при відкритті, тому видаляємо ланцюг і просимо
 * Excel зробити повний перерахунок без перепакування книги редактором.
 */
const touchedSheetNamesForOps = (ops: EjoosSyncOp[]) => {
  const names = new Set<string>();
  for (const op of ops) {
    if (op.kind === "shpo_occupant") {
      names.add("1. ШПО");
      names.add("6. Табель");
      if (op.payload.excludedSourceExcelRow) names.add("2. ООС");
      if (op.payload.clearExcludedExcelRow) names.add("3. Виключені");
    } else if (op.kind === "timesheet_day") {
      names.add("6. Табель");
      if (op.payload.clearExcludedExcelRow) names.add("3. Виключені");
    } else if (op.kind === "absent_upsert") {
      names.add("5. Тимчасово відсутні");
      names.add("6. Табель");
    } else if (op.kind === "absent_close") {
      names.add("5. Тимчасово відсутні");
      names.add("6. Табель");
    } else if (op.kind === "exclude_transfer") {
      names.add("1. ШПО");
      names.add("2. ООС");
      names.add("3. Виключені");
      names.add("6. Табель");
      if (op.payload.arrivalExcelRow) names.add("4. Тимчасово прибулі");
    } else if (op.kind === "position_change") {
      names.add("1. ШПО");
      names.add("2. ООС");
      names.add("4. Тимчасово прибулі");
      names.add("6. Табель");
      if (positionCloseWritesExcluded(op.payload)) names.add("3. Виключені");
      if (op.payload.clearExcludedExcelRow) names.add("3. Виключені");
    } else if (
      op.kind === "other_manual" &&
      op.payload.type === "TRANSFER_CANCELLED" &&
      op.payload.reviewReason !== "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH"
    ) {
      names.add("3. Виключені");
      if (op.payload.restoreShpo === "1") names.add("1. ШПО");
      if (op.payload.restoreOos === "1") names.add("2. ООС");
      if (op.payload.restoreTimesheet === "1") names.add("6. Табель");
    } else if (
      op.kind === "other_manual" &&
      op.payload.type === "TRANSFER_CANCELLED" &&
      op.payload.excludedExcelRow
    ) {
      names.add("3. Виключені");
    }
  }
  return names;
};

const workbookSheetPaths = async (zip: JSZip) => {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  const paths = new Map<string, string>();
  if (!workbookXml || !relsXml) return paths;
  const targets = new Map<string, string>();
  for (const relation of relsXml.matchAll(
    /<Relationship\b(?=[^>]*\bId="([^"]+)")(?=[^>]*\bTarget="([^"]+)")[^>]*\/>/gi,
  )) {
    targets.set(relation[1], relation[2]);
  }
  for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*)>/gi)) {
    const attrs = sheet[1];
    const name = attrs.match(/\bname="([^"]+)"/)?.[1];
    const relationId = attrs.match(/\br:id="([^"]+)"/)?.[1];
    const target = relationId ? targets.get(relationId) : undefined;
    if (!name || !target) continue;
    paths.set(
      name,
      target.startsWith("/")
        ? target.slice(1)
        : target.startsWith("xl/")
          ? target
          : `xl/${target.replace(/^\.\//, "")}`,
    );
  }
  return paths;
};

async function finalizeEjoosWorkbookBlob(
  file: Blob,
  sourceFile: Blob | File,
  touchedSheetNames: Set<string>,
) {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sourceZip = await JSZip.loadAsync(await sourceFile.arrayBuffer());

  // xlsx-populate переписує навіть аркуші, яких операція не торкалася.
  // Повертаємо їхній XML дослівно з базової справної версії.
  const sourceSheetPaths = await workbookSheetPaths(sourceZip);
  for (const [sheetName, sheetPath] of sourceSheetPaths) {
    if (touchedSheetNames.has(sheetName)) continue;
    const sourceSheet = sourceZip.file(sheetPath);
    if (!sourceSheet) continue;
    zip.file(sheetPath, await sourceSheet.async("uint8array"));
  }

  for (const path of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(path)) continue;
    const worksheet = zip.file(path);
    if (!worksheet) continue;
    const xml = await worksheet.async("string");
    const cleaned = repairBrokenCellOpenTags(
      stripInvalidNumericCellValues(xml),
    );
    if (cleaned !== xml) zip.file(path, cleaned);
  }

  // Нові значення використовують наявні стилі клітинок. Не дозволяємо
  // бібліотеці глобально перебудовувати style table всієї книги.
  const sourceStyles = sourceZip.file("xl/styles.xml");
  if (sourceStyles) {
    const sourceStylesXml = await sourceStyles.async("string");
    zip.file("xl/styles.xml", sourceStylesXml);
    const cellXfsBody =
      sourceStylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ??
      "";
    const sourceStyleCount = (cellXfsBody.match(/<xf\b/gi) ?? []).length;
    if (sourceStyleCount > 0) {
      for (const path of Object.keys(zip.files)) {
        if (!/^xl\/worksheets\/sheet\d+\.xml$/i.test(path)) continue;
        const worksheet = zip.file(path);
        if (!worksheet) continue;
        const xml = await worksheet.async("string");
        const repaired = xml.replace(/<c\b([^>]*)>/gi, (cellTag, attrs) => {
          const style = String(attrs).match(/\bs="(\d+)"/i);
          if (!style || Number(style[1]) < sourceStyleCount) return cellTag;
          return `<c${String(attrs).replace(/\s+s="\d+"/i, "")}>`;
        });
        if (repaired !== xml) zip.file(path, repaired);
      }
    }
  }

  zip.remove("xl/calcChain.xml");
  zip.remove("xl/ejoosAppliedChanges.json");

  const relsPath = "xl/_rels/workbook.xml.rels";
  const rels = await zip.file(relsPath)?.async("string");
  if (rels) {
    zip.file(
      relsPath,
      rels.replace(
        /<Relationship\b(?=[^>]*(?:Type="[^"]*\/calcChain"|Target="(?:\.\.\/)?calcChain\.xml"))[^>]*\/>/gi,
        "",
      ),
    );
  }

  const contentTypesPath = "[Content_Types].xml";
  const contentTypes = await zip.file(contentTypesPath)?.async("string");
  if (contentTypes) {
    zip.file(
      contentTypesPath,
      contentTypes.replace(
        /<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/gi,
        "",
      ),
    );
  }

  const workbookPath = "xl/workbook.xml";
  const workbookXml = await zip.file(workbookPath)?.async("string");
  if (workbookXml) {
    const calcPr =
      '<calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>';
    const nextWorkbookXml = /<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/i.test(
      workbookXml,
    )
      ? workbookXml.replace(
          /<calcPr\b[^>]*(?:\/>|>[\s\S]*?<\/calcPr>)/i,
          calcPr,
        )
      : workbookXml.replace("</workbook>", `${calcPr}</workbook>`);
    zip.file(workbookPath, nextWorkbookXml);
  }

  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  });
}

type SheetLike = {
  usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
  cell: (
    r: number,
    c: number,
  ) => {
    value: (v?: unknown) => unknown;
    style?: {
      (name: string): unknown;
      (names: string[]): Record<string, unknown>;
      (styles: Record<string, unknown>): unknown;
      (name: string, value: unknown): unknown;
    };
  };
  row?: (r: number) => { height: (h?: number) => number };
};

type WorkbookLike = {
  sheet: (name: string) => SheetLike | undefined;
  outputAsync: (type: "blob") => Promise<Blob>;
};

/** ПЕРЕВ з обліку: Виключені → Табель → ШПО/ООС. Відкрите тимч. прибуття закриваємо, якщо лишилось після пропущеної внутрішньої ПОСАДИ. */
function applyExcludeTransfer(input: {
  op: EjoosSyncOp;
  plan: EjoosSyncPlan;
  excluded: SheetLike | undefined;
  timesheet: SheetLike | undefined;
  shpo: SheetLike | undefined;
  oos: SheetLike | undefined;
  arrivals: SheetLike | undefined;
  timesheetStyleCopies: TimesheetStyleCopy[];
}) {
  const { op, plan, excluded, timesheet, shpo, oos, arrivals, timesheetStyleCopies } =
    input;
  const personId = op.payload.fromPersonId || op.personId;
  const fullName = op.payload.fromName || op.fullName;
  const rank = op.payload.fromRank || op.rank;
  const positionIndex =
    op.payload.fromPositionIndex ||
    op.payload.previousIndex ||
    op.positionIndex;
  const destination =
    op.payload.destination || op.payload.documentsDest || "";
  const documentsDestination =
    op.payload.documentsDest || op.payload.changeText || "";
  const timesheetDestination = op.payload.timesheetDestination || destination;
  const excludeDateLabel =
    op.payload.excludeDate;
  const departDay = dayFromOrderLabel(excludeDateLabel);

  let shpoRow = Number(op.payload.shpoExcelRow || 0);
  if (shpo && shpoRow <= 0) {
    shpoRow = findPersonRowFlexible(shpo, {
      personId,
      fullName,
      positionIndex: op.payload.occupiedPositionIndex || positionIndex,
      startRow: 7,
      nameCol: col("G"),
      idCol: col("H"),
      indexCol: col("A"),
      matchByIndex: excludeWritePlan(op.payload).matchShpoByIndex,
    });
  }

  let timesheetRow = Number(op.payload.timesheetExcelRow || 0);
  if (timesheet && timesheetRow <= 0) {
    timesheetRow = findPersonRowFlexible(timesheet, {
      personId,
      fullName,
      positionIndex: op.payload.occupiedPositionIndex || positionIndex,
      startRow: 7,
      nameCol: col("G"),
      idCol: col("H"),
      indexCol: col("B"),
      matchByIndex: excludeWritePlan(op.payload).matchTimesheetByIndex,
    });
  }
  const timesheetWritePlan = excludeWritePlan(op.payload);
  if (
    timesheet &&
    timesheetRow > 0 &&
    timesheetRowHasDeparture(timesheet, timesheetRow) &&
    !timesheetWritePlan.replaceInPlace
  ) {
    timesheetRow = 0;
  }

  // 1) Виключені — A:AA переносимо з ООС, AB:AF дописуємо з Рух
  if (excluded && !op.payload.excludedExcelRow) {
    const targetRow = nextExcludedManualRow(excluded, 6);
    // Не передаємо styles.xml через API xlsx-populate: саме копіювання стилів
    // створювало сотні нових font/fill/border/xf та ламало наступну версію.
    copyRowHeight(excluded, 6, targetRow);
    const oosRow =
      oos && personId
        ? findPersonIdRow(oos, personId, 6, col("C"))
        : oos && fullName
          ? findPersonRowFlexible(oos, {
              personId: "",
              fullName,
              positionIndex: "",
              startRow: 6,
              nameCol: col("B"),
              idCol: col("C"),
              indexCol: col("D"),
            })
          : 0;

    if (oos && oosRow > 0) {
      if (rank) oos.cell(oosRow, col("A")).value(rank);
      if (op.payload.lastRankOrderDate) {
        oos.cell(oosRow, col("O")).value(op.payload.lastRankOrderDate);
      }
      if (op.payload.lastRankOrderNumber) {
        oos.cell(oosRow, col("P")).value(op.payload.lastRankOrderNumber);
      }
      fillExcludedBaseValuesFromOos(oos, oosRow, excluded, targetRow);
    } else if (shpo && shpoRow > 0) {
      // Fallback без ООС: переносимо тільки те, що є у ШПО.
      copyCell(shpo, shpoRow, col("F"), excluded, targetRow, col("A"));
      copyCell(shpo, shpoRow, col("G"), excluded, targetRow, col("B"));
      copyCell(shpo, shpoRow, col("H"), excluded, targetRow, col("C"));
      copyCell(shpo, shpoRow, col("A"), excluded, targetRow, col("D"));
    } else {
      if (rank) excluded.cell(targetRow, col("A")).value(rank);
      if (fullName) excluded.cell(targetRow, col("B")).value(fullName);
      if (personId) excluded.cell(targetRow, col("C")).value(personId);
    }
    // D в ООС інколи є формульною/нестандартною клітинкою, яку xlsx-populate
    // повертає не як просте значення. План уже зчитав актуальний індекс із ООС,
    // тому закріплюємо його явно після переносу базової картки.
    if (positionIndex) excluded.cell(targetRow, col("D")).value(positionIndex);
    // Звання вже визначене у плані з ЕЖООС; РУХ тут лише fallback.
    if (rank) excluded.cell(targetRow, col("A")).value(rank);
    if (op.payload.lastRankOrderDate) {
      excluded.cell(targetRow, col("L")).value(op.payload.lastRankOrderDate);
    }
    if (op.payload.lastRankOrderNumber) {
      excluded.cell(targetRow, col("M")).value(op.payload.lastRankOrderNumber);
    }
    // Пишемо готовий dd.mm.yyyy як текст без створення нових style records.
    // Оформлення нового рядка буде відновлене окремим OOXML-шаром.
    excluded.cell(targetRow, col("AB")).value(excludeDateLabel || null);
    excluded
      .cell(targetRow, col("AC"))
      .value(op.payload.orderDate || null);
    excluded.cell(targetRow, col("AD")).value(op.payload.orderNumber || null);
    excluded
      .cell(targetRow, col("AE"))
      .value(formatExcludedDestination(documentsDestination) || null);
    excluded.cell(targetRow, col("AF")).value(formatExcludedListBasis(op.payload));
    excluded.cell(targetRow, col("AF")).style?.("wrapText", true);
    excluded.cell(targetRow, col("Z")).style?.("wrapText", true);
    excluded.cell(targetRow, col("E")).style?.("wrapText", true);
  }

  // 2) Табель: закритий епізод копіюємо вниз (як при ПОСАДА), штат лишає індекс/ВОС/тариф.
  if (timesheet) {
    const existingRows = findTimesheetPersonRows(timesheet, personId, fullName);
    const plannedKeep = timesheetRow;
    const openRows = existingRows.filter(
      (row) => !timesheetRowHasDeparture(timesheet, row),
    );
    let keepRow =
      (plannedKeep && openRows.includes(plannedKeep) ? plannedKeep : 0) ||
      openRows[0] ||
      (timesheetWritePlan.replaceInPlace &&
      plannedKeep &&
      existingRows.includes(plannedKeep)
        ? plannedKeep
        : 0) ||
      (timesheetWritePlan.replaceInPlace ? existingRows[0] : 0) ||
      plannedKeep ||
      0;
    const writeClosedEpisode = (rowNumber: number) => {
      if (rank) timesheet.cell(rowNumber, col("F")).value(rank);
      if (fullName) timesheet.cell(rowNumber, col("G")).value(fullName);
      if (personId) timesheet.cell(rowNumber, col("H")).value(personId);
      writeTimesheetTransferHistory(timesheet, rowNumber, {
        departDay,
        destination: timesheetDestination,
        payload: op.payload,
        orderNumber: op.payload.orderNumber,
        orderDate: op.payload.orderDate,
        lastDay: Math.min(31, Math.max(departDay, plan.timesheetDay)),
        absenceSpans: parseTimesheetAbsenceSpans(
          op.payload.timesheetAbsenceSpans || "",
        ),
        activeFromDay: dayFromOrderLabel(op.payload.timesheetActiveFrom || ""),
      });
    };
    const clearTimesheetOccupant = (rowNumber: number) => {
      timesheet.cell(rowNumber, col("F")).value(null);
      timesheet.cell(rowNumber, col("G")).value(null);
      timesheet.cell(rowNumber, col("H")).value(null);
      timesheet.cell(rowNumber, col("AN")).value(null);
      for (let day = 1; day <= 31; day += 1) {
        timesheet.cell(rowNumber, col("I") + day - 1).value(null);
      }
    };
    if (keepRow > 0 && timesheetWritePlan.replaceInPlace) {
      writeClosedEpisode(keepRow);
    } else if (keepRow > 0 && !timesheetWritePlan.replaceInPlace) {
      const historyRow = findTimesheetAppendNear(timesheet, keepRow);
      copyTimesheetHistoryValues(
        timesheet,
        keepRow,
        historyRow,
        timesheetStyleCopies,
      );
      writeClosedEpisode(historyRow);
      clearTimesheetOccupant(keepRow);
      keepRow = historyRow;
    } else if (timesheetWritePlan.createTimesheetHistory) {
      keepRow = findTimesheetAppendNear(timesheet, 7);
      recordTimesheetStyleCopy(
        timesheetStyleCopies,
        occupiedTimesheetStyleRow(timesheet, keepRow),
        keepRow,
      );
      if (positionIndex) timesheet.cell(keepRow, col("B")).value(positionIndex);
      writeClosedEpisode(keepRow);
    }
    clearOtherTimesheetPersonRows(timesheet, {
      personId,
      fullName,
      keepRow,
      keepHistory: true,
    });
  }

  // 3) ШПО — прибрати звання/ПІБ/ID, лишити посаду
  if (shpo && shpoRow > 0) {
    shpo.cell(shpoRow, col("F")).value(null);
    shpo.cell(shpoRow, col("G")).value(null);
    shpo.cell(shpoRow, col("H")).value(null);
    shpo.cell(shpoRow, col("R")).value(null);
  }

  // 4) ООС — видалити рядок
  if (oos && personId) {
    const oosRow = findPersonIdRow(oos, personId, 6, col("C"));
    if (oosRow) clearOosRow(oos, oosRow);
  } else if (oos && fullName) {
    const oosRow = findPersonRowFlexible(oos, {
      personId: "",
      fullName,
      positionIndex: "",
      startRow: 6,
      nameCol: col("B"),
      idCol: col("C"),
      indexCol: col("D"),
    });
    if (oosRow) clearOosRow(oos, oosRow);
  }

  const arrivalRow = Number(op.payload.arrivalExcelRow || 0);
  if (arrivals && arrivalRow > 0) {
    const departDate =
      op.payload.arrivalDepartDate || op.payload.appointmentOrderDate || "";
    const orderNumber =
      op.payload.arrivalDepartOrderNumber ||
      op.payload.appointmentOrderNumber ||
      op.payload.orderNumber;
    const orderDate =
      op.payload.arrivalDepartOrderDate ||
      op.payload.appointmentOrderDate ||
      "";
    const closeColumns = findArrivalCloseColumns(arrivals);
    if (closeColumns.departDateCol) {
      arrivals.cell(arrivalRow, closeColumns.departDateCol).value(departDate || null);
    }
    if (closeColumns.orderDateCol) {
      arrivals.cell(arrivalRow, closeColumns.orderDateCol).value(orderDate || null);
    }
    if (closeColumns.orderNumberCol) {
      arrivals
        .cell(arrivalRow, closeColumns.orderNumberCol)
        .value(orderNumber ? `№${orderNumber}` : null);
    }
    if (closeColumns.orderCombinedCol) {
      arrivals.cell(arrivalRow, closeColumns.orderCombinedCol).value(
        [orderNumber ? `№${orderNumber}` : "", orderDate ? `від ${orderDate}` : ""]
          .filter(Boolean)
          .join(" "),
      );
    }
  }
}

function copyRowHeight(
  sheet: SheetLike,
  sourceRow: number,
  targetRow: number,
) {
  const height = sheet.row?.(sourceRow).height();
  if (typeof height === "number") sheet.row?.(targetRow).height(height);
}

function nextExcludedManualRow(sheet: SheetLike, minRow: number) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? minRow;
  let lastManual = minRow - 1;
  for (let row = minRow; row <= end; row += 1) {
    const hasManualValue = [
      col("A"),
      col("B"),
      col("AB"),
      col("AC"),
      col("AD"),
      col("AE"),
      col("AF"),
    ].some((column) => {
      const value = sheet.cell(row, column).value();
      return value !== undefined && value !== null && String(value).trim() !== "";
    });
    if (hasManualValue) lastManual = row;
  }
  return lastManual + 1;
}

function fillExcludedBaseValuesFromOos(
  oos: SheetLike,
  oosRow: number,
  excluded: SheetLike,
  targetRow: number,
) {
  OOS_TO_EXCLUDED_BASE.forEach(([fromCol, toCol]) => {
    const value = oos.cell(oosRow, fromCol).value();
    if (isMeaningfulCellValue(value)) {
      const next =
        toCol === 5 ? formatExcludedPositionDates(value) || value : value;
      excluded.cell(targetRow, toCol).value(next);
      if (toCol === 5) excluded.cell(targetRow, toCol).style?.("wrapText", true);
    } else {
      excluded.cell(targetRow, toCol).value(null);
    }
  });
}

function isMeaningfulCellValue(value: unknown) {
  if (value === undefined || value === null) return false;
  if (value instanceof Date) return !Number.isNaN(value.getTime());
  if (typeof value === "object") {
    const rich = value as { text?: () => string };
    if (typeof rich.text === "function") {
      return rich.text().trim() !== "";
    }
    return false;
  }
  return String(value).trim() !== "";
}

function copyCell(
  from: SheetLike,
  fromRow: number,
  fromCol: number,
  to: SheetLike,
  toRow: number,
  toCol: number,
) {
  const value = from.cell(fromRow, fromCol).value();
  if (isMeaningfulCellValue(value)) {
    to.cell(toRow, toCol).value(value);
  }
}

function copySheetRowValues(
  sheet: SheetLike,
  sourceRow: number,
  targetRow: number,
  endCol: number,
) {
  copyRowHeight(sheet, sourceRow, targetRow);
  for (let column = 1; column <= endCol; column += 1) {
    sheet
      .cell(targetRow, column)
      .value(sheet.cell(sourceRow, column).value());
  }
}

type TimesheetStyleCopy = { sourceRow: number; targetRow: number };

function recordTimesheetStyleCopy(
  jobs: TimesheetStyleCopy[],
  sourceRow: number,
  targetRow: number,
) {
  if (sourceRow >= 7 && targetRow >= 7 && sourceRow !== targetRow) {
    jobs.push({ sourceRow, targetRow });
  }
}

function copyTimesheetHistoryValues(
  timesheet: SheetLike,
  sourceRow: number,
  targetRow: number,
  jobs: TimesheetStyleCopy[],
) {
  copySheetRowValues(timesheet, sourceRow, targetRow, col("AN"));
  recordTimesheetStyleCopy(jobs, sourceRow, targetRow);
}

function occupiedTimesheetStyleRow(timesheet: SheetLike, skipRow: number) {
  for (let row = skipRow - 1; row >= 7; row -= 1) {
    const name = String(timesheet.cell(row, col("G")).value() ?? "").trim();
    if (name) return row;
  }
  const end = timesheet.usedRange()?.endCell().rowNumber() ?? skipRow;
  for (let row = 7; row <= end; row += 1) {
    if (row === skipRow) continue;
    const name = String(timesheet.cell(row, col("G")).value() ?? "").trim();
    if (name) return row;
  }
  return 7;
}

function timesheetRowHasDeparture(sheet: SheetLike, rowNumber: number) {
  for (let day = 1; day <= 31; day += 1) {
    if (
      isTimesheetDepartureMark(
        sheet.cell(rowNumber, col("I") + day - 1).value(),
      )
    ) {
      return true;
    }
  }
  return false;
}

function writeTimesheetTransferHistory(
  sheet: SheetLike,
  rowNumber: number,
  opts: {
    departDay: number;
    destination: string;
    payload?: {
      type?: string;
      exclusionReason?: string;
      destination?: string;
      documentsDest?: string;
      timesheetDestination?: string;
      changeText?: string;
    };
    orderNumber: string;
    orderDate: string;
    lastDay: number;
    absenceSpans?: TimesheetAbsenceSpan[];
    activeFromDay?: number;
  },
) {
  const absenceSpans = opts.absenceSpans ?? [];
  const activeFromDay = opts.activeFromDay || 1;
  const departureText = formatTimesheetTransferMark(
    opts.payload || {
      destination: opts.destination,
      timesheetDestination: opts.destination,
    },
  );
  let presentDays = 0;
  for (let day = 1; day <= 31; day += 1) {
    const cell = sheet.cell(rowNumber, col("I") + day - 1);
    const mark = timesheetTransferMarkForDay({
      day,
      departDay: opts.departDay,
      lastDay: opts.lastDay,
      activeFromDay,
      absenceSpans,
      departMark: departureText,
    });
    cell.value(mark);
    if (mark === "+") presentDays += 1;
  }
  sheet.cell(rowNumber, col("AN")).value(presentDays);
}

function findTimesheetPersonRows(
  sheet: SheetLike,
  personId: string,
  fullName: string,
) {
  const usedEnd = sheet.usedRange()?.endCell().rowNumber() ?? 7;
  const grid: unknown[][] = [];
  let empty = 0;
  for (let row = 6; row <= usedEnd + 400; row += 1) {
    const cells: unknown[] = [];
    cells[1] = sheet.cell(row, col("B")).value();
    cells[5] = sheet.cell(row, col("F")).value();
    cells[6] = sheet.cell(row, col("G")).value();
    cells[7] = sheet.cell(row, col("H")).value();
    grid[row - 1] = cells;
    const named =
      String(cells[6] ?? "").trim() || String(cells[7] ?? "").trim();
    if (row <= usedEnd) continue;
    empty = named ? 0 : empty + 1;
    if (empty >= 80) break;
  }
  return findTimesheetPersonRowsInGrid(grid, personId, fullName);
}

function clearOtherTimesheetPersonRows(
  sheet: SheetLike,
  input: {
    personId: string;
    fullName: string;
    keepRow: number;
    keepHistory?: boolean;
  },
) {
  const rows = findTimesheetPersonRows(sheet, input.personId, input.fullName);
  let keep = input.keepRow;
  if (!keep) {
    keep =
      rows.find((row) => !timesheetRowHasDeparture(sheet, row)) ||
      rows.find((row) => timesheetRowHasDeparture(sheet, row)) ||
      rows[0] ||
      0;
  }
  for (const row of rows) {
    if (row === keep) continue;
    if (input.keepHistory && timesheetRowHasDeparture(sheet, row)) continue;
    sheet.cell(row, col("F")).value(null);
    sheet.cell(row, col("G")).value(null);
    sheet.cell(row, col("H")).value(null);
    sheet.cell(row, col("AN")).value(null);
    for (let day = 1; day <= 31; day += 1) {
      sheet.cell(row, col("I") + day - 1).value(null);
    }
  }
}

function findTimesheetAppendNear(sheet: SheetLike, sourceRow: number) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? sourceRow;
  for (let row = sourceRow + 1; row <= end + 30; row += 1) {
    const index = String(sheet.cell(row, col("B")).value() ?? "").trim();
    const name = String(sheet.cell(row, col("G")).value() ?? "").trim();
    const id = String(sheet.cell(row, col("H")).value() ?? "").trim();
    if (!index && !name && !id) return row;
  }
  return nextAppendRow(sheet, sourceRow + 1);
}

function findPersonRowFlexible(
  sheet: SheetLike,
  opts: {
    personId: string;
    fullName: string;
    positionIndex: string;
    startRow: number;
    nameCol: number;
    idCol: number;
    indexCol: number;
    matchByIndex?: boolean;
  },
) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? opts.startRow;
  const nameKey = opts.fullName
    .trim()
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ");
  const idKey = String(opts.personId ?? "").trim();
  let indexMatch = 0;
  for (let row = opts.startRow; row <= end; row += 1) {
    const id = String(sheet.cell(row, opts.idCol).value() ?? "").trim();
    const name = String(sheet.cell(row, opts.nameCol).value() ?? "")
      .trim()
      .toLowerCase()
      .replace(/[''`´']/g, "")
      .replace(/\s+/g, " ");
    const index = String(sheet.cell(row, opts.indexCol).value() ?? "").trim();
    if (idKey && id === idKey) return row;
    if (nameKey && name === nameKey) return row;
    if (
      opts.matchByIndex !== false &&
      opts.positionIndex &&
      index === opts.positionIndex
    ) {
      if (id || name) indexMatch = row;
      else if (!indexMatch) indexMatch = row;
    }
  }
  // Якщо ПІБ/ID не збіглися (розбіжності написання) — посада з людиною на індексі.
  return opts.matchByIndex === false ? 0 : indexMatch;
}

const timesheetRowIsHistoryOnly = (
  timesheet: {
    cell: (row: number, column: number) => { value: () => unknown };
  },
  rowNumber: number,
  activeFromDay: number,
) => {
  let departed = false;
  let plusAfter = false;
  for (let day = 1; day <= 31; day += 1) {
    const value = timesheet.cell(rowNumber, col("I") + day - 1).value();
    if (isTimesheetDepartureMark(value)) departed = true;
    if (String(value ?? "").trim() === "+" && day >= Math.max(1, activeFromDay)) {
      plusAfter = true;
    }
  }
  return departed && !plusAfter;
};

const paintTimesheetArchiveDays = (
  timesheet: {
    cell: (
      row: number,
      column: number,
    ) => { value: (next?: unknown) => unknown };
  },
  rowNumber: number,
  plan: EjoosSyncPlan,
  payload: Record<string, string>,
) => {
  if (rowNumber <= 0) return;
  const activeFromDay = dayFromOrderLabel(payload.timesheetActiveFrom || "");
  if (
    timesheetRowIsHistoryOnly(timesheet, rowNumber, activeFromDay) &&
    !payload.transferCancelOrder
  ) {
    return;
  }
  const rawSpans = parseTimesheetAbsenceSpans(payload.timesheetAbsenceSpans || "");
  const spans =
    activeFromDay > 1
      ? clipAbsenceSpansToActiveEpisode(rawSpans, activeFromDay)
      : rawSpans;
  const fillBeforeActive = activeFromDay > 1;
  const spanEnd = spans.reduce((max, span) => Math.max(max, span.toDay), 0);
  const lastDay = Math.max(plan.timesheetDay, spanEnd);
  let presentDays = 0;
  for (let day = 1; day <= 31; day += 1) {
    const mark = timesheetMarkFromArchive(day, {
      activeFromDay: activeFromDay || 1,
      lastDay,
      spans,
      fillBeforeActive,
    });
    if (mark == null) {
      if (day > lastDay) {
        timesheet.cell(rowNumber, col("I") + day - 1).value(null);
      }
      continue;
    }
    timesheet.cell(rowNumber, col("I") + day - 1).value(mark);
    if (day <= lastDay && mark === "+") presentDays += 1;
  }
  timesheet.cell(rowNumber, col("AN")).value(presentDays);

  const historyRow = Number(payload.historyTimesheetExcelRow || 0);
  const historySpans = parseTimesheetAbsenceSpans(
    payload.historyTimesheetAbsenceSpans || "",
  );
  if (historyRow > 0 && historySpans.length && historyRow !== rowNumber) {
    for (let day = 1; day <= 31; day += 1) {
      const code = timesheetCodeOnDay(day, historySpans);
      if (!code) continue;
      const current = timesheet.cell(historyRow, col("I") + day - 1).value();
      if (isTimesheetDepartureMark(current)) continue;
      timesheet.cell(historyRow, col("I") + day - 1).value(code);
    }
  }
};

const repairHistoryTimesheetRow = (
  timesheet: SheetLike,
  op: EjoosSyncOp,
  lastDay: number,
) => {
  const historyRow = Number(op.payload.historyTimesheetExcelRow || 0);
  const departDay = dayFromOrderLabel(op.payload.historyDepartDate || "");
  if (historyRow <= 0 || departDay <= 0) return;
  if (timesheetRowHasDeparture(timesheet, historyRow)) return;
  writeTimesheetTransferHistory(timesheet, historyRow, {
    departDay,
    destination: op.payload.historyDepartDest || "",
    payload: {
      destination: op.payload.historyDepartDest || "",
      timesheetDestination: op.payload.historyDepartDest || "",
    },
    orderNumber: op.payload.historyOrderNumber || "",
    orderDate: op.payload.historyDepartDate || "",
    lastDay,
  });
};

const writeAbsenceReturn = (
  absent: {
    cell: (
      row: number,
      column: number,
    ) => { value: (next?: unknown) => unknown };
  },
  rowNumber: number,
  payload: Record<string, string>,
  fallbackDate: string,
) => {
  const returnDate = payload.returnDate || fallbackDate;
  if (returnDate) absent.cell(rowNumber, col("M")).value(returnDate);
  const returnOrderDate = payload.returnOrderDate || "";
  const returnOrderNumber = payload.returnOrderNumber || "";
  absent.cell(rowNumber, col("N")).value(null);
  if (returnOrderDate) absent.cell(rowNumber, col("O")).value(returnOrderDate);
  if (returnOrderNumber) {
    absent.cell(rowNumber, col("P")).value(returnOrderNumber);
  }
};

function nextAppendRow(
  sheet: {
    usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
  },
  minRow: number,
) {
  const used = sheet.usedRange();
  return Math.max(minRow, (used?.endCell().rowNumber() ?? minRow - 1) + 1);
}

function writeOosWrappedCell(
  sheet: SheetLike,
  row: number,
  column: number,
  value: string,
) {
  sheet.cell(row, column).value(value || null);
}

function appendCellHistory(
  sheet: SheetLike,
  row: number,
  column: number,
  nextValue: string,
) {
  const merged = mergeOosHistoryValue(sheet.cell(row, column).value(), nextValue);
  if (!merged) return;
  writeOosWrappedCell(sheet, row, column, merged);
}

function nextEmptyAbsentRow(sheet: {
  usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
  cell: (r: number, c: number) => { value: () => unknown };
}) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? 6;
  for (let row = 6; row <= end + 5; row += 1) {
    const name = String(sheet.cell(row, 2).value() ?? "").trim();
    if (!name) return row;
  }
  return end + 1;
}

function findPersonIdRow(
  sheet: {
    usedRange: () => { endCell: () => { rowNumber: () => number } } | undefined;
    cell: (r: number, c: number) => { value: () => unknown };
  },
  personId: string,
  startRow: number,
  idCol: number,
) {
  const end = sheet.usedRange()?.endCell().rowNumber() ?? startRow;
  for (let row = startRow; row <= end; row += 1) {
    if (String(sheet.cell(row, idCol).value() ?? "").trim() === personId) {
      return row;
    }
  }
  return 0;
}

function clearDuplicateTimesheetRow(
  sheet: SheetLike | undefined,
  op: EjoosSyncOp,
) {
  if (!sheet) return;
  const rowNumber = Number(op.payload.duplicateTimesheetExcelRow || 0);
  if (rowNumber <= 0) return;
  const keepRow =
    Number(op.payload.timesheetExcelRow || 0) ||
    Number(op.payload.keepTimesheetExcelRow || 0);
  if (keepRow > 0 && rowNumber === keepRow) return;
  if (
    op.kind === "timesheet_day" &&
    Number(op.payload.excelRow || 0) === rowNumber
  ) {
    return;
  }
  if (op.payload.clearTimesheetIndex === "1") {
    sheet.cell(rowNumber, col("B")).value(null);
  }
  sheet.cell(rowNumber, col("F")).value(null);
  sheet.cell(rowNumber, col("G")).value(null);
  sheet.cell(rowNumber, col("H")).value(null);
  for (let day = 1; day <= 31; day += 1) {
    sheet.cell(rowNumber, col("I") + day - 1).value(null);
  }
  sheet.cell(rowNumber, col("AN")).value(null);
}

function clearOosRow(
  sheet: {
    cell: (r: number, c: number) => { value: (v?: unknown) => unknown };
  },
  row: number,
) {
  for (let c = 1; c <= 40; c += 1) {
    sheet.cell(row, c).value(null);
  }
}

export async function blobToBase64(blob: Blob): Promise<string> {
  const buffer = await blob.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return btoa(binary);
}

export function downloadTextFile(fileName: string, text: string) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

export function downloadBlobFile(fileName: string, blob: Blob) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}

export async function fileToBase64(file: File): Promise<string> {
  return blobToBase64(file);
}

export function base64ToFile(base64: string, fileName: string): File {
  const raw = base64.includes(",") ? base64.split(",")[1] : base64;
  const binary = atob(raw);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new File(
    [bytes],
    fileName,
    {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    },
  );
}
