import { createMovementKey } from "../../parse/pb";
import {
  absenceSpansBeforeEpisode,
  clipAbsenceSpansToActiveEpisode,
  encodeTimesheetAbsenceSpans,
  extractTimesheetDestinationFromPosition,
  journalDayFromDateMs,
  shouldUseArchiveReturnForStaffEpisode,
} from "../../../ejoosTimesheetText";
import { isOwnUnitStaffMove } from "../../../ejoosMovementRules";
import { positionCloseWritesExcluded } from "../../../ejoosExcludePolicy";
import type { PbMovement } from "../../types/pb";
import type { EjoosSyncOp } from "../../types/syncOp";
import {
  hasActualReturn,
  isPositionIndex,
  positionChangeDestination,
} from "../movementContext";
import { opId } from "../opId";

export type MovementPlanResult = { ops: EjoosSyncOp[]; handled: boolean };

export type PositionChangeMovementInput = {
  event: PbMovement;
  absenceSpansBeforeEpisode: unknown;
  alreadyVacatedForAbsence: unknown;
  archiveAll: unknown;
  arrivalOf: unknown;
  augustAbsenceSpansFor: unknown;
  byPersonName: unknown;
  cancelledTransferOf: unknown;
  chainedPositionRows: unknown;
  clipAbsenceSpansToActiveEpisode: unknown;
  createMovementKey: unknown;
  dateMs: unknown;
  dayById: unknown;
  dayByIndex: unknown;
  ejoosDays: unknown;
  ejoosOos: unknown;
  encodeTimesheetAbsenceSpans: unknown;
  extractTimesheetDestinationFromPosition: unknown;
  findLatestExcludedRow: unknown;
  findMovementExcludedRow: unknown;
  findNonStaffOccupantExcelRow: unknown;
  inboundStaffDateFor: unknown;
  isInternalStaffIndexHop: unknown;
  isOwnUnitStaffMove: unknown;
  isPositionChangeRow: (excelRow: number) => boolean;
  isPositionIndex: unknown;
  isSamePerson: unknown;
  isVacantStaffRow: unknown;
  journalDayFromDateMs: unknown;
  journalMonthStartLabel: unknown;
  latestPositionByName: unknown;
  leadWindowStart: unknown;
  movementPersonKey: unknown;
  normKey: unknown;
  oosById: unknown;
  oosByName: unknown;
  openAbsenceOf: unknown;
  opId: unknown;
  ownUnitIndexHistoryOf: unknown;
  ownUnitMoveSuperseded: unknown;
  positionChainByPerson: unknown;
  positionChangeDestination: unknown;
  positionCloseWritesExcluded: unknown;
  priorEpisodeTimesheetOf: unknown;
  shPersonById: unknown;
  shPersonByName: unknown;
  shpoByIndex: unknown;
  shpoPersonById: unknown;
  shpoPersonByName: unknown;
  shpoSheet: unknown;
  staffAppointmentDateFor: unknown;
  staffIndexTimesheetForPerson: unknown;
  timesheetEpisodeStartFor: unknown;
  timesheetNeedsTransferCancelSplit: unknown;
  timesheetPeople: unknown;
  timesheetRowsOf: unknown;
  timesheetSheet: unknown;
  transferCancelOf: unknown;
  wasMovementProcessed: unknown;
};

export const planPositionChangeMovementOps = (
  input: PositionChangeMovementInput,
): MovementPlanResult => {
  const ops: EjoosSyncOp[] = [];
  const event = input.event;
  if (!input.isPositionChangeRow(event.excelRow)) {
    return { ops, handled: false };
  }
    if (input.ownUnitMoveSuperseded(event)) return { ops, handled: true };
    const chained = input.chainedPositionRows.has(event.excelRow);
    const personChain = input.positionChainByPerson.get(input.movementPersonKey(event));
    // Ланцюг має пріоритет: інакше по одній особі вийшло б дві різні
    // операції зміни посади з різних рядків РУХ.
    if (personChain?.length && !chained) return { ops, handled: true };
    if (personChain?.length) {
      const last = personChain[personChain.length - 1];
      if (event.excelRow !== last.excelRow) return { ops, handled: true };
    }
    if (
      !chained &&
      input.latestPositionByName.get(input.normKey(event.fullName))?.excelRow !==
        event.excelRow
    ) {
      return { ops, handled: true };
    }
    const shPerson =
      (event.personId && input.shPersonById.get(event.personId)) ||
      input.byPersonName(input.shPersonByName, event.personId, event.fullName) ||
      null;
    // `sh` is authoritative for the current occupant and usually contains the
    // stable ID which may be absent in РУХ.
    const personId = shPerson?.personId || event.personId;
    const fullName = shPerson?.fullName || event.fullName;
    const rank = shPerson?.rank || event.rank;
    const arrival = input.arrivalOf(personId, fullName);
    // Відкрита відсутність не скасовує постановку з «Тимчасово прибулі».
    if (input.alreadyVacatedForAbsence(personId, fullName) && !arrival) {
      return { ops, handled: true };
    }
    // У кроці ланцюга особи вже може не бути в sh на цьому індексі
    // (далі за датою — розпорядження), тому цільовий індекс беремо з РУХ.
    const shIndex = input.isPositionIndex(shPerson?.positionIndex || "")
      ? shPerson!.positionIndex
      : "";
    const nextIndex = chained
      ? event.nextIndex || shIndex || event.previousIndex
      : shIndex || event.nextIndex || event.previousIndex;
    // Без коректного нового індексу писати нову посаду нікуди.
    if (!input.isPositionIndex(nextIndex)) return { ops, handled: true };
    const targetShpo = nextIndex
      ? (input.shpoByIndex.get(nextIndex) ?? null)
      : null;
    const indexTimesheet = nextIndex
      ? (input.dayByIndex.get(nextIndex) ?? null)
      : null;
    const existingOos =
      (personId && input.oosById.get(personId)) ||
      input.byPersonName(input.oosByName, personId, fullName) ||
      input.ejoosOos.find((row) => input.isSamePerson({ personId, fullName }, row)) ||
      null;
    const excludedSource = existingOos
      ? null
      : input.findLatestExcludedRow(personId, fullName);
    const transferCancel = input.transferCancelOf(event);
    const cancelledTransfer = input.cancelledTransferOf({ personId, fullName });
    const personTimesheetRows = input.timesheetRowsOf(personId, fullName);
    const personStaffTimesheet = input.staffIndexTimesheetForPerson(
      personId,
      fullName,
      nextIndex,
    );
    const personActiveTimesheet =
      [...personTimesheetRows]
        .filter((row) => !row.hasDepartureText)
        .sort(
          (left, right) => right.plusDays.length - left.plusDays.length,
        )[0] ?? null;
    const targetTimesheet = (() => {
      const vacantStaff = input.isVacantStaffRow(indexTimesheet)
        ? indexTimesheet
        : null;
      if (transferCancel) {
        if (personStaffTimesheet) return personStaffTimesheet;
        if (personActiveTimesheet) return personActiveTimesheet;
        if (vacantStaff) return vacantStaff;
        if (
          indexTimesheet &&
          input.isSamePerson({ personId, fullName }, indexTimesheet)
        ) {
          return indexTimesheet;
        }
        return indexTimesheet;
      }
      const staffScan = personStaffTimesheet
        ? input.timesheetPeople.find(
            (row) => row.excelRow === personStaffTimesheet.excelRow,
          )
        : undefined;
      // Реальне вибуття: історичний «вибув» не зафарбовуємо — новий епізод
      // пишемо на вакантний штатний рядок.
      if (vacantStaff && staffScan?.hasDepartureText) return vacantStaff;
      if (personStaffTimesheet && !staffScan?.hasDepartureText) {
        return personStaffTimesheet;
      }
      if (personActiveTimesheet) return personActiveTimesheet;
      if (vacantStaff) return vacantStaff;
      if (personStaffTimesheet) return personStaffTimesheet;
      if (
        indexTimesheet &&
        input.isSamePerson({ personId, fullName }, indexTimesheet)
      ) {
        return indexTimesheet;
      }
      return indexTimesheet;
    })();
    const actualExclusionEvidence = Boolean(
      cancelledTransfer && input.findMovementExcludedRow(cancelledTransfer),
    );
    const historyTimesheet =
      personTimesheetRows.find(
        (row) =>
          row.hasDepartureText && row.excelRow !== targetTimesheet?.excelRow,
      ) ?? null;
    const currentTimesheet =
      historyTimesheet ||
      (personId && input.dayById.get(personId)) ||
      input.ejoosDays.find((row) => input.isSamePerson({ personId, fullName }, row)) ||
      null;
    const samePersonRow = (
      row: { personId: string; fullName: string } | null | undefined,
    ) => Boolean(row && input.isSamePerson({ personId, fullName }, row));
    const previousShpo = event.previousIndex
      ? (input.shpoByIndex.get(event.previousIndex) ?? null)
      : null;
    const previousIndexTimesheet = event.previousIndex
      ? (input.dayByIndex.get(event.previousIndex) ?? null)
      : null;
    // Рядок особи в ШПО може бути без індексу (лише ПІБ) — його теж треба
    // звільнити, але цільовий рядок нової посади не чіпаємо.
    const personShpoRow =
      (personId && input.shpoPersonById.get(personId)) ||
      input.byPersonName(input.shpoPersonByName, personId, fullName) ||
      null;
    const oldShpo = samePersonRow(previousShpo)
      ? previousShpo
      : personShpoRow && personShpoRow.excelRow !== targetShpo?.excelRow
        ? personShpoRow
        : null;
    const previousShpoTimesheet = samePersonRow(previousIndexTimesheet)
      ? previousIndexTimesheet
      : null;
    // Стару штатну посаду звільняємо в ШПО/Табелі. У «Виключені» рядок
    // пишемо лише коли це не внутрішній стрибок 1ПБ→1ПБ.
    const closeOldPosition = Boolean(
      event.previousIndex &&
      event.previousIndex !== nextIndex &&
      (oldShpo || previousShpoTimesheet),
    );
    const internalStaffHop = input.isInternalStaffIndexHop(event);
    const writeExcludedOnClose = input.positionCloseWritesExcluded({
      closeOldPosition: closeOldPosition ? "1" : "",
      internalStaffHop: internalStaffHop ? "1" : "",
    });
    const returningToStaffIndex = Boolean(
      shPerson &&
      nextIndex === shPerson.positionIndex &&
      historyTimesheet?.hasDepartureText &&
      input.isOwnUnitStaffMove(event) &&
      !samePersonRow(targetShpo) &&
      !transferCancel,
    );
    // Якщо особа вже стоїть на цільовому індексі в ЕЖООС і стару посаду
    // закривати не треба — у РУХ лише історія. Внутрішній стрибок теж
    // не вигадує «Виключені»: залишки старого Табеля чистить детектор дублів.
    if (
      samePersonRow(targetShpo) &&
      !returningToStaffIndex &&
      (!closeOldPosition || internalStaffHop)
    ) {
      const arrivalStillOpen = Boolean(arrival);
      const oosMissing = !existingOos;
      if (
        !arrivalStillOpen &&
        !oosMissing &&
        (!transferCancel ||
          !input.timesheetNeedsTransferCancelSplit(
            personId,
            fullName,
            nextIndex,
            transferCancel.orderDate,
          ))
      ) {
        return { ops, handled: true };
      }
    }
    // Старий рядок «Тимчасово прибулі» лишається історією. Якщо особа вже
    // на штатній посаді, це не постановка з тимчасового прибуття.
    const isTempArrivalPlacement = Boolean(arrival) && !closeOldPosition;
    const openAbsence = input.openAbsenceOf(personId, fullName);
    const dispositionShpoExcelRow = input.findNonStaffOccupantExcelRow(
      input.shpoSheet,
      personId,
      fullName,
      { index: 0, name: 6, id: 7 },
    );
    const dispositionTimesheetExcelRow = input.findNonStaffOccupantExcelRow(
      input.timesheetSheet,
      personId,
      fullName,
      { index: 1, name: 6, id: 7 },
    );
    const returningFromDisposition = Boolean(
      dispositionShpoExcelRow ||
      dispositionTimesheetExcelRow ||
      openAbsence ||
      /розпорядж/iu.test(
        `${personChain?.[0]?.previousIndex || ""} ${event.previousIndex}`,
      ),
    );
    const indexHistory = (() => {
      const fromChain = (personChain?.length ? personChain : [event])
        .filter((item) => input.isPositionIndex(item.nextIndex))
        .map((item) => ({
          index: item.nextIndex,
          date: item.orderDate || item.basisDate || event.orderDate,
        }));
      if (fromChain.length > 1) return fromChain;
      const extra = input.ownUnitIndexHistoryOf(personId, fullName);
      return extra.length ? extra : fromChain;
    })();
    const newestFirstHistory = [...indexHistory].reverse();
    const oosHistoryIndexes = newestFirstHistory
      .map((item) => item.index)
      .filter((index, idx, all) => all.indexOf(index) === idx)
      .join("\n");
    const oosHistoryDates = newestFirstHistory
      .filter(
        (item, idx, all) =>
          all.findIndex((other) => other.index === item.index) === idx,
      )
      .map((item) => item.date)
      .join("\n");
    const staffTimesheetFrom =
      input.timesheetEpisodeStartFor(personId, fullName, nextIndex) ||
      input.inboundStaffDateFor(personId, fullName) ||
      event.orderDate;
    const monthSpans = input.augustAbsenceSpansFor(personId, fullName);
    const latestArchiveReturn = returningFromDisposition
      ? [...input.archiveAll]
          .filter((period) => input.isSamePerson({ personId, fullName }, period))
          .filter((period) => hasActualReturn(period.returnDate))
          .sort(
            (left, right) =>
              input.dateMs(right.returnDate) - input.dateMs(left.returnDate),
          )[0]
      : null;
    const carryAbsenceFromMonthStart =
      !returningFromDisposition &&
      monthSpans.some((span) => span.fromDay === 1);
    const archiveReturnDate =
      latestArchiveReturn?.returnDate &&
      hasActualReturn(latestArchiveReturn.returnDate)
        ? latestArchiveReturn.returnDate
        : "";
    const useArchiveReturnForEpisode =
      returningFromDisposition &&
      archiveReturnDate &&
      shouldUseArchiveReturnForStaffEpisode({
        archiveReturnDate,
        monthStartMs: input.leadWindowStart,
        openAbsenceGround: openAbsence?.ground || "",
        monthSpans,
        hasReturn: hasActualReturn,
      });
    const timesheetActiveFrom = returningFromDisposition
      ? (useArchiveReturnForEpisode ? archiveReturnDate : "") ||
        event.orderDate ||
        archiveReturnDate ||
        (openAbsence?.actualReturn &&
        hasActualReturn(openAbsence.actualReturn)
          ? openAbsence.actualReturn
          : "")
      : carryAbsenceFromMonthStart
        ? input.journalMonthStartLabel
        : returningToStaffIndex
          ? event.orderDate
          : staffTimesheetFrom;
    const timesheetPreserveHistory =
      transferCancel ||
      carryAbsenceFromMonthStart ||
      returningFromDisposition ||
      input.journalDayFromDateMs(input.dateMs(timesheetActiveFrom), input.leadWindowStart) <= 1
        ? ""
        : "1";
    const staleTimesheet =
      input.timesheetRowsOf(personId, fullName).find(
        (row) =>
          row.excelRow !== targetTimesheet?.excelRow && !row.hasDepartureText,
      ) ?? null;
    const newPositionText = input.positionChangeDestination(event);
    const timesheetDestination =
      input.extractTimesheetDestinationFromPosition(newPositionText) ||
      newPositionText;
    const chainIndexes = indexHistory
      .map((item) => item.index)
      .filter(Boolean);
    const chainNote =
      chainIndexes.length > 1 ? `Ланцюг ${chainIndexes.join(" → ")}: ` : "";
    const canApply = Boolean(
      (shPerson || chained || returningFromDisposition) &&
      nextIndex &&
      targetShpo &&
      targetTimesheet,
    );
    const applyNow = canApply;
    const processedBefore = input.wasMovementProcessed(event);

    ops.push({
      id: input.opId(["pos", event.movementNumber || String(event.excelRow)]),
      kind: "position_change",
      class: applyNow ? "ready" : "needs_input",
      sheet: returningFromDisposition
        ? "5. Тимч. відсутні → 1. ШПО / 2. ООС / 6. Табель"
        : writeExcludedOnClose
          ? "3. Виключені → 6. Табель → 1. ШПО / 2. ООС"
          : closeOldPosition
            ? "1. ШПО / 2. ООС / 6. Табель"
            : isTempArrivalPlacement
              ? "4. Тимч. прибулі → 1. ШПО / 2. ООС / 6. Табель"
              : "1. ШПО / 2. ООС / 6. Табель",
      personId,
      fullName,
      positionIndex: nextIndex,
      rank,
      before: returningFromDisposition
        ? `у розпорядженні / ${openAbsence?.ground || "СЗЧ"}`
        : closeOldPosition
          ? `штатна посада ${event.previousIndex}`
          : isTempArrivalPlacement
            ? `тимчасово прибулий${arrival?.positionIndex ? ` · інд. ${arrival.positionIndex}` : ""}`
            : event.previousIndex || "?",
      after: `штатна посада ${nextIndex || event.changeText || "?"}${
        oosHistoryIndexes && oosHistoryIndexes.includes("\n")
          ? ` · ООС ${oosHistoryIndexes.replaceAll("\n", " → ")}`
          : ""
      }`,
      sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
      why:
        chainNote +
        (processedBefore
          ? "Подія є в історії застосувань, але фактичний стан ШПО/ООС/Табеля потребує повторного проведення"
          : !canApply
            ? "Не знайдено однозначний рядок нового індексу в ШПО/Табелі — потрібна ручна перевірка"
            : returningFromDisposition
              ? `Повернення із СЗЧ / розпорядження: закрити період, прибрати з блоку «у розпорядженні», поставити на ${nextIndex}${
                  chainIndexes.length > 1
                    ? `, історія ООС ${[...chainIndexes].reverse().join(", ")}`
                    : ""
                }`
              : closeOldPosition
                ? writeExcludedOnClose
                  ? `зміна посади в межах 1ПБ: закрити ${event.previousIndex} у Виключених і Табелі, поставити на ${nextIndex}, ООС лишити активним`
                  : `внутрішня зміна посади 1ПБ ${event.previousIndex} → ${nextIndex}: оновити ШПО/ООС/Табель, у «Виключені» не пишемо`
                : isTempArrivalPlacement
                  ? `ПОСАДА №${event.orderNumber || "?"} від ${event.orderDate || "?"}: закрити тимчасове прибуття, поставити на штат ${nextIndex}; Табель з ${staffTimesheetFrom || event.orderDate || "дати наказу"}, коди archive в цьому ж кроці`
                  : transferCancel
                    ? `Серпневий ланцюг: ПОСАДА ${nextIndex} з ${input.staffAppointmentDateFor(personId, fullName, nextIndex) || event.orderDate || "?"}; ПЕРЕВ №${cancelledTransfer?.orderNumber || "?"} скасовано №${transferCancel.orderNumber || "?"} — один рядок Табеля, запис у «Виключені» прибираємо`
                    : returningToStaffIndex
                      ? `Повернення на ${nextIndex} з ${event.orderDate || "?"}: історичний рядок з вибуттям лишаємо, новий активний з «+» від дати наказу`
                      : event.type === "ПЕРЕВ"
                        ? "Внутрішній ПЕРЕВ у межах 1ПБ підтверджений поточним sh: змінити штатну позицію, ООС і Табель"
                        : "ПОСАДА підтверджена поточним sh: оновити штатну позицію, ООС і Табель"),
      confidence: applyNow ? "high" : "manual",
      payload: {
        movementNumber: event.movementNumber,
        previousIndex: event.previousIndex,
        nextIndex,
        changeText: event.changeText,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
        basisNumber: event.basisNumber,
        basisDate: event.basisDate,
        nextName: fullName,
        nextRank: rank,
        nextPersonId: personId,
        positionTitle: shPerson?.positionTitle || "",
        statusRaw: shPerson?.status || event.status,
        isTempArrivalPlacement: isTempArrivalPlacement ? "1" : "",
        arrivalExcelRow: arrival ? String(arrival.excelRow) : "",
        oosExcelRow: existingOos ? String(existingOos.excelRow) : "",
        oosHistoryIndexes,
        oosHistoryDates,
        shpoExcelRow: targetShpo ? String(targetShpo.excelRow) : "",
        timesheetExcelRow: targetTimesheet
          ? String(targetTimesheet.excelRow)
          : "",
        previousTimesheetExcelRow: transferCancel
          ? ""
          : returningToStaffIndex && historyTimesheet
            ? String(historyTimesheet.excelRow)
            : currentTimesheet &&
                currentTimesheet.excelRow !== targetTimesheet?.excelRow
              ? String(currentTimesheet.excelRow)
              : "",
        clearTimesheetExcelRow: staleTimesheet
          ? String(staleTimesheet.excelRow)
          : "",
        timesheetActiveFrom,
        timesheetSkipHistory: "1",
        timesheetPreserveHistory,
        timesheetBindStaffIndex: isTempArrivalPlacement ? nextIndex : "",
        returningToStaffIndex: returningToStaffIndex ? "1" : "",
        returningFromDisposition: returningFromDisposition ? "1" : "",
        dispositionShpoExcelRow: dispositionShpoExcelRow
          ? String(dispositionShpoExcelRow)
          : "",
        dispositionTimesheetExcelRow: dispositionTimesheetExcelRow
          ? String(dispositionTimesheetExcelRow)
          : "",
        openAbsenceExcelRow: openAbsence ? String(openAbsence.excelRow) : "",
        timesheetAbsenceSpans: input.encodeTimesheetAbsenceSpans(
          (() => {
            const activeDay =
              input.journalDayFromDateMs(
                input.dateMs(timesheetActiveFrom),
                input.leadWindowStart,
              ) || 1;
            if (carryAbsenceFromMonthStart) return monthSpans;
            return input.clipAbsenceSpansToActiveEpisode(monthSpans, activeDay);
          })(),
        ),
        historyTimesheetExcelRow: (() => {
          if (transferCancel || carryAbsenceFromMonthStart) return "";
          const fromDay = input.journalDayFromDateMs(
            input.dateMs(timesheetActiveFrom),
            input.leadWindowStart,
          );
          if (fromDay <= 1) return "";
          const prior = input.priorEpisodeTimesheetOf(
            personId,
            fullName,
            targetTimesheet?.excelRow || 0,
            nextIndex,
          );
          return prior ? String(prior.excelRow) : "";
        })(),
        historyTimesheetAbsenceSpans: input.encodeTimesheetAbsenceSpans(
          carryAbsenceFromMonthStart
            ? []
            : input.absenceSpansBeforeEpisode(
                monthSpans,
                input.journalDayFromDateMs(
                  input.dateMs(timesheetActiveFrom),
                  input.leadWindowStart,
                ) || 1,
              ),
        ),
        arrivedFrom: event.arrivedFrom || arrival?.fromUnit || "",
        arrivalDepartDate: arrival ? event.orderDate : "",
        arrivalDepartOrderNumber: arrival ? event.orderNumber : "",
        arrivalDepartOrderDate: arrival ? event.orderDate : "",
        cancelledTransferOrder: cancelledTransfer?.orderNumber || "",
        cancelledTransferDate: cancelledTransfer?.orderDate || "",
        cancelledTransferDest:
          cancelledTransfer?.destination ||
          cancelledTransfer?.changeText ||
          "",
        actualExclusionEvidence: actualExclusionEvidence ? "1" : "",
        chainStep: personChain?.length ? String(personChain.length) : "",
        chainTotal: personChain?.length ? String(personChain.length) : "",
        closeOldPosition: closeOldPosition ? "1" : "",
        internalStaffHop: internalStaffHop ? "1" : "",
        previousShpoExcelRow:
          closeOldPosition && oldShpo ? String(oldShpo.excelRow) : "",
        previousIndexTimesheetExcelRow:
          closeOldPosition && previousShpoTimesheet
            ? String(previousShpoTimesheet.excelRow)
            : "",
        excludeDate: writeExcludedOnClose ? event.orderDate : "",
        documentsDest: writeExcludedOnClose
          ? event.changeText || newPositionText
          : "",
        timesheetDestination: closeOldPosition ? timesheetDestination : "",
        exclusionReason: writeExcludedOnClose ? "ПЕРЕВЕДЕННЯ 1 ПБ" : "",
        fromRank: oldShpo?.rank || previousShpoTimesheet?.rank || rank,
        fromName: oldShpo?.fullName || fullName,
        fromPersonId: oldShpo?.personId || personId,
        fromPositionIndex: event.previousIndex,
        excludedSourceExcelRow: excludedSource
          ? String(excludedSource.excelRow)
          : "",
        transferCancelOrder: transferCancel?.orderNumber || "",
        transferCancelDate: transferCancel?.orderDate || "",
      },
      movementKey: input.createMovementKey(event),
      checkedDefault: applyNow,
    });
    return { ops, handled: true };
  return { ops, handled: true };
};
