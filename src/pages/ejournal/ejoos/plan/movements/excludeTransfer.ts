import { createMovementKey, parseRankPromotion } from "../../parse/pb";
import { canonicalName, dateMs, norm } from "../../parse/cellText";
import {
  absenceSpansBeforeEpisode,
  encodeTimesheetAbsenceSpans,
  extractTimesheetDestinationFromPosition,
  formatDispositionTimesheetDeparture,
  journalDayFromDateMs,
} from "../../../ejoosTimesheetText";
import {
  isDispositionAbsenceStatus,
  isOwnUnitStaffMove,
  isOutboundStaffMove,
} from "../../../ejoosMovementRules";
import {
  absenceOnlyBlocksExclusion,
  excludedTimesheetWrite,
  externalTransferProcessState,
  isUnrecordedSameMonthTransit,
  laterReturnSupersedesOutbound,
  skipExternalIfAlreadyProcessed,
} from "../../../ejoosExcludePolicy";
import { timesheetRowInExpectedUnitSection } from "../../../ejoosTimesheetUnitSections";
import type { PbMovement } from "../../types/pb";
import type { EjoosSyncOp } from "../../types/syncOp";
import {
  formatTransferDestinationForTimesheetMark as formatTransferDestinationForTimesheet,
  hasActualReturn,
  isPositionIndex,
  positionChangeDestination,
  unitCodeFromMovement,
} from "../movementContext";
import { opId } from "../opId";

export type MovementPlanResult = { ops: EjoosSyncOp[]; handled: boolean };

export type ExcludeTransferMovementInput = {
  event: PbMovement;
  existingOps: EjoosSyncOp[];
  absenceOnlyBlocksExclusion: unknown;
  archiveAll: unknown;
  arrivalOf: unknown;
  byPersonName: unknown;
  cancelledExternalTransferRows: unknown;
  createMovementKey: unknown;
  dayByIndex: unknown;
  ejoosShpo: unknown;
  excludedTimesheetWrite: unknown;
  externalTransferProcessState: unknown;
  extractTimesheetDestinationFromPosition: unknown;
  findMovementExcludedRow: unknown;
  formatTransferDestinationForTimesheet: unknown;
  inboundStaffDateFor: unknown;
  inboundStaffPlacementBefore: unknown;
  internalPositionMovementRows: unknown;
  isOutboundStaffMove: unknown;
  isOwnFirstPbDestination: unknown;
  isPositionChangeRow: unknown;
  isPositionIndex: unknown;
  isSamePerson: unknown;
  isUnrecordedSameMonthTransit: unknown;
  isVacantStaffRow: unknown;
  latestRankEventByPerson: unknown;
  movementEventTime: unknown;
  movementPersonKey: unknown;
  onStaffOos: unknown;
  onStaffShpo: unknown;
  oosPersonById: unknown;
  oosPersonByName: unknown;
  openDispositionAbsenceMs: unknown;
  opId: unknown;
  parseRankPromotion: unknown;
  pendingRankByPerson: unknown;
  personIdFromShpo: unknown;
  personStillInEjoos: unknown;
  personStillInSh: unknown;
  positionChangeDestination: unknown;
  shpoByIndex: unknown;
  skipExternalIfAlreadyProcessed: unknown;
  sourcePositionTitleForOutbound: unknown;
  staffEpisodePaintPayload: unknown;
  staleExcludedForMovement: unknown;
  staleExcludedClearPayload: (
    stale: Array<{ excelRow: number }>,
  ) => Record<string, string>;
  timesheetClosedFor: unknown;
  timesheetRowInExpectedUnitSection: unknown;
  timesheetRowsOf: unknown;
  timesheetSheet: unknown;
  transferCancelOf: unknown;
  unitCodeFromMovement: unknown;
  wasMovementProcessed: unknown;
  laterReturnSupersedesOutbound: unknown;
};

export const planExcludeTransferMovementOps = (
  input: ExcludeTransferMovementInput,
): MovementPlanResult => {
  const ops: EjoosSyncOp[] = [];
  const event = input.event;
  const isExternalUnitDeparture =
    input.isOutboundStaffMove(event) &&
    !(
      event.type === "ПЕРЕВ" &&
      input.internalPositionMovementRows.has(event.excelRow)
    ) &&
    !(event.type === "ПОСАДА" && input.isPositionChangeRow(event.excelRow));
  if (!isExternalUnitDeparture) {
    return { ops, handled: false };
  }
  const existingExcludedEarly = input.findMovementExcludedRow(event);
    const sourcePositionTitle = input.sourcePositionTitleForOutbound(event);
    if (input.wasMovementProcessed(event) && existingExcludedEarly) {
      if (
        input.timesheetClosedFor(
          event.personId,
          event.fullName,
          sourcePositionTitle,
        )
      ) {
        return { ops, handled: true };
      }
    }
    const inboundPlacement = input.inboundStaffPlacementBefore(event);
    const unrecordedSameMonthTransit = input.isUnrecordedSameMonthTransit({
      hasInboundPlacement: Boolean(inboundPlacement),
      alreadyExcluded: Boolean(existingExcludedEarly),
      stillInSh: input.personStillInSh(event.personId, event.fullName),
      stillInEjoos: input.personStillInEjoos(event.personId, event.fullName),
    });
    const transferProcessState = input.externalTransferProcessState({
      onStaffShpo: input.onStaffShpo(event.personId, event.fullName),
      onStaffOos: input.onStaffOos(event.personId, event.fullName),
      hasMatchingExcluded: Boolean(existingExcludedEarly),
      timesheetClosed: input.timesheetClosedFor(
        event.personId,
        event.fullName,
        sourcePositionTitle,
      ),
    });
    if (
      input.skipExternalIfAlreadyProcessed({
        stillInEjoos: input.personStillInEjoos(event.personId, event.fullName),
        unrecordedTransit: unrecordedSameMonthTransit,
        processState: transferProcessState,
      })
    ) {
      return { ops, handled: true };
    }
    // Скасоване переведення не є чинним: новий рядок у Виключені не пишемо,
    // навіть якщо історичний запис від того наказу вже є.
    if (input.cancelledExternalTransferRows.has(event.excelRow)) return;
    {
      const cancel = input.transferCancelOf(event);
      if (cancel && input.movementEventTime(cancel) >= input.movementEventTime(event)) {
        return { ops, handled: true };
      }
    }
    // Вибуття підтверджуємо не лише рядком РУХ: людини вже не повинно бути
    // в актуальному sh. Повернення 07.08 після вибуття 05.08 теж скасовує
    // виключення — Атрахов знову чинний occupant.
    if (
      input.laterReturnSupersedesOutbound({
        stillInSh: input.personStillInSh(event.personId, event.fullName),
        returnedAfterOutbound: Boolean(
          input.inboundStaffDateFor(event.personId, event.fullName),
        ),
      }) ||
      input.absenceOnlyBlocksExclusion({
        absenceAt: input.openDispositionAbsenceMs(event.personId, event.fullName),
        outboundAt: input.movementEventTime(event),
      })
    ) {
      return { ops, handled: true };
    }
    const existingExcluded = existingExcludedEarly;
    const alreadyExcluded = Boolean(existingExcluded);
    const destinationRaw = event.destination || "";
    const destinationUpper = destinationRaw.toUpperCase();
    const rawDestination =
      !destinationRaw ||
      destinationUpper.includes("РОЗПОР") ||
      destinationUpper.includes("ПЕРЕВ") ||
      destinationUpper === event.type ||
      input.isOwnFirstPbDestination(destinationRaw)
        ? ""
        : destinationRaw;
    // Якщо «Куди» = 0 або лишилось 1ПБ, військова частина часто в S / примітці (А7400).
    const unitFromS = input.unitCodeFromMovement(event);
    const exclusionPlace =
      rawDestination ||
      (/[АA]\s*\d{4}/iu.test(unitFromS) ? unitFromS : "") ||
      event.note ||
      event.changeText;
    // Табель: «вибув до/у» + підрозділ з «Яка зміна» без назви посади.
    // Код в/ч А#### лишається для Виключених (AE / підстава), не для Табеля.
    const timesheetDestination = (() => {
      const positionSource =
        input.positionChangeDestination(event) || event.changeText;
      const fromPosition =
        input.extractTimesheetDestinationFromPosition(positionSource);
      if (fromPosition) return fromPosition;
      const unitPhrase =
        input.formatTransferDestinationForTimesheet(
          [rawDestination, event.note, unitFromS, event.changeText]
            .filter(Boolean)
            .join(" "),
        ) || rawDestination;
      if (
        /[АA]\s*\d{4}/iu.test(unitPhrase) ||
        /в\s*\/\s*ч/iu.test(unitPhrase)
      ) {
        return event.note && /[АA]\s*\d{4}|в\s*\/\s*ч/iu.test(event.note)
          ? event.note
          : unitPhrase;
      }
      return unitPhrase;
    })();
    const exclusionReason =
      event.type === "ПЕРЕВ" || event.type === "ПОСАДА"
        ? "ПЕРЕВЕДЕННЯ"
        : event.type === "ЗВІЛЬН"
          ? "ЗВІЛЬНЕННЯ"
          : "Розпорядження";
    const shpoAtOldIndex =
      event.previousIndex || inboundPlacement?.nextIndex
        ? (input.shpoByIndex.get(
            event.previousIndex || inboundPlacement?.nextIndex || "",
          ) ?? null)
        : null;
    // Індекс після вибуття міг зайняти хтось інший (Хубаєв → Атрахов на
    // 2103764). ШПО/Табель чіпаємо лише якщо рядок цієї самої особи.
    const shpo =
      (event.personId &&
        input.ejoosShpo.find((row) => row.personId === event.personId)) ||
      input.ejoosShpo.find((row) => input.isSamePerson(event, row)) ||
      (shpoAtOldIndex && input.isSamePerson(event, shpoAtOldIndex)
        ? shpoAtOldIndex
        : null);
    const oos =
      (event.personId && input.oosPersonById.get(event.personId)) ||
      input.byPersonName(input.oosPersonByName, event.personId, event.fullName) ||
      null;
    const staffIndexLeft =
      (input.isPositionIndex(event.previousIndex) && event.previousIndex) ||
      (input.isPositionIndex(inboundPlacement?.nextIndex || "") &&
        inboundPlacement!.nextIndex) ||
      "";
    // Для вибуття у «Виключені» базові персональні дані мають іти з ЕЖООС.
    // Якщо в серпні було ЗВАННЯ раніше за цей ПЕРЕВ — беремо нове звання,
    // навіть якщо картка ООС/ШПО ще не оновлена.
    const rankBeforeTransfer = [...input.latestRankEventByPerson.values()].find(
      (rankEvent) =>
        input.isSamePerson(event, rankEvent) &&
        input.movementEventTime(rankEvent) <= input.movementEventTime(event),
    );
    const pendingRank =
      input.pendingRankByPerson.get(input.movementPersonKey(event)) ||
      (event.personId
        ? input.pendingRankByPerson.get(`id:${event.personId}`)
        : null);
    const promotedRank =
      pendingRank?.payload.nextRank ||
      (rankBeforeTransfer
        ? input.parseRankPromotion(rankBeforeTransfer).nextRank
        : "");
    const fromRank =
      promotedRank || oos?.rank || shpo?.rank || event.rank || "";
    const fromName = oos?.fullName || shpo?.fullName || event.fullName || "";
    const occupiedIndex =
      shpo?.positionIndex ||
      String(oos?.positionIndex || "").match(/\d{5,}/)?.[0] ||
      oos?.positionIndex ||
      "";
    // У «Виключені» пишемо посаду, з якої вибув із 1ПБ (РУХ), навіть якщо
    // внутрішню ПОСАДУ в ЕЖООС ще не провели.
    const fromIndex = staffIndexLeft || occupiedIndex;
    const fromId = input.personIdFromShpo(input.ejoosShpo, {
      fullName: fromName,
      positionIndex: occupiedIndex || shpo?.positionIndex || "",
      personId: oos?.personId || event.personId || shpo?.personId,
    });
    const arrival =
      input.arrivalOf(fromId, fromName) ||
      input.arrivalOf(event.personId, event.fullName);
    const namedTimesheetRows = [
      ...input.timesheetRowsOf(fromId, fromName),
      ...input.timesheetRowsOf(event.personId, event.fullName),
    ].filter(
      (row, index, rows) =>
        rows.findIndex((other) => other.excelRow === row.excelRow) === index,
    );
    const timesheetWrite = (() => {
      const base = input.excludedTimesheetWrite(
        namedTimesheetRows,
        fromIndex && input.isVacantStaffRow(input.dayByIndex.get(fromIndex))
          ? input.dayByIndex.get(fromIndex)?.excelRow || 0
          : 0,
      );
      const historyRow = namedTimesheetRows.find(
        (row) => row.hasDepartureText,
      );
      if (
        base.replaceInPlace &&
        historyRow &&
        input.timesheetSheet &&
        sourcePositionTitle &&
        !input.timesheetRowInExpectedUnitSection(
          input.timesheetSheet,
          historyRow.excelRow,
          sourcePositionTitle,
        )
      ) {
        return {
          createHistory: false,
          replaceInPlace: false,
          sourceExcelRow: historyRow.excelRow,
        };
      }
      return base;
    })();
    // Колонки «наказ» і перша «дата» — стройовий наказ. Окремі реквізити
    // підстави (наприклад 668-РС від 03.08.2026) сюди не підставляємо.
    const excludeDate = event.orderDate;
    const hasRequiredExcludedFields = Boolean(
      exclusionPlace &&
      exclusionReason &&
      excludeDate &&
      event.orderNumber &&
      event.orderDate,
    );
    const staleExcluded = input.staleExcludedForMovement(event);
    const staleClear = input.staleExcludedClearPayload(staleExcluded);
    const staleClearNote = staleExcluded.length
      ? ` Попередн${staleExcluded.length === 1 ? "ій рядок" : "і рядки"} Виключені R${staleExcluded.map((row) => row.excelRow).join(", R")} прибираємо — лишаємо чинне ПЕРЕВ.`
      : "";
    const excludeEpisodePaint =
      fromIndex && timesheetWrite.sourceExcelRow
        ? input.staffEpisodePaintPayload(
            fromId || event.personId,
            fromName,
            fromIndex,
            timesheetWrite.sourceExcelRow,
          )
        : null;

    ops.push({
      id: input.opId([
        "excl",
        event.movementNumber || String(event.excelRow),
        event.type,
      ]),
      kind: "exclude_transfer",
      class:
        hasRequiredExcludedFields && Boolean(shpo || fromName)
          ? "ready"
          : "needs_input",
      sheet: "Виключені → Табель → ШПО/ООС",
      personId: fromId || event.personId,
      fullName: fromName,
      positionIndex: fromIndex,
      rank: fromRank,
      before: shpo
        ? occupiedIndex && occupiedIndex !== fromIndex
          ? `ШПО R${shpo.excelRow}: ${fromRank} ${fromName} · інд. ${occupiedIndex} (кінцева посада РУХ ${fromIndex})`
          : `ШПО R${shpo.excelRow}: ${fromRank} ${fromName} · інд. ${fromIndex}`
        : unrecordedSameMonthTransit
          ? `немає в ШПО/ООС · транзит ${inboundPlacement?.orderDate || "?"} → ${event.orderDate || "?"}`
          : "в обліку (ШПО не знайдено — перевірте вручну)",
      after: `${alreadyExcluded ? "доробити очищення після виключення" : "виключити"}: ${event.type} → ${exclusionPlace || "(куди?)"} · табель: ${timesheetDestination || "(куди?)"} · ${exclusionReason} · наказ №${event.orderNumber || "?"} від ${event.orderDate || "?"}`,
      sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
      why: alreadyExcluded
        ? `Рядок у Виключених уже є — не дублюємо його, доробляємо Табель → ШПО/ООС. Тимч. відсутні/прибулі не чіпаємо.${staleClearNote}`
        : unrecordedSameMonthTransit
          ? `Постановка ${inboundPlacement?.orderDate || "?"} на ${fromIndex || "штабний індекс"} і вибуття ${event.orderDate || "?"} у цьому місяці, але в ЕЖООС особи не було. Пишемо Виключені + історичний Табель; зайнятий індекс ШПО/ООС не чіпаємо.`
          : hasRequiredExcludedFields
            ? pendingRank
              ? `Спочатку звання ${pendingRank.payload.nextRank} (№${pendingRank.payload.orderNumber || "?"} від ${pendingRank.payload.orderDate || "?"}), потім ${event.type} → ${exclusionPlace}: Виключені → Табель → очистка ШПО/ООС`
              : inboundPlacement
                ? `Ланцюг: ${inboundPlacement.orderDate || "?"} ${inboundPlacement.type} ${inboundPlacement.previousIndex || "?"} → ${inboundPlacement.nextIndex || fromIndex} (лишився б у ООС); далі ${event.type} №${event.orderNumber || "?"} від ${event.orderDate || "?"} — вибув з 1ПБ, в sh немає. Виключені з інд. ${fromIndex}; внутрішню постановку на штат не проводимо.`
                : `${event.type} з «куди»: алгоритм Виключені → Табель (історія) → очистка ШПО/ООС. Відкритий рядок «Тимч. прибулі» закриваємо.${staleClearNote}`
            : "У РУХ не вистачає фактичного місця вибуття або реквізитів стройового наказу",
      confidence:
        hasRequiredExcludedFields && (shpo || unrecordedSameMonthTransit)
          ? "high"
          : "manual",
      payload: {
        movementNumber: event.movementNumber,
        type: event.type,
        destination: exclusionPlace,
        timesheetDestination,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
        excludeDate,
        timesheetActiveFrom:
          excludeEpisodePaint?.timesheetActiveFrom ||
          (unrecordedSameMonthTransit || arrival
            ? inboundPlacement?.orderDate || ""
            : ""),
        timesheetAbsenceSpans: excludeEpisodePaint?.timesheetAbsenceSpans || "",
        timesheetPreserveHistory:
          excludeEpisodePaint?.timesheetPreserveHistory || "",
        historyTimesheetExcelRow:
          excludeEpisodePaint?.historyTimesheetExcelRow || "",
        historyTimesheetAbsenceSpans:
          excludeEpisodePaint?.historyTimesheetAbsenceSpans || "",
        timesheetCreateHistory: timesheetWrite.createHistory ? "1" : "",
        timesheetReplaceInPlace: timesheetWrite.replaceInPlace ? "1" : "",
        transitSameMonth: unrecordedSameMonthTransit ? "1" : "",
        basisNumber: event.basisNumber,
        basisDate: event.basisDate,
        previousIndex: staffIndexLeft || event.previousIndex,
        nextIndex: event.nextIndex,
        occupiedPositionIndex: occupiedIndex,
        priorPlacementRow: inboundPlacement
          ? String(inboundPlacement.excelRow)
          : "",
        priorPlacementType: inboundPlacement?.type || "",
        priorPlacementDate: inboundPlacement?.orderDate || "",
        priorPlacementOrder: inboundPlacement?.orderNumber || "",
        priorPlacementFromIndex: inboundPlacement?.previousIndex || "",
        priorPlacementToIndex: inboundPlacement?.nextIndex || "",
        arrivedFrom:
          inboundPlacement?.previousIndex ||
          inboundPlacement?.destination ||
          "",
        appointmentOrderNumber: inboundPlacement?.orderNumber || "",
        appointmentOrderDate: inboundPlacement?.orderDate || "",
        arrivalExcelRow: arrival ? String(arrival.excelRow) : "",
        arrivalDepartDate:
          inboundPlacement?.orderDate || event.orderDate || "",
        arrivalDepartOrderNumber:
          inboundPlacement?.orderNumber || event.orderNumber || "",
        arrivalDepartOrderDate:
          inboundPlacement?.orderDate || event.orderDate || "",
        excludedExcelRow: existingExcluded
          ? String(existingExcluded.excelRow)
          : "",
        oosExcelRow: oos ? String(oos.excelRow) : "",
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        timesheetExcelRow: timesheetWrite.sourceExcelRow
          ? String(timesheetWrite.sourceExcelRow)
          : "",
        fromRank,
        fromName,
        fromPersonId: fromId || event.personId,
        fromPositionIndex: fromIndex,
        lastRankOrderNumber:
          pendingRank?.payload.orderNumber ||
          rankBeforeTransfer?.orderNumber ||
          "",
        lastRankOrderDate:
          pendingRank?.payload.orderDate ||
          rankBeforeTransfer?.orderDate ||
          "",
        // AE «Куди вибув / Куди направлені документи»: дослівно з РУХ
        // «Яка зміна». Підрозділ/вч з «Куди» лишається для Табеля/підстави.
        documentsDest: event.changeText || "",
        changeText: event.changeText || "",
        positionTitle: sourcePositionTitle || "",
        exclusionReason,
        awaitRankChange: pendingRank ? "1" : "",
        ...staleClear,
      },
      movementKey: input.createMovementKey(event),
      checkedDefault:
        hasRequiredExcludedFields &&
        Boolean(shpo || unrecordedSameMonthTransit),
    });
    if (
      /відсутн.*архів|в архіві/iu.test(`${event.status} ${event.note}`) &&
      !input.archiveAll.some((period) => input.isSamePerson(event, period)) &&
      !input.existingOps.some(
        (op) =>
          op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING" &&
          input.isSamePerson(event, op),
      )
    ) {
      ops.push({
        id: input.opId(["archive_missing", event.personId || event.fullName]),
        kind: "data_mismatch",
        class: "needs_input",
        sheet: "Дані джерел / archive",
        personId: fromId || event.personId,
        fullName: fromName,
        positionIndex: fromIndex,
        rank: fromRank,
        before: event.status || "ВІДСУТНІЙ в АРХІВІ",
        after: "ARCHIVE_REFERENCE_MISSING → перевірити archive",
        sourceRef: `Рух!R${event.excelRow} СТАТУС=«${event.status}»`,
        why: "У РУХ стоїть «ВІДСУТНІЙ в АРХІВІ», але в archive немає запису за ПІБ чи ID. ЛІК / ВІД / СЗЧ не вигадуємо — кадровий маршрут вибуття це не скасовує.",
        confidence: "review",
        payload: {
          type: "ARCHIVE_REFERENCE_MISSING",
          mismatchKind: "ARCHIVE_REFERENCE_MISSING",
          statusRaw: event.status,
        },
        checkedDefault: false,
      });
    }
  return { ops, handled: true };
};
