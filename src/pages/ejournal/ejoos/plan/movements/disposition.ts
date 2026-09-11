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

export type DispositionMovementInput = {
  event: PbMovement;
  activeMovementsAll: unknown;
  activeSzchPersonIds: unknown;
  activeSzchPersonNames: unknown;
  activeTimesheetRowOf: unknown;
  alreadyVacatedForAbsence: unknown;
  archiveAll: unknown;
  augustAbsenceSpansFor: unknown;
  absenceSpansBeforeEpisode: unknown;
  byPersonName: unknown;
  canonicalName: unknown;
  createMovementKey: unknown;
  dateMs: unknown;
  dayByIndex: unknown;
  ejoosAbsents: unknown;
  ejoosDays: unknown;
  encodeTimesheetAbsenceSpans: unknown;
  eventInLeadWindow: unknown;
  formatDispositionTimesheetDeparture: unknown;
  hasActualReturn: unknown;
  isDispositionAbsenceStatus: unknown;
  isOwnUnitStaffMove: unknown;
  isSamePerson: unknown;
  isTimesheetStaffPositionRow: unknown;
  isVacantStaffRow: unknown;
  journalDayFromDateMs: unknown;
  leadWindowStart: unknown;
  mapStatus: unknown;
  movementEventTime: unknown;
  movementPersonKey: unknown;
  onStaffShpo: unknown;
  oosPersonById: unknown;
  oosPersonByName: unknown;
  opId: unknown;
  personInTextRows: unknown;
  personNameKeys: unknown;
  personStillInSh: unknown;
  positionChainByPerson: unknown;
  priorMonthDispositionMonthLabel: unknown;
  shpoByPositionIndex: unknown;
  shpoDispositionRows: unknown;
  shpoPersonById: unknown;
  shpoPersonByName: unknown;
  shpoStrayRowByName: unknown;
  shpoSzchRows: unknown;
  staffIndexTimesheetForPerson: unknown;
  staffPositionTitleForIndex: unknown;
  timesheetDay: unknown;
  timesheetDayLabel: unknown;
  timesheetDispositionStaffRows: unknown;
  timesheetRowByCanonicalName: unknown;
  timesheetSzchRows: unknown;
  norm: unknown;
};

export const planDispositionMovementOps = (
  input: DispositionMovementInput,
): MovementPlanResult => {
  const ops: EjoosSyncOp[] = [];
  const event = input.event;
    const samePerson = (personId: string, fullName: string) =>
      input.isSamePerson(event, { personId, fullName });
    const inCurrentSh = input.personStillInSh(event.personId, event.fullName);
    const inActiveShpo = input.onStaffShpo(event.personId, event.fullName);
    const hasOpenAbsence = input.ejoosAbsents.some(
      (row) =>
        (samePerson(row.personId, row.fullName) ||
          (event.fullName &&
            row.fullName &&
            input.canonicalName(row.fullName) === input.canonicalName(event.fullName))) &&
        !row.actualReturn,
    );
    const inActiveOos = Boolean(
      (event.personId && input.oosPersonById.has(event.personId)) ||
      input.byPersonName(input.oosPersonByName, event.personId, event.fullName),
    );

    // Стара подія РУХ не є завданням сама по собі. Якщо людини вже немає
    // в актуальних джерелах і немає відкритої відсутності, стан вважаємо
    // відпрацьованим раніше (ALREADY_PROCESSED / NO_ACTION).
    if (!inCurrentSh && !inActiveShpo && !inActiveOos && !hasOpenAbsence) {
      return { ops, handled: true };
    }

    const oldPosition = input.shpoByPositionIndex.get(event.previousIndex);
    const oldPositionFreed =
      !oldPosition || !samePerson(oldPosition.personId, oldPosition.fullName);
    const remainsInOos = inActiveOos;
    const hasSzchContext = Boolean(
      (event.personId && input.activeSzchPersonIds.has(event.personId)) ||
      [...input.personNameKeys(event.personId, event.fullName)].some((key) =>
        input.activeSzchPersonNames.has(key),
      ),
    );
    const szchRemains = input.ejoosAbsents.some(
      (row) =>
        (samePerson(row.personId, row.fullName) ||
          (event.fullName &&
            row.fullName &&
            input.canonicalName(row.fullName) === input.canonicalName(event.fullName))) &&
        input.isDispositionAbsenceStatus(row.ground) &&
        !row.actualReturn,
    );
    const dispositionInShpo = input.personInTextRows(event, input.shpoDispositionRows);
    const dispositionInTimesheet = input.personInTextRows(
      event,
      input.timesheetDispositionStaffRows,
    );
    const szchReflectedElsewhere =
      input.personInTextRows(event, input.shpoSzchRows) ||
      input.personInTextRows(event, input.timesheetSzchRows);
    const absenceStateReflected =
      !hasSzchContext || szchRemains || szchReflectedElsewhere;
    const dispositionStateReflected =
      dispositionInShpo || dispositionInTimesheet;
    const openArchivePeriods = input.archiveAll
      .filter(
        (period) =>
          samePerson(period.personId, period.fullName) &&
          !input.hasActualReturn(period.returnDate),
      )
      .sort((a, b) => b.excelRow - a.excelRow);
    // Стан у РУХ (СЗЧ / БЕЗВІСТИ) головніший за випадковий останній період
    // архіву: інакше в блок розпорядження потрапляє «ЛІКУВАННЯ» чи службовий
    // рядок «ВНЕСЕННЯ ДАНИХ».
    const activeArchivePeriod =
      openArchivePeriods.find((period) =>
        input.isDispositionAbsenceStatus(period.absenceType),
      ) ?? openArchivePeriods[0];
    const timesheetRow =
      input.activeTimesheetRowOf(
        event.personId,
        event.fullName,
        event.previousIndex,
      ) ||
      input.timesheetRowByCanonicalName(event.fullName) ||
      (event.previousIndex
        ? (() => {
            const scan = input.staffIndexTimesheetForPerson(
              event.personId,
              event.fullName,
              event.previousIndex,
            );
            if (!scan) return null;
            return (
              input.ejoosDays.find((row) => row.excelRow === scan.excelRow) ?? {
                excelRow: scan.excelRow,
                personId: scan.personId || event.personId,
                fullName: scan.fullName || event.fullName,
                rank: scan.rank || event.rank,
                positionIndex: scan.positionIndex || event.previousIndex,
                dayValue: "",
              }
            );
          })()
        : null);
    // Рядок особи в ШПО/Табелі може бути вже в блоці розпорядження — тоді
    // його не чіпаємо. Закривати треба лише штатний рядок.
    const shpoDispositionExcelRows = new Set(
      input.shpoDispositionRows.map((row) => row.excelRow),
    );
    const timesheetDispositionExcelRows = new Set(
      input.timesheetDispositionStaffRows.map((row) => row.excelRow),
    );
    const personShpoRow =
      (event.personId && input.shpoPersonById.get(event.personId)) ||
      input.byPersonName(input.shpoPersonByName, event.personId, event.fullName) ||
      input.byPersonName(input.shpoStrayRowByName, event.personId, event.fullName) ||
      null;
    const staffShpoRow =
      oldPosition && samePerson(oldPosition.personId, oldPosition.fullName)
        ? oldPosition
        : personShpoRow &&
            !shpoDispositionExcelRows.has(personShpoRow.excelRow)
          ? personShpoRow
          : null;
    const staffTimesheetRow =
      (timesheetRow &&
      (!timesheetDispositionExcelRows.has(timesheetRow.excelRow) ||
        input.isTimesheetStaffPositionRow(timesheetRow.excelRow))
        ? timesheetRow
        : null) ||
      (event.previousIndex &&
      input.dayByIndex.get(event.previousIndex) &&
      (input.isVacantStaffRow(input.dayByIndex.get(event.previousIndex)!) ||
        input.isSamePerson(
          { personId: event.personId, fullName: event.fullName },
          input.dayByIndex.get(event.previousIndex)!,
        ))
        ? input.dayByIndex.get(event.previousIndex)!
        : null);

    // РОЗПОРЯДЖ не є виключенням зі списків частини: ООС та СЗЧ
    // зберігаються, а до «3. Виключені» особа не переноситься.
    if (
      oldPositionFreed &&
      !staffShpoRow &&
      !staffTimesheetRow &&
      ((absenceStateReflected && dispositionStateReflected) ||
        (remainsInOos && hasOpenAbsence) ||
        input.alreadyVacatedForAbsence(event.personId, event.fullName))
    ) {
      return { ops, handled: true };
    }

    // Розпорядження проводимо лише після непроведених змін посади:
    // спочатку історія 2103791 → 2103179, і лише потім вивід у розпорядження.
    const pendingPositionSteps = (
      input.positionChainByPerson.get(input.movementPersonKey(event)) ?? []
    ).filter((pending) => pending.excelRow < event.excelRow);
    if (pendingPositionSteps.length) {
      const first = pendingPositionSteps[0];
      const last = pendingPositionSteps[pendingPositionSteps.length - 1];
      const total = pendingPositionSteps.length + 1;
      ops.push({
        id: input.opId([
          "disposition-after-position",
          event.movementNumber || String(event.excelRow),
        ]),
        kind: "other_manual",
        class: "needs_input",
        sheet: "ШПО → розпорядження / Тимчасово відсутні / Табель",
        personId: event.personId,
        fullName: event.fullName,
        rank: event.rank,
        positionIndex: event.previousIndex,
        before: `у ЕЖООС ще штатна посада ${first.previousIndex}`,
        after: `спочатку зміна посади ${first.previousIndex} → ${last.nextIndex}, потім розпорядження з ${event.previousIndex || "—"}`,
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: `Крок ${total} з ${total}: розпорядження проводимо після зміни посади, інакше зникне історія переходу ${first.previousIndex} → ${last.nextIndex}.`,
        confidence: "manual",
        payload: {
          type: event.type,
          chainWaiting: "1",
          chainStep: String(total),
          chainTotal: String(total),
          awaitFromIndex: first.previousIndex,
          awaitToIndex: last.nextIndex,
          awaitOrderNumber: first.orderNumber,
          awaitOrderDate: first.orderDate,
          previousIndex: event.previousIndex,
          destination: event.destination || event.changeText,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
        },
        checkedDefault: false,
      });
      return { ops, handled: true };
    }

    // Запис у «5. Тимчасово відсутні» веде sync з archive (absent_upsert),
    // не move_to_disposition — інакше дубль «ДОДАТИ РЯДОК» поверх archive.
    const openAbsentRow = input.ejoosAbsents.find(
      (row) =>
        (samePerson(row.personId, row.fullName) ||
          (event.fullName &&
            row.fullName &&
            input.canonicalName(row.fullName) === input.canonicalName(event.fullName))) &&
        !row.actualReturn,
    );
    const needsTimesheetClose = Boolean(staffTimesheetRow);
    const keepOpenAbsenceTimesheet = Boolean(
      hasSzchContext || szchRemains || openAbsentRow,
    );
    const timesheetNeedsCreate =
      keepOpenAbsenceTimesheet &&
      !staffTimesheetRow &&
      Boolean(event.previousIndex);
    const dispositionAbsenceHint =
      (openAbsentRow?.ground &&
      input.isDispositionAbsenceStatus(openAbsentRow.ground)
        ? input.norm(openAbsentRow.ground)
        : "") ||
      activeArchivePeriod?.absenceType ||
      event.status ||
      "";
    const keepTimesheetRowInPlace = /безвіст/i.test(dispositionAbsenceHint);
    const indexTimesheet = event.previousIndex
      ? input.dayByIndex.get(event.previousIndex)
      : undefined;
    const vacateTimesheetStaffSlot =
      keepOpenAbsenceTimesheet &&
      !keepTimesheetRowInPlace &&
      Boolean(
        staffTimesheetRow &&
        indexTimesheet &&
        staffTimesheetRow.excelRow === indexTimesheet.excelRow,
      );
    const canMoveToDisposition = Boolean(
      staffShpoRow || needsTimesheetClose || timesheetNeedsCreate,
    );
    if (canMoveToDisposition) {
      const orderMs = input.dateMs(event.orderDate || event.basisDate);
      const orderInJournalMonth =
        !orderMs ||
        !input.leadWindowStart ||
        input.journalDayFromDateMs(orderMs, input.leadWindowStart) > 0;
      const journalMonthBlocked = Boolean(
        orderMs && input.leadWindowStart && !orderInJournalMonth,
      );
      const laterStaffPlacementInWindow = input.activeMovementsAll.some(
        (movement) =>
          samePerson(movement.personId, movement.fullName) &&
          movement.type === "ПОСАДА" &&
          input.isOwnUnitStaffMove(movement) &&
          input.eventInLeadWindow(movement) &&
          (input.movementEventTime(movement) > input.movementEventTime(event) ||
            (input.movementEventTime(movement) === input.movementEventTime(event) &&
              movement.excelRow > event.excelRow)),
      );
      if (
        journalMonthBlocked &&
        (dispositionStateReflected || laterStaffPlacementInWindow)
      ) {
        return { ops, handled: true };
      }
      const targetMonthLabel = input.priorMonthDispositionMonthLabel(
        event.orderDate || event.basisDate || "",
      );
      const absenceStatus =
        (openAbsentRow?.ground &&
        input.isDispositionAbsenceStatus(openAbsentRow.ground)
          ? input.norm(openAbsentRow.ground)
          : "") ||
        (input.isDispositionAbsenceStatus(event.status)
          ? input.norm(event.status)
          : "") ||
        activeArchivePeriod?.absenceType ||
        event.status ||
        "РОЗПОРЯДЖЕННЯ";
      const mappedAbsence = input.mapStatus(absenceStatus);
      const absenceCode =
        mappedAbsence.timesheetCode ||
        (/безвіст/iu.test(absenceStatus) ? "ЗБ" : "") ||
        (/сзч|самовіл/iu.test(absenceStatus) ? "СЗЧ" : "") ||
        absenceStatus;
      const dispositionTimesheetDeparture =
        input.formatDispositionTimesheetDeparture(
          event.destination || event.changeText,
          event.orderNumber,
          event.orderDate,
        );
      const openAbsenceLabel = /безвіст/iu.test(absenceStatus)
        ? "БЕЗВІСТИ"
        : /сзч|самовіл/iu.test(absenceStatus)
          ? "СЗЧ"
          : "";
      ops.push({
        id: input.opId([
          "move-to-disposition",
          event.movementNumber || String(event.excelRow),
        ]),
        kind: "move_to_disposition",
        // Без даних архіву запис відсутності заповнити нічим.
        class: journalMonthBlocked ? "needs_input" : "ready",
        sheet: journalMonthBlocked
          ? "ШПО → розпорядження / Табель (місяць наказу)"
          : "ШПО → розпорядження / Тимчасово відсутні / Табель",
        personId: event.personId || staffShpoRow?.personId || "",
        fullName: event.fullName || staffShpoRow?.fullName || "",
        rank: event.rank || staffShpoRow?.rank || "",
        positionIndex: event.previousIndex,
        before: `штатна посада ${event.previousIndex}`,
        after: [
          event.destination || event.changeText || "у розпорядження",
          absenceStatus,
          keepOpenAbsenceTimesheet
            ? `Табель 01–${String(input.timesheetDay).padStart(2, "0")} ${absenceCode}; ${event.orderDate || "у дату наказу"} — ${dispositionTimesheetDeparture}; далі «-»`
            : staffTimesheetRow || timesheetNeedsCreate
              ? `Табель до ${event.orderDate || "наказу"} «+»; у дату наказу — ${dispositionTimesheetDeparture}; далі до кінця місяця «-»`
              : "",
          openAbsentRow || szchRemains
            ? `${openAbsenceLabel || "відсутність"} лишається відкритою`
            : hasSzchContext
              ? "відсутність додасть sync з archive"
              : "",
          !keepOpenAbsenceTimesheet &&
          !staffTimesheetRow &&
          !timesheetNeedsCreate &&
          !dispositionInTimesheet
            ? "штатний рядок Табеля не знайдено — перевірити"
            : timesheetNeedsCreate
              ? `додати рядок у блок «ВИБУВ У РОЗПОРЯДЖЕННЯ…» з ${absenceCode}`
              : "",
          !remainsInOos ? "в ООС активного запису немає" : "",
        ]
          .filter(Boolean)
          .join(" · "),
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: journalMonthBlocked
          ? `Наказ у ${targetMonthLabel}, а зараз «станом на» ${input.timesheetDayLabel}. Змініть дату на ${event.orderDate || "місяць наказу"} і перебудуйте — тоді застосуйте розпорядження.`
          : keepOpenAbsenceTimesheet
          ? openAbsenceLabel === "БЕЗВІСТИ"
            ? "БЕЗВІСТИ → РОЗПОРЯДЖ: звільнити ШПО (фінальний sh wins), ООС і відкритий БЕЗВІСТИ лишити; у Табелі до дати наказу ЗБ, у дату наказу — вибуття в розпорядження, далі «-». Виключені не змінюються."
            : "СЗЧ → РОЗПОРЯДЖ: звільнити ШПО (фінальний sh wins), ООС і відкритий СЗЧ лишити; у Табелі до дати наказу СЗЧ, у дату наказу — вибуття в розпорядження, далі «-». Виключені не змінюються."
          : "РОЗПОРЯДЖ звільняє стару штатну посаду, але залишає особу в ООС. Виключені не змінюються.",
        confidence: journalMonthBlocked
          ? "manual"
          : activeArchivePeriod
            ? "high"
            : "review",
        payload: {
          type: event.type,
          previousIndex: event.previousIndex,
          destination: event.destination || event.changeText,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
          basisNumber: event.basisNumber,
          basisDate: event.basisDate,
          shpoExcelRow: String(staffShpoRow?.excelRow || ""),
          timesheetExcelRow: String(staffTimesheetRow?.excelRow || ""),
          positionTitle: input.staffPositionTitleForIndex(
            staffShpoRow?.positionIndex || event.previousIndex,
          ),
          skipShpoDisposition: dispositionInShpo ? "1" : "",
          absenceExcelRow: String(openAbsentRow?.excelRow || ""),
          needsAbsenceRecord: "",
          absenceType: absenceStatus,
          absenceCode,
          absenceDate: activeArchivePeriod?.departDate || "",
          absencePlace: activeArchivePeriod?.place || "",
          absenceOrderNumber: activeArchivePeriod?.orderNumber || "",
          absenceOrderDate: activeArchivePeriod?.orderDate || "",
          plannedReturn: activeArchivePeriod?.plannedReturn || "",
          remainsInOos: String(remainsInOos),
          timesheetFound: String(
            Boolean(staffTimesheetRow) || timesheetNeedsCreate,
          ),
          timesheetCreateRow: timesheetNeedsCreate ? "1" : "",
          timesheetStaffIndex: event.previousIndex || "",
          restorePerson:
            staffTimesheetRow &&
            input.isVacantStaffRow(staffTimesheetRow) &&
            !staffTimesheetRow.fullName &&
            !staffTimesheetRow.personId
              ? "1"
              : "",
          keepOpenSzchTimesheet: keepOpenAbsenceTimesheet ? "1" : "",
          vacateTimesheetStaffSlot: vacateTimesheetStaffSlot ? "1" : "",
          hasSzchContext: String(hasSzchContext),
          szchRemains: String(szchRemains),
          szchReflectedElsewhere: String(szchReflectedElsewhere),
          dispositionInShpo: String(dispositionInShpo),
          dispositionInTimesheet: String(dispositionInTimesheet),
          timesheetAbsenceSpans: journalMonthBlocked
            ? ""
            : input.encodeTimesheetAbsenceSpans(
                (() => {
                  const spans = input.augustAbsenceSpansFor(
                    event.personId || staffShpoRow?.personId || "",
                    event.fullName || staffShpoRow?.fullName || "",
                  );
                  if (!keepOpenAbsenceTimesheet) return spans;
                  const orderDay =
                    orderMs && input.leadWindowStart
                      ? input.journalDayFromDateMs(orderMs, input.leadWindowStart)
                      : 0;
                  return orderDay > 1
                    ? input.absenceSpansBeforeEpisode(spans, orderDay)
                    : spans;
                })(),
              ),
          journalMonthBlocked: journalMonthBlocked ? "1" : "",
          suggestedAsOfDate: event.orderDate || "",
          targetMonthLabel,
        },
        movementKey: input.createMovementKey(event),
        checkedDefault: !journalMonthBlocked,
      });
      return { ops, handled: true };
    }

    const missing = [
      !dispositionStateReflected &&
        "перевірити відображення розпорядження у ШПО або Табелі",
      hasSzchContext &&
        !absenceStateReflected &&
        "перевірити чинний запис відсутності у Тимчасово відсутніх",
    ].filter(Boolean);
    ops.push({
      id: input.opId([
        "disposition-review",
        event.movementNumber || String(event.excelRow),
      ]),
      kind: "other_manual",
      class: "conflict",
      sheet: "ШПО / ООС / Тимчасово відсутні / Табель",
      personId: event.personId,
      fullName: event.fullName,
      rank: event.rank,
      positionIndex: event.previousIndex,
      before: `РОЗПОРЯДЖ зі штатної посади ${event.previousIndex || "—"}`,
      after: missing.join("; "),
      sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
      why: "РОЗПОРЯДЖ не видаляє особу з ООС і не додає її до Виключених без окремої події виключення зі списків частини.",
      confidence: "manual",
      payload: {
        type: event.type,
        previousIndex: event.previousIndex,
        destination: event.destination || event.changeText,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
        oldPositionFreed: String(oldPositionFreed),
        remainsInOos: String(remainsInOos),
        hasSzchContext: String(hasSzchContext),
        szchRemains: String(szchRemains),
        szchReflectedElsewhere: String(szchReflectedElsewhere),
        dispositionInShpo: String(dispositionInShpo),
        dispositionInTimesheet: String(dispositionInTimesheet),
        dispositionStateReflected: String(dispositionStateReflected),
      },
      checkedDefault: false,
    });
    return { ops, handled: true };
  return { ops, handled: true };
};
