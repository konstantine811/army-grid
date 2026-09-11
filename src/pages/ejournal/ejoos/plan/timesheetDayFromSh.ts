import {
  isTimesheetAbsenceCode,
  timesheetHorizonFillDays,
} from "../../ejoosTimesheetText";
import type { EjoosStatusMapping } from "../../ejoosStatusMap";
import type { PbArchivePeriod, PbShPerson } from "../types/pb";
import type {
  EjoosArrivalRow,
  EjoosExcludedRow,
  EjoosShpoRow,
  EjoosTimesheetPersonScan,
  EjoosTimesheetRow,
} from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import type { StaffEpisodePaintPayload } from "./tempArrivalClose";
import { hasActualReturn } from "./movementContext";
import { opId } from "./opId";

export type MapShStatus = (raw: string) => EjoosStatusMapping;

export type TimesheetDayFromShInput = {
  shPeople: PbShPerson[];
  ejoosDays: EjoosTimesheetRow[];
  archiveAll: PbArchivePeriod[];
  timesheetDay: number;
  timesheetDayLabel: string;
  dayByIndex: Map<string, EjoosTimesheetRow>;
  shpoByIndex: Map<string, EjoosShpoRow>;
  timesheetScanByRow: Map<number, EjoosTimesheetPersonScan>;
  arrivalById: Map<string, EjoosArrivalRow>;
  arrivalByName: Map<string, EjoosArrivalRow>;
  mapStatus: MapShStatus;
  alreadyVacatedForAbsence: (personId: string, fullName: string) => boolean;
  positionEventForShPerson: (person: PbShPerson) => unknown;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => T | null | undefined;
  timesheetRowsOf: (
    personId: string,
    fullName: string,
  ) => EjoosTimesheetPersonScan[];
  activeTimesheetRowOf: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => EjoosTimesheetRow | null | undefined;
  augustAbsenceSpansFor: (
    personId: string,
    fullName: string,
  ) => Array<{ fromDay: number; toDay: number; code: string }>;
  staffEpisodePaintPayload: (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => StaffEpisodePaintPayload;
  findFalseHopExcludedRow: (
    personId: string,
    fullName: string,
  ) => EjoosExcludedRow | null | undefined;
};

/** timesheet_day зі статусу sh (окремо від position_change та archive paint). */
export const planTimesheetDayFromShOps = (
  input: TimesheetDayFromShInput,
): EjoosSyncOp[] => {
  const ops: EjoosSyncOp[] = [];

  for (const person of input.shPeople) {
    if (!person.status || person.status === "0") continue;
    const inTempArrivals = Boolean(
      (person.personId && input.arrivalById.get(person.personId)) ||
        input.byPersonName(
          input.arrivalByName,
          person.personId,
          person.fullName,
        ),
    );
    if (
      input.alreadyVacatedForAbsence(person.personId, person.fullName) &&
      !inTempArrivals
    ) {
      continue;
    }
    const mapped = input.mapStatus(person.status);

    // Постановку на штат фарбує position_change; окремий timesheet_day тут
    // лише дублює і може зіпсувати коди СЗЧ на початку місяця.
    if (Boolean(input.positionEventForShPerson(person))) continue;

    const archiveControlsReportDay = input
      .augustAbsenceSpansFor(person.personId, person.fullName)
      .some(
        (span) =>
          Boolean(span.code) &&
          span.fromDay <= input.timesheetDay &&
          span.toDay >= input.timesheetDay,
      );
    if (archiveControlsReportDay) continue;

    const personTimesheetScan =
      input.timesheetRowsOf(person.personId, person.fullName).find(
        (row) => !row.hasDepartureText,
      ) || input.timesheetRowsOf(person.personId, person.fullName)[0];
    const personTimesheet =
      input.activeTimesheetRowOf(
        person.personId,
        person.fullName,
        person.positionIndex,
      ) ||
      (personTimesheetScan
        ? input.ejoosDays.find(
            (row) => row.excelRow === personTimesheetScan.excelRow,
          ) || {
            excelRow: personTimesheetScan.excelRow,
            personId: personTimesheetScan.personId,
            fullName: personTimesheetScan.fullName,
            rank: personTimesheetScan.rank,
            positionIndex: personTimesheetScan.positionIndex,
            dayValue: String(
              personTimesheetScan.dayCodes[input.timesheetDay] || "",
            ).trim(),
          }
        : null);
    const indexTimesheet = person.positionIndex
      ? (input.dayByIndex.get(person.positionIndex) ?? null)
      : null;
    const timesheetRow = personTimesheet || indexTimesheet;
    const timesheetScan = timesheetRow
      ? input.timesheetScanByRow.get(timesheetRow.excelRow)
      : undefined;
    const confirmedReturn = input.archiveAll.some(
      (period) =>
        input.isSamePerson(person, period) && hasActualReturn(period.returnDate),
    );
    const before = personTimesheet?.dayValue || "—";
    const afterCode = mapped.timesheetCode;
    const horizonFills = timesheetHorizonFillDays({
      dayCodes: timesheetScan?.dayCodes ?? [],
      horizon: input.timesheetDay,
      reportCode: afterCode || "+",
      confirmedReturn,
    });
    const continuedMark =
      horizonFills.find((item) => item.day === input.timesheetDay)?.mark || "";
    const onMatchingShpo = Boolean(
      person.positionIndex &&
        input.shpoByIndex.get(person.positionIndex) &&
        input.isSamePerson(
          person,
          input.shpoByIndex.get(person.positionIndex)!,
        ),
    );
    const indexTakenByOther = Boolean(
      indexTimesheet &&
        (indexTimesheet.personId || indexTimesheet.fullName) &&
        !input.isSamePerson(person, indexTimesheet),
    );
    const missingFromTimesheet =
      !personTimesheet && onMatchingShpo && mapped.timesheetCode === "+";
    if (missingFromTimesheet) {
      const restorePaint = input.staffEpisodePaintPayload(
        person.personId,
        person.fullName,
        person.positionIndex,
        indexTimesheet?.excelRow || 0,
      );
      ops.push({
        id: opId(["ts_restore", person.personId || person.positionIndex]),
        kind: "timesheet_day",
        class: "needs_input",
        sheet: "6. Табель",
        personId: person.personId,
        fullName: person.fullName,
        positionIndex: person.positionIndex,
        rank: person.rank,
        before: "немає в Табелі",
        after: `відновити на ${person.positionIndex}${
          indexTakenByOther ? " — рядок зайнятий іншою особою" : ""
        }`,
        sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}» · ШПО інд. ${person.positionIndex}`,
        why: indexTakenByOther
          ? "У ШПО особа є і в строю, але рядок Табеля на цьому індексі зайнятий іншою людиною — не перезаписуємо автоматично"
          : "У ШПО особа вже стоїть на актуальній посаді і в строю, а в Табелі її немає. Кадровий рух не проводимо — лише відновити рядок Табеля після перевірки",
        confidence: "review",
        payload: {
          type: "PAINT_ARCHIVE",
          timesheetCode: afterCode || "+",
          day: String(input.timesheetDay),
          excelRow:
            !indexTakenByOther && indexTimesheet
              ? String(indexTimesheet.excelRow)
              : "",
          statusRaw: person.status,
          restorePerson: "1",
          nextName: person.fullName,
          nextPersonId: person.personId,
          nextRank: input.shpoByIndex.get(person.positionIndex)?.rank || "",
          timesheetActiveFrom: restorePaint.timesheetActiveFrom,
          timesheetPreserveHistory: restorePaint.timesheetPreserveHistory,
          timesheetAbsenceSpans: restorePaint.timesheetAbsenceSpans,
          historyTimesheetExcelRow: restorePaint.historyTimesheetExcelRow,
          historyTimesheetAbsenceSpans:
            restorePaint.historyTimesheetAbsenceSpans,
          timesheetSkipHistory: "1",
        },
        checkedDefault: false,
      });
    }

    if (!missingFromTimesheet && afterCode) {
      const after =
        afterCode === "+" && continuedMark && continuedMark !== "+"
          ? continuedMark
          : afterCode;
      const needsHorizonFill = horizonFills.length > 0;
      const falseHopExcluded = input.findFalseHopExcludedRow(
        person.personId,
        person.fullName,
      );
      if (before !== after || needsHorizonFill || falseHopExcluded) {
        const isReady = mapped.confidence === "high" && Boolean(timesheetRow);
        const fillingGaps =
          needsHorizonFill &&
          (before === after ||
            before === "—" ||
            !String(before || "").trim() ||
            (after === "+" && before === "+"));
        ops.push({
          id: opId([
            "ts",
            person.personId || person.positionIndex,
            String(input.timesheetDay),
            after,
          ]),
          kind: "timesheet_day",
          class: !timesheetRow
            ? "needs_input"
            : mapped.confidence === "manual"
              ? "needs_input"
              : mapped.confidence === "review"
                ? "conflict"
                : "ready",
          sheet: falseHopExcluded ? "6. Табель / 3. Виключені" : "6. Табель",
          personId: person.personId,
          fullName: person.fullName,
          positionIndex: person.positionIndex,
          rank: person.rank,
          before,
          after,
          sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
          why: falseHopExcluded
            ? `Внутрішня ПОСАДА 1ПБ не є вибуттям — прибрати хибний рядок Виключені R${falseHopExcluded.excelRow}${
                before !== after ? `; Табель ${before} → ${after}` : ""
              }`
            : fillingGaps
              ? `У Табелі порожні дні до зрізу ${input.timesheetDayLabel} — продовжуємо «${after}», не лише день ${input.timesheetDay}`
              : isTimesheetAbsenceCode(after) && afterCode === "+"
                ? `Відкритий «${after}» без фактичного прибуття — лишаємо той самий статус по ${input.timesheetDayLabel}`
                : mapped.reason,
          confidence: mapped.confidence,
          payload: {
            timesheetCode: after,
            day: String(input.timesheetDay),
            excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
            statusRaw: person.status,
            confirmedReturn: confirmedReturn ? "1" : "",
            clearExcludedExcelRow: falseHopExcluded
              ? String(falseHopExcluded.excelRow)
              : "",
          },
          checkedDefault: isReady,
        });
      }
    } else if (
      !missingFromTimesheet &&
      mapped.confidence !== "high" &&
      mapped.ruleId !== "absent_archive"
    ) {
      const alreadyPresent = before === "+";
      const continueAbsence =
        continuedMark &&
        continuedMark !== "+" &&
        isTimesheetAbsenceCode(continuedMark);
      const defaultsToPresent =
        Boolean(timesheetRow) &&
        !continueAbsence &&
        (before === "—" || alreadyPresent);
      if (continueAbsence) {
        ops.push({
          id: opId([
            "ts_continue",
            person.personId || person.fullName,
            continuedMark,
          ]),
          kind: "timesheet_day",
          class: "ready",
          sheet: "6. Табель",
          personId: person.personId,
          fullName: person.fullName,
          positionIndex: person.positionIndex,
          rank: person.rank,
          before,
          after: continuedMark,
          sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
          why: `Відкритий «${continuedMark}» без прибуття — продовжуємо по ${input.timesheetDayLabel}, не ставимо «+»`,
          confidence: "high",
          payload: {
            day: String(input.timesheetDay),
            excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
            statusRaw: person.status,
            timesheetCode: continuedMark,
          },
          checkedDefault: true,
        });
      } else if (horizonFills.length && alreadyPresent) {
        ops.push({
          id: opId([
            "ts_fill",
            person.personId || person.fullName,
            String(input.timesheetDay),
          ]),
          kind: "timesheet_day",
          class: "ready",
          sheet: "6. Табель",
          personId: person.personId,
          fullName: person.fullName,
          positionIndex: person.positionIndex,
          rank: person.rank,
          before,
          after: "+",
          sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
          why: `У Табелі порожні дні до зрізу ${input.timesheetDayLabel} — добиваємо «+»`,
          confidence: "high",
          payload: {
            day: String(input.timesheetDay),
            excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
            statusRaw: person.status,
            timesheetCode: "+",
          },
          checkedDefault: true,
        });
      } else if (!alreadyPresent) {
        ops.push({
          id: opId([
            "ts_manual",
            person.personId || person.fullName,
            person.status,
          ]),
          kind: "timesheet_day",
          class: defaultsToPresent ? "ready" : "needs_input",
          sheet: "6. Табель",
          personId: person.personId,
          fullName: person.fullName,
          positionIndex: person.positionIndex,
          rank: person.rank,
          before,
          after: defaultsToPresent ? "+" : "(оберіть код)",
          sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
          why: defaultsToPresent
            ? "У Табелі за поточний день порожньо — за замовчуванням ставимо «+»"
            : mapped.reason,
          confidence: defaultsToPresent ? "high" : "manual",
          payload: {
            day: String(input.timesheetDay),
            excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
            statusRaw: person.status,
            timesheetCode: defaultsToPresent ? "+" : "",
          },
          checkedDefault: defaultsToPresent,
        });
      }
    }
  }

  return ops;
};
