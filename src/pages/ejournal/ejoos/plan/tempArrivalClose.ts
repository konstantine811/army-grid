import { isOwnUnitStaffMove } from "../../ejoosMovementRules";
import { canonicalName } from "../parse/cellText";
import type { PbMovement, PbShPerson } from "../types/pb";
import type {
  EjoosArrivalRow,
  EjoosOosRow,
  EjoosShpoRow,
  EjoosTimesheetRow,
} from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { opId } from "./opId";

export type StaffEpisodePaintPayload = {
  timesheetActiveFrom: string;
  timesheetPreserveHistory: string;
  timesheetAbsenceSpans: string;
  historyTimesheetExcelRow: string;
  historyTimesheetAbsenceSpans: string;
};

export type TempArrivalClosePlanInput = {
  shPeople: PbShPerson[];
  movementsAll: PbMovement[];
  shpoByIndex: Map<string, EjoosShpoRow>;
  ejoosOos: EjoosOosRow[];
  dayByIndex: Map<string, EjoosTimesheetRow>;
  oosById: Map<string, EjoosOosRow>;
  oosByName: Map<string, EjoosOosRow>;
  journalMonthStartLabel: string;
  existingOps: EjoosSyncOp[];
  arrivalOf: (personId: string, fullName: string) => EjoosArrivalRow | null;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  eventInLeadWindow: (event: PbMovement) => boolean;
  movementEventTime: (event: PbMovement) => number;
  isPositionIndex: (value: string) => boolean;
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => T | null | undefined;
  staffIndexTimesheetForPerson: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => EjoosTimesheetRow | null | undefined;
  staffAppointmentDateFor: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => string;
  timesheetEpisodeStartFor: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => string;
  inboundStaffDateFor: (personId: string, fullName: string) => string;
  staffEpisodePaintPayload: (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => StaffEpisodePaintPayload;
};

/** Закриття «Тимчасово прибулі», коли sh/ШПО вже відповідають штатній посаді. */
export const planTempArrivalCloseOps = (
  input: TempArrivalClosePlanInput,
): EjoosSyncOp[] => {
  const ops: EjoosSyncOp[] = [];
  for (const person of input.shPeople) {
    const nextIndex = person.positionIndex;
    if (!nextIndex || !input.isPositionIndex(nextIndex) || !person.fullName) {
      continue;
    }
    const personId = person.personId;
    const fullName = person.fullName;
    const arrival = input.arrivalOf(personId, fullName);
    if (!arrival) continue;
    if (
      input.existingOps.some(
        (op) =>
          op.kind === "position_change" &&
          input.isSamePerson({ personId, fullName }, op),
      ) ||
      ops.some(
        (op) =>
          op.kind === "position_change" &&
          input.isSamePerson({ personId, fullName }, op),
      )
    ) {
      continue;
    }
    const targetShpo = input.shpoByIndex.get(nextIndex) ?? null;
    if (!targetShpo || !input.isSamePerson({ personId, fullName }, targetShpo)) {
      continue;
    }
    const placementEvent =
      [...input.movementsAll]
        .filter(
          (event) =>
            input.isSamePerson({ personId, fullName }, event) &&
            input.eventInLeadWindow(event) &&
            (event.type === "ПРИБУВ" || isOwnUnitStaffMove(event)) &&
            (event.nextIndex === nextIndex ||
              String(event.changeText || "").includes(nextIndex) ||
              event.type === "ПРИБУВ"),
        )
        .sort(
          (left, right) =>
            input.movementEventTime(right) - input.movementEventTime(left) ||
            right.excelRow - left.excelRow,
        )[0] ?? null;
    const existingOos =
      (personId && input.oosById.get(personId)) ||
      input.byPersonName(input.oosByName, personId, fullName) ||
      input.ejoosOos.find((row) =>
        input.isSamePerson({ personId, fullName }, row),
      ) ||
      null;
    const indexTimesheet = input.dayByIndex.get(nextIndex) ?? null;
    const personStaffTimesheet = input.staffIndexTimesheetForPerson(
      personId,
      fullName,
      nextIndex,
    );
    const targetTimesheet =
      personStaffTimesheet ||
      (indexTimesheet &&
      input.isSamePerson({ personId, fullName }, indexTimesheet)
        ? indexTimesheet
        : null) ||
      indexTimesheet;
    const rank = person.rank || placementEvent?.rank || targetShpo.rank;
    const orderDate =
      placementEvent?.orderDate ||
      placementEvent?.basisDate ||
      input.staffAppointmentDateFor(personId, fullName, nextIndex) ||
      arrival.arriveDate ||
      "";
    const orderNumber = placementEvent?.orderNumber || "";
    const staffTimesheetFrom =
      input.timesheetEpisodeStartFor(personId, fullName, nextIndex) ||
      input.inboundStaffDateFor(personId, fullName) ||
      orderDate;
    const timesheetActiveFrom =
      staffTimesheetFrom || orderDate || input.journalMonthStartLabel;
    const episodePaint = input.staffEpisodePaintPayload(
      personId,
      fullName,
      nextIndex,
      targetTimesheet?.excelRow || 0,
    );
    const canApply = Boolean(targetShpo && targetTimesheet);
    ops.push({
      id: opId([
        "temp-arrival-close",
        personId || canonicalName(fullName),
        nextIndex,
      ]),
      kind: "position_change",
      class: canApply ? "ready" : "needs_input",
      sheet: "4. Тимч. прибулі → 1. ШПО / 2. ООС / 6. Табель",
      personId,
      fullName,
      positionIndex: nextIndex,
      rank,
      before: `тимчасово прибулий${arrival.fromUnit ? ` · ${arrival.fromUnit}` : ""}`,
      after: `штатна посада ${nextIndex}`,
      sourceRef: placementEvent
        ? `Рух!R${placementEvent.excelRow} №${placementEvent.movementNumber}`
        : `sh!R${person.excelRow} · ${nextIndex}`,
      why: placementEvent
        ? `ПРИБУВ/постановка №${orderNumber || "?"} від ${orderDate || "?"}: закрити тимчасове прибуття, зафіксувати в ООС штат ${nextIndex}`
        : `Особа вже на штаті ${nextIndex} за sh/ШПО — закрити відкритий рядок «Тимчасово прибулі» та оновити ООС`,
      confidence: canApply ? "high" : "manual",
      payload: {
        movementNumber: placementEvent?.movementNumber || "",
        previousIndex:
          placementEvent?.previousIndex || arrival.fromUnit || "БРЕЗ",
        nextIndex,
        changeText: placementEvent?.changeText || "",
        orderNumber,
        orderDate,
        basisNumber: placementEvent?.basisNumber || "",
        basisDate: placementEvent?.basisDate || "",
        nextName: fullName,
        nextRank: rank,
        nextPersonId: personId,
        positionTitle: person.positionTitle || "",
        statusRaw: person.status || placementEvent?.status || "",
        isTempArrivalPlacement: "1",
        arrivalExcelRow: String(arrival.excelRow),
        arrivalDepartDate: orderDate,
        arrivalDepartOrderNumber: orderNumber,
        arrivalDepartOrderDate: orderDate,
        oosExcelRow: existingOos ? String(existingOos.excelRow) : "",
        oosHistoryIndexes: nextIndex,
        oosHistoryDates: orderDate,
        shpoExcelRow: String(targetShpo.excelRow),
        timesheetExcelRow: targetTimesheet
          ? String(targetTimesheet.excelRow)
          : "",
        timesheetActiveFrom:
          episodePaint.timesheetActiveFrom || timesheetActiveFrom,
        timesheetSkipHistory: "1",
        timesheetPreserveHistory: episodePaint.timesheetPreserveHistory || "",
        timesheetBindStaffIndex: nextIndex,
        reconcileTempArrival: "1",
        ...episodePaint,
      },
      checkedDefault: canApply,
    });
  }
  return ops;
};
