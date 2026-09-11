import { createMovementKey, isSzchCancellation } from "../parse/pb";
import { normKey } from "../parse/cellText";
import type { PbMovement, PbShPerson } from "../types/pb";
import type { EjoosAbsentRow, EjoosTimesheetRow } from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { opId } from "./opId";

export type SzchAbsentClosePlanInput = {
  activeMovementsAll: PbMovement[];
  shPeople: PbShPerson[];
  ejoosDays: EjoosTimesheetRow[];
  timesheetDayLabel: string;
  openById: Map<string, EjoosAbsentRow>;
  openByName: Map<string, EjoosAbsentRow>;
  dayByIndex: Map<string, EjoosTimesheetRow>;
  eventInLeadWindow: (event: PbMovement) => boolean;
  activeTimesheetRowOf: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => EjoosTimesheetRow | null | undefined;
};

/** Скасування СЗЧ → absent_close + відновлення «+» у Табелі. */
export const planSzchAbsentCloseOps = (
  input: SzchAbsentClosePlanInput,
): { ops: EjoosSyncOp[]; absenceRowsClosedByMovement: Set<number> } => {
  const ops: EjoosSyncOp[] = [];
  const absenceRowsClosedByMovement = new Set<number>();
  const latestSzchCancellationByPerson = new Map<string, PbMovement>();
  for (const event of input.activeMovementsAll) {
    if (!isSzchCancellation(event)) continue;
    const key = event.personId
      ? `id:${event.personId}`
      : `name:${normKey(event.fullName)}`;
    if (!key.endsWith(":")) latestSzchCancellationByPerson.set(key, event);
  }
  for (const event of latestSzchCancellationByPerson.values()) {
    if (!input.eventInLeadWindow(event)) continue;
    const open =
      (event.personId && input.openById.get(event.personId)) ||
      input.openByName.get(normKey(event.fullName)) ||
      null;
    const person = input.shPeople.find(
      (candidate) =>
        (event.personId &&
          candidate.personId &&
          event.personId === candidate.personId) ||
        normKey(candidate.fullName) === normKey(event.fullName),
    );
    const timesheetRow =
      input.activeTimesheetRowOf(
        event.personId,
        event.fullName,
        person?.positionIndex || "",
      ) ||
      (person?.positionIndex &&
        input.dayByIndex.get(person.positionIndex)) ||
      input.ejoosDays.find(
        (row) => normKey(row.fullName) === normKey(event.fullName),
      ) ||
      null;
    if (!open && timesheetRow?.dayValue === "+") continue;
    const returnDate = event.orderDate || input.timesheetDayLabel;
    if (open) absenceRowsClosedByMovement.add(open.excelRow);
    ops.push({
      id: opId([
        "szch_cancel",
        event.personId || event.fullName,
        String(event.excelRow),
      ]),
      kind: "absent_close",
      class: timesheetRow ? "ready" : "needs_input",
      sheet: open ? "5. Тимчасово відсутні / 6. Табель" : "6. Табель",
      personId: event.personId || open?.personId || "",
      fullName: event.fullName || open?.fullName || "",
      positionIndex: person?.positionIndex || open?.positionIndex || "",
      rank: person?.rank || event.rank || open?.rank || "",
      before: open
        ? `СЗЧ відкрито з ${open.departDate || "?"}`
        : "відкритого періоду СЗЧ у «Тимч. відсутні» немає",
      after: `скасування СЗЧ / повернення ${returnDate}`,
      sourceRef: `Рух!R${event.excelRow} СТАТУС=«${event.status}»`,
      why: open
        ? "Скасування СЗЧ закриває період тимчасової відсутності та відновлює «+» у Табелі"
        : "Відкритого рядка СЗЧ немає: ШПО/ООС не змінюємо, відновлюємо «+» у Табелі",
      confidence: timesheetRow ? "high" : "manual",
      payload: {
        excelRow: open ? String(open.excelRow) : "",
        returnDate,
        timesheetExcelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
        returnDay: returnDate,
        movementNumber: event.movementNumber,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
      },
      movementKey: createMovementKey(event),
      checkedDefault: Boolean(timesheetRow),
    });
  }
  return { ops, absenceRowsClosedByMovement };
};
