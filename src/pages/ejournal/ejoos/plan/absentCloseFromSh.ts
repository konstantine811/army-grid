import { encodeTimesheetAbsenceSpans } from "../../ejoosTimesheetText";
import { canonicalName } from "../parse/cellText";
import type { PbArchivePeriod, PbShPerson } from "../types/pb";
import type { EjoosAbsentRow, EjoosArrivalRow } from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { hasActualReturn } from "./movementContext";
import { opId } from "./opId";
import type { EjoosStatusMapping } from "../../ejoosStatusMap";

export type AbsentCloseFromShInput = {
  shPeople: PbShPerson[];
  archive: PbArchivePeriod[];
  openById: Map<string, EjoosAbsentRow>;
  openByName: Map<string, EjoosAbsentRow>;
  arrivalById: Map<string, EjoosArrivalRow>;
  arrivalByName: Map<string, EjoosArrivalRow>;
  absenceRowsClosedByMovement: Set<number>;
  timesheetDayLabel: string;
  mapStatus: (raw: string) => EjoosStatusMapping;
  dateInLeadWindow: (date: string) => boolean;
  alreadyVacatedForAbsence: (personId: string, fullName: string) => boolean;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => T | null | undefined;
  activeTimesheetRowOf: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => { excelRow: number } | null | undefined;
  augustAbsenceSpansFor: (
    personId: string,
    fullName: string,
  ) => Array<{ fromDay: number; toDay: number; code: string }>;
};

/** Закриття «Тимч. відсутні», коли в sh знову «в строю» (+). */
export const planAbsentCloseFromShOps = (
  input: AbsentCloseFromShInput,
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
    if (mapped.timesheetCode !== "+") continue;

    const open =
      (person.personId && input.openById.get(person.personId)) ||
      input.byPersonName(input.openByName, person.personId, person.fullName) ||
      input.openByName.get(canonicalName(person.fullName)) ||
      null;
    const archiveReturn = input.archive.find(
      (period) =>
        input.isSamePerson(person, period) &&
        hasActualReturn(period.returnDate) &&
        (canonicalName(period.absenceType) ===
          canonicalName(open?.ground || "") ||
          !open?.ground),
    );
    if (
      !open ||
      input.absenceRowsClosedByMovement.has(open.excelRow) ||
      !(
        input.dateInLeadWindow(open.departDate) ||
        (archiveReturn && input.dateInLeadWindow(archiveReturn.returnDate))
      )
    ) {
      continue;
    }

    const returnDate = archiveReturn?.returnDate || input.timesheetDayLabel;
    const returnOrder = [
      archiveReturn?.returnOrderDate,
      archiveReturn?.returnOrderNumber,
    ]
      .filter(Boolean)
      .join(" ");
    ops.push({
      id: opId([
        "absent_close",
        open.personId || open.fullName,
        String(open.excelRow),
      ]),
      kind: "absent_close",
      class: "ready",
      sheet: "5. Тимчасово відсутні",
      personId: person.personId || open.personId,
      fullName: person.fullName || open.fullName,
      positionIndex: person.positionIndex || open.positionIndex,
      rank: person.rank,
      before: `відкрито: ${open.ground || "—"} з ${open.departDate || "?"} → ${open.place || "?"}`,
      after: returnOrder
        ? `фактичне прибуття: ${returnDate} · ${returnOrder}`
        : `фактичне прибуття: ${returnDate}`,
      sourceRef: archiveReturn
        ? `archive!R${archiveReturn.excelRow} → повернення ${returnDate}`
        : `sh!R${person.excelRow} → В СТРОЮ; ЕЖООС sheet5 R${open.excelRow}`,
      why: archiveReturn
        ? "У archive є фактичне повернення — закриваємо «Тимч. відсутні» цією датою і наказом, не днем зрізу sh"
        : "У 1ПБ знову «в строю», а в «Тимч. відсутні» період ще відкритий",
      confidence: "high",
      payload: {
        excelRow: String(open.excelRow),
        returnDate,
        returnOrderNumber: archiveReturn?.returnOrderNumber || "",
        returnOrderDate: archiveReturn?.returnOrderDate || "",
        timesheetExcelRow: String(
          input.activeTimesheetRowOf(
            person.personId,
            person.fullName,
            person.positionIndex,
          )?.excelRow || "",
        ),
        returnDay: returnDate,
        timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
          input.augustAbsenceSpansFor(person.personId, person.fullName),
        ),
        timesheetActiveFrom: "",
        timesheetSkipHistory: "1",
      },
      checkedDefault: true,
    });
    input.absenceRowsClosedByMovement.add(open.excelRow);
  }

  return ops;
};
