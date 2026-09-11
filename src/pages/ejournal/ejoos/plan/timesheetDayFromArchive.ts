import {
  clipAbsenceSpansToActiveEpisode,
  encodeTimesheetAbsenceSpans,
  journalDayFromDateMs,
  sameTimesheetDayMark,
  timesheetMarkFromArchive,
} from "../../ejoosTimesheetText";
import { dateMs } from "../parse/cellText";
import type { PbShPerson } from "../types/pb";
import type { EjoosTimesheetPersonScan, EjoosTimesheetRow } from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import type { StaffEpisodePaintPayload } from "./tempArrivalClose";
import { opId } from "./opId";

export type TimesheetDayFromArchiveInput = {
  shPeople: PbShPerson[];
  existingOps: EjoosSyncOp[];
  timesheetDay: number;
  leadWindowStart: number;
  timesheetScanByRow: Map<number, EjoosTimesheetPersonScan>;
  personStillInSh: (personId: string, fullName: string) => boolean;
  positionEventForShPerson: (person: PbShPerson) => unknown;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  augustAbsenceSpansFor: (
    personId: string,
    fullName: string,
  ) => Array<{ fromDay: number; toDay: number; code: string }>;
  staffAppointmentDateFor: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => string;
  inboundStaffDateFor: (personId: string, fullName: string) => string;
  activeTimesheetRowOf: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => EjoosTimesheetRow | null | undefined;
  staffEpisodePaintPayload: (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => StaffEpisodePaintPayload;
};

/** PAINT_ARCHIVE: табель за хронологією archive, не зі статусу sh. */
export const planTimesheetDayFromArchiveOps = (
  input: TimesheetDayFromArchiveInput,
): EjoosSyncOp[] => {
  const ops: EjoosSyncOp[] = [];
  const timesheetAlreadyPainted = (personId: string, fullName: string) =>
    [...input.existingOps, ...ops].some(
      (op) =>
        input.isSamePerson({ personId, fullName }, op) &&
        Boolean(
          op.payload.timesheetAbsenceSpans ||
            op.payload.timesheetActiveFrom ||
            op.payload.restoreTimesheet === "1",
        ),
    );

  for (const person of input.shPeople) {
    if (!input.personStillInSh(person.personId, person.fullName)) continue;
    if (Boolean(input.positionEventForShPerson(person))) continue;
    if (timesheetAlreadyPainted(person.personId, person.fullName)) continue;
    if (
      [...input.existingOps, ...ops].some(
        (op) =>
          op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT" &&
          input.isSamePerson(person, op),
      )
    ) {
      continue;
    }

    const spans = input.augustAbsenceSpansFor(person.personId, person.fullName);
    const appointmentDate = input.staffAppointmentDateFor(
      person.personId,
      person.fullName,
      person.positionIndex,
    );
    const inboundDate = input.inboundStaffDateFor(
      person.personId,
      person.fullName,
    );
    const active = input.activeTimesheetRowOf(
      person.personId,
      person.fullName,
      person.positionIndex,
    );
    if (!active) continue;

    const episodePaint = input.staffEpisodePaintPayload(
      person.personId,
      person.fullName,
      person.positionIndex,
      active.excelRow,
    );
    const activeFromLabel =
      episodePaint.timesheetActiveFrom || appointmentDate || inboundDate || "";
    let activeFromDay =
      journalDayFromDateMs(dateMs(activeFromLabel), input.leadWindowStart) || 1;
    const scan = input.timesheetScanByRow.get(active.excelRow);
    if (
      scan?.hasDepartureText &&
      !scan.plusDays.some((day) => day >= Math.max(1, activeFromDay))
    ) {
      continue;
    }
    if (activeFromDay <= 1 && scan?.plusDays.length) {
      const firstPlus = Math.min(...scan.plusDays);
      const absenceEndedBeforePlus =
        spans.length > 0 &&
        firstPlus > 1 &&
        spans.every((span) => span.toDay < firstPlus);
      const prefixInactive =
        firstPlus > 1 &&
        Array.from({ length: firstPlus - 1 }, (_, index) => index + 1).every(
          (day) => {
            const actual = (scan.dayCodes[day] || "").trim();
            return (
              !actual || actual === "вибув" || sameTimesheetDayMark(actual, "-")
            );
          },
        );
      if (absenceEndedBeforePlus && prefixInactive) {
        activeFromDay = firstPlus;
      }
    }
    const episodeSpans =
      activeFromDay > 1
        ? clipAbsenceSpansToActiveEpisode(spans, activeFromDay)
        : spans;
    const firstPlusDay = scan?.plusDays.length ? Math.min(...scan.plusDays) : 0;
    const falseInactivePrefix =
      activeFromDay <= 1 &&
      !spans.length &&
      firstPlusDay > 1 &&
      Array.from({ length: firstPlusDay - 1 }, (_, index) => index + 1).some(
        (day) => sameTimesheetDayMark(scan?.dayCodes[day] || "", "-"),
      );
    if (
      !episodeSpans.length &&
      activeFromDay <= 1 &&
      !spans.length &&
      !falseInactivePrefix
    ) {
      continue;
    }

    let mismatch = false;
    for (let day = 1; day <= input.timesheetDay; day += 1) {
      const expected = timesheetMarkFromArchive(day, {
        activeFromDay,
        lastDay: input.timesheetDay,
        spans: episodeSpans,
        fillBeforeActive: activeFromDay > 1,
      });
      if (!expected) continue;
      const actual = (scan?.dayCodes[day] || "").trim();
      if (actual === "вибув") continue;
      if (sameTimesheetDayMark(actual, expected)) continue;
      mismatch = true;
      break;
    }
    if (!mismatch) continue;

    ops.push({
      id: opId(["ts_archive", person.personId || person.fullName]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: person.personId,
      fullName: person.fullName,
      positionIndex: person.positionIndex,
      rank: person.rank,
      before: "позначки днів не збігаються з archive",
      after:
        activeFromDay > 1
          ? `штатний рядок з ${activeFromLabel}: до постановки «-», далі «+»`
          : episodeSpans.length
            ? `фактичні коди з archive (${episodeSpans.map((span) => `${span.fromDay}–${span.toDay}:${span.code}`).join(", ")})`
            : `активний рядок з ${activeFromLabel}`,
      sourceRef: `archive + sh · Табель R${active.excelRow}`,
      why:
        activeFromDay > 1
          ? "Новий штатний епізод: коди відсутності на цей рядок не переносимо. СЗЧ лишається в «Тимч. відсутні» та на історичному рядку Табеля, якщо він є."
          : "Табель ведемо за фактичною хронологією archive, не з РУХ. Історичний рядок з вибуттям не перераховуємо.",
      confidence: "high",
      payload: {
        type: "PAINT_ARCHIVE",
        excelRow: String(active.excelRow),
        ...episodePaint,
        timesheetActiveFrom:
          episodePaint.timesheetActiveFrom || activeFromLabel,
        timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(episodeSpans),
        timesheetPreserveHistory:
          activeFromDay > 1 ? "1" : episodePaint.timesheetPreserveHistory,
        historyDepartDate: "",
        historyOrderNumber: "",
        historyDepartDest: "",
      },
      checkedDefault: true,
    });
  }

  return ops;
};
