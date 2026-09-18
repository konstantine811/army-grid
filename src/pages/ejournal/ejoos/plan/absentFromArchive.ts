import { archiveReturnContradictsCurrentSh } from "../../ejoosTimesheetText";
import { isDispositionAbsenceStatus } from "../../ejoosMovementRules";
import { canonicalName, dateMs, normKey } from "../parse/cellText";
import type { PbArchivePeriod, PbShPerson } from "../types/pb";
import type { EjoosAbsentRow } from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { hasActualReturn } from "./movementContext";
import { opId } from "./opId";
import type { StaffEpisodePaintPayload } from "./tempArrivalClose";

export type MapPbStatus = (raw: string) => {
  timesheetCode: string;
  ruleId?: string;
  reason?: string;
};

export type AbsentArchivePlanInput = {
  archive: PbArchivePeriod[];
  ejoosAbsents: EjoosAbsentRow[];
  openById: Map<string, EjoosAbsentRow>;
  openByName: Map<string, EjoosAbsentRow>;
  shPersonById: Map<string, PbShPerson>;
  shPersonByName: Map<string, PbShPerson>;
  absenceRowsClosedByMovement: Set<number>;
  existingOps: EjoosSyncOp[];
  mapStatus: MapPbStatus;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => T | null | undefined;
  laterArchivePeriodOf: (seed: {
    personId: string;
    fullName: string;
    departDate?: string;
  }) => PbArchivePeriod | null;
  alreadyVacatedForAbsence: (personId: string, fullName: string) => boolean;
  inboundStaffPlacementThisMonth: (personId: string, fullName: string) => boolean;
  activeTimesheetRowOf: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => { excelRow: number } | null | undefined;
  staffEpisodePaintPayload: (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => StaffEpisodePaintPayload;
};

/** absent_upsert / absent_close з archive та застарілих відкритих рядків «Тимч. відсутні». */
export const planAbsentArchiveOps = (input: AbsentArchivePlanInput): EjoosSyncOp[] => {
  const ops: EjoosSyncOp[] = [];
  const allOps = () => [...input.existingOps, ...ops];
  const hasAbsentCloseForRow = (excelRow: number) =>
    allOps().some(
      (op) =>
        op.kind === "absent_close" &&
        Number(op.payload.excelRow || 0) === excelRow,
    );

  input.archive.forEach((period) => {
    const open =
      (period.personId && input.openById.get(period.personId)) ||
      input.byPersonName(input.openByName, period.personId, period.fullName) ||
      input.openByName.get(normKey(period.fullName)) ||
      null;
    const sameAbsentPeriod = (row: EjoosAbsentRow) => {
      if (!input.isSamePerson(period, row)) return false;
      const sameKind =
        canonicalName(row.ground) === canonicalName(period.absenceType) ||
        (isDispositionAbsenceStatus(row.ground) &&
          isDispositionAbsenceStatus(period.absenceType));
      const sameDepart =
        Boolean(dateMs(row.departDate)) &&
        dateMs(row.departDate) === dateMs(period.departDate);
      return sameKind && (sameDepart || !row.departDate);
    };
    const recorded = input.ejoosAbsents.find(sameAbsentPeriod) ?? null;
    const reuseAbsent =
      recorded || (open && sameAbsentPeriod(open) ? open : null);
    if (
      recorded?.actualReturn &&
      dateMs(recorded.actualReturn) === dateMs(period.returnDate)
    ) {
      return;
    }
    const complete = Boolean(
      period.absenceType && (period.departDate || period.plannedReturn),
    );
    const before = reuseAbsent
      ? `${reuseAbsent.ground} / ${reuseAbsent.place} / ${reuseAbsent.departDate}`
      : "(немає запису в «Тимч. відсутні»)";
    const returnOrderText = [
      period.returnOrderDate,
      period.returnOrderNumber
        ? `№${period.returnOrderNumber.replace(/^№/i, "")}`
        : "",
    ]
      .filter(Boolean)
      .join(" ");
    const after = [
      `${period.absenceType || "?"} → ${period.place || "?"} з ${period.departDate || "?"}`,
      period.returnDate
        ? `повернення ${period.returnDate}${returnOrderText ? ` · ${returnOrderText}` : ""}`
        : "",
    ]
      .filter(Boolean)
      .join(" · ");
    const shPersonForPaint =
      (period.personId && input.shPersonById.get(period.personId)) ||
      input.byPersonName(
        input.shPersonByName,
        period.personId,
        period.fullName,
      ) ||
      null;
    const activeTimesheetForPaint = input.activeTimesheetRowOf(
      period.personId,
      period.fullName,
      shPersonForPaint?.positionIndex || "",
    );
    const episodePaintForRefresh = activeTimesheetForPaint?.excelRow
      ? input.staffEpisodePaintPayload(
          period.personId,
          period.fullName,
          shPersonForPaint?.positionIndex || "",
          activeTimesheetForPaint.excelRow,
        )
      : null;
    const needsTimesheetRefresh = Boolean(
      activeTimesheetForPaint?.excelRow &&
        (episodePaintForRefresh?.timesheetAbsenceSpans ||
          input.mapStatus(period.absenceType).timesheetCode),
    );
    const needsClose =
      hasActualReturn(period.returnDate) &&
      Boolean(reuseAbsent) &&
      dateMs(reuseAbsent?.actualReturn || "") !== dateMs(period.returnDate);
    const sameAbsenceKind =
      open &&
      (canonicalName(open.ground) === canonicalName(period.absenceType) ||
        (isDispositionAbsenceStatus(open.ground) &&
          isDispositionAbsenceStatus(period.absenceType)));
    const newerDifferentAbsence =
      open &&
      !sameAbsenceKind &&
      !hasActualReturn(open.actualReturn) &&
      Boolean(dateMs(open.departDate)) &&
      Boolean(dateMs(period.departDate)) &&
      dateMs(period.departDate) > dateMs(open.departDate);
    if (
      newerDifferentAbsence &&
      !input.absenceRowsClosedByMovement.has(open.excelRow) &&
      !hasAbsentCloseForRow(open.excelRow)
    ) {
      ops.push({
        id: opId([
          "close-superseded-absence",
          period.personId || period.fullName,
          String(open.excelRow),
          period.departDate,
        ]),
        kind: "absent_close",
        class: "ready",
        sheet: "5. Тимчасово відсутні",
        personId: period.personId || open.personId,
        fullName: period.fullName || open.fullName,
        positionIndex: open.positionIndex,
        rank: period.rank || open.rank,
        before: `${open.ground} з ${open.departDate || "?"} ще відкритий`,
        after: `закрити ${period.departDate} — далі ${period.absenceType || "інший період"}`,
        sourceRef: `archive!R${period.excelRow} · «Тимч. відсутні» R${open.excelRow}`,
        why: `Новий archive-період «${period.absenceType || "відсутність"}» починається ${period.departDate}; попередній «${open.ground || "період"}» не може залишатися відкритим паралельно.`,
        confidence: "high",
        payload: {
          type: "CLOSE_SUPERSEDED_OPEN_ABSENCE",
          excelRow: String(open.excelRow),
          returnDate: period.departDate,
          returnDay: period.departDate,
          timesheetSkipHistory: "1",
        },
        checkedDefault: true,
      });
    }
    if (sameAbsenceKind && !hasActualReturn(period.returnDate)) {
      const later = input.laterArchivePeriodOf(period);
      const shNow =
        (period.personId && input.shPersonById.get(period.personId)) ||
        input.byPersonName(
          input.shPersonByName,
          period.personId,
          period.fullName,
        ) ||
        null;
      const shPresent =
        Boolean(shNow) && input.mapStatus(shNow!.status).timesheetCode === "+";
      const openRow = reuseAbsent || open;
      if (
        later &&
        shPresent &&
        openRow &&
        !hasActualReturn(openRow.actualReturn)
      ) {
        ops.push({
          id: opId([
            "close-stale-abs",
            period.personId || period.fullName,
            String(openRow.excelRow),
          ]),
          kind: "absent_close",
          class: "ready",
          sheet: "5. Тимчасово відсутні",
          personId: period.personId || openRow.personId,
          fullName: period.fullName || openRow.fullName,
          positionIndex: shNow?.positionIndex || openRow.positionIndex,
          rank: shNow?.rank || period.rank || openRow.rank,
          before: `${openRow.ground} з ${openRow.departDate || "?"} ще відкритий`,
          after: `закрити ${later.departDate} — далі вже ${later.absenceType || "інший період"}`,
          sourceRef: `archive!R${period.excelRow} · «Тимч. відсутні» R${openRow.excelRow}`,
          why: `Старий відкритий «${openRow.ground}» перекритий пізнішим періодом ${later.absenceType || ""} з ${later.departDate}. У sh особа в строю — лікування в Табелі після відпустки не продовжуємо.`,
          confidence: "high",
          payload: {
            type: "CLOSE_SUPERSEDED_OPEN_ABSENCE",
            excelRow: String(openRow.excelRow),
            returnDate: later.departDate,
            returnDay: later.departDate,
            timesheetSkipHistory: "1",
          },
          checkedDefault: true,
        });
      }
      if (!needsTimesheetRefresh) return;
    }
    if (
      reuseAbsent &&
      input.absenceRowsClosedByMovement.has(reuseAbsent.excelRow)
    ) {
      return;
    }
    if (
      input.alreadyVacatedForAbsence(period.personId, period.fullName) &&
      open &&
      !input.inboundStaffPlacementThisMonth(period.personId, period.fullName)
    ) {
      return;
    }
    if (open && before === after && !needsTimesheetRefresh) return;
    if (
      recorded &&
      !hasActualReturn(period.returnDate) &&
      !needsTimesheetRefresh
    ) {
      return;
    }
    const shPerson = shPersonForPaint;
    const archiveReturnVsSh = archiveReturnContradictsCurrentSh(
      shPerson ? input.mapStatus(shPerson.status).timesheetCode : "",
      input.mapStatus(period.absenceType).timesheetCode,
      hasActualReturn(period.returnDate),
    );
    if (archiveReturnVsSh) {
      ops.push({
        id: opId([
          "archive_sh_return",
          period.personId || period.fullName,
          String(period.excelRow),
        ]),
        kind: "absent_upsert",
        class: "needs_input",
        sheet: "5. Тимч. відсутні / archive vs sh",
        personId: shPerson?.personId || period.personId,
        fullName: shPerson?.fullName || period.fullName,
        positionIndex: shPerson?.positionIndex || "",
        rank: period.rank,
        before,
        after: `archive: повернення ${period.returnDate}; sh досі ${shPerson?.status || period.absenceType}`,
        sourceRef: `archive!R${period.excelRow} · sh`,
        why: `NEEDS_REVIEW: archive вже має повернення ${period.returnDate}, а поточний sh досі ${shPerson?.status || "відсутній"}. Не закриваємо період і не тягнемо СЗЧ/ЗБ до дня звіту — перевірте, що саме застаріло.`,
        confidence: "manual",
        payload: {
          mismatchKind: "ARCHIVE_RETURN_SH_STILL_ABSENT",
          absenceType: period.absenceType,
          place: period.place,
          departDate: period.departDate,
          returnDate: period.returnDate,
          statusRaw: shPerson?.status || "",
          existingExcelRow: reuseAbsent ? String(reuseAbsent.excelRow) : "",
        },
        checkedDefault: false,
      });
      return;
    }
    const activeTs = activeTimesheetForPaint;
    const episodePaint =
      episodePaintForRefresh ??
      input.staffEpisodePaintPayload(
        period.personId,
        period.fullName,
        shPerson?.positionIndex || "",
        activeTs?.excelRow || 0,
      );
    const returnOrderNo = period.returnOrderNumber
      ? `№${period.returnOrderNumber.replace(/^№/i, "")}`
      : "";

    ops.push({
      id: opId([
        "absent_up",
        period.personId || period.fullName,
        period.periodNumber || String(period.excelRow),
      ]),
      kind: "absent_upsert",
      class: complete ? "ready" : "needs_input",
      sheet: "5. Тимч. відсутні",
      personId: shPerson?.personId || period.personId,
      fullName: shPerson?.fullName || period.fullName,
      positionIndex: shPerson?.positionIndex || "",
      rank: period.rank,
      before,
      after,
      sourceRef: `archive!R${period.excelRow} №${period.periodNumber || "—"}`,
      why: complete
        ? period.returnDate
          ? `Закрити період ${period.absenceType || "відсутності"} фактичним поверненням ${period.returnDate}${returnOrderNo ? ` ${returnOrderNo}` : ""}. Коди відсутності — у «Тимч. відсутні»${episodePaint.historyTimesheetExcelRow ? " і на історичному рядку Табеля" : ""}, не на новому штатному епізоді.`
          : needsTimesheetRefresh && before === after
            ? `«Тимч. відсутні» вже збігається з archive — оновити коди в Табелі (${episodePaint.timesheetAbsenceSpans || input.mapStatus(period.absenceType).timesheetCode || "відсутність"})`
            : "Відкритий період цього місяця з archive — внести у «Тимч. відсутні»"
        : "В archive неповні поля (дата/підстава) — дозаповніть перед застосуванням",
      confidence: complete ? "high" : "manual",
      payload: {
        absenceType: period.absenceType,
        place: period.place,
        departDate: period.departDate,
        orderNumber: period.orderNumber,
        orderDate: period.orderDate,
        plannedReturn: period.plannedReturn || "?",
        returnDate: period.returnDate,
        returnOrderNumber: period.returnOrderNumber,
        returnOrderDate: period.returnOrderDate,
        periodNumber: period.periodNumber,
        positionTitle: period.positionTitle,
        timesheetExcelRow: activeTs ? String(activeTs.excelRow) : "",
        timesheetSkipHistory: "1",
        ...episodePaint,
        historyDepartDate: "",
        historyOrderNumber: "",
        historyDepartDest: "",
        timesheetCode:
          dateMs(period.departDate) &&
          dateMs(period.returnDate) &&
          dateMs(period.returnDate) <= dateMs(period.departDate)
            ? ""
            : input.mapStatus(period.absenceType).timesheetCode || "",
        existingExcelRow: reuseAbsent ? String(reuseAbsent.excelRow) : "",
      },
      checkedDefault:
        complete &&
        (!reuseAbsent || needsClose || needsTimesheetRefresh),
    });
  });

  for (const row of input.ejoosAbsents.filter((item) => !item.actualReturn)) {
    if (input.absenceRowsClosedByMovement.has(row.excelRow)) continue;
    if (hasAbsentCloseForRow(row.excelRow)) continue;
    const shNow =
      (row.personId && input.shPersonById.get(row.personId)) ||
      input.byPersonName(input.shPersonByName, row.personId, row.fullName) ||
      null;
    if (!shNow || input.mapStatus(shNow.status).timesheetCode !== "+") continue;
    const later = input.laterArchivePeriodOf({
      personId: row.personId,
      fullName: row.fullName,
      departDate: row.departDate,
    });
    if (!later) continue;
    ops.push({
      id: opId([
        "close-stale-abs",
        row.personId || row.fullName,
        String(row.excelRow),
      ]),
      kind: "absent_close",
      class: "ready",
      sheet: "5. Тимчасово відсутні",
      personId: row.personId || shNow.personId,
      fullName: row.fullName || shNow.fullName,
      positionIndex: shNow.positionIndex || row.positionIndex,
      rank: shNow.rank || row.rank,
      before: `${row.ground} з ${row.departDate || "?"} ще відкритий`,
      after: `закрити ${later.departDate} — далі вже ${later.absenceType || "інший період"}`,
      sourceRef: `«Тимч. відсутні» R${row.excelRow} · archive далі ${later.absenceType || ""}`,
      why: `Старий відкритий «${row.ground}» перекритий пізнішим періодом ${later.absenceType || ""} з ${later.departDate}. У sh особа в строю — після відпустки в Табелі має бути «+», не «лік».`,
      confidence: "high",
      payload: {
        type: "CLOSE_SUPERSEDED_OPEN_ABSENCE",
        excelRow: String(row.excelRow),
        returnDate: later.departDate,
        returnDay: later.departDate,
        timesheetSkipHistory: "1",
      },
      checkedDefault: true,
    });
  }

  return ops;
};
