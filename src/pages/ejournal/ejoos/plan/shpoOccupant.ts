import { normKey } from "../parse/cellText";
import type { PbMovement, PbShPerson } from "../types/pb";
import type { EjoosExcludedRow, EjoosOosRow, EjoosShpoRow, EjoosTimesheetRow } from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { isOwnUnitStaffMove } from "../../ejoosMovementRules";
import type { StaffEpisodePaintPayload } from "./tempArrivalClose";
import { opId } from "./opId";

export type ShpoOccupantPlanInput = {
  shPeople: PbShPerson[];
  existingOps: EjoosSyncOp[];
  shpoByIndex: Map<string, EjoosShpoRow>;
  dayByIndex: Map<string, EjoosTimesheetRow>;
  arrivalById: Map<string, unknown>;
  arrivalByName: Map<string, unknown>;
  oosPersonById: Map<string, EjoosOosRow>;
  oosPersonByName: Map<string, EjoosOosRow>;
  positionEventForShPerson: (person: PbShPerson) => unknown;
  alreadyVacatedForAbsence: (personId: string, fullName: string) => boolean;
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
  findLatestExcludedRow: (
    personId: string,
    fullName: string,
  ) => EjoosExcludedRow | null | undefined;
  personStillInEjoos: (personId: string, fullName: string) => boolean;
  transferCancelForPerson: (
    personId: string,
    fullName: string,
  ) => PbMovement | null | undefined;
  staffEpisodePaintPayload: (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => StaffEpisodePaintPayload;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
};

export type ShpoReconcilePlanInput = {
  shOccupantByIndex: Map<string, PbShPerson>;
  existingOps: EjoosSyncOp[];
  activeMovementsAll: PbMovement[];
  shpoByIndex: Map<string, EjoosShpoRow>;
  dayByIndex: Map<string, EjoosTimesheetRow>;
  arrivalById: Map<string, unknown>;
  arrivalByName: Map<string, unknown>;
  alreadyVacatedForAbsence: (personId: string, fullName: string) => boolean;
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
  eventInLeadWindow: (event: PbMovement) => boolean;
  movementEventTime: (event: PbMovement) => number;
  staffEpisodePaintPayload: (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => StaffEpisodePaintPayload;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
};

/** Звірка зайнятості ШПО/Табель з sh (до absent/timesheet фаз). */
export const planShpoOccupantOps = (input: ShpoOccupantPlanInput): EjoosSyncOp[] => {
  const ops: EjoosSyncOp[] = [];

  for (const person of input.shPeople) {
    if (!person.positionIndex || !person.fullName) continue;
    const pendingPositionOp = input.existingOps.some(
      (op) =>
        op.kind === "position_change" &&
        input.isSamePerson(person, op) &&
        op.positionIndex === person.positionIndex &&
        op.class === "ready",
    );
    if (input.positionEventForShPerson(person) || pendingPositionOp) continue;

    const shpo = input.shpoByIndex.get(person.positionIndex);
    const ts =
      input.staffIndexTimesheetForPerson(
        person.personId,
        person.fullName,
        person.positionIndex,
      ) || input.dayByIndex.get(person.positionIndex);
    if (!shpo && !ts) continue;

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

    const beforeName = shpo?.fullName || (!shpo ? ts?.fullName : "") || "—";
    const beforeRank = shpo?.rank || (!shpo ? ts?.rank : "") || "—";
    const beforeId = shpo?.personId || (!shpo ? ts?.personId : "") || "—";
    const afterName = person.fullName;
    const afterRank = person.rank || "—";
    const afterId = person.personId || "—";
    const occupantOos =
      (person.personId && input.oosPersonById.get(person.personId)) ||
      input.byPersonName(
        input.oosPersonByName,
        person.personId,
        person.fullName,
      ) ||
      null;
    const occupantExcluded = occupantOos
      ? null
      : input.findLatestExcludedRow(person.personId, person.fullName);

    const nameChanged = normKey(beforeName) !== normKey(afterName);
    const rankChanged = normKey(beforeRank) !== normKey(afterRank);
    const idChanged =
      Boolean(person.personId) &&
      beforeId !== "—" &&
      beforeId !== person.personId;
    const needsRestoreFromExcluded =
      !input.personStillInEjoos(person.personId, person.fullName) &&
      Boolean(
        occupantExcluded ||
          input.transferCancelForPerson(person.personId, person.fullName),
      );

    if (
      !nameChanged &&
      !rankChanged &&
      !idChanged &&
      !needsRestoreFromExcluded
    ) {
      continue;
    }
    if (rankChanged && !nameChanged && !idChanged) continue;

    if (idChanged && nameChanged) {
      ops.push({
        id: opId([
          "shpo_conflict",
          person.positionIndex,
          person.personId || afterName,
        ]),
        kind: "shpo_occupant",
        class: "conflict",
        sheet: "1. ШПО / 6. Табель",
        personId: person.personId,
        fullName: person.fullName,
        positionIndex: person.positionIndex,
        rank: person.rank,
        before: `${beforeRank} ${beforeName} (ID ${beforeId})`,
        after: `${afterRank} ${afterName} (ID ${afterId})`,
        sourceRef: `sh!R${person.excelRow} індекс ${person.positionIndex}`,
        why: "На тому ж індексі посади інший ID/ПІБ — потрібна ручна перевірка перед зміною зайнятості",
        confidence: "review",
        payload: {
          shpoExcelRow: shpo ? String(shpo.excelRow) : "",
          timesheetExcelRow: ts ? String(ts.excelRow) : "",
          nextName: afterName,
          nextRank: afterRank,
          nextPersonId: afterId,
          excludedSourceExcelRow: occupantExcluded
            ? String(occupantExcluded.excelRow)
            : "",
        },
        checkedDefault: false,
      });
      continue;
    }

    const episodePaint = input.staffEpisodePaintPayload(
      person.personId,
      person.fullName,
      person.positionIndex,
      ts?.excelRow || 0,
    );
    ops.push({
      id: opId(["shpo", person.positionIndex, person.personId || afterName]),
      kind: "shpo_occupant",
      class: "ready",
      sheet: "1. ШПО / 6. Табель",
      personId: person.personId,
      fullName: person.fullName,
      positionIndex: person.positionIndex,
      rank: person.rank,
      before: `${beforeRank} ${beforeName} (ID ${beforeId})`,
      after: `${afterRank} ${afterName} (ID ${afterId})`,
      sourceRef: `sh!R${person.excelRow} індекс ${person.positionIndex}`,
      why: needsRestoreFromExcluded
        ? `Особа є в sh на ${person.positionIndex}, але в активних ШПО/ООС її немає. Відновити зайнятість${occupantExcluded ? ` і картку ООС з «Виключені» R${occupantExcluded.excelRow}` : ""}`
        : "Зайнятість посади в ЕЖООС відрізняється від sh — оновити ПІБ/звання/ID",
      confidence: "high",
      payload: {
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
        nextName: afterName,
        nextRank: afterRank,
        nextPersonId: person.personId,
        excludedSourceExcelRow: occupantExcluded
          ? String(occupantExcluded.excelRow)
          : "",
        ...episodePaint,
        timesheetActiveFrom:
          episodePaint.timesheetActiveFrom ||
          (needsRestoreFromExcluded
            ? input.transferCancelForPerson(person.personId, person.fullName)
                ?.orderDate || ""
            : ""),
        timesheetPreserveHistory: needsRestoreFromExcluded
          ? "1"
          : episodePaint.timesheetPreserveHistory,
      },
      checkedDefault: true,
    });
  }

  return ops;
};

/** Фінальна звірка зайнятості штатного індексу з sh. */
export const planShpoReconcileOps = (
  input: ShpoReconcilePlanInput,
): EjoosSyncOp[] => {
  const ops: EjoosSyncOp[] = [];

  for (const [positionIndex, shPerson] of input.shOccupantByIndex) {
    const inTempArrivals = Boolean(
      (shPerson.personId && input.arrivalById.get(shPerson.personId)) ||
        input.byPersonName(
          input.arrivalByName,
          shPerson.personId,
          shPerson.fullName,
        ),
    );
    if (
      input.alreadyVacatedForAbsence(shPerson.personId, shPerson.fullName) &&
      !inTempArrivals
    ) {
      continue;
    }
    const shpo = input.shpoByIndex.get(positionIndex);
    const shpoMatches =
      Boolean(shpo) &&
      input.isSamePerson(shPerson, shpo!) &&
      normKey(shpo?.fullName || "") === normKey(shPerson.fullName);
    if (shpoMatches) continue;
    const hasShpoOp = input.existingOps.some(
      (op) =>
        (op.kind === "shpo_occupant" || op.kind === "position_change") &&
        input.isSamePerson(shPerson, op) &&
        op.positionIndex === positionIndex &&
        op.class === "ready",
    );
    if (hasShpoOp) continue;
    const ts =
      input.staffIndexTimesheetForPerson(
        shPerson.personId,
        shPerson.fullName,
        positionIndex,
      ) || input.dayByIndex.get(positionIndex);
    const returnEvent = [...input.activeMovementsAll]
      .filter(
        (event) =>
          input.isSamePerson(shPerson, event) &&
          input.eventInLeadWindow(event) &&
          isOwnUnitStaffMove(event) &&
          (event.nextIndex === positionIndex ||
            String(event.changeText || "").includes(positionIndex)),
      )
      .sort(
        (left, right) =>
          input.movementEventTime(right) - input.movementEventTime(left) ||
          right.excelRow - left.excelRow,
      )[0];
    const episodePaint = input.staffEpisodePaintPayload(
      shPerson.personId,
      shPerson.fullName,
      positionIndex,
      ts?.excelRow || 0,
    );
    ops.push({
      id: opId([
        "shpo-reconcile",
        positionIndex,
        shPerson.personId || shPerson.fullName,
      ]),
      kind: "shpo_occupant",
      class: "ready",
      sheet: "1. ШПО / 6. Табель",
      personId: shPerson.personId,
      fullName: shPerson.fullName,
      positionIndex,
      rank: shPerson.rank,
      before: shpo?.fullName
        ? `${shpo.rank || "?"} ${shpo.fullName} (ID ${shpo.personId || "—"})`
        : "вакантно",
      after: `${shPerson.rank || "?"} ${shPerson.fullName} (ID ${shPerson.personId || "—"})`,
      sourceRef: `sh!R${shPerson.excelRow} · фінальна зайнятість ${positionIndex}`,
      why: "Після всіх рухів на цьому індексі актуальна sh визначає поточного військовослужбовця. Відновлюємо ШПО/Табель за sh, історію попередніх осіб не чіпаємо.",
      confidence: "high",
      payload: {
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
        nextName: shPerson.fullName,
        nextRank: shPerson.rank,
        nextPersonId: shPerson.personId,
        reconcileFromSh: "1",
        ...episodePaint,
        timesheetActiveFrom:
          episodePaint.timesheetActiveFrom || returnEvent?.orderDate || "",
      },
      checkedDefault: true,
    });
  }

  return ops;
};
