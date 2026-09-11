import { createMovementKey, parseContractDatesFromChangeText, parseRankPromotion } from "../parse/pb";
import { dateMs, norm, normKey } from "../parse/cellText";
import type { PbMovement } from "../types/pb";
import type { EjoosOosRow, EjoosShpoRow, EjoosTimesheetRow } from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { isRankAssignmentEvent } from "./movementContext";
import { opId } from "./opId";

export type RankAndContractPlanInput = {
  activeMovementsAll: PbMovement[];
  ejoosDays: EjoosTimesheetRow[];
  eventInLeadWindow: (event: PbMovement) => boolean;
  movementPersonKey: (event: PbMovement) => string;
  movementEventTime: (event: PbMovement) => number;
  personStillInEjoos: (personId: string, fullName: string) => boolean;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => T | null | undefined;
  shpoPersonById: Map<string, EjoosShpoRow>;
  shpoPersonByName: Map<string, EjoosShpoRow>;
  oosPersonById: Map<string, EjoosOosRow>;
  oosPersonByName: Map<string, EjoosOosRow>;
  dayById: Map<string, EjoosTimesheetRow>;
  staffIndexTimesheetForPerson: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => EjoosTimesheetRow | null | undefined;
  isContractMovementType: (type: string) => boolean;
};

export type RankAndContractPlanResult = {
  ops: EjoosSyncOp[];
  pendingRankByPerson: Map<string, EjoosSyncOp>;
  latestRankEventByPerson: Map<string, PbMovement>;
};

export const planRankAndContractOps = (
  input: RankAndContractPlanInput,
): RankAndContractPlanResult => {
  const ops: EjoosSyncOp[] = [];
  const pendingRankByPerson = new Map<string, EjoosSyncOp>();

  const latestRankEventByPerson = new Map<string, PbMovement>();
  for (const event of input.activeMovementsAll) {
    if (!isRankAssignmentEvent(event) || !input.eventInLeadWindow(event)) continue;
    const key = input.movementPersonKey(event);
    if (!key) continue;
    const previous = latestRankEventByPerson.get(key);
    if (
      !previous ||
      input.movementEventTime(event) > input.movementEventTime(previous) ||
      (input.movementEventTime(event) === input.movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      latestRankEventByPerson.set(key, event);
    }
  }
  for (const event of latestRankEventByPerson.values()) {
    if (!input.personStillInEjoos(event.personId, event.fullName)) continue;
    const { previousRank, nextRank } = parseRankPromotion(event);
    if (!nextRank) continue;
    const shpo =
      (event.personId && input.shpoPersonById.get(event.personId)) ||
      input.byPersonName(
        input.shpoPersonByName,
        event.personId,
        event.fullName,
      ) ||
      null;
    const oos =
      (event.personId && input.oosPersonById.get(event.personId)) ||
      input.byPersonName(input.oosPersonByName, event.personId, event.fullName) ||
      null;
    const ts =
      input.staffIndexTimesheetForPerson(
        event.personId,
        event.fullName,
        shpo?.positionIndex || "",
      ) ||
      (event.personId && input.dayById.get(event.personId)) ||
      input.ejoosDays.find((row) => input.isSamePerson(event, row)) ||
      null;
    const currentRank = oos?.rank || shpo?.rank || ts?.rank || "";
    if (currentRank && normKey(currentRank) === normKey(nextRank)) continue;
    const canApply = Boolean(
      event.orderNumber && event.orderDate && (shpo || oos || ts),
    );
    const rankOp: EjoosSyncOp = {
      id: opId([
        "rank",
        event.personId || event.fullName,
        event.orderNumber || String(event.excelRow),
      ]),
      kind: "rank_change",
      class: canApply ? "ready" : "needs_input",
      sheet: "1. ШПО / 2. ООС / 6. Табель",
      personId: shpo?.personId || oos?.personId || event.personId,
      fullName: shpo?.fullName || oos?.fullName || event.fullName,
      positionIndex:
        shpo?.positionIndex ||
        oos?.positionIndex ||
        ts?.positionIndex ||
        event.previousIndex ||
        event.nextIndex,
      rank: nextRank,
      before: currentRank || previousRank || "—",
      after: `${nextRank} · наказ №${event.orderNumber || "?"} від ${event.orderDate || "?"}`,
      sourceRef: `Рух!R${event.excelRow} №${event.movementNumber} · ЗВАННЯ`,
      why: canApply
        ? "Серпневе присвоєння звання треба провести до виключення, інакше в «Виключені» піде старе звання"
        : "У Рух не вистачає номера/дати наказу про присвоєння звання",
      confidence: canApply ? "high" : "manual",
      payload: {
        previousRank: currentRank || previousRank,
        nextRank,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        oosExcelRow: oos ? String(oos.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
      },
      movementKey: createMovementKey(event),
      checkedDefault: canApply,
    };
    ops.push(rankOp);
    const key = input.movementPersonKey(event);
    if (key) pendingRankByPerson.set(key, rankOp);
    if (event.personId) pendingRankByPerson.set(`id:${event.personId}`, rankOp);
  }

  const latestContractEventByPerson = new Map<string, PbMovement>();
  for (const event of input.activeMovementsAll) {
    if (!input.isContractMovementType(event.type) || !input.eventInLeadWindow(event)) {
      continue;
    }
    const key = input.movementPersonKey(event);
    if (!key) continue;
    const previous = latestContractEventByPerson.get(key);
    if (
      !previous ||
      input.movementEventTime(event) > input.movementEventTime(previous) ||
      (input.movementEventTime(event) === input.movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      latestContractEventByPerson.set(key, event);
    }
  }
  for (const event of latestContractEventByPerson.values()) {
    if (!input.personStillInEjoos(event.personId, event.fullName)) continue;
    const oos =
      (event.personId && input.oosPersonById.get(event.personId)) ||
      input.byPersonName(input.oosPersonByName, event.personId, event.fullName) ||
      null;
    const parsedDates = parseContractDatesFromChangeText(event.changeText);
    const isMotivationContract = /МОТИВАЦ.*КОНТР/iu.test(event.type);
    const contractFrom =
      parsedDates.contractFrom || event.basisDate || event.orderDate;
    const contractTo =
      parsedDates.contractTo ||
      (isMotivationContract ? "" : norm(event.changeText));
    const serviceType = "контракт";
    const alreadyApplied = Boolean(
      oos &&
        normKey(oos.serviceType) === normKey(serviceType) &&
        dateMs(oos.contractFrom) === dateMs(contractFrom) &&
        normKey(oos.contractTo) === normKey(contractTo),
    );
    if (alreadyApplied) continue;
    const canApply = Boolean(oos && contractFrom && contractTo);
    ops.push({
      id: opId([
        "contract",
        event.personId || event.fullName,
        event.orderNumber || String(event.excelRow),
      ]),
      kind: "contract_update",
      class: canApply ? "ready" : "needs_input",
      sheet: "2. ООС",
      personId: oos?.personId || event.personId,
      fullName: oos?.fullName || event.fullName,
      positionIndex:
        oos?.positionIndex || event.nextIndex || event.previousIndex,
      rank: oos?.rank || event.rank,
      before: oos
        ? `${oos.serviceType || "вид служби не вказано"} · ${oos.contractFrom || "дата не вказана"} · ${oos.contractTo || "строк не вказаний"}`
        : "рядок ООС не знайдено",
      after: `${serviceType} · ${contractFrom} · ${contractTo}`,
      sourceRef: `Рух!R${event.excelRow} №${event.movementNumber} · КОНТРАКТ`,
      why: canApply
        ? "Подія КОНТРАКТ заповнює вид служби та строки контракту в колонках 17–19 аркуша ООС"
        : "Для події КОНТРАКТ не знайдено рядок ООС або дату/строк контракту",
      confidence: canApply ? "high" : "manual",
      payload: {
        oosExcelRow: oos ? String(oos.excelRow) : "",
        serviceType,
        contractFrom,
        contractTo,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
        basisNumber: event.basisNumber,
        basisDate: event.basisDate,
      },
      movementKey: createMovementKey(event),
      checkedDefault: canApply,
    });
  }

  return { ops, pendingRankByPerson, latestRankEventByPerson };
};
