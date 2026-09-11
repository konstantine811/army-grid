import { createMovementKey } from "../parse/pb";
import { canonicalName, normKey } from "../parse/cellText";
import { isFalseInternalHopExclusion } from "../../ejoosExcludePolicy";
import { findDuplicateTimesheetExtras } from "../../ejoosTimesheetDuplicates";
import type { PbMovement, PbShPerson } from "../types/pb";
import type {
  EjoosExcludedRow,
  EjoosTimesheetPersonScan,
  EjoosTimesheetRow,
} from "../types/workbookRows";
import type { EjoosSyncOp } from "../types/syncOp";
import { opId } from "./opId";

export type TimesheetDayCleanupInput = {
  ops: EjoosSyncOp[];
  timesheetPeople: EjoosTimesheetPersonScan[];
  ejoosExcluded: EjoosExcludedRow[];
  dayByIndex: Map<string, EjoosTimesheetRow>;
  timesheetScanByRow: Map<number, EjoosTimesheetPersonScan>;
  shPeople: PbShPerson[];
  shPersonById: Map<string, PbShPerson>;
  shPersonByName: Map<string, PbShPerson>;
  excludeAlreadyPlanned: (personId: string, fullName: string) => boolean;
  finalExternalTransferFor: (
    personId: string,
    fullName: string,
  ) => PbMovement | null;
  considerMovement: (event: PbMovement) => void;
  personStillInSh: (personId: string, fullName: string) => boolean;
  transferCancelForPerson: (
    personId: string,
    fullName: string,
  ) => PbMovement | null | undefined;
  eventInLeadWindow: (event: PbMovement) => boolean;
  absenceOnlyBlocksExclusion: (input: {
    absenceAt: number;
    outboundAt: number;
  }) => boolean;
  openDispositionAbsenceMs: (personId: string, fullName: string) => number;
  movementEventTime: (event: PbMovement) => number;
  timesheetRowsOf: (
    personId: string,
    fullName: string,
  ) => EjoosTimesheetPersonScan[];
  excludedRowsToClear: (payload: EjoosSyncOp["payload"]) => number[];
  staleExcludedForMovement: (event: PbMovement) => EjoosExcludedRow[];
  staleExcludedClearPayload: (
    stale: Array<{ excelRow: number }>,
  ) => Record<string, string>;
  latestOutboundTransferOf: (
    personId: string,
    fullName: string,
  ) => PbMovement | null;
  activeTimesheetRowOf: (
    personId: string,
    fullName: string,
    staffIndex: string,
  ) => EjoosTimesheetRow | null | undefined;
  byPersonName: <T>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => T | null | undefined;
  personStillInEjoos: (personId: string, fullName: string) => boolean;
  isSamePerson: (
    left: { personId?: string; fullName?: string },
    right: { personId?: string; fullName?: string },
  ) => boolean;
  movementPersonKey: (event: { personId: string; fullName: string }) => string;
};

/** Cleanup timesheet_day: stale tab rows, duplicate exclusions, dup tab rows. */
export const planTimesheetDayCleanup = (input: TimesheetDayCleanupInput): void => {
  const staffTimesheetRows = new Set(
    [...input.dayByIndex.values()].map((row) => row.excelRow),
  );

  for (const row of input.timesheetPeople) {
    if (!staffTimesheetRows.has(row.excelRow)) continue;
    if (input.personStillInSh(row.personId, row.fullName)) continue;
    if (input.transferCancelForPerson(row.personId, row.fullName)) continue;
    const transfer = input.finalExternalTransferFor(row.personId, row.fullName);
    if (
      input.absenceOnlyBlocksExclusion({
        absenceAt: input.openDispositionAbsenceMs(row.personId, row.fullName),
        outboundAt: transfer ? input.movementEventTime(transfer) : 0,
      })
    ) {
      continue;
    }
    if (!transfer) continue;
    if (!input.excludeAlreadyPlanned(row.personId, row.fullName)) {
      input.considerMovement(transfer);
    }
    if (input.excludeAlreadyPlanned(row.personId, row.fullName)) continue;
    const history = input.timesheetRowsOf(row.personId, row.fullName).find(
      (other) => other.excelRow !== row.excelRow && other.hasDepartureText,
    );
    input.ops.push({
      id: opId([
        "stale-tab",
        row.personId || row.fullName,
        String(row.excelRow),
      ]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: row.personId,
      fullName: row.fullName,
      positionIndex: row.positionIndex,
      rank: row.rank,
      before: `штатний рядок R${row.excelRow}: ${row.fullName || "ПІБ"} ще стоїть на ${row.positionIndex}`,
      after:
        "прибрати ПІБ/ID зі штатної позиції; історичний рядок з вибуттям лишити",
      sourceRef: history
        ? `Табель!R${row.excelRow} · історія R${history.excelRow}`
        : `Табель!R${row.excelRow}`,
      why: "Особа відсутня в актуальному sh і має чинне зовнішнє переведення, але персональні дані лишилися на старій штатній позиції Табеля. Історичний рядок — не поточний стан; штатний рядок треба очистити.",
      confidence: "high",
      payload: {
        type: "STALE_TAB_PERSON_ROW",
        clearStalePerson: "1",
        excelRow: String(row.excelRow),
        keepHistoryRow: history ? String(history.excelRow) : "",
      },
      movementKey: createMovementKey(transfer),
      checkedDefault: true,
    });
  }

  const excludedClearAlreadyPlanned = (excelRow: number) =>
    input.ops.some((op) => input.excludedRowsToClear(op.payload).includes(excelRow));

  const seenStaleExcludedPeople = new Set<string>();
  for (const row of input.ejoosExcluded) {
    const key = input.movementPersonKey(row);
    if (!key || seenStaleExcludedPeople.has(key)) continue;
    seenStaleExcludedPeople.add(key);
    if (input.personStillInSh(row.personId, row.fullName)) continue;
    const latest = input.latestOutboundTransferOf(row.personId, row.fullName);
    if (!latest) continue;
    const stale = input.staleExcludedForMovement(latest).filter(
      (item) => !excludedClearAlreadyPlanned(item.excelRow),
    );
    if (!stale.length) continue;
    input.ops.push({
      id: opId([
        "clear-stale-excl",
        row.personId || row.fullName,
        stale.map((item) => String(item.excelRow)).join("-"),
      ]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "3. Виключені",
      personId: row.personId || latest.personId,
      fullName: row.fullName || latest.fullName,
      positionIndex: latest.previousIndex || "",
      rank: latest.rank || "",
      before: `Виключені ${stale.map((item) => `R${item.excelRow}`).join(", ")} — старі ПЕРЕВ`,
      after: `прибрати дублі; лишити чинне ПЕРЕВ №${latest.orderNumber || "?"} від ${latest.orderDate || "?"}`,
      sourceRef: stale.map((item) => `Виключені!R${item.excelRow}`).join(" · "),
      why: "Особа вже має чинне зовнішнє ПЕРЕВ. Попередні рядки «Виключені» не видалились — лишаємо один актуальний запис.",
      confidence: "high",
      payload: {
        type: "CLEAR_STALE_EXCLUSION_DUPLICATE",
        ...input.staleExcludedClearPayload(stale),
      },
      checkedDefault: true,
    });
  }

  for (const row of input.ejoosExcluded) {
    if (excludedClearAlreadyPlanned(row.excelRow)) continue;
    if (!isFalseInternalHopExclusion(row)) continue;
    if (!input.personStillInSh(row.personId, row.fullName)) continue;
    if (!input.personStillInEjoos(row.personId, row.fullName)) continue;
    const shPerson =
      (row.personId && input.shPersonById.get(row.personId)) ||
      input.byPersonName(input.shPersonByName, row.personId, row.fullName) ||
      null;
    const timesheetRow = input.activeTimesheetRowOf(
      row.personId,
      row.fullName,
      shPerson?.positionIndex || "",
    );
    input.ops.push({
      id: opId([
        "clear-hop-excl",
        row.personId || row.fullName,
        String(row.excelRow),
      ]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "3. Виключені",
      personId: row.personId || shPerson?.personId || "",
      fullName: row.fullName || shPerson?.fullName || "",
      positionIndex: shPerson?.positionIndex || "",
      rank: shPerson?.rank || "",
      before: `Виключені R${row.excelRow}: ${row.note || "ПЕРЕВЕДЕННЯ 1 ПБ"}`,
      after: "прибрати — особа досі в 1ПБ на штатній посаді",
      sourceRef: `Виключені!R${row.excelRow}`,
      why: "Внутрішня зміна посади 1ПБ потрапила в «Виключені» помилково. Людина є в актуальному sh і в ШПО/ООС — рядок виключення прибираємо.",
      confidence: "high",
      payload: {
        type: "CLEAR_INTERNAL_HOP_EXCLUSION",
        clearExcludedExcelRow: String(row.excelRow),
        excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
      },
      checkedDefault: true,
    });
  }

  for (const row of input.timesheetPeople) {
    if (staffTimesheetRows.has(row.excelRow)) continue;
    const cancel = input.transferCancelForPerson(row.personId, row.fullName);
    if (!cancel || !input.eventInLeadWindow(cancel)) continue;
    if (!input.personStillInSh(row.personId, row.fullName)) continue;
    const staff = row.positionIndex
      ? input.dayByIndex.get(row.positionIndex)
      : undefined;
    const staffScan = staff
      ? input.timesheetScanByRow.get(staff.excelRow)
      : undefined;
    if (staffScan?.hasDepartureText) continue;
    const shPerson =
      (row.personId && input.shPersonById.get(row.personId)) ||
      input.byPersonName(input.shPersonByName, row.personId, row.fullName) ||
      null;
    const personId = row.personId || shPerson?.personId || "";
    const fullName = row.fullName || shPerson?.fullName || "";
    const host =
      input.ops.find(
        (op) =>
          op.kind === "position_change" &&
          input.isSamePerson({ personId, fullName }, op),
      ) ||
      input.ops.find(
        (op) =>
          op.kind === "absent_upsert" &&
          input.isSamePerson({ personId, fullName }, op),
      ) ||
      input.ops.find(
        (op) =>
          op.kind !== "data_mismatch" &&
          op.kind !== "timesheet_day" &&
          input.isSamePerson({ personId, fullName }, op),
      );
    if (host) {
      host.payload.duplicateTimesheetExcelRow = String(row.excelRow);
      host.payload.clearStalePerson = "1";
      host.payload.clearTimesheetIndex = "1";
      continue;
    }
    input.ops.push({
      id: opId(["dup-tab-cancel", personId || fullName, String(row.excelRow)]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId,
      fullName,
      positionIndex: row.positionIndex || shPerson?.positionIndex || "",
      rank: row.rank || shPerson?.rank || "",
      before: `дубль R${row.excelRow}: ${fullName || "ПІБ"} · «вибув» після скасованого переведення`,
      after: "прибрати другий рядок — за серпень лишається один активний запис",
      sourceRef: `Табель!R${row.excelRow}`,
      why: "Переведення скасовано в тому ж місяці. Другий рядок Табеля з «вибув» не є окремою фактичною наявністю і не потрібен.",
      confidence: "high",
      payload: {
        type: "DUPLICATE_TAB_AFTER_CANCEL",
        clearStalePerson: "1",
        clearTimesheetIndex: "1",
        excelRow: String(row.excelRow),
      },
      movementKey: createMovementKey(cancel),
      checkedDefault: true,
    });
  }

  const timesheetOccupantByIndex = new Map(
    input.shPeople
      .filter((person) => person.positionIndex)
      .map((person) => [
        person.positionIndex,
        { personId: person.personId, fullName: person.fullName },
      ]),
  );
  const clearedTimesheetRows = new Set(
    input.ops
      .filter(
        (op) =>
          op.kind === "timesheet_day" &&
          op.payload.clearStalePerson === "1" &&
          Number(op.payload.excelRow || 0) > 0,
      )
      .map((op) => Number(op.payload.excelRow)),
  );
  const namesForTimesheetId = new Map<string, Set<string>>();
  for (const row of input.timesheetPeople) {
    const id = String(row.personId || "").trim();
    const name = canonicalName(row.fullName);
    if (!id || !name) continue;
    const names = namesForTimesheetId.get(id) ?? new Set<string>();
    names.add(name);
    namesForTimesheetId.set(id, names);
  }
  const uniqueTimesheetPersonId = (personId: string) => {
    const id = String(personId || "").trim();
    if (!id) return "";
    const names = namesForTimesheetId.get(id);
    if (!names || names.size > 1) return "";
    return id;
  };
  for (const item of findDuplicateTimesheetExtras(
    input.timesheetPeople,
    timesheetOccupantByIndex,
  )) {
    if (clearedTimesheetRows.has(item.extra.excelRow)) continue;
    clearedTimesheetRows.add(item.extra.excelRow);
    const keepName =
      item.keep.fullName || item.keep.personId || "канонічний рядок";
    input.ops.push({
      id: opId([
        "dup-tab-row",
        item.extra.personId || item.extra.fullName,
        String(item.extra.excelRow),
      ]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: uniqueTimesheetPersonId(item.extra.personId),
      fullName: item.extra.fullName,
      positionIndex: item.extra.positionIndex,
      rank: "",
      before: `дубль R${item.extra.excelRow}: ${item.extra.fullName || "ПІБ"} на ${item.extra.positionIndex || "індексі"}`,
      after: `прибрати рядок — лишається R${item.keep.excelRow} (${keepName})`,
      sourceRef: `Табель!R${item.extra.excelRow} · канон R${item.keep.excelRow}`,
      why:
        item.reason === "leftover_history"
          ? "Після вибуття в Табелі лишився другий іменний рядок. Історію з «вибув» лишаємо, копію з ПІБ прибираємо."
          : item.reason === "internal_hop"
            ? "Внутрішня зміна посади в 1ПБ: один місячний рядок. Копію з «вибув до відділення/взводу» прибираємо."
            : item.reason === "same_person"
              ? "У Табелі не може бути двох активних записів однієї особи. Повторне застосування раніше дописувало копію замість штатного рядка."
              : "Один штатний індекс — один активний рядок Табеля. Історія з «вибув» лишається, зайві копії з «+» прибираємо.",
      confidence: "high",
      payload: {
        type: "DUPLICATE_TAB_ROW",
        clearStalePerson: "1",
        clearTimesheetIndex: "1",
        excelRow: String(item.extra.excelRow),
        keepTimesheetExcelRow: String(item.keep.excelRow),
      },
      checkedDefault: true,
    });
  }
};
