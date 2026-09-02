/**
 * Політика виключення / транзиту. Нові фічі чіпають цей файл,
 * а не розкидані `if` у плані та apply.
 */

import { normKey } from "./ejoosIdentity";
import {
  isInternalStaffIndexHop,
  mentionsExternalMilitaryUnit,
  mentionsForeignUnit,
} from "./ejoosMovementRules";

export const isUnrecordedSameMonthTransit = (input: {
  hasInboundPlacement: boolean;
  alreadyExcluded: boolean;
  stillInSh: boolean;
  stillInEjoos: boolean;
}) =>
  input.hasInboundPlacement &&
  !input.alreadyExcluded &&
  !input.stillInSh &&
  !input.stillInEjoos;

/**
 * Внутрішня ПОСАДА / 1ПБ→1ПБ не ставиться на штат, якщо далі по ланцюжку є
 * вибуття з частини, а в актуальній sh людини вже немає.
 */
export const ownUnitMoveSupersededByOutbound = (input: {
  stillInSh: boolean;
  hasLaterOutbound: boolean;
}) => input.hasLaterOutbound && !input.stillInSh;

/**
 * Закриття старої штатної посади всередині 1ПБ (ПОСАДА / ПЕРЕВ 1ПБ→1ПБ)
 * не створює рядок у «Виключені». Особа лишається в ООС.
 */
export const positionCloseWritesExcluded = (payload: {
  closeOldPosition?: string;
  internalStaffHop?: string;
  previousIndex?: string;
  fromPositionIndex?: string;
  nextIndex?: string;
  changeText?: string;
  documentsDest?: string;
  destination?: string;
  statusRaw?: string;
}) => {
  if (payload.closeOldPosition !== "1") return false;
  if (payload.internalStaffHop === "1") return false;
  // Старі протоколи могли не записати прапорець — індекси 1ПБ→1ПБ все одно
  // не є виключенням зі списків.
  if (
    isInternalStaffIndexHop({
      type: "ПОСАДА",
      previousIndex: payload.previousIndex || payload.fromPositionIndex || "",
      nextIndex: payload.nextIndex || "",
      destination: payload.statusRaw || payload.destination || "",
      changeText: payload.changeText || payload.documentsDest || "",
      note: "",
    })
  ) {
    return false;
  }
  return true;
};

/**
 * Рядок «Виключені» від внутрішньої ПОСАДИ 1ПБ (ПЕРЕВЕДЕННЯ 1 ПБ),
 * поки людина досі в актуальному sh. Це не вибуття з частини.
 */
export const isFalseInternalHopExclusion = (input: {
  destination?: string;
  note?: string;
}) => {
  const destination = String(input.destination || "");
  const note = String(input.note || "");
  const blob = `${destination} ${note}`;
  if (mentionsExternalMilitaryUnit(blob) || mentionsForeignUnit(blob)) {
    return false;
  }
  return /ПЕРЕВЕДЕНН/iu.test(note) && /1\s*ПБ/iu.test(blob);
};

export type ExternalTransferProcessState =
  | "ALREADY_PROCESSED"
  | "PARTIALLY_PROCESSED"
  | "NOT_PROCESSED";

/**
 * Зовнішнє ПЕРЕВ готове лише коли виконані всі 4 інваріанти.
 * 3/4 — PARTIALLY_PROCESSED: треба дописати відсутній аркуш, не NO_ACTION.
 */
export const externalTransferProcessState = (input: {
  onStaffShpo: boolean;
  onStaffOos: boolean;
  hasMatchingExcluded: boolean;
  timesheetClosed: boolean;
}): ExternalTransferProcessState => {
  const done = [
    !input.onStaffShpo,
    !input.onStaffOos,
    input.hasMatchingExcluded,
    input.timesheetClosed,
  ].filter(Boolean).length;
  if (done === 4) return "ALREADY_PROCESSED";
  if (done === 0) return "NOT_PROCESSED";
  return "PARTIALLY_PROCESSED";
};

/** Немає в ШПО/ООС = уже проведено лише якщо всі аркуші збіглися. */
export const skipExternalIfAlreadyProcessed = (input: {
  stillInEjoos: boolean;
  unrecordedTransit: boolean;
  processState?: ExternalTransferProcessState;
}) => {
  if (input.unrecordedTransit) return false;
  if (input.processState) return input.processState === "ALREADY_PROCESSED";
  return !input.stillInEjoos;
};

/**
 * Відкрита відсутність + немає на штаті ШПО = проведений вивід,
 * лише якщо людини вже немає в актуальному sh. Якщо sh каже «в строю» —
 * це повернення, і СЗЧ треба закрити, а не вважати стан завершеним.
 */
export const isStaleVacatedAbsence = (input: {
  onStaffShpo: boolean;
  hasOpenAbsence: boolean;
  stillInSh: boolean;
}) => input.hasOpenAbsence && !input.onStaffShpo && !input.stillInSh;

/** Вибуття 05.08 не виключає, якщо 07.08 людина вже повернулась у частину. */
export const laterReturnSupersedesOutbound = (input: {
  stillInSh: boolean;
  returnedAfterOutbound: boolean;
}) => input.stillInSh || input.returnedAfterOutbound;

/**
 * Відкритий БЕЗВІСТИ / СЗЧ не виключає зі списків.
 * Пізніший зовнішній ПЕРЕВ / ЗВІЛЬН після дати відсутності — окрема кадрова подія.
 */
export const absenceOnlyBlocksExclusion = (input: {
  absenceAt: number;
  outboundAt: number;
}) =>
  input.absenceAt > 0 &&
  (input.outboundAt <= 0 || input.absenceAt >= input.outboundAt);

/**
 * Попередні рядки «Виключені» тієї ж особи, які не збігаються з чинним
 * наказом ПЕРЕВ. При новому переведенні їх треба прибрати — лишається один.
 */
export const staleExcludedRowsToClear = (input: {
  rows: Array<{ excelRow: number; orderNumber?: string; orderDate?: string }>;
  currentOrderNumber?: string;
  currentOrderDate?: string;
}) => {
  const order = String(input.currentOrderNumber || "").trim();
  const date = String(input.currentOrderDate || "").trim();
  const keep =
    [...input.rows]
      .filter((row) => {
        if (order && row.orderNumber) {
          return normKey(row.orderNumber) === normKey(order);
        }
        if (!order && date && row.orderDate) {
          return normKey(row.orderDate) === normKey(date);
        }
        return false;
      })
      .sort((left, right) => right.excelRow - left.excelRow)[0] ?? null;
  return input.rows
    .filter((row) => row.excelRow !== keep?.excelRow)
    .sort((left, right) => left.excelRow - right.excelRow);
};

export const excludedRowsToClear = (payload: {
  clearExcludedExcelRow?: string;
  clearExcludedExcelRows?: string;
}) => {
  const rows = new Set<number>();
  const one = Number(payload.clearExcludedExcelRow || 0);
  if (one > 0) rows.add(one);
  for (const part of String(payload.clearExcludedExcelRows || "").split(
    /[,\s;]+/,
  )) {
    const row = Number(part);
    if (row > 0) rows.add(row);
  }
  return [...rows].sort((left, right) => left - right);
};

export const excludeWritePlan = (payload: Record<string, string>) => {
  const transitSameMonth = payload.transitSameMonth === "1";
  const createTimesheetHistory = payload.timesheetCreateHistory === "1";
  const replaceInPlace = payload.timesheetReplaceInPlace === "1";
  return {
    transitSameMonth,
    createTimesheetHistory,
    replaceInPlace,
    matchShpoByIndex: !transitSameMonth,
    matchTimesheetByIndex: !createTimesheetHistory,
  };
};

/**
 * Табель при виключенні — як при внутрішньому ПЕРЕВ/ПОСАДА:
 * закритий епізод копіюємо вниз у межах роти, штатний рядок лишає
 * індекс / ВОС / тарифний план без особи й днів.
 * Уже наявний «вибув»-екстра оновлюємо на місці, щоб не плодити дублі.
 */
export const excludedTimesheetWrite = (
  namedRows: Array<{
    excelRow: number;
    hasDepartureText: boolean;
  }>,
  vacantStaffExcelRow = 0,
) => {
  const history = namedRows.filter((row) => row.hasDepartureText);
  const active = namedRows.filter((row) => !row.hasDepartureText);
  if (active[0]) {
    return {
      createHistory: false,
      replaceInPlace: false,
      sourceExcelRow: active[0].excelRow,
    };
  }
  if (history[0]) {
    return {
      createHistory: false,
      replaceInPlace: true,
      sourceExcelRow: history[0].excelRow,
    };
  }
  if (vacantStaffExcelRow > 0) {
    return {
      createHistory: false,
      replaceInPlace: false,
      sourceExcelRow: vacantStaffExcelRow,
    };
  }
  return { createHistory: true, replaceInPlace: false, sourceExcelRow: 0 };
};
