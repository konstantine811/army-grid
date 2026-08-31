/**
 * Політика виключення / транзиту. Нові фічі чіпають цей файл,
 * а не розкидані `if` у плані та apply.
 */

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

/** Немає в ШПО/ООС = уже проведено, окрім транзиту цього місяця. */
export const skipExternalIfAlreadyProcessed = (input: {
  stillInEjoos: boolean;
  unrecordedTransit: boolean;
}) => !input.stillInEjoos && !input.unrecordedTransit;

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

export const excludeWritePlan = (payload: Record<string, string>) => {
  const transitSameMonth = payload.transitSameMonth === "1";
  const createTimesheetHistory = payload.timesheetCreateHistory === "1";
  return {
    transitSameMonth,
    createTimesheetHistory,
    matchShpoByIndex: !transitSameMonth,
    matchTimesheetByIndex: !createTimesheetHistory,
  };
};
