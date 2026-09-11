import type { EjoosSyncPlan } from "../types/syncOp";

export const SOURCE_DATE_UNKNOWN_MESSAGE =
  "SOURCE_DATE_UNKNOWN: не вдалося визначити дату 1ПБ. Вкажіть «станом на» вручну.";

export const TIMESHEET_MONTH_HEADER_UNKNOWN_MESSAGE =
  "TIMESHEET_MONTH_HEADER_UNKNOWN: заголовок місяця в I2 не прочитано. Місяць беремо з дати 1ПБ.";

export const MONTH_ROLLOVER_BLOCK_MESSAGE =
  "MONTH_ROLLOVER_REQUIRED: місяць Табеля ЕЖООС не збігається з місяцем 1ПБ. Заголовок не перейменовуємо — потрібен новий файл ЕЖООС на цей місяць.";

export const workbookApplyBlockMessage = (
  plan:
    | Pick<
        EjoosSyncPlan,
        | "monthRolloverRequired"
        | "sourceDateUnknown"
        | "timesheetMonthHeaderUnknown"
      >
    | null
    | undefined,
) => {
  if (!plan) return "";
  if (plan.sourceDateUnknown) return SOURCE_DATE_UNKNOWN_MESSAGE;
  return "";
};

export const planBlocksWorkbookApply = (
  plan:
    | Pick<
        EjoosSyncPlan,
        | "monthRolloverRequired"
        | "sourceDateUnknown"
        | "timesheetMonthHeaderUnknown"
      >
    | null
    | undefined,
) => Boolean(workbookApplyBlockMessage(plan));

/**
 * Горизонт Табеля = дата джерела 1ПБ. Не дотягуємо до «сьогодні»:
 * у файлі за 25.08 немає даних за 26–31, тож `+` туди не вигадуємо.
 */
export const refreshPlanTimesheetHorizon = (
  plan: Pick<EjoosSyncPlan, "timesheetDay" | "timesheetDayLabel">,
  _now = new Date(),
): Pick<EjoosSyncPlan, "timesheetDay" | "timesheetDayLabel"> => ({
  timesheetDay: plan.timesheetDay,
  timesheetDayLabel: plan.timesheetDayLabel,
});
