/**
 * Прийменник для Табеля: «вибув до …» або «вибув у …».
 *
 * до — підрозділ / частина / місце (часто родовий відмінок):
 *   «до МЕХАНІЗОВАНОГО ВІДДІЛЕННЯ … БАТАЛЬЙОНУ», «до в/ч А4784»
 *
 * у — стан або напрямок (знахідний):
 *   «у розпорядження», «у відрядження», «у відпустку»
 */
const LEADING_DIRECTION = /^(?:вибув(?:\s+(?:до|у|в|на))?|до|у|в|на)\s+/i;

const USES_U =
  /^(?:розпорядженн[яіеєю]|підпорядкуванн[яіеєю]|відрядженн[яіеєю]|відпустк[уаі]|лікуванн[яі]|навчанн[яі]|резерв\b|штат\b|сзч\b)/i;

const USES_DO =
  /(?:відділенн|взвод|рот[аиуы]|батальйон|батаре|дивізіон|служб|комплект|командуванн|управлінн|в[/.]?\s*ч|військов|[АA]\s*\d{4}|(?:ого|ої|ення|взводу|роти|батальйону|батареї)\b)/i;

export const stripTimesheetDirectionPrefix = (value: string) =>
  String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(LEADING_DIRECTION, "")
    .trim();

/** «у» або «до» за текстом місця призначення. */
export const timesheetDeparturePreposition = (
  rawDestination: string,
): "у" | "до" => {
  const destination = stripTimesheetDirectionPrefix(rawDestination);
  if (!destination) return "до";
  if (USES_U.test(destination)) return "у";
  if (USES_DO.test(destination)) return "до";
  return "до";
};

export const formatTimesheetDeparture = (rawDestination: string) => {
  const destination = stripTimesheetDirectionPrefix(rawDestination);
  if (!destination) return "вибув";
  return `вибув ${timesheetDeparturePreposition(destination)} ${destination}`;
};

/** Історична позначка вибуття в дні Табеля, а не поточний стан. */
export const isTimesheetDepartureMark = (value: unknown) => {
  const text = String(value ?? "").replace(/\s+/g, " ");
  return /вибув|переведення|розпорядження/iu.test(text);
};

export const dayFromOrderLabel = (value: string) => {
  const day = Number(String(value || "").match(/(\d{1,2})[./-]/)?.[1] ?? 0);
  return day >= 1 && day <= 31 ? day : 0;
};

/** День місяця журналу (1–31) або 0, якщо дата в іншому місяці/році. */
export const journalDayFromDateMs = (
  ms: number,
  monthStartMs: number,
): number => {
  if (!ms || !monthStartMs) return 0;
  const anchor = new Date(monthStartMs);
  const date = new Date(ms);
  if (
    date.getUTCFullYear() !== anchor.getUTCFullYear() ||
    date.getUTCMonth() !== anchor.getUTCMonth()
  ) {
    return 0;
  }
  return date.getUTCDate();
};

/** Дата в межах вибраного місяця журналу [monthStart … monthEnd]. */
export const dateMsInJournalMonth = (
  ms: number,
  monthStartMs: number,
  monthEndMs: number,
): boolean =>
  Boolean(
    ms && monthStartMs && monthEndMs && ms >= monthStartMs && ms <= monthEndMs,
  );

/** Табель: період **почався** у вибраному місяці. */
export const archivePeriodAffectsJournalTimesheet = (
  departMs: number,
  monthStartMs: number,
  monthEndMs: number,
): boolean => dateMsInJournalMonth(departMs, monthStartMs, monthEndMs);

/**
 * Табель: період перетинає місяць журналу — у т.ч. відсутність з кінця попереднього
 * місяця (МЕДРОТА 31.07 → повернення 03.08 дає коди на 01–02.08).
 *
 * Старі закриті періоди (напр. відпустка 09.10.2025–07.11.2025) не перетинають
 * серпень 2026. Період без дати повернення, що почався **до** місяця журналу,
 * теж не тягнемо вперед — інакше archive 2025 року фарбує весь поточний місяць.
 */
export const archivePeriodOverlapsJournalTimesheet = (
  departMs: number,
  returnMs: number | null,
  monthStartMs: number,
  monthEndMs: number,
): boolean => {
  if (!monthStartMs || !monthEndMs || !departMs) return false;
  if (departMs > monthEndMs) return false;

  const effectiveEndMs =
    returnMs ?? (departMs >= monthStartMs ? monthEndMs : 0);

  return effectiveEndMs >= monthStartMs;
};

/**
 * Archive для «Тимч. відсутні» / закриття: торкається місяця, якщо вибуття або
 * повернення в межах [monthStart … monthEnd]. Старі відкриті періоди до місяця — ні.
 */
export const archivePeriodTouchesJournalMonth = (
  departMs: number,
  returnMs: number | null,
  monthStartMs: number,
  monthEndMs: number,
): boolean => {
  if (!monthStartMs || !monthEndMs) return false;
  if (dateMsInJournalMonth(departMs, monthStartMs, monthEndMs)) return true;
  if (returnMs && dateMsInJournalMonth(returnMs, monthStartMs, monthEndMs)) {
    return true;
  }
  return false;
};

/**
 * Новий активний рядок після скасування вибуття: до дати відновлення «-»,
 * далі «+». Старий рядок з «вибув» не перезаписуємо.
 */
export const timesheetMarkAfterRestore = (
  day: number,
  activeFromDay: number,
  lastDay: number,
): "+" | "-" | null => {
  if (activeFromDay < 1 || day < 1) return null;
  if (day < activeFromDay) return "-";
  if (day <= lastDay) return "+";
  return null;
};

export type TimesheetAbsenceSpan = {
  fromDay: number;
  toDay: number;
  code: string;
};

/** «11-18:лік|19-20:СЗЧ» — перекриття днів: пізніший період має пріоритет. */
export const encodeTimesheetAbsenceSpans = (
  spans: TimesheetAbsenceSpan[],
) =>
  spans
    .filter((span) => span.fromDay > 0 && span.toDay >= span.fromDay && span.code)
    .map((span) => `${span.fromDay}-${span.toDay}:${span.code}`)
    .join("|");

export const parseTimesheetAbsenceSpans = (
  raw: string,
): TimesheetAbsenceSpan[] =>
  String(raw || "")
    .split("|")
    .map((part) => {
      const match = part.trim().match(/^(\d{1,2})-(\d{1,2}):(.+)$/);
      if (!match) return null;
      const fromDay = Number(match[1]);
      const toDay = Number(match[2]);
      const code = match[3].trim();
      if (fromDay < 1 || toDay < fromDay || !code) return null;
      return { fromDay, toDay, code };
    })
    .filter((span): span is TimesheetAbsenceSpan => Boolean(span));

export const timesheetCodeOnDay = (
  day: number,
  spans: TimesheetAbsenceSpan[],
) => {
  let code = "";
  for (const span of spans) {
    if (day >= span.fromDay && day <= span.toDay) code = span.code;
  }
  return code;
};

/** Коди відсутності на активному штатному епізоді — лише з дати постановки. */
export const clipAbsenceSpansToActiveEpisode = (
  spans: TimesheetAbsenceSpan[],
  activeFromDay: number,
): TimesheetAbsenceSpan[] => {
  if (activeFromDay <= 1) return spans;
  return spans
    .map((span) => ({
      ...span,
      fromDay: Math.max(span.fromDay, activeFromDay),
    }))
    .filter((span) => span.toDay >= span.fromDay);
};

/** Коди попереднього облікового епізоду — дні до нової штатної постановки. */
export const absenceSpansBeforeEpisode = (
  spans: TimesheetAbsenceSpan[],
  activeFromDay: number,
): TimesheetAbsenceSpan[] => {
  if (activeFromDay <= 1) return [];
  return spans
    .map((span) => ({
      ...span,
      toDay: Math.min(span.toDay, activeFromDay - 1),
    }))
    .filter((span) => span.fromDay > 0 && span.toDay >= span.fromDay);
};

/** Позначка дня **до** вибуття (ПЕРЕВ): archive-код або «+». */
export const timesheetMarkBeforeDeparture = (
  day: number,
  departDay: number,
  spans: TimesheetAbsenceSpan[],
): string => {
  if (day < 1 || day >= departDay) return "+";
  if (!spans.length) return "+";
  return (
    timesheetMarkFromArchive(day, {
      activeFromDay: 1,
      lastDay: Math.max(1, departDay - 1),
      spans,
      fillBeforeActive: true,
    }) ?? "+"
  );
};

export const countTimesheetPresentDaysBeforeDeparture = (
  departDay: number,
  spans: TimesheetAbsenceSpan[],
): number => {
  let count = 0;
  for (let day = 1; day < departDay; day += 1) {
    if (timesheetMarkBeforeDeparture(day, departDay, spans) === "+") {
      count += 1;
    }
  }
  return count;
};

/** Один алгоритм днів вибуття: ZIP, populate і preview. */
export const timesheetTransferMarkForDay = (input: {
  day: number;
  departDay: number;
  lastDay: number;
  activeFromDay?: number;
  absenceSpans?: TimesheetAbsenceSpan[];
  departMark: string;
}): string | null => {
  const { day, departDay, lastDay, departMark } = input;
  const activeFrom =
    input.activeFromDay && input.activeFromDay > 1 ? input.activeFromDay : 1;
  const spans = input.absenceSpans ?? [];
  if (departDay < 1) return null;
  if (day < departDay) {
    if (activeFrom > 1 && day < activeFrom) return "-";
    return spans.length
      ? timesheetMarkBeforeDeparture(day, departDay, spans)
      : "+";
  }
  if (day === departDay) return departMark;
  if (day <= lastDay) return "-";
  return null;
};

export type ArchiveAbsencePeriodInput = {
  departDate: string;
  returnDate: string;
  absenceType: string;
  excelRow?: number;
  departMs?: number;
  returnMs?: number | null;
};

/** Ланцюжок archive: лік → СЗЧ тощо; дні лише в межах місяця журналу. */
export const buildTimesheetAbsenceSpans = (
  periods: ArchiveAbsencePeriodInput[],
  options: {
    timesheetDay: number;
    monthStartMs: number;
    monthEndMs: number;
    reportDayMs?: number;
    mapCode: (absenceType: string) => string;
    hasReturn: (returnDate: string) => boolean;
  },
): TimesheetAbsenceSpan[] => {
  const sorted = [...periods]
    .filter((period) => (period.departMs ?? 0) > 0)
    .sort((left, right) => {
      const leftMs = left.departMs ?? 0;
      const rightMs = right.departMs ?? 0;
      if (leftMs !== rightMs) return leftMs - rightMs;
      return (left.excelRow ?? 0) - (right.excelRow ?? 0);
    });

  return sorted
    .map((period, index) => {
      const departMs = period.departMs ?? 0;
      const returnMs = period.returnMs ?? null;
      const { monthStartMs, monthEndMs, timesheetDay } = options;
      if (
        !archivePeriodOverlapsJournalTimesheet(
          departMs,
          returnMs,
          monthStartMs,
          monthEndMs,
        )
      ) {
        return null;
      }

      let fromDay = journalDayFromDateMs(departMs, monthStartMs);
      if (fromDay <= 0 && departMs > 0 && departMs < monthStartMs) {
        fromDay = 1;
      }
      if (fromDay <= 0) return null;

      const nextPeriod = sorted[index + 1];
      const nextDepartMs = nextPeriod?.departMs ?? 0;
      const nextFromDay =
        nextDepartMs > 0
          ? journalDayFromDateMs(nextDepartMs, monthStartMs)
          : 0;

      const monthLastDay =
        journalDayFromDateMs(monthEndMs, monthStartMs) || timesheetDay;
      let toDay = monthLastDay;
      if (returnMs) {
        if (returnMs > monthEndMs) {
          toDay = monthLastDay;
        } else if (options.reportDayMs && returnMs > options.reportDayMs) {
          toDay = monthLastDay;
        } else {
          const returnDayInMonth = journalDayFromDateMs(returnMs, monthStartMs);
          if (returnDayInMonth > 0) {
            toDay =
              returnDayInMonth > fromDay ? returnDayInMonth - 1 : returnDayInMonth;
          }
        }
      } else if (nextFromDay > fromDay) {
        toDay = nextFromDay - 1;
      }

      if (nextFromDay > fromDay && toDay >= nextFromDay) {
        toDay = nextFromDay - 1;
      }
      toDay = Math.min(toDay, monthLastDay);

      const code = options.mapCode(period.absenceType);
      return { fromDay, toDay, code };
    })
    .filter(
      (span): span is TimesheetAbsenceSpan =>
        Boolean(span) &&
        span.fromDay > 0 &&
        span.toDay >= span.fromDay &&
        Boolean(span.code),
    );
};

/** Якщо sh досі СЗЧ/ЗБ, а archive вже має «повернення» — Табель ведемо до дня звіту. */
export const extendDispositionSpanToReportDay = (
  spans: TimesheetAbsenceSpan[],
  options: {
    timesheetDay: number;
    shTimesheetCode: string | null;
    latestArchiveCode: string | null;
  },
): TimesheetAbsenceSpan[] => {
  const { timesheetDay, shTimesheetCode, latestArchiveCode } = options;
  if (
    !spans.length ||
    !shTimesheetCode ||
    shTimesheetCode === "+" ||
    shTimesheetCode !== latestArchiveCode
  ) {
    return spans;
  }
  const next = spans.map((span) => ({ ...span }));
  const target = [...next].reverse().find((span) => span.code === shTimesheetCode);
  if (target && target.toDay < timesheetDay) {
    target.toDay = timesheetDay;
  }
  return next;
};

/**
 * Фактична позначка дня: archive (лік/СЗЧ/…) має пріоритет над «+».
 * «СКАСОВАНО» в клітинку не пишемо — скасування лише визначає, з якого дня рядок знову активний.
 * Якщо штатний рядок починається з середини місяця (`fillBeforeActive` + `activeFromDay > 1`),
 * коди відсутності до цієї дати лишаються на історичному рядку — тут пишемо «-».
 */
export const timesheetMarkFromArchive = (
  day: number,
  options: {
    activeFromDay: number;
    lastDay: number;
    spans: TimesheetAbsenceSpan[];
    fillBeforeActive?: boolean;
  },
): string | null => {
  const {
    activeFromDay,
    lastDay,
    spans,
    fillBeforeActive = true,
  } = options;
  if (day < 1 || day > lastDay) return null;
  const absence = timesheetCodeOnDay(day, spans);
  if (activeFromDay > 1 && day < activeFromDay) {
    return fillBeforeActive ? "-" : null;
  }
  if (absence) return absence;
  if (day >= Math.max(1, activeFromDay || 1)) return "+";
  return null;
};

const DASH_MARK = /^[-–—−·]+$/;

/** «-» / «–» / «—» на Табелі — та сама позначка відсутності епізоду. */
export const sameTimesheetDayMark = (actual: string, expected: string) => {
  const left = String(actual || "").trim().toLocaleLowerCase("uk-UA");
  const right = String(expected || "").trim().toLocaleLowerCase("uk-UA");
  if (!right) return !left;
  if (left === right) return true;
  return right === "-" && DASH_MARK.test(left);
};
