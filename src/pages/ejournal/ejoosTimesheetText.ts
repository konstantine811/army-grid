/**
 * Прийменник для Табеля: «вибув до …» або «вибув у …».
 *
 * до — підрозділ / частина / місце (часто родовий відмінок):
 *   «до МЕХАНІЗОВАНОГО ВІДДІЛЕННЯ … БАТАЛЬЙОНУ», «до в/ч А4784»
 *
 * у — стан або напрямок (знахідний):
 *   «у розпорядження», «у відрядження», «у відпустку»
 */
import {
  mentionsExternalMilitaryUnit,
  mentionsForeignUnit,
} from "./ejoosMovementRules";
import { EJOOS_TIMESHEET_CODES } from "./ejoosRules";

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
  const unitCodes = [
    ...destination.matchAll(/[АA]\s*(\d{4})(?!\d)/giu),
  ].map((match) => `А${match[1]}`);
  const uniqueUnits = [...new Set(unitCodes)];
  const hasSubunit =
    /відділенн|взвод|рот[аиуы]|батальйон|батаре|дивізіон/i.test(destination);
  if (uniqueUnits.length === 1 && !hasSubunit) {
    return `вибув у в/ч ${uniqueUnits[0]}`;
  }
  return `вибув ${timesheetDeparturePreposition(destination)} ${destination}`;
};

/** РОЗПОРЯДЖ: повна фраза в день наказу; до неї — код відсутності, після — «-». */
export const formatDispositionTimesheetDeparture = (
  rawDestination: string,
  orderNumber: string,
  orderDate: string,
) => {
  const destination =
    String(rawDestination || "")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/^[,;:\s]+/, "")
      .replace(/^який\s+знаходиться\s+/iu, "")
      .replace(/^у\s+розпорядженні(?=\s|$)/iu, "у розпорядження")
      .replace(/^розпорядженні(?=\s|$)/iu, "у розпорядження") ||
    "у розпорядження командира військової частини А4862";
  const number = String(orderNumber || "")
    .replace(/^\s*(?:наказ\s*)?№\s*/iu, "")
    .trim();
  const date = String(orderDate || "").trim();
  return [
    `вибув ${destination}`,
    number ? `наказ №${number}` : "",
    date ? `від ${date}` : "",
  ]
    .filter(Boolean)
    .join(" ");
};

/** РОЗПОРЯДЖ: до дати наказу — код відсутності, у дату — «вибув…», далі «-». */
export const timesheetMarkForOpenDispositionDay = (
  day: number,
  options: {
    dispositionOrderDay: number;
    dispositionDeparture: string;
    absenceCode: string;
    lastDay: number;
    absenceSpans?: TimesheetAbsenceSpan[];
    fillBeforeActive?: boolean;
  },
): string => {
  const {
    dispositionOrderDay,
    dispositionDeparture,
    absenceCode,
    lastDay,
    absenceSpans,
  } = options;
  if (dispositionOrderDay > 0 && day === dispositionOrderDay) {
    return dispositionDeparture;
  }
  if (dispositionOrderDay > 0 && day > dispositionOrderDay) {
    return "-";
  }
  if (absenceSpans?.length) {
    const fromSpan = timesheetMarkFromArchive(day, {
      activeFromDay: 1,
      lastDay,
      spans: absenceSpans,
      fillBeforeActive: options.fillBeforeActive ?? true,
    });
    if (fromSpan) return fromSpan;
  }
  return absenceCode;
};

/**
 * Хвіст «Яка зміна»: «2103200 Стрілець → 2103xxx Стрілець МЕХАНІЗОВАНОГО…»
 * → нова посада без службового індексу.
 */
export const positionTitleTail = (value: string) => {
  const text = String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
  if (!text) return "";
  const tail = text.split(/→|->|=>/).pop()?.trim() || text;
  return tail.replace(/^\d{5,}[\s.:;-]*/, "").trim() || tail;
};

const STRUCTURE_POSITION_WORDS = new Set([
  "ВІДДІЛЕННЯ",
  "ВІДДІЛЕННІ",
  "ВЗВОДУ",
  "ВЗВОДІ",
  "РОТИ",
  "РОТІ",
  "БАТАЛЬЙОНУ",
  "БАТАЛЬЙОНІ",
  "БАТАРЕЇ",
  "ДИВІЗІОНУ",
  "УПРАВЛІННЯ",
  "ШТАБУ",
  "СЕКЦІЇ",
  "СЛУЖБИ",
  "ВІДДІЛУ",
  "ГРУПИ",
  "КОМАНДИ",
]);

const isStructureModifier = (token: string) => {
  const word = token.toUpperCase().replace(/[^0-9А-ЯІЇЄҐA-Z]/gu, "");
  if (!word) return false;
  if (/^\d+$/.test(word)) return true;
  return (
    word.endsWith("ОГО") ||
    word.endsWith("ЬКОГО") ||
    word.endsWith("НОГО") ||
    word.endsWith("ОВОГО") ||
    word.endsWith("ЕВОГО") ||
    word.endsWith("ОЇ") ||
    word.endsWith("ЬКОЇ") ||
    word.endsWith("НОЇ") ||
    word.endsWith("ОВОЇ") ||
    word.endsWith("ЕВОЇ") ||
    word.endsWith("ИХ") ||
    word.endsWith("ЬКИХ") ||
    word.endsWith("НИХ") ||
    word.endsWith("ОВИХ") ||
    word.endsWith("ЕВИХ")
  );
};

/**
 * З назви посади лишаємо лише ланцюг підрозділу:
 * «Стрілець МЕХАНІЗОВАНОГО ВІДДІЛЕННЯ … БАТАЛЬЙОНУ»
 * → «МЕХАНІЗОВАНОГО ВІДДІЛЕННЯ … БАТАЛЬЙОНУ».
 * Без слів відділення/взвод/рота/батальйон — порожньо (не підставляти посаду).
 */
export const extractTimesheetDestinationFromPosition = (
  positionTitle: string,
) => {
  const text = positionTitleTail(positionTitle);
  if (!text) return "";
  const tokens = text.split(" ");
  const structureIndex = tokens.findIndex((token) => {
    const word = token.toUpperCase().replace(/[^А-ЯІЇЄҐA-Z]/gu, "");
    return STRUCTURE_POSITION_WORDS.has(word);
  });
  if (structureIndex < 0) return "";

  let start = structureIndex;
  while (start > 0 && isStructureModifier(tokens[start - 1])) {
    start -= 1;
  }
  return tokens.slice(start).join(" ").trim();
};

/** Історична позначка вибуття в дні Табеля, а не поточний стан. */
export const isTimesheetDepartureMark = (value: unknown) => {
  const text = String(value ?? "").replace(/\s+/g, " ");
  return /вибув|переведення|розпорядження/iu.test(text);
};

/**
 * «вибув до 3 піхотного відділення … взводу» — внутрішня зміна посади.
 * Не плутати з вибуттям у 2ПБ / в/ч А####.
 */
export const isInternalStaffTimesheetDeparture = (value: unknown) => {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  if (!isTimesheetDepartureMark(text)) return false;
  if (/розпорядж/iu.test(text)) return false;
  if (mentionsExternalMilitaryUnit(text) || mentionsForeignUnit(text)) {
    return false;
  }
  if (/батальйон/iu.test(text) && !/1\s*піхотн/iu.test(text)) return false;
  return /відділенн|взвод|рот[аиуы]|батаре|1\s*піхотн/iu.test(text);
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
 * Відкритий archive з попереднього місяця фарбує журнал лише коли актуальний
 * sh досі підтверджує ту саму відсутність (СЗЧ з 25.07 → серпень = СЗЧ).
 * Старий відкритий 2025 при sh «В СТРОЮ» не тягнемо.
 */
export const currentStatusConfirmsOpenAbsence = (
  shTimesheetCode: string | null | undefined,
  archiveTimesheetCode: string | null | undefined,
) => {
  const sh = String(shTimesheetCode || "").trim();
  const archive = String(archiveTimesheetCode || "").trim();
  if (!sh || !archive || sh === "+") return false;
  return sh.toLocaleLowerCase("uk-UA") === archive.toLocaleLowerCase("uk-UA");
};

export type ArchiveCarryOptions = {
  /** Відкритий період, що почався до місяця, продовжити від 1 числа. */
  carryOpen?: boolean;
};

/**
 * Табель: період перетинає місяць журналу — у т.ч. відсутність з кінця попереднього
 * місяця (МЕДРОТА 31.07 → повернення 03.08 дає коди на 01–02.08).
 *
 * Старі закриті періоди (напр. відпустка 09.10.2025–07.11.2025) не перетинають
 * серпень 2026. Відкритий період до місяця — лише якщо `carryOpen` (sh досі СЗЧ/ЗБ).
 */
export const archivePeriodOverlapsJournalTimesheet = (
  departMs: number,
  returnMs: number | null,
  monthStartMs: number,
  monthEndMs: number,
  options?: ArchiveCarryOptions,
): boolean => {
  if (!monthStartMs || !monthEndMs || !departMs) return false;
  if (departMs > monthEndMs) return false;
  if (returnMs != null) return returnMs >= monthStartMs;
  if (departMs >= monthStartMs) return true;
  return Boolean(options?.carryOpen);
};

/**
 * Archive для «Тимч. відсутні» / закриття: торкається місяця, якщо вибуття або
 * повернення в межах [monthStart … monthEnd], або відкритий період підтверджений sh.
 */
export const archivePeriodTouchesJournalMonth = (
  departMs: number,
  returnMs: number | null,
  monthStartMs: number,
  monthEndMs: number,
  options?: ArchiveCarryOptions,
): boolean => {
  if (!monthStartMs || !monthEndMs) return false;
  if (dateMsInJournalMonth(departMs, monthStartMs, monthEndMs)) return true;
  if (returnMs && dateMsInJournalMonth(returnMs, monthStartMs, monthEndMs)) {
    return true;
  }
  if (
    !returnMs &&
    departMs &&
    departMs < monthStartMs &&
    options?.carryOpen
  ) {
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

/** Коди на історичному рядку закритої посади — до дня вибуття, не суцільні «+». */
export const historyAbsenceSpansForClosedEpisode = (
  payload: {
    historyTimesheetAbsenceSpans?: string;
    timesheetAbsenceSpans?: string;
  },
  departDay: number,
): TimesheetAbsenceSpan[] => {
  const history = parseTimesheetAbsenceSpans(
    payload.historyTimesheetAbsenceSpans || "",
  );
  if (history.length) return history;
  return absenceSpansBeforeEpisode(
    parseTimesheetAbsenceSpans(payload.timesheetAbsenceSpans || ""),
    departDay,
  );
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
  personId?: string;
  fullName?: string;
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
    confirmOpenCarry?: (period: ArchiveAbsencePeriodInput) => boolean;
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
      const laterPeriod = sorted
        .slice(index + 1)
        .find((candidate) => (candidate.departMs ?? 0) > departMs);
      const laterDepartMs = laterPeriod?.departMs ?? 0;
      // Старий відкритий МЕДРОТА 2025 не фарбує серпень, якщо далі вже була
      // відпустка / інший період, що почався до місяця журналу.
      if (
        !returnMs &&
        departMs < monthStartMs &&
        laterDepartMs > departMs &&
        laterDepartMs <= monthStartMs
      ) {
        return null;
      }
      if (
        !archivePeriodOverlapsJournalTimesheet(
          departMs,
          returnMs,
          monthStartMs,
          monthEndMs,
          {
            carryOpen: Boolean(options.confirmOpenCarry?.(period)),
          },
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
      const horizonDay = timesheetDay > 0 ? timesheetDay : monthLastDay;
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
      } else {
        toDay = horizonDay;
      }

      if (nextFromDay > fromDay && toDay >= nextFromDay) {
        toDay = nextFromDay - 1;
      }
      toDay = Math.min(toDay, monthLastDay, horizonDay);

      const code = options.mapCode(period.absenceType);
      return { fromDay, toDay, code };
    })
    .filter((span): span is TimesheetAbsenceSpan => {
      if (!span) return false;
      return span.fromDay > 0 && span.toDay >= span.fromDay && Boolean(span.code);
    });
};

/** Якщо sh досі СЗЧ/ЗБ, а archive вже має «повернення» — не здогадуємось. */
export const archiveReturnContradictsCurrentSh = (
  shTimesheetCode: string | null | undefined,
  archiveTimesheetCode: string | null | undefined,
  hasReturn: boolean,
) => hasReturn && currentStatusConfirmsOpenAbsence(shTimesheetCode, archiveTimesheetCode);

/**
 * Відкритий СЗЧ/ЛК у sh і в archive — тягнемо код до дня зрізу 1ПБ.
 * Суперечність archive-return vs sh-still-absent лишається в NEEDS_REVIEW.
 */
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

const KNOWN_ABSENCE_CODES = new Set(
  EJOOS_TIMESHEET_CODES.filter((code) => code !== "+").map((code) =>
    code.toLocaleLowerCase("uk-UA"),
  ),
);

/** ЛК — історичний синонім «лік» у Табелі. */
KNOWN_ABSENCE_CODES.add("лк");

/** Лише відомі коди відсутності. Число «19» (сума «+») сюди не входить. */
export const isTimesheetAbsenceCode = (code: string) => {
  const mark = String(code || "").trim();
  if (!mark || mark === "+") return false;
  if (DASH_MARK.test(mark)) return false;
  if (isTimesheetDepartureMark(mark)) return false;
  return KNOWN_ABSENCE_CODES.has(mark.toLocaleLowerCase("uk-UA"));
};

/** Позначка дня Табеля, яку можна тягнути по горизонту. */
export const isPlausibleTimesheetDayMark = (code: string) => {
  const mark = String(code || "").trim();
  if (!mark) return false;
  if (mark === "+") return true;
  if (DASH_MARK.test(mark)) return true;
  if (isTimesheetDepartureMark(mark)) return true;
  return isTimesheetAbsenceCode(mark);
};

/**
 * Порожні дні після останньої позначки добиваємо до дня зрізу 1ПБ.
 * Відкритий ЛК/СЗЧ без фактичного повернення лишається і на день зрізу —
 * не ставимо «+» на 25, якщо прибуття не було.
 */
export const timesheetHorizonFillDays = (input: {
  dayCodes: Array<string | null | undefined>;
  horizon: number;
  reportCode: string;
  confirmedReturn?: boolean;
}): Array<{ day: number; mark: string }> => {
  const horizon = Math.min(31, Math.max(0, Number(input.horizon) || 0));
  const report = String(input.reportCode || "").trim();
  if (horizon < 1 || !report || report === "(оберіть код)") return [];

  const markAt = (day: number) => String(input.dayCodes[day] ?? "").trim();
  const keepOpenAbsence = (seen: string) =>
    isTimesheetAbsenceCode(seen) && report === "+" && !input.confirmedReturn;

  const writes: Array<{ day: number; mark: string }> = [];
  let lastSeen = "";

  for (let day = 1; day <= horizon; day += 1) {
    const existing = markAt(day);
    if (existing && !DASH_MARK.test(existing)) {
      if (isTimesheetDepartureMark(existing)) {
        lastSeen = "вибув";
        continue;
      }
      if (!isPlausibleTimesheetDayMark(existing)) {
        if (day === horizon && report) {
          writes.push({ day, mark: report });
          lastSeen = report;
        }
        continue;
      }
      if (day === horizon && existing === "+" && keepOpenAbsence(lastSeen)) {
        writes.push({ day, mark: lastSeen });
        continue;
      }
      if (day === horizon && report && !sameTimesheetDayMark(existing, report)) {
        if (keepOpenAbsence(existing) || keepOpenAbsence(lastSeen)) {
          lastSeen = isTimesheetAbsenceCode(existing)
            ? existing
            : lastSeen || existing;
          continue;
        }
        writes.push({ day, mark: report });
        lastSeen = report;
        continue;
      }
      lastSeen = existing;
      continue;
    }

    if (!lastSeen || lastSeen === "вибув") {
      if (day === horizon) {
        writes.push({ day, mark: report });
        lastSeen = report;
      }
      continue;
    }

    const mark =
      day === horizon && report !== lastSeen && !keepOpenAbsence(lastSeen)
        ? report
        : lastSeen;
    if (!sameTimesheetDayMark(existing, mark)) {
      writes.push({ day, mark });
    }
    lastSeen = mark;
  }
  return writes;
};

const UK_MONTH_TITLES = [
  "Січень",
  "Лютий",
  "Березень",
  "Квітень",
  "Травень",
  "Червень",
  "Липень",
  "Серпень",
  "Вересень",
  "Жовтень",
  "Листопад",
  "Грудень",
] as const;

const TIMESHEET_MONTH_HEADER_RE =
  /(січень|лютий|березень|квітень|травень|червень|липень|серпень|вересень|жовтень|листопад|грудень)\s+(\d{4})\s*р\.?/iu;

/** Заголовок вкладки «6. Табель»: «Серпень 2026 р.» з дати журналу. */
export const formatTimesheetMonthHeader = (label: string) => {
  const match = String(label || "").match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (!match) return "";
  const month = Number(match[2]);
  const yearRaw = Number(match[3]);
  const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
  if (month < 1 || month > 12 || !Number.isFinite(year)) return "";
  return `${UK_MONTH_TITLES[month - 1]} ${year} р.`;
};

export const parseTimesheetMonthHeaderText = (value: string) => {
  const match = String(value || "").match(TIMESHEET_MONTH_HEADER_RE);
  if (!match) return null;
  const idx = UK_MONTH_TITLES.findIndex(
    (name) => name.toLocaleLowerCase("uk-UA") === match[1].toLocaleLowerCase("uk-UA"),
  );
  if (idx < 0) return null;
  return { month: idx + 1, year: Number(match[2]), matched: match[0] };
};

const excelSerialToDate = (value: number) => {
  if (!Number.isFinite(value) || value <= 20000 || value >= 80000) return null;
  const utc = Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000;
  const date = new Date(utc);
  if (Number.isNaN(date.getTime())) return null;
  return date;
};

/** Текст заголовка місяця: рядок «Серпень 2026 р.» або дата Excel у I2. */
export const timesheetMonthHeaderTextFromCell = (value: unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${UK_MONTH_TITLES[value.getMonth()]} ${value.getFullYear()} р.`;
  }
  if (typeof value === "number") {
    const date = excelSerialToDate(value);
    if (date) {
      return `${UK_MONTH_TITLES[date.getUTCMonth()]} ${date.getUTCFullYear()} р.`;
    }
  }
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
};

export const findTimesheetMonthHeaderCell = (
  rows: Array<Array<unknown> | undefined>,
) => {
  const hits: Array<{
    row: number;
    column: number;
    text: string;
    month: number;
    year: number;
    matched: string;
  }> = [];
  for (let row = 0; row < Math.min(6, rows.length); row += 1) {
    const cells = rows[row] ?? [];
    for (let column = 0; column < Math.min(40, cells.length); column += 1) {
      const text = timesheetMonthHeaderTextFromCell(cells[column]);
      const parsed = parseTimesheetMonthHeaderText(text);
      if (!parsed) continue;
      hits.push({
        row: row + 1,
        column: column + 1,
        text,
        month: parsed.month,
        year: parsed.year,
        matched: parsed.matched,
      });
    }
  }
  if (!hits.length) return null;
  const atAnchor = hits.find((hit) => hit.row === 2 && hit.column === 9);
  if (atAnchor) return atAnchor;
  const overDays = hits.filter((hit) => hit.row <= 3 && hit.column >= 9);
  if (overDays.length) return overDays[overDays.length - 1];
  // «Січень 2026 р.» у колонці A — залишок шаблону, не місяць Табеля.
  return null;
};

export const replaceTimesheetMonthHeaderText = (
  current: string,
  nextHeader: string,
) => {
  if (!nextHeader) return current;
  if (!current) return nextHeader;
  if (!TIMESHEET_MONTH_HEADER_RE.test(current)) return nextHeader;
  return current.replace(TIMESHEET_MONTH_HEADER_RE, nextHeader);
};
