import { allBasisOrderOptions } from "./ubdBasisOrdersDirectory";
import {
  UBD_BASIS_ORDER_OPTIONS,
  type UbdBasisOrderOption,
} from "./ubdBasisOrdersData";

export type { UbdBasisOrderOption };
export { UBD_BASIS_ORDER_OPTIONS };

const parseUaDate = (value: string): Date | null => {
  const text = String(value ?? "")
    .trim()
    .replaceAll("/", ".")
    .replaceAll("-", ".");
  const match = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += 2000;
  if (
    day < 1 ||
    day > 31 ||
    month < 1 ||
    month > 12 ||
    year < 2000 ||
    year > 2100
  ) {
    return null;
  }
  const date = new Date(year, month - 1, day);
  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  return date;
};

const dateKey = (date: Date) =>
  date.getFullYear() * 10_000 + (date.getMonth() + 1) * 100 + date.getDate();

const OPEN_END = 99_999_999;

export const normalizeBasisLocation = (value: string) =>
  String(value ?? "")
    .toLocaleLowerCase("uk-UA")
    .replace(/[ʼ’']/g, "")
    .replace(/\b(н\.?\s*п\.?|смт|селище|село|місто)\b/gi, " ")
    .replace(/(^|[\s(])(с|м)\.(?=\s|$)/gi, " ")
    .replace(/[.,;:№#"()/\\]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const basisLocationsMatch = (left: string, right: string) => {
  const a = normalizeBasisLocation(left);
  const b = normalizeBasisLocation(right);
  if (!a || !b) return false;
  return a === b || a.includes(b) || b.includes(a);
};

const optionRange = (option: UbdBasisOrderOption) => {
  const from = parseUaDate(option.validFrom || option.date);
  const to = parseUaDate(option.validTo || "");
  return { from, to };
};

const inRange = (date: Date, from: Date | null, to: Date | null) => {
  if (!from) return false;
  const key = dateKey(date);
  if (key < dateKey(from)) return false;
  if (to && key > dateKey(to)) return false;
  return true;
};

const rangesOverlap = (
  stayFrom: Date,
  stayTo: Date,
  from: Date | null,
  to: Date | null,
) => {
  if (!from) return false;
  const start = Math.max(dateKey(stayFrom), dateKey(from));
  const end = Math.min(dateKey(stayTo), to ? dateKey(to) : OPEN_END);
  return start <= end;
};

/** Перша дата з «Період завдань» (`з 29.07.2026-11.08.2026` → 29.07.2026). */
export const parseUbdTaskPeriodStartDate = (
  taskPeriod: string,
): Date | null => {
  const text = String(taskPeriod ?? "").trim();
  if (!text) return null;

  const range = text.match(
    /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})\s*[-–—]\s*(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/,
  );
  if (range) return parseUaDate(range[1]);

  const single = text.match(/(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/);
  return single ? parseUaDate(single[1]) : null;
};

export const parseUbdTaskPeriodEndDate = (taskPeriod: string): Date | null => {
  const text = String(taskPeriod ?? "").trim();
  if (!text) return null;
  const range = text.match(
    /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})\s*[-–—]\s*(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/,
  );
  return range ? parseUaDate(range[2]) : null;
};

export const ubdBasisOrderOptionKey = (option: UbdBasisOrderOption) =>
  `${option.number}@@${option.date}`;

export const findUbdBasisOrderByKey = (key: string) => {
  const [number, date] = String(key ?? "").split("@@");
  if (!number || !date) return null;
  return (
    allBasisOrderOptions().find(
      (item) => item.number === number && item.date === date,
    ) ??
    UBD_BASIS_ORDER_OPTIONS.find(
      (item) => item.number === number && item.date === date,
    ) ??
    null
  );
};

const pickByDateOnly = (
  taskPeriod: string,
  options: UbdBasisOrderOption[],
): UbdBasisOrderOption | null => {
  if (!options.length) return null;

  const start = parseUbdTaskPeriodStartDate(taskPeriod);
  if (!start) return options[0] ?? null;

  const target = dateKey(start);
  let best: UbdBasisOrderOption | null = null;
  let bestKey = -1;

  for (const option of options) {
    const parsed = parseUaDate(option.date);
    if (!parsed) continue;
    const key = dateKey(parsed);
    if (key > target) continue;
    if (key > bestKey) {
      best = option;
      bestKey = key;
    }
  }

  return best ?? options[0] ?? null;
};

/** Усі БР, що перекривають період на цій локації. */
export const findUbdBasisOrdersForLocationPeriod = (
  location: string,
  taskPeriod: string,
  options: UbdBasisOrderOption[] = allBasisOrderOptions(),
): UbdBasisOrderOption[] => {
  const place = String(location ?? "").trim();
  const start = parseUbdTaskPeriodStartDate(taskPeriod);
  const end = parseUbdTaskPeriodEndDate(taskPeriod) || start;
  if (!place || !start || !end) return [];

  const located = options.filter(
    (option) => option.location && basisLocationsMatch(option.location, place),
  );
  if (!located.length) return [];

  return located
    .filter((option) => {
      const { from, to } = optionRange(option);
      return rangesOverlap(start, end, from, to);
    })
    .sort((left, right) => {
      const leftFrom = optionRange(left).from;
      const rightFrom = optionRange(right).from;
      return (leftFrom ? dateKey(leftFrom) : 0) - (rightFrom ? dateKey(rightFrom) : 0);
    });
};

/**
 * Дефолтний № БР.
 * Якщо є локація в довіднику — локація + дата (або всі БР за період).
 * Інакше як раніше: дата розпорядження = «з» періоду / найближча раніша.
 */
export const pickUbdBasisOrderForTaskPeriod = (
  taskPeriod: string,
  options: UbdBasisOrderOption[] = allBasisOrderOptions(),
  location = "",
): UbdBasisOrderOption | null => {
  const matches = findUbdBasisOrdersForLocationPeriod(
    location,
    taskPeriod,
    options,
  );
  if (matches.length) {
    const start = parseUbdTaskPeriodStartDate(taskPeriod);
    if (start) {
      const covering = matches.find((option) => {
        const { from, to } = optionRange(option);
        return inRange(start, from, to);
      });
      if (covering) return covering;
    }
    return matches[0];
  }
  return pickByDateOnly(taskPeriod, options);
};

export const resolveUbdBasisForTask = (
  taskPeriod: string,
  location = "",
  options: UbdBasisOrderOption[] = allBasisOrderOptions(),
) => {
  const matches = findUbdBasisOrdersForLocationPeriod(
    location,
    taskPeriod,
    options,
  );
  if (matches.length > 1) {
    return {
      number: matches.map((item) => item.number).join(", "),
      date: matches[0]?.date || "",
      matches,
    };
  }
  const picked =
    matches[0] || pickUbdBasisOrderForTaskPeriod(taskPeriod, options, location);
  if (!picked) return null;
  return {
    number: picked.number,
    date: picked.date,
    matches: matches.length ? matches : [picked],
  };
};

export const formatUbdBasisOrderLabel = (option: UbdBasisOrderOption) => {
  const place = option.location?.trim();
  const till = option.validTo?.trim();
  const from = (option.validFrom || option.date).trim();
  if (place && from) {
    return till
      ? `${option.number} · ${place} · ${from}–${till}`
      : `${option.number} · ${place} · з ${from}`;
  }
  return `${option.number} · ${option.date}`;
};

const formatUaDateKey = (date: Date) => {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${day}.${month}.${date.getFullYear()}`;
};

const normalizeUaDateKey = (value: string) =>
  String(value ?? "")
    .trim()
    .replaceAll("/", ".")
    .replaceAll("-", ".");

/** Чи дата БР збігається з датою «з» періоду завдань (або є локаційний збіг). */
export const ubdBasisDateMatchesTaskPeriod = (
  taskPeriod: string,
  basisDate: string,
  location = "",
) => {
  if (
    location &&
    findUbdBasisOrdersForLocationPeriod(location, taskPeriod).length
  ) {
    return true;
  }
  const start = parseUbdTaskPeriodStartDate(taskPeriod);
  if (!start) return true;
  const wanted = formatUaDateKey(start);
  const got = normalizeUaDateKey(basisDate);
  if (!got) return false;
  return wanted === got;
};

/** Чи є в списку хоч один БР саме на дату «з» періоду. */
export const ubdHasExactBasisForTaskPeriod = (
  taskPeriod: string,
  location = "",
) => {
  if (
    location &&
    findUbdBasisOrdersForLocationPeriod(location, taskPeriod).length
  ) {
    return true;
  }
  const start = parseUbdTaskPeriodStartDate(taskPeriod);
  if (!start) return true;
  const wanted = formatUaDateKey(start);
  return allBasisOrderOptions().some((item) => item.date === wanted);
};

/** БР ще не готовий: явный прапор або дата БР ≠ «з» періоду. */
export const ubdBasisIsNotReady = (
  taskPeriod: string,
  basisDate: string,
  explicitFlag?: boolean | string | null,
  location = "",
) => {
  if (explicitFlag === true || explicitFlag === "true") return true;
  if (explicitFlag === false || explicitFlag === "false") return false;
  if (!String(taskPeriod ?? "").trim()) return false;
  return !ubdBasisDateMatchesTaskPeriod(taskPeriod, basisDate, location);
};
