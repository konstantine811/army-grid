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

export const ubdBasisOrderOptionKey = (option: UbdBasisOrderOption) =>
  `${option.number}@@${option.date}`;

export const findUbdBasisOrderByKey = (key: string) => {
  const [number, date] = String(key ?? "").split("@@");
  if (!number || !date) return null;
  return (
    UBD_BASIS_ORDER_OPTIONS.find(
      (item) => item.number === number && item.date === date,
    ) ?? null
  );
};

/**
 * Дефолтний № БР: дата розпорядження = дата «з» періоду завдань,
 * інакше найближча раніша; якщо раніше немає — найраніша в списку.
 * Якщо на дату кілька номерів — беремо перший.
 */
export const pickUbdBasisOrderForTaskPeriod = (
  taskPeriod: string,
  options: UbdBasisOrderOption[] = UBD_BASIS_ORDER_OPTIONS,
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

export const formatUbdBasisOrderLabel = (option: UbdBasisOrderOption) =>
  `${option.number} · ${option.date}`;

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

/** Чи дата БР збігається з датою «з» періоду завдань. */
export const ubdBasisDateMatchesTaskPeriod = (
  taskPeriod: string,
  basisDate: string,
) => {
  const start = parseUbdTaskPeriodStartDate(taskPeriod);
  if (!start) return true;
  const wanted = formatUaDateKey(start);
  const got = normalizeUaDateKey(basisDate);
  if (!got) return false;
  return wanted === got;
};

/** Чи є в списку хоч один БР саме на дату «з» періоду. */
export const ubdHasExactBasisForTaskPeriod = (taskPeriod: string) => {
  const start = parseUbdTaskPeriodStartDate(taskPeriod);
  if (!start) return true;
  const wanted = formatUaDateKey(start);
  return UBD_BASIS_ORDER_OPTIONS.some((item) => item.date === wanted);
};

/** БР ще не готовий: явный прапор або дата БР ≠ «з» періоду. */
export const ubdBasisIsNotReady = (
  taskPeriod: string,
  basisDate: string,
  explicitFlag?: boolean | string | null,
) => {
  if (explicitFlag === true || explicitFlag === "true") return true;
  if (explicitFlag === false || explicitFlag === "false") return false;
  if (!String(taskPeriod ?? "").trim()) return false;
  return !ubdBasisDateMatchesTaskPeriod(taskPeriod, basisDate);
};
