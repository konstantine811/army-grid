import { parseAsOfDateLabel } from "./ejoosSyncPlan";

/** `DD.MM.YYYY` → `YYYY-MM-DD` для `<input type="date">`. */
export const uaLabelToIsoDate = (label: string): string => {
  const parsed = parseAsOfDateLabel(String(label || "").trim());
  if (parsed.sourceDateUnknown || !parsed.label) return "";
  const match = parsed.label.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!match) return "";
  return `${match[3]}-${match[2]}-${match[1]}`;
};

/** `YYYY-MM-DD` → `DD.MM.YYYY`. */
export const isoDateToUaLabel = (iso: string): string => {
  const parsed = parseAsOfDateLabel(String(iso || "").trim());
  return parsed.sourceDateUnknown ? "" : parsed.label;
};

/** Маска набору: лише цифри, автоматичні крапки. */
export const formatUaDateTyping = (raw: string): string => {
  const digits = raw.replace(/\D/g, "").slice(0, 8);
  if (digits.length <= 2) return digits;
  if (digits.length <= 4) {
    return `${digits.slice(0, 2)}.${digits.slice(2)}`;
  }
  return `${digits.slice(0, 2)}.${digits.slice(2, 4)}.${digits.slice(4)}`;
};

export const isCompleteUaDate = (text: string): boolean => {
  const parsed = parseAsOfDateLabel(String(text || "").trim());
  return !parsed.sourceDateUnknown && Boolean(parsed.label);
};

export const uaDateToIso = (text: string): string | null => {
  if (!isCompleteUaDate(text)) return null;
  return uaLabelToIsoDate(text.trim());
};

const ISO_DATE_RE = /^(\d{4})-(\d{2})-(\d{2})$/;

export const parseIsoDateParts = (iso: string) => {
  const match = String(iso || "").trim().match(ISO_DATE_RE);
  if (!match) return null;
  return {
    year: Number(match[1]),
    month: Number(match[2]),
    day: Number(match[3]),
  };
};

/**
 * Нативний `<input type="date">` стріляє `change`, щойно крутять місяць/рік
 * (день лишається той самий або обрізається). Застосовуємо дату лише коли
 * вибрали число в уже показаному місяці.
 */
export const isAsOfPickerDayCommit = (
  prevIso: string,
  nextIso: string,
): boolean => {
  const next = parseIsoDateParts(nextIso);
  if (!next) return false;
  const prev = parseIsoDateParts(prevIso);
  if (!prev) return true;
  if (prev.year !== next.year || prev.month !== next.month) return false;
  return prev.day !== next.day;
};

const ASOF_STORAGE_KEY = "army-grid:ejoos-source-as-of";

/** ISO `YYYY-MM-DD` з явної дати «станом на» (ISO або ДД.ММ.РРРР). */
export const normalizeAsOfIso = (value: string): string => {
  const trimmed = String(value || "").trim();
  if (parseIsoDateParts(trimmed)) return trimmed;
  return uaLabelToIsoDate(trimmed);
};

export const readRememberedAsOfIso = (): string => {
  if (typeof localStorage === "undefined") return "";
  try {
    return normalizeAsOfIso(localStorage.getItem(ASOF_STORAGE_KEY) || "");
  } catch {
    return "";
  }
};

export const writeRememberedAsOfIso = (value: string): string => {
  const iso = normalizeAsOfIso(value);
  if (typeof localStorage === "undefined") return iso;
  try {
    if (iso) localStorage.setItem(ASOF_STORAGE_KEY, iso);
    else localStorage.removeItem(ASOF_STORAGE_KEY);
  } catch {
    /* quota / private mode */
  }
  return iso;
};
