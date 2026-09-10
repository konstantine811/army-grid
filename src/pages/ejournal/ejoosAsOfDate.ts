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
