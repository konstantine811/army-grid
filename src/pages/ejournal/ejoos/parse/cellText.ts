import type { CellValue } from "../../../../excelRoundTrip";
import {
  formatUkDate,
  tryParseExcelSerialDate,
} from "../../../../shared/format";

export const norm = (value: CellValue | unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatUkDate(value);
  }
  const parsedSerial = tryParseExcelSerialDate(value);
  if (parsedSerial) return formatUkDate(parsedSerial);
  if (
    typeof value === "number" &&
    Number.isFinite(value) &&
    value > 20000 &&
    value < 80000
  ) {
    const utc = Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000;
    const date = new Date(utc);
    const day = String(date.getUTCDate()).padStart(2, "0");
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    return `${day}.${month}.${date.getUTCFullYear()}`;
  }
  if (value && typeof value === "object" && !(value instanceof Date)) {
    const maybe = value as { text?: () => string; value?: () => unknown };
    if (typeof maybe.text === "function") return norm(maybe.text());
    if (typeof maybe.value === "function") return norm(maybe.value());
    return "";
  }
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
};

export const normKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

/** Уточнення на кшталт «(08.02.1985 р.н.)» — не інше ПІБ, а лише позначка. */
export const canonicalName = (value: string) =>
  normKey(
    value
      .replace(/\([^)]*\)/g, " ")
      .replace(/[.,;]/g, " ")
      .replace(/\s+/g, " "),
  );

export const dateMs = (value: string) => {
  const match = String(value ?? "").match(
    /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
  );
  if (!match) return 0;
  const year =
    Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
  const result = Date.UTC(year, Number(match[2]) - 1, Number(match[1]));
  return Number.isFinite(result) ? result : 0;
};

/** Штатний ID 1ПБ/ЕЖООС — коротке число, не РНОКПП/ІПН з 8+ цифр. */
export const isJournalPersonId = (value: string) => /^\d{1,7}$/.test(value.trim());

/** ID is never an Excel date. Reject date-shaped fallback values instead of showing them as IDs. */
export const normId = (value: CellValue | unknown) => {
  if (value instanceof Date) return "";
  const text =
    typeof value === "number" && Number.isFinite(value)
      ? String(Math.trunc(value))
      : String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text)) return "";
  return text === "0" || text === "0.0" ? "" : text;
};
