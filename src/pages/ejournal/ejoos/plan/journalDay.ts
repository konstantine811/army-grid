import type { JournalTimesheetDay } from "../types/workbookRows";

const journalDayFromParts = (
  day: number,
  month: number,
  year: number,
): JournalTimesheetDay => ({
  day,
  label: `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`,
  sourceDateUnknown: false,
});

export const parseTimesheetDayFromPbName = (
  fileName: string,
  _fallback?: Date,
): JournalTimesheetDay => {
  const match = fileName.match(/(\d{2})[._-](\d{2})[._-](\d{2,4})/);
  if (match) {
    const day = Number(match[1]);
    if (day >= 1 && day <= 31) {
      const month = Number(match[2]);
      const yearRaw = Number(match[3]);
      const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
      return journalDayFromParts(day, month, year);
    }
  }
  const compact = fileName.match(/_(\d{2})(\d{2})(\d{4})/);
  if (compact) {
    const day = Number(compact[1]);
    const month = Number(compact[2]);
    const year = Number(compact[3]);
    return journalDayFromParts(day, month, year);
  }
  return { day: 0, label: "", sourceDateUnknown: true };
};

/** ISO `YYYY-MM-DD` або `DD.MM.YYYY` — явна дата «станом на». */
export const parseAsOfDateLabel = (value: string): JournalTimesheetDay => {
  const iso = String(value || "")
    .trim()
    .match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return journalDayFromParts(day, month, year);
    }
  }
  return parseTimesheetDayFromPbName(value);
};

/** День табеля з назви файлу 1ПБ / ЄЖООС. Без дати — SOURCE_DATE_UNKNOWN. */
export const resolveJournalTimesheetDay = parseTimesheetDayFromPbName;

export const sourceTimesheetHorizonNote = (timesheetDayLabel: string) =>
  `Джерело 1ПБ станом на ${timesheetDayLabel}. Табель буде оновлено лише по ${timesheetDayLabel}.`;
