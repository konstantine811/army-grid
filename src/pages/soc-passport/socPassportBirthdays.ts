import { isPresentStatus, isTransiterDestination } from "./socPassportFields";
import type { SocPerson } from "./socPassportTypes";

export const UKRAINIAN_MONTH_NAMES = [
  "січень",
  "лютий",
  "березень",
  "квітень",
  "травень",
  "червень",
  "липень",
  "серпень",
  "вересень",
  "жовтень",
  "листопад",
  "грудень",
] as const;

export type BirthDateParts = {
  day: number;
  month: number;
  year: number;
};

export type BirthdayPersonRow = {
  name: string;
  callsign: string;
  birthDate: string;
  location: string;
  day: number;
  note?: string;
};

export const missingBirthDateReason = (
  person: Pick<SocPerson, "birthDate" | "match">,
) => {
  if (parseBirthDateParts(person.birthDate)) return "";
  if (!person.match.oos) return "немає в ООС ЕЖООС";
  if (!String(person.birthDate ?? "").trim()) {
    return "немає дати в ООС і в штатці";
  }
  return "дату не розпізнано";
};

const toBirthdayRow = (
  person: SocPerson,
  extras: Partial<BirthdayPersonRow> = {},
): BirthdayPersonRow => ({
  name: person.name,
  callsign: person.callsign.trim(),
  birthDate: formatBirthDateDisplay(person.birthDate),
  location: (person.morningLocation || person.oosDislocation || "").trim(),
  day: 0,
  ...extras,
});

export const ukrainianMonthLabel = (asOf = new Date()) =>
  UKRAINIAN_MONTH_NAMES[asOf.getMonth()] ?? "";

export const parseBirthDateParts = (value: string): BirthDateParts | null => {
  const text = String(value ?? "").trim();
  if (!text) return null;

  const serialMatch = text.match(/^(\d{4,5})(?:\.0+)?$/);
  if (serialMatch) {
    const serial = Number(serialMatch[1]);
    if (Number.isFinite(serial) && serial > 20000) {
      const date = new Date((serial - 25569) * 86400000);
      const year = date.getUTCFullYear();
      const month = date.getUTCMonth() + 1;
      const day = date.getUTCDate();
      if (year >= 1950 && year <= 2010) return { day, month, year };
    }
  }

  const dotted = text.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})/);
  if (dotted) {
    let year = Number(dotted[3]);
    if (year < 100) year += year >= 50 ? 1900 : 2000;
    const month = Number(dotted[2]);
    const day = Number(dotted[1]);
    if (
      year >= 1950 &&
      year <= 2010 &&
      month >= 1 &&
      month <= 12 &&
      day >= 1 &&
      day <= 31
    ) {
      return { day, month, year };
    }
  }

  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    const year = Number(iso[1]);
    const month = Number(iso[2]);
    const day = Number(iso[3]);
    if (year >= 1950 && year <= 2010) return { day, month, year };
  }

  return null;
};

export const pickUsableBirthDate = (...values: Array<string | undefined>) => {
  for (const value of values) {
    const text = String(value ?? "").trim();
    if (parseBirthDateParts(text)) return text;
  }
  return (
    values
      .map((value) => String(value ?? "").trim())
      .find(Boolean) ?? ""
  );
};

export const formatBirthDateDisplay = (value: string) => {
  const parts = parseBirthDateParts(value);
  if (!parts) return String(value ?? "").trim();
  return [
    String(parts.day).padStart(2, "0"),
    String(parts.month).padStart(2, "0"),
    String(parts.year),
  ].join(".");
};

export const isBirthdayInMonth = (value: string, asOf = new Date()) => {
  const parts = parseBirthDateParts(value);
  return Boolean(parts && parts.month === asOf.getMonth() + 1);
};

export const isOnDutyNonTransiter = (person: Pick<
  SocPerson,
  "isTransiter" | "morningStatus" | "morningDestination" | "morningAbsenceNotes"
>) => {
  if (person.isTransiter) return false;
  if (isTransiterDestination(person.morningDestination)) return false;
  if (isTransiterDestination(person.morningAbsenceNotes)) return false;
  return isPresentStatus(person.morningStatus);
};

export const listCurrentMonthBirthdays = (
  people: SocPerson[],
  asOf = new Date(),
): BirthdayPersonRow[] =>
  people
    .filter(
      (person) =>
        isOnDutyNonTransiter(person) &&
        isBirthdayInMonth(person.birthDate, asOf),
    )
    .map((person) => {
      const parts = parseBirthDateParts(person.birthDate);
      return toBirthdayRow(person, { day: parts?.day ?? 0 });
    })
    .sort((left, right) => {
      if (left.day !== right.day) return left.day - right.day;
      return left.name.localeCompare(right.name, "uk");
    });

/** У строю зі штатки (ранковий), але дати народження з ЕЖООС немає. */
export const listOnDutyMissingBirthDates = (
  people: SocPerson[],
): BirthdayPersonRow[] =>
  people
    .filter(
      (person) =>
        isOnDutyNonTransiter(person) && !parseBirthDateParts(person.birthDate),
    )
    .map((person) =>
      toBirthdayRow(person, {
        birthDate: "",
        note: missingBirthDateReason(person),
      }),
    )
    .sort((left, right) => left.name.localeCompare(right.name, "uk"));
