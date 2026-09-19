import type { CellValue, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { tryParseExcelSerialDate } from "../../shared/format";
import { parseExcelLabState } from "../excel-lab/ExcelLabStateParser";
import { parsePbMovements } from "../ejournal/ejoosSyncPlan";
import {
  buildCombatLossesFromPb,
  buildDeparturesFromEjoos,
  buildDispositionFromArchive,
  buildSzchFromRuh,
  type CombatLossReason,
  type DepartureCategory,
} from "../soc-passport/socPassportDepartures";
import { normalizeLooseText } from "../soc-passport/socPassportFields";
import type { RankGroup } from "../soc-passport/socPassportTypes";

const cellText = (value: CellValue | undefined) => String(value ?? "").trim();

const parseEventDate = (value: CellValue | undefined): Date | null => {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const fromSerial = tryParseExcelSerialDate(value);
  if (fromSerial) return fromSerial;
  const text = cellText(value);
  const match = text.match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (!match) return null;
  const day = Number(match[1]);
  const month = Number(match[2]);
  const yearRaw = Number(match[3]);
  const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(Date.UTC(year, month - 1, day));
  return Number.isNaN(date.getTime()) ? null : date;
};

export const isDateInMonth = (
  dateValue: string,
  month: number,
  year: number,
): boolean => {
  const date = parseEventDate(dateValue);
  if (!date) return false;
  return date.getUTCMonth() + 1 === month && date.getUTCFullYear() === year;
};

export const monthSinceDate = (month: number, year: number) =>
  `01.${String(month).padStart(2, "0")}.${year}`;

export const monthUntilDate = (month: number, year: number) => {
  const lastDay = new Date(Date.UTC(year, month, 0)).getUTCDate();
  return `${String(lastDay).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`;
};

export type SergeantsUnitMetrics = {
  combatKilled: number;
  combatSanitary: number;
  combatMissing: number;
  combatPow: number;
  nonCombatIrreversible: number;
  nonCombatSanitary: number;
  onTreatment: number;
  szchMonth: number;
};

export type SergeantsByUnitReport = {
  month: number;
  year: number;
  totals: SergeantsUnitMetrics;
  byUnit: Array<{ unit: string } & SergeantsUnitMetrics>;
};

type PersonLike = {
  fullName: string;
  personId?: string;
  rankGroup: RankGroup;
};

const emptyMetrics = (): SergeantsUnitMetrics => ({
  combatKilled: 0,
  combatSanitary: 0,
  combatMissing: 0,
  combatPow: 0,
  nonCombatIrreversible: 0,
  nonCombatSanitary: 0,
  onTreatment: 0,
  szchMonth: 0,
});

const isSergeant = (person: PersonLike) => person.rankGroup === "sergeant";

const personKey = (person: PersonLike) =>
  person.personId || normalizeLooseText(person.fullName);

const isRozporadzhennyaMovement = (...parts: string[]) => {
  const text = normalizeLooseText(parts.filter(Boolean).join(" "));
  return /(розпорядж|у\s*розпор)/.test(text);
};

/** Особи, які за 1ПБ «Рух» уже виведені у розпорядження. */
export const buildRozporadzhennyaPersonKeys = (
  workbook: ExcelWorkbookSnapshot,
) => {
  const keys = new Set<string>();
  for (const movement of parsePbMovements(workbook)) {
    if (
      !isRozporadzhennyaMovement(
        movement.type,
        movement.status,
        movement.note,
        movement.destination,
        movement.changeText,
        movement.nextIndex,
      )
    ) {
      continue;
    }
    keys.add(movement.personId || normalizeLooseText(movement.fullName));
  }
  return keys;
};

const buildStaffUnitLookup = (staff: ExcelWorkbookSnapshot) => {
  const byName = new Map<string, string>();
  const state = parseExcelLabState(staff.sheets);
  for (const [fullName, entry] of Object.entries(state)) {
    byName.set(normalizeLooseText(fullName), entry.unit || "Без підрозділу");
  }
  return (person: PersonLike) =>
    byName.get(normalizeLooseText(person.fullName)) ?? "Без підрозділу";
};

const bump = (
  bucket: Map<string, SergeantsUnitMetrics>,
  unit: string,
  patch: Partial<SergeantsUnitMetrics>,
) => {
  const current = bucket.get(unit) ?? emptyMetrics();
  bucket.set(unit, {
    combatKilled: current.combatKilled + (patch.combatKilled ?? 0),
    combatSanitary: current.combatSanitary + (patch.combatSanitary ?? 0),
    combatMissing: current.combatMissing + (patch.combatMissing ?? 0),
    combatPow: current.combatPow + (patch.combatPow ?? 0),
    nonCombatIrreversible:
      current.nonCombatIrreversible + (patch.nonCombatIrreversible ?? 0),
    nonCombatSanitary:
      current.nonCombatSanitary + (patch.nonCombatSanitary ?? 0),
    onTreatment: current.onTreatment + (patch.onTreatment ?? 0),
    szchMonth: current.szchMonth + (patch.szchMonth ?? 0),
  });
};

const sumMetrics = (rows: SergeantsUnitMetrics[]): SergeantsUnitMetrics =>
  rows.reduce(
    (acc, row) => ({
      combatKilled: acc.combatKilled + row.combatKilled,
      combatSanitary: acc.combatSanitary + row.combatSanitary,
      combatMissing: acc.combatMissing + row.combatMissing,
      combatPow: acc.combatPow + row.combatPow,
      nonCombatIrreversible:
        acc.nonCombatIrreversible + row.nonCombatIrreversible,
      nonCombatSanitary: acc.nonCombatSanitary + row.nonCombatSanitary,
      onTreatment: acc.onTreatment + row.onTreatment,
      szchMonth: acc.szchMonth + row.szchMonth,
    }),
    emptyMetrics(),
  );

export const classifySergeantsCombatLoss = (
  ...parts: string[]
): keyof Pick<
  SergeantsUnitMetrics,
  "combatKilled" | "combatSanitary" | "combatMissing" | "combatPow"
> => {
  const text = normalizeLooseText(parts.filter(Boolean).join(" "));
  if (/(полон|полонен|захоплен)/.test(text)) return "combatPow";
  if (/(безвіст|зникл.*безв|\bзб\b|\b500\b)/.test(text)) return "combatMissing";
  if (/(загибл|загинул|смерт|убит|\b200\b)/.test(text)) return "combatKilled";
  if (/(поранен|\b300\b|санітар|травм|контуз)/.test(text)) return "combatSanitary";
  return "combatSanitary";
};

const combatPatchForReason = (
  reason: CombatLossReason,
  matchParts: string[],
): Partial<SergeantsUnitMetrics> => {
  const bucket = classifySergeantsCombatLoss(...matchParts);
  if (bucket !== "combatSanitary") {
    return { [bucket]: 1 };
  }
  if (reason === "killed") return { combatKilled: 1 };
  if (reason === "missing") return { combatMissing: 1 };
  return { combatSanitary: 1 };
};

const nonCombatPatch = (
  category: DepartureCategory,
  ground: string,
): Partial<SergeantsUnitMetrics> => {
  if (category === "transfer") return {};
  const text = normalizeLooseText(ground);
  if (/(поранен|\b300\b|санітар|травм|контуз|здоров)/.test(text)) {
    return { nonCombatSanitary: 1 };
  }
  return { nonCombatIrreversible: 1 };
};

export const buildSergeantsByUnitReport = (input: {
  pb: ExcelWorkbookSnapshot;
  ejoos: ExcelWorkbookSnapshot;
  staff: ExcelWorkbookSnapshot;
  month: number;
  year: number;
}): SergeantsByUnitReport => {
  const { pb, ejoos, staff, month, year } = input;
  const resolveUnit = buildStaffUnitLookup(staff);
  const since = monthSinceDate(month, year);
  const until = monthUntilDate(month, year);
  const byUnit = new Map<string, SergeantsUnitMetrics>();

  const combat = buildCombatLossesFromPb(pb, { sinceDate: since });
  for (const person of combat.people) {
    if (!isSergeant(person)) continue;
    if (!isDateInMonth(person.orderDate, month, year)) continue;
    bump(
      byUnit,
      resolveUnit(person),
      combatPatchForReason(person.reason, [
        person.typeOrAbsence,
        person.status,
        person.placeOrNote,
        person.matchNote,
      ]),
    );
  }

  const departures = buildDeparturesFromEjoos(ejoos, { sinceDate: since });
  for (const person of departures.people) {
    if (!isSergeant(person)) continue;
    const eventDate = person.excludeDate || person.orderDate;
    if (!isDateInMonth(eventDate, month, year)) continue;
    bump(
      byUnit,
      resolveUnit(person),
      nonCombatPatch(person.category, `${person.ground} ${person.type}`),
    );
  }

  const disposition = buildDispositionFromArchive(pb, {
    sinceDate: "01.01.2026",
    openOnly: true,
  });
  const rozporKeys = buildRozporadzhennyaPersonKeys(pb);
  for (const person of disposition.people) {
    if (!isSergeant(person)) continue;
    if (person.reason !== "treatment") continue;
    if (rozporKeys.has(personKey(person))) continue;
    bump(byUnit, resolveUnit(person), { onTreatment: 1 });
  }

  const szch = buildSzchFromRuh(pb, { sinceDate: since, untilDate: until });
  for (const person of szch.people) {
    if (!isSergeant(person)) continue;
    bump(byUnit, resolveUnit(person), { szchMonth: 1 });
  }

  const rows = [...byUnit.entries()]
    .map(([unit, metrics]) => ({ unit, ...metrics }))
    .sort((a, b) => a.unit.localeCompare(b.unit, "uk"));

  return {
    month,
    year,
    totals: sumMetrics(rows),
    byUnit: rows,
  };
};
