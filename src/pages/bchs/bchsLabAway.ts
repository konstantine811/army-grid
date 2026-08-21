import {
  BCHS_EXCEL_COMMAND_ROSTER_UNITS,
  BCHS_RANK_OFFICER,
  BCHS_RANK_SERGEANT,
  BCHS_RANK_SOLDIER,
  computeBchsUnitAwayStats,
  isBchsDetachedStatus,
  matchesBchsRosterUnit,
  normalizeBchsText,
} from "./bchsCalc";
import type { BchsLabPersonRow } from "./bchsLabParse";

/** @deprecated use BCHS_EXCEL_COMMAND_ROSTER_UNITS from bchsCalc */
export const BCHS_LAB_COMMAND_ROSTER_UNITS = BCHS_EXCEL_COMMAND_ROSTER_UNITS;

export type BchsLabAwayRank = "оф." | "серж." | "солд." | "all";

export type BchsLabAwayCountOptions = {
  /** U = *Відком. за межі ПБ*. Default: true. */
  requireDetached?: boolean;
  /** Точний match по колонці B (як Excel). */
  rosterUnits?: string[];
  /** Рядок Аркуш1 — через matchesBchsRosterUnit (ширше за Excel). */
  arkush1Unit?: string;
  rank?: BchsLabAwayRank;
};

export type BchsLabAwayBucket = {
  officers: number;
  sergeants: number;
  soldiers: number;
  total: number;
  people: BchsLabPersonRow[];
};

export type BchsLabAwayCountResult = BchsLabAwayBucket & {
  byRosterUnit: Record<string, BchsLabAwayBucket>;
};

const emptyBucket = (): BchsLabAwayBucket => ({
  officers: 0,
  sergeants: 0,
  soldiers: 0,
  total: 0,
  people: [],
});

const addPersonToBucket = (bucket: BchsLabAwayBucket, person: BchsLabPersonRow) => {
  const rank = person.normalized.rankCategory;
  if (rank === BCHS_RANK_OFFICER) bucket.officers += 1;
  else if (rank === BCHS_RANK_SERGEANT) bucket.sergeants += 1;
  else if (rank === BCHS_RANK_SOLDIER) bucket.soldiers += 1;
  else return;

  bucket.total += 1;
  bucket.people.push(person);
};

const matchesOptions = (
  person: BchsLabPersonRow,
  options: BchsLabAwayCountOptions,
) => {
  // Lab завжди рахує лише battalion A = «нова».
  if (person.normalized.battalion !== "нова") return false;
  if (options.requireDetached !== false && !isBchsDetachedStatus(person.status))
    return false;

  if (options.rosterUnits?.length) {
    const allowed = new Set(
      options.rosterUnits.map((unit) => normalizeBchsText(unit)),
    );
    if (!allowed.has(person.normalized.rosterUnit)) return false;
  } else if (options.arkush1Unit) {
    if (!matchesBchsRosterUnit(options.arkush1Unit, person.rosterUnit))
      return false;
  }

  if (options.rank && options.rank !== "all") {
    if (person.normalized.rankCategory !== options.rank) return false;
  }

  return true;
};

/** COUNTIFS-style відкомандировані: A=нова + B + U + I. */
export const countBchsLabAway = (
  novaPeople: BchsLabPersonRow[],
  options: BchsLabAwayCountOptions = {},
): BchsLabAwayCountResult => {
  const result = emptyBucket();
  const byRosterUnit: Record<string, BchsLabAwayBucket> = {};

  novaPeople.forEach((person) => {
    if (!matchesOptions(person, options)) return;

    addPersonToBucket(result, person);

    const unitKey = person.rosterUnit.trim() || "(порожній B)";
    const bucket = byRosterUnit[unitKey] ?? emptyBucket();
    addPersonToBucket(bucket, person);
    byRosterUnit[unitKey] = bucket;
  });

  return { ...result, byRosterUnit };
};

/** Сума трьох COUNTIFS для Командування (ж + штаб + група БС), завжди нова. */
export const countBchsLabAwayCommandExcel = (
  novaPeople: BchsLabPersonRow[],
  options: Omit<BchsLabAwayCountOptions, "rosterUnits" | "arkush1Unit"> = {},
) => {
  const merged = emptyBucket();
  const byFormulaUnit: Record<string, BchsLabAwayBucket> = {};

  BCHS_LAB_COMMAND_ROSTER_UNITS.forEach((rosterUnit) => {
    const part = countBchsLabAway(novaPeople, {
      ...options,
      rosterUnits: [rosterUnit],
    });
    byFormulaUnit[rosterUnit] = part;
    part.people.forEach((person) => addPersonToBucket(merged, person));
  });

  return { ...merged, byFormulaUnit };
};

export const logBchsLabAwayReport = (
  novaPeople: BchsLabPersonRow[],
  arkush1Unit = "Командування",
) => {
  const excelExact = countBchsLabAwayCommandExcel(novaPeople, {
    requireDetached: true,
  });
  const appUnit = countBchsLabAway(novaPeople, { arkush1Unit });
  const prodStats = computeBchsUnitAwayStats(novaPeople, arkush1Unit);

  console.group(`[BCHS Lab] AK–AN · ${arkush1Unit} · A=нова`);
  console.log("COUNTIFS (ж+штаб+група БС, нова, U=відком):", {
    AK: excelExact.officers,
    AL: excelExact.sergeants,
    AM: excelExact.soldiers,
    AN: excelExact.total,
  });
  console.log("By B (formula units):", excelExact.byFormulaUnit);
  console.log("App matchesBchsRosterUnit (production):", {
    AK: appUnit.officers,
    AL: appUnit.sergeants,
    AM: appUnit.soldiers,
    AN: appUnit.total,
  });
  console.log("computeBchsUnitAwayStats (export):", prodStats);
  console.table(
    excelExact.people.map((person) => ({
      row: person.excelRowNumber,
      A: person.battalion,
      B: person.rosterUnit,
      I: person.rankCategory,
      U: person.status,
      AC: person.destination,
      ПІБ: person.fullName,
    })),
  );

  if (excelExact.soldiers === 0) {
    console.info(
      "[BCHS Lab] AM=0: серед відкомандированих (нова) у ж/штаб/група БС немає I=«солд.»",
    );
  }

  console.groupEnd();

  return { excelExact, appUnit, prodStats };
};

export type BchsLabConsoleApi = {
  countAway: (options?: BchsLabAwayCountOptions) => BchsLabAwayCountResult;
  awayCommand: (
    options?: Omit<BchsLabAwayCountOptions, "rosterUnits" | "arkush1Unit">,
  ) => ReturnType<typeof countBchsLabAwayCommandExcel>;
  logAway: (arkush1Unit?: string) => ReturnType<typeof logBchsLabAwayReport>;
};

export const createBchsLabConsoleApi = (
  novaPeople: BchsLabPersonRow[],
): BchsLabConsoleApi => ({
  countAway: (options = {}) => {
    const result = countBchsLabAway(novaPeople, options);
    console.log("[BCHS Lab] countAway · A=нова", options, result);
    return result;
  },
  awayCommand: (options = {}) => {
    const result = countBchsLabAwayCommandExcel(novaPeople, options);
    console.log("[BCHS Lab] awayCommand · A=нова", options, result);
    return result;
  },
  logAway: (arkush1Unit = "Командування") =>
    logBchsLabAwayReport(novaPeople, arkush1Unit),
});
