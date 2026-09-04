import type { BackendEjournalImport } from "../../api";
import type {
  CellValue,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { getColumnLabel, valueToDisplay } from "../../excelRoundTrip";
import type { AnalyticsMetric } from "../analytics/analyticsData";
import type {
  BchsAnalyticsRow,
  BchsAnalyticsSnapshot,
  BchsAnalyticsTableColumn,
  BchsAnalyticsTableRow,
  BchsComparisonRow,
  BchsDataIssue,
  BchsPersonnelAwayPerson,
  BchsPersonnelStayStats,
  BchsSupplementRow,
  BchsSupplementSnapshot,
  BchsUnitAbsenceCategoryStats,
  BchsUnitAttachedStats,
  BchsUnitAwayStats,
  BchsUnitCombatComponentStats,
} from "./bchsTypes";

export const BCHS_PIB_FILL_VALUE_KEY = "__pibFill";

export const columnLetterToIndex = (letter: string) =>
  letter
    .toUpperCase()
    .split("")
    .reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;

export const getSheetValue = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
  rowNumber: number,
  columnLetter: string,
) => {
  if (!sheet) return undefined;

  const originalColumnIndex = columnLetterToIndex(columnLetter);
  const normalizedColumnIndex = sheet.columnIndexes.findIndex(
    (index) => index === originalColumnIndex,
  );

  return sheet.rawRows[rowNumber - 1]?.[
    normalizedColumnIndex >= 0 ? normalizedColumnIndex : originalColumnIndex
  ];
};

export const bchsToNumber = (value: unknown) => {
  const numberValue = Number(value);
  if (Number.isFinite(numberValue)) return numberValue;

  const textValue = valueToDisplay(value as CellValue).trim();
  if (!textValue) return 0;

  const parsedValue = Number(
    textValue.replace(",", ".").replace(/[^\d.-]/g, ""),
  );
  if (!Number.isFinite(parsedValue)) return 0;

  return textValue.includes("%") && Math.abs(parsedValue) > 1
    ? parsedValue / 100
    : parsedValue;
};

export const toPercent = (value: unknown) =>
  `${Math.round(bchsToNumber(value) * 100)}%`;

export const formatRatioPercent = (value: number, total: number) => {
  if (!total) return "0%";
  return `${Math.round((value / total) * 100)}%`;
};

export const normalizeBchsPercentValue = (
  value: unknown,
  numerator: number,
  denominator: number,
) => {
  const percent = bchsToNumber(value);
  if (percent > 0 || !denominator) return percent;

  return numerator / denominator;
};

export const emptyBchsRow: BchsAnalyticsRow = {
  rowNumber: 0,
  unit: "Усього",
  staff: 0,
  staffOfficers: 0,
  staffSergeants: 0,
  staffSoldiers: 0,
  listed: 0,
  listedOfficers: 0,
  listedSergeants: 0,
  listedSoldiers: 0,
  staffedPercent: 0,
  available: 0,
  availableOfficers: 0,
  availableSergeants: 0,
  availableSoldiers: 0,
  shortage: 0,
  shortagePercent: 0,
  absent: 0,
  businessTrip: 0,
  absentOfficers: 0,
  absentSergeants: 0,
  absentSoldiers: 0,
  training: 0,
  hospitalWounded: 0,
  hospitalIllness: 0,
  vacation: 0,
  awol: 0,
  missing: 0,
  killed: 0,
  medWounded: 0,
  medIllness: 0,
  inRanksActually: 0,
  actualPercent: 0,
  combatComponent: 0,
};

export const BCHS_ANALYTICS_START_COLUMN = columnLetterToIndex("B");
export const BCHS_ANALYTICS_END_COLUMN = columnLetterToIndex("BL");
export const BCHS_PERCENT_COLUMNS = new Set(["K", "U", "BA"]);

/**
 * Аркуш1 layout for "Виконують завдання в інших підрозділах полку":
 * AK=Офіцери, AL=Сержанти, AM=Солдати, AN=Всього, AO=Куди відкомандировані
 *
 * Counted from Аркуш2 like Excel COUNTIFS:
 * - A (battalion) = "нова"
 * - U (status) contains "Відком. за межі ПБ"
 * - B (підрозділ) matches the Аркуш1 unit aliases
 * - I (rank) = Оф. / Серж. / солд.
 * - destinations from AC "В якому підрозділі"
 *
 * Аркуш1 "В розташуванні з інших підрозділів полку":
 * AP=Офіцери, AQ=Сержанти, AR=Солдати, AS=AP+AQ+AR, AT=Звідки прикомандировані
 *
 * Лише два рядки Аркуш1 (решта = 0):
 * - Інженерно-саперне відділення: B=*ПРИКОМАНДИРОВАНІ*, U=в строю|Новоприбулий, звання з M
 * - Відділення РРЕБ: X=*БРЕЗ*, U=в строю|Новоприбулий, звання з M
 * Без фільтра A=«нова» (прикомандировані з інших батальйонів).
 */

export const normalizeBchsText = (value: unknown) =>
  valueToDisplay(value as CellValue)
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase()
    .replace(/ё/g, "е");

/**
 * Для зіставлення звань: українська «і/ї», російська «и» і латинські
 * lookalike-літери зводяться до кирилиці (інакше COUNTIFS *капитан*
 * не бачить «капітан» / «капitan»).
 */
export const normalizeBchsRankMatchText = (value: unknown) =>
  normalizeBchsText(value)
    .replace(/[іїïíìıi]/g, "и")
    .replace(/[àáâäa]/g, "а")
    .replace(/[èéêëe]/g, "е")
    .replace(/[òóôöo]/g, "о")
    .replace(/[pρ]/g, "р")
    .replace(/c/g, "с")
    .replace(/y/g, "у")
    .replace(/t/g, "т")
    .replace(/n/g, "н")
    .replace(/m/g, "м")
    .replace(/k/g, "к")
    .replace(/['’`´]/g, "");

/** Exact B values from Excel COUNTIFS for row «Командування» (AK–AN). */
export const BCHS_EXCEL_COMMAND_ROSTER_UNITS = [
  "ж",
  "штаб",
  "група безпілотних систем",
] as const;

/** Аркуш1 sub-rows → exact B on Аркуш2 (Управління = ж, …). */
export const BCHS_SUBUNIT_ROSTER_ALIASES: Record<string, readonly string[]> = {
  управління: ["ж"],
  штаб: ["штаб"],
  "група безпілотних систем": ["група безпілотних систем"],
};

export const BCHS_COMMAND_ROSTER_ALIASES = [
  ...BCHS_EXCEL_COMMAND_ROSTER_UNITS,
  "управління",
  "командування",
] as const;

/** Є реальний ПІБ — вакантні / порожні рядки штатки не рахуємо як людей. */
export const hasBchsFullName = (person: { fullName?: unknown }) => {
  const name = normalizeBchsText(person.fullName);
  return Boolean(name) && name !== "-" && name !== "без піб" && name !== "(без піб)";
};

/** Перший фільтр у всьому pipeline БЧС: battalion A = «нова» і є ПІБ. */
export const filterBchsNovaPeople = <T extends { battalion: string }>(
  people: T[],
) =>
  people.filter((person) => {
    if (normalizeBchsText(person.battalion) !== "нова") return false;
    if (!("fullName" in person)) return true;
    return hasBchsFullName(person);
  });

/** Взяти roster з джерел і одразу застосувати фільтр «нова» (перший крок перед розрахунками). */
export const resolveBchsRosterPeople = (
  ...sources: Array<BchsPersonnelAwayPerson[] | undefined | null>
): BchsPersonnelAwayPerson[] => {
  for (const source of sources) {
    if (source?.length) return filterBchsNovaPeople(source);
  }
  return [];
};

/** Повний roster без фільтра «нова» — для AP–AT (ПРИКОМАНДИРОВАНІ / БРЕЗ). */
export const resolveBchsFullRosterPeople = (
  ...sources: Array<BchsPersonnelAwayPerson[] | undefined | null>
): BchsPersonnelAwayPerson[] => {
  const lists = sources.filter(
    (source): source is BchsPersonnelAwayPerson[] => Boolean(source?.length),
  );
  if (!lists.length) return [];

  const base = lists.reduce((best, list) =>
    list.length > best.length ? list : best,
  );
  const withFills =
    lists.find((list) =>
      list.some((person) => Boolean(person.pibHighlightRgb)),
    ) ?? null;
  if (!withFills || withFills === base) return base;

  const fillByName = new Map<string, string>();
  withFills.forEach((person) => {
    const name = normalizeBchsText(person.fullName);
    if (name && person.pibHighlightRgb && !fillByName.has(name)) {
      fillByName.set(name, person.pibHighlightRgb);
    }
  });

  return base.map((person) => {
    if (person.pibHighlightRgb) return person;
    const fill = fillByName.get(normalizeBchsText(person.fullName));
    return fill ? { ...person, pibHighlightRgb: fill } : person;
  });
};

export const BCHS_RANK_OFFICER = "оф.";
export const BCHS_RANK_SERGEANT = "серж.";
export const BCHS_RANK_SOLDIER = "солд.";

export const normalizeBchsDestinationLabel = (value: string) => {
  const label = value.replace(/\s+/g, " ").trim();
  const normalized = label.toLowerCase().replace(/ё/g, "е");

  if (normalized === "тринзитер" || normalized === "транзитер")
    return "Транзитер";
  if (normalized === "швал" || normalized === "шквал") return "Шквал";
  if (normalized === "полк") return "Полк";
  if (
    normalized === "рреб" ||
    normalized === "реб" ||
    normalized === "рота реб" ||
    normalized === "рота рреб"
  )
    return "РРЕБ";
  if (normalized === "знахарь" || normalized === "знахар") return "Знахарь";
  if (normalized === "полігон б" || normalized === "полигон б")
    return "Полігон Б";
  if (normalized === "рекрутинг") return "Рекрутинг";
  if (normalized === "рбпак") return "РБпАК";
  if (normalized === "ппо") return "ППО";
  if (normalized === "2пб") return "2ПБ";
  if (normalized === "2ббс") return "2ББС";
  if (normalized === "бмз") return "БМЗ";

  return label;
};

export const parseBchsDestinationText = (value: unknown) => {
  const text = valueToDisplay(value as CellValue);
  const metrics = new Map<string, number>();
  const pattern = /([^,;\n\r]+?)\s*[-–—−‒]\s*(\d+)/g;
  let match: RegExpExecArray | null;

  while ((match = pattern.exec(text)) !== null) {
    const label = normalizeBchsDestinationLabel(match[1] ?? "");
    const count = Number(match[2]);

    if (label && Number.isFinite(count) && count > 0)
      metrics.set(label, (metrics.get(label) ?? 0) + count);
  }

  return metrics;
};

export const sumBchsDestinationCounts = (metrics: Map<string, number>) =>
  Array.from(metrics.values()).reduce((sum, value) => sum + value, 0);

export const formatBchsDestinationText = (metrics: Map<string, number>) =>
  Array.from(metrics.entries())
    .filter(([, value]) => value > 0)
    .sort(
      (left, right) =>
        right[1] - left[1] || left[0].localeCompare(right[0], "uk"),
    )
    .map(([label, value]) => `${label}-${value}`)
    .join("\n");

export const isBchsCommandUnit = (unitName: string) =>
  /команд/i.test(unitName.trim());

export const isBchsTotalUnit = (unitName: string) =>
  /усього|всего|разом/i.test(unitName.trim());

export const normalizeBchsRosterUnitLabel = (unitName: string) => {
  const label = unitName.replace(/\s+/g, " ").trim();
  const normalized = normalizeBchsText(label);
  if (!normalized) return "Без підрозділу";
  if (
    normalized === "ж" ||
    normalized === "штаб" ||
    normalized === "управління" ||
    normalized === "командування" ||
    normalized.includes("група безпілотних систем")
  )
    return "Командування";
  return label;
};

export const isBchsDetachedStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  // Require "за межі" so "в межах ПБ" is not counted.
  return normalized.includes("відком") && normalized.includes("за межі");
};

/** Excel V:Y «Відсутні» = U contains «Відрядження» (not «Відком. за межі ПБ»). */
export const isBchsBusinessTripStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized.includes("відряд") && !isBchsDetachedStatus(status);
};

export const emptyBchsUnitAwayStats = (): BchsUnitAwayStats => ({
  officers: 0,
  sergeants: 0,
  soldiers: 0,
  total: 0,
  destinations: new Map(),
  destinationText: "",
});

export const matchesBchsRosterUnit = (arkush1Unit: string, rosterUnit: string) => {
  const unit = normalizeBchsText(arkush1Unit);
  const rawRoster = normalizeBchsText(rosterUnit);
  if (!unit || !rawRoster) return false;

  const subunitAliases = BCHS_SUBUNIT_ROSTER_ALIASES[unit];
  if (subunitAliases) {
    return subunitAliases.some((alias) => rawRoster === normalizeBchsText(alias));
  }

  const roster = normalizeBchsText(normalizeBchsRosterUnitLabel(rosterUnit));

  if (isBchsCommandUnit(unit)) {
    // Excel COUNTIFS: exact match on column B (ж / штаб / група БС).
    return BCHS_EXCEL_COMMAND_ROSTER_UNITS.some(
      (alias) => rawRoster === normalizeBchsText(alias),
    );
  }

  if (roster === unit) return true;

  // Avoid false positives from short prefixes like "рота".
  const minPrefixLength = 12;
  if (
    roster.length >= minPrefixLength &&
    unit.length >= minPrefixLength &&
    (unit.startsWith(roster) || roster.startsWith(unit))
  ) {
    return true;
  }

  const compactUnit = unit.slice(0, 36);
  const compactRoster = roster.slice(0, 36);
  return (
    compactUnit.length >= minPrefixLength &&
    compactUnit === compactRoster
  );
};

export const logBchsDetachedPeopleDebug = (people: BchsPersonnelAwayPerson[]) => {
  const detached = filterBchsNovaPeople(people).filter((person) =>
    isBchsDetachedStatus(person.status),
  );

  const byUnit = new Map<
    string,
    Array<{
      піб: string;
      rankCategory: string;
      status: string;
      destination: string;
      destinationNormalized: string;
    }>
  >();

  detached.forEach((person) => {
    const unit = person.rosterUnit.trim() || "(порожній підрозділ)";
    const list = byUnit.get(unit) ?? [];
    list.push({
      піб: person.fullName || "(без ПІБ)",
      rankCategory: person.rankCategory,
      status: person.status,
      destination: person.destination,
      destinationNormalized:
        normalizeBchsDestinationLabel(person.destination) || "(порожньо)",
    });
    byUnit.set(unit, list);
  });

  const asObject = Object.fromEntries(
    Array.from(byUnit.entries())
      .sort((left, right) => left[0].localeCompare(right[0], "uk"))
      .map(([unit, items]) => [unit, items]),
  );

  const uavKey =
    Object.keys(asObject).find((unit) =>
      normalizeBchsText(unit).includes("рота безпілотних авіаційних"),
    ) ?? "Рота безпілотних авіаційних комплексів";

  console.groupCollapsed(
    `[BCHS] Відкомандировані · нова · усього ${detached.length}`,
  );
  console.log("Усі підрозділи (масив по підрозділах):", asObject);
  console.log(
    `Рота безпілотних авіаційних комплексів (${(asObject[uavKey] ?? []).length}):`,
    asObject[uavKey] ?? [],
  );
  console.table(
    detached.map((person) => ({
      ПІБ: person.fullName || "(без ПІБ)",
      підрозділ: person.rosterUnit,
      rank: person.rankCategory,
      статус: person.status,
      "В якому підрозділі": person.destination,
    })),
  );
  console.groupEnd();
};

export const computeBchsUnitAwayStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitAwayStats => {
  if (isBchsTotalUnit(unitName)) return emptyBchsUnitAwayStats();

  const stats = emptyBchsUnitAwayStats();
  const novaPeople = filterBchsNovaPeople(people);

  novaPeople.forEach((person) => {
    if (!isBchsDetachedStatus(person.status)) return;
    if (!matchesBchsRosterUnit(unitName, person.rosterUnit)) return;

    const rank = normalizeBchsText(person.rankCategory);
    if (rank === BCHS_RANK_OFFICER) stats.officers += 1;
    else if (rank === BCHS_RANK_SERGEANT) stats.sergeants += 1;
    else if (rank === BCHS_RANK_SOLDIER) stats.soldiers += 1;
    else return;

    const destination =
      normalizeBchsDestinationLabel(person.destination) || "невідомо";
    stats.destinations.set(
      destination,
      (stats.destinations.get(destination) ?? 0) + 1,
    );
  });

  stats.total = stats.officers + stats.sergeants + stats.soldiers;
  stats.destinationText = formatBchsDestinationText(stats.destinations);
  return stats;
};

export type BchsUnitRankBucketStats = {
  officers: number;
  sergeants: number;
  soldiers: number;
  total: number;
};

export type BchsUnitRankListedAvailableStats = {
  /** Excel G/H/I — COUNTIFS(I, B, M<>). */
  listed: BchsUnitRankBucketStats;
  /** Excel L/M/N — «в строю» + «Відком. за межі ПБ». */
  available: BchsUnitRankBucketStats;
  /** Excel V/W/X — listed − available. */
  absent: BchsUnitRankBucketStats;
};

const emptyBchsUnitRankBucketStats = (): BchsUnitRankBucketStats => ({
  officers: 0,
  sergeants: 0,
  soldiers: 0,
  total: 0,
});

/** Excel «за списком» / V:Y: M (звання) заповнене. */
export const isBchsListedWithRankTitle = (person: BchsPersonnelAwayPerson) =>
  Boolean(person.rankTitle.trim());

/** Excel subtract part for V/W/X: «в строю» + «Відком. за межі ПБ». */
export const isBchsAvailableForAbsentFormula = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized.includes("в строю") || isBchsDetachedStatus(status);
};

const classifyBchsRankCategory = (
  rankCategory: string,
): "officers" | "sergeants" | "soldiers" | null => {
  const rank = normalizeBchsText(rankCategory);
  if (rank === BCHS_RANK_OFFICER) return "officers";
  if (rank === BCHS_RANK_SERGEANT) return "sergeants";
  if (rank === BCHS_RANK_SOLDIER || rank.includes("солд")) return "soldiers";
  return null;
};

const finalizeBchsUnitRankBucketStats = (
  stats: Omit<BchsUnitRankBucketStats, "total">,
): BchsUnitRankBucketStats => ({
  ...stats,
  total: stats.officers + stats.sergeants + stats.soldiers,
});

/** Excel G–I / L–N / V–X / O (= SUM L:N), AJ (= O). */
export const computeBchsUnitRankListedAvailableStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitRankListedAvailableStats => {
  if (isBchsTotalUnit(unitName)) {
    const empty = emptyBchsUnitRankBucketStats();
    return { listed: empty, available: empty, absent: empty };
  }

  const listed = { officers: 0, sergeants: 0, soldiers: 0 };
  const available = { officers: 0, sergeants: 0, soldiers: 0 };

  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsListedWithRankTitle(person)) return;
    if (!matchesBchsRosterUnit(unitName, person.rosterUnit)) return;

    const bucket = classifyBchsRankCategory(person.rankCategory);
    if (!bucket) return;

    listed[bucket] += 1;
    if (isBchsAvailableForAbsentFormula(person.status)) {
      available[bucket] += 1;
    }
  });

  const listedStats = finalizeBchsUnitRankBucketStats(listed);
  const availableStats = finalizeBchsUnitRankBucketStats(available);

  return {
    listed: listedStats,
    available: availableStats,
    absent: finalizeBchsUnitRankBucketStats({
      officers: Math.max(0, listed.officers - available.officers),
      sergeants: Math.max(0, listed.sergeants - available.sergeants),
      soldiers: Math.max(0, listed.soldiers - available.soldiers),
    }),
  };
};

/** Excel V/W/X: COUNTIFS(I,M<>)-COUNTIFS(в строю+відком), A=нова. */
export const computeBchsUnitAbsentRankStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitRankBucketStats =>
  computeBchsUnitRankListedAvailableStats(people, unitName).absent;

export const applyBchsRankListedAvailableToComparisonRow = (
  row: BchsComparisonRow,
  stats: BchsUnitRankListedAvailableStats,
): BchsComparisonRow =>
  createBchsComparisonRow({
    ...row,
    listedOfficers: stats.listed.officers,
    listedSergeants: stats.listed.sergeants,
    listedSoldiers: stats.listed.soldiers,
    listed: stats.listed.total,
    availableOfficers: stats.available.officers,
    availableSergeants: stats.available.sergeants,
    availableSoldiers: stats.available.soldiers,
    available: stats.available.total,
    absentOfficers: stats.absent.officers,
    absentSergeants: stats.absent.sergeants,
    absentSoldiers: stats.absent.soldiers,
    absent: stats.absent.total,
  });

export const sumBchsComparisonRankListedAvailable = (
  rows: BchsComparisonRow[],
): BchsUnitRankListedAvailableStats => {
  const listed = finalizeBchsUnitRankBucketStats({
    officers: rows.reduce((sum, row) => sum + row.listedOfficers, 0),
    sergeants: rows.reduce((sum, row) => sum + row.listedSergeants, 0),
    soldiers: rows.reduce((sum, row) => sum + row.listedSoldiers, 0),
  });
  const available = finalizeBchsUnitRankBucketStats({
    officers: rows.reduce((sum, row) => sum + row.availableOfficers, 0),
    sergeants: rows.reduce((sum, row) => sum + row.availableSergeants, 0),
    soldiers: rows.reduce((sum, row) => sum + row.availableSoldiers, 0),
  });
  const absent = finalizeBchsUnitRankBucketStats({
    officers: rows.reduce((sum, row) => sum + row.absentOfficers, 0),
    sergeants: rows.reduce((sum, row) => sum + row.absentSergeants, 0),
    soldiers: rows.reduce((sum, row) => sum + row.absentSoldiers, 0),
  });

  return { listed, available, absent };
};

/** Excel Z: COUNTIFS(U;"*Відрядження*",B,M<>), A=нова — без розбивки по I. */
export const computeBchsUnitBusinessTripCount = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): number => {
  if (isBchsTotalUnit(unitName)) return 0;

  let count = 0;
  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsListedWithRankTitle(person)) return;
    if (!matchesBchsRosterUnit(unitName, person.rosterUnit)) return;
    if (!isBchsBusinessTripStatus(person.status)) return;
    count += 1;
  });

  return count;
};

export const emptyBchsUnitAbsenceCategoryStats =
  (): BchsUnitAbsenceCategoryStats => ({
    training: 0,
    hospitalWounded: 0,
    hospitalIllness: 0,
    vacation: 0,
    awol: 0,
    missing: 0,
    killed: 0,
    medWounded: 0,
    medIllness: 0,
  });

/** Excel U: «*Лікування*». */
export const isBchsTreatmentStatus = (status: string) =>
  normalizeBchsText(status).includes("лікування");

/** Excel AA: «*Від-ні навчання*». */
export const isBchsTrainingStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized.includes("від") && normalized.includes("навч");
};

/** Excel AD: «*Відпустка*» (колонка U). */
export const isBchsVacationStatus = (status: string) =>
  normalizeBchsText(status).includes("відпуст");

/** Excel AE: «*СЗЧ*» + «*Не в сторою*». */
export const isBchsAwolStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized.includes("сзч") || normalized.includes("не в сторою");
};

export const isBchsMissingStatus = (status: string) =>
  normalizeBchsText(status).includes("зникл");

export const isBchsKilledStatus = (status: string) =>
  normalizeBchsText(status).includes("загиб");

/**
 * Універсальний маркер поранення в примітках / статусі.
 * Ловить «по пораненню», «поранення», «після поранення», окремий код 300.
 */
export const BCHS_WOUNDED_NOTE_RE =
  /(?:по\s*)?поран(?:ен\w*|\.)|(?:^|[^\d])300(?:[^\d]|$)/i;

/** Excel AF/T / колонка «Примітки»: поранення в будь-якому рядку клітинки. */
export const isBchsWoundedByExcelNote = (note: string) =>
  BCHS_WOUNDED_NOTE_RE.test(normalizeBchsText(note));

const isBchsExcelHospitalPlace = (medicalPlace: string) =>
  normalizeBchsText(medicalPlace).includes("шпитал");

const isBchsExcelMedicalLeavePlace = (medicalPlace: string) => {
  const normalized = normalizeBchsText(medicalPlace);
  return (
    normalized.includes("відпустка лікув") ||
    (normalized.includes("відпуст") && normalized.includes("лікув"))
  );
};

const isBchsExcelMedRotaPlace = (medicalPlace: string) => {
  const normalized = normalizeBchsText(medicalPlace);
  return normalized.includes("мед") && normalized.includes("рота");
};

const getBchsWoundedNoteForUnit = (
  person: BchsPersonnelAwayPerson,
  unitName: string,
) =>
  isBchsCommandUnit(unitName) ? person.treatmentNote : person.medicalNote;

const countBchsUnitListedPeople = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
  matches: (person: BchsPersonnelAwayPerson) => boolean,
): number => {
  if (isBchsTotalUnit(unitName)) return 0;

  let count = 0;
  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsListedWithRankTitle(person)) return;
    if (!matchesBchsRosterUnit(unitName, person.rosterUnit)) return;
    if (matches(person)) count += 1;
  });
  return count;
};

/** Excel AA–AI: COUNTIFS по U/AE/AF (+ AC/AI = total − wounded). */
export const computeBchsUnitAbsenceCategoryStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitAbsenceCategoryStats => {
  if (isBchsTotalUnit(unitName)) return emptyBchsUnitAbsenceCategoryStats();

  const isTreatmentPerson = (person: BchsPersonnelAwayPerson) =>
    isBchsTreatmentStatus(person.status);

  const isWoundedPerson = (person: BchsPersonnelAwayPerson) =>
    isBchsWoundedByExcelNote(getBchsWoundedNoteForUnit(person, unitName));

  const hospitalWounded =
    countBchsUnitListedPeople(
      people,
      unitName,
      (person) =>
        isTreatmentPerson(person) &&
        isBchsExcelHospitalPlace(person.medicalPlace) &&
        isWoundedPerson(person),
    ) +
    countBchsUnitListedPeople(
      people,
      unitName,
      (person) =>
        isTreatmentPerson(person) &&
        isBchsExcelMedicalLeavePlace(person.medicalPlace) &&
        isWoundedPerson(person),
    );

  const hospitalTotal =
    countBchsUnitListedPeople(
      people,
      unitName,
      (person) =>
        isTreatmentPerson(person) &&
        isBchsExcelHospitalPlace(person.medicalPlace),
    ) +
    countBchsUnitListedPeople(
      people,
      unitName,
      (person) =>
        isTreatmentPerson(person) &&
        isBchsExcelMedicalLeavePlace(person.medicalPlace),
    );

  const medWounded = countBchsUnitListedPeople(
    people,
    unitName,
    (person) =>
      isTreatmentPerson(person) &&
      isBchsExcelMedRotaPlace(person.medicalPlace) &&
      isWoundedPerson(person),
  );

  const medTotal = countBchsUnitListedPeople(
    people,
    unitName,
    (person) =>
      isTreatmentPerson(person) &&
      isBchsExcelMedRotaPlace(person.medicalPlace),
  );

  return {
    training: countBchsUnitListedPeople(people, unitName, (person) =>
      isBchsTrainingStatus(person.status),
    ),
    hospitalWounded,
    hospitalIllness: Math.max(0, hospitalTotal - hospitalWounded),
    vacation: countBchsUnitListedPeople(people, unitName, (person) =>
      isBchsVacationStatus(person.status),
    ),
    awol:
      countBchsUnitListedPeople(people, unitName, (person) =>
        normalizeBchsText(person.status).includes("сзч"),
      ) +
      countBchsUnitListedPeople(people, unitName, (person) =>
        normalizeBchsText(person.status).includes("не в сторою"),
      ),
    missing: countBchsUnitListedPeople(people, unitName, (person) =>
      isBchsMissingStatus(person.status),
    ),
    killed: countBchsUnitListedPeople(people, unitName, (person) =>
      isBchsKilledStatus(person.status),
    ),
    medWounded,
    medIllness: Math.max(0, medTotal - medWounded),
  };
};

export const extractBchsAwayPeopleFromSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
): BchsPersonnelAwayPerson[] => {
  if (!sheet?.rawRows?.length) return [];

  const columnIndex = (originalIndex: number) => {
    const normalized = sheet.columnIndexes.findIndex(
      (index) => index === originalIndex,
    );
    return normalized >= 0 ? normalized : originalIndex;
  };

  const battalionCol = columnIndex(0);
  const rosterCol = columnIndex(1);
  const positionCol = columnIndex(4);
  const shpkFactCol = columnIndex(7);
  const rankCol = columnIndex(8);
  const rankTitleCol = columnIndex(12);
  const fullNameCol = columnIndex(13);
  const callsignCol = columnIndex(14);
  const birthDateCol = columnIndex(15);
  const statusCol = columnIndex(20);
  const roleTypeCol = columnIndex(21);
  const combatReadinessCol = columnIndex(22);
  const bzvpStatusCol = columnIndex(23);
  const brezAssignmentCol = columnIndex(26);
  const mobilizationCol = columnIndex(11);
  const treatmentNoteCol = columnIndex(19);
  const destinationCol = columnIndex(28);
  const medicalPlaceCol = columnIndex(30);
  const medicalNoteCol = columnIndex(31);
  const directionCol = columnIndex(32);
  const medicalPlaceAltCol = columnIndex(39);
  const combatReadinessAltCol = columnIndex(41);

  const sourceRows =
    sheet.rows.length > 0
      ? sheet.rows.map((row) => ({
          values: row.values,
          excelRowNumber: row.excelRowNumber,
        }))
      : sheet.rawRows.slice(1).map((values, index) => ({
          values,
          excelRowNumber: index + 2,
        }));

  return sourceRows
    .map(({ values: row, excelRowNumber }) => ({
      battalion: valueToDisplay(row[battalionCol] as CellValue),
      rosterUnit: valueToDisplay(row[rosterCol] as CellValue),
      position: valueToDisplay(row[positionCol] as CellValue),
      shpkFact: valueToDisplay(row[shpkFactCol] as CellValue),
      rankCategory: valueToDisplay(row[rankCol] as CellValue),
      rankTitle: valueToDisplay(row[rankTitleCol] as CellValue),
      fullName: valueToDisplay(row[fullNameCol] as CellValue),
      callsign: valueToDisplay(row[callsignCol] as CellValue),
      birthDate: valueToDisplay(row[birthDateCol] as CellValue),
      status: valueToDisplay(row[statusCol] as CellValue),
      roleType: valueToDisplay(row[roleTypeCol] as CellValue),
      combatReadiness:
        valueToDisplay(row[combatReadinessCol] as CellValue) ||
        valueToDisplay(row[combatReadinessAltCol] as CellValue),
      bzvpStatus: valueToDisplay(row[bzvpStatusCol] as CellValue),
      brezAssignment: valueToDisplay(row[brezAssignmentCol] as CellValue),
      treatmentNote: valueToDisplay(row[treatmentNoteCol] as CellValue),
      mobilizationContract: valueToDisplay(row[mobilizationCol] as CellValue),
      destination: valueToDisplay(row[destinationCol] as CellValue),
      medicalPlace:
        valueToDisplay(row[medicalPlaceCol] as CellValue) ||
        valueToDisplay(row[medicalPlaceAltCol] as CellValue),
      medicalNote: valueToDisplay(row[medicalNoteCol] as CellValue),
      direction: valueToDisplay(row[directionCol] as CellValue),
      pibHighlightRgb: sheet.pibFillByExcelRow?.[excelRowNumber] ?? null,
    }))
    .filter(
      (person) =>
        person.rosterUnit ||
        person.position ||
        person.status ||
        person.rankCategory ||
        person.rankTitle ||
        person.fullName ||
        person.roleType ||
        person.combatReadiness ||
        person.bzvpStatus ||
        person.brezAssignment ||
        person.destination ||
        person.treatmentNote ||
        person.medicalPlace ||
        person.medicalNote ||
        person.direction,
    );
};

const pickBchsDbFirstNonEmpty = (
  row: Record<string, unknown>,
  candidates: Array<string | undefined>,
) => {
  const keys = candidates.filter((key): key is string => Boolean(key?.trim()));
  const withRosterPrefix = keys.flatMap((key) =>
    key.startsWith("roster__") ? [key] : [key, `roster__${key}`],
  );
  for (const key of withRosterPrefix) {
    const value = pickBchsDbCellValue(row, [key]).trim();
    if (value) return value;
  }
  return "";
};

type BchsDbColumnMeta = {
  key: string;
  letter?: string;
  originalIndex?: number;
};

const pickBchsDbCellValue = (
  row: Record<string, unknown>,
  candidates: Array<string | undefined>,
) => {
  const wanted = candidates
    .filter((key): key is string => Boolean(key && key.trim()))
    .map((key) => key.trim());

  for (const key of wanted) {
    if (Object.prototype.hasOwnProperty.call(row, key)) {
      return valueToDisplay(row[key] as CellValue);
    }
  }

  const lowerToActual = new Map(
    Object.keys(row).map((key) => [key.toLowerCase(), key] as const),
  );
  for (const key of wanted) {
    const actual = lowerToActual.get(key.toLowerCase());
    if (!actual) continue;
    return valueToDisplay(row[actual] as CellValue);
  }

  return "";
};

const resolveBchsDbColumnKeyByLetter = (
  columns: BchsDbColumnMeta[] | undefined,
  letter: string,
) => {
  if (!columns?.length) return undefined;
  const target = letter.trim().toUpperCase();
  const byLetter = columns.find(
    (column) => column.letter?.trim().toUpperCase() === target,
  );
  if (byLetter?.key) return byLetter.key;

  const byIndex = columns.find(
    (column) =>
      typeof column.originalIndex === "number" &&
      getColumnLabel(column.originalIndex).toUpperCase() === target,
  );
  return byIndex?.key;
};

/** Map DB row values (header keys and/or Excel letters) back to roster people. */
export const extractBchsAwayPeopleFromDbRows = (
  rows: Array<Record<string, unknown>>,
  columns?: BchsDbColumnMeta[],
): BchsPersonnelAwayPerson[] => {
  const keyA = resolveBchsDbColumnKeyByLetter(columns, "A");
  const keyB = resolveBchsDbColumnKeyByLetter(columns, "B");
  const keyE = resolveBchsDbColumnKeyByLetter(columns, "E");
  const keyH = resolveBchsDbColumnKeyByLetter(columns, "H");
  const keyI = resolveBchsDbColumnKeyByLetter(columns, "I");
  const keyM = resolveBchsDbColumnKeyByLetter(columns, "M");
  const keyN = resolveBchsDbColumnKeyByLetter(columns, "N");
  const keyO = resolveBchsDbColumnKeyByLetter(columns, "O");
  const keyP = resolveBchsDbColumnKeyByLetter(columns, "P");
  const keyL = resolveBchsDbColumnKeyByLetter(columns, "L");
  const keyU = resolveBchsDbColumnKeyByLetter(columns, "U");
  const keyV = resolveBchsDbColumnKeyByLetter(columns, "V");
  const keyW = resolveBchsDbColumnKeyByLetter(columns, "W");
  const keyX = resolveBchsDbColumnKeyByLetter(columns, "X");
  const keyAA = resolveBchsDbColumnKeyByLetter(columns, "AA");
  const keyT = resolveBchsDbColumnKeyByLetter(columns, "T");
  const keyAC = resolveBchsDbColumnKeyByLetter(columns, "AC");
  const keyAE = resolveBchsDbColumnKeyByLetter(columns, "AE");
  const keyAF = resolveBchsDbColumnKeyByLetter(columns, "AF");

  return rows.map((row) => ({
    // A: батальйон (нова) — без цього після F5 усі COUNTIFS по nova = 0.
    battalion: pickBchsDbCellValue(row, [
      keyA,
      "A",
      "батальйон",
      "battalion",
      "column_1",
      "№",
    ]),
    rosterUnit: pickBchsDbCellValue(row, [
      keyB,
      "B",
      "підрозділ",
      "rosterUnit",
      "column_2",
    ]),
    position: pickBchsDbCellValue(row, [
      keyE,
      "E",
      "посада",
      "position",
      "column_5",
    ]),
    shpkFact: pickBchsDbCellValue(row, [
      keyH,
      "H",
      "шпк",
      "shpkFact",
      "column_8",
    ]),
    rankCategory: pickBchsDbCellValue(row, [
      keyI,
      "I",
      "категорія",
      "rankCategory",
      "column_9",
    ]),
    rankTitle: pickBchsDbCellValue(row, [
      keyM,
      "M",
      "звання",
      "rankTitle",
      "column_13",
    ]),
    fullName: pickBchsDbCellValue(row, [
      keyN,
      "N",
      "піб",
      "fullName",
      "column_14",
    ]),
    callsign: pickBchsDbCellValue(row, [
      keyO,
      "O",
      "позивн",
      "callsign",
      "column_15",
    ]),
    birthDate: pickBchsDbCellValue(row, [
      keyP,
      "P",
      "дата народження",
      "народжен",
      "birthDate",
      "column_16",
    ]),
    status: pickBchsDbCellValue(row, [
      keyU,
      "U",
      "статус",
      "status",
      "column_21",
    ]),
    roleType: pickBchsDbCellValue(row, [
      keyV,
      "V",
      "тип_в_с",
      "roleType",
      "column_22",
    ]),
    combatReadiness: pickBchsDbFirstNonEmpty(row, [
      keyW,
      "W",
      "статус_бг",
      "combatReadiness",
      "column_23",
      "column_42",
    ]),
    bzvpStatus: pickBchsDbCellValue(row, [
      keyX,
      "X",
      "бзвп_брез",
      "bzvpStatus",
      "column_24",
    ]),
    brezAssignment: pickBchsDbCellValue(row, [
      keyAA,
      "AA",
      "відрядження_брез",
      "brezAssignment",
      "column_27",
    ]),
    treatmentNote: pickBchsDbCellValue(row, [
      keyT,
      "T",
      "примітки_лікування",
      "treatmentNote",
      "column_20",
    ]),
    mobilizationContract: pickBchsDbCellValue(row, [
      keyL,
      "L",
      "мобілізація",
      "mobilizationContract",
      "column_12",
    ]),
    destination: pickBchsDbCellValue(row, [
      keyAC,
      "AC",
      "в_якому_підрозділі",
      "externalUnit",
      "column_29",
    ]),
    medicalPlace: pickBchsDbFirstNonEmpty(row, [
      keyAE,
      "AE",
      "місце_перебування",
      "medicalPlace",
      "column_31",
      "column_40",
    ]),
    medicalNote: pickBchsDbCellValue(row, [
      keyAF,
      "AF",
      "примітки",
      "medicalNote",
      "column_32",
    ]),
    direction: pickBchsDbFirstNonEmpty(row, [
      "напрямок",
      "direction",
      "column_33",
    ]),
    fighterExitDate: pickBchsDbFirstNonEmpty(row, [
      "fighter_status_exit_date",
      "дата_виходу",
    ]),
    fighterReturnDate: pickBchsDbFirstNonEmpty(row, [
      "fighter_status_return_date",
      "дата_повернення",
    ]),
    pibHighlightRgb:
      pickBchsDbCellValue(row, [
        BCHS_PIB_FILL_VALUE_KEY,
        "pibHighlightRgb",
        "__pibFill",
      ]) || null,
  }));
};

/** Extract + перший фільтр «нова» — стандартний вхід для analytics/export. */
export const extractBchsNovaPeopleFromSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) => filterBchsNovaPeople(extractBchsAwayPeopleFromSheet(sheet));

export const extractBchsNovaPeopleFromDbRows = (
  rows: Array<Record<string, unknown>>,
  columns?: BchsDbColumnMeta[],
) => filterBchsNovaPeople(extractBchsAwayPeopleFromDbRows(rows, columns));

export const BCHS_EMPTY_STAY_PLACE_LABEL = "Не вказано";

export const isBchsBattleReadyPerson = (person: BchsPersonnelAwayPerson) =>
  normalizeBchsText(person.combatReadiness) === "бг";

const hasActiveFighterExit = (person: BchsPersonnelAwayPerson) => {
  const exitDate = String(person.fighterExitDate ?? "").trim();
  const returnDate = String(person.fighterReturnDate ?? "").trim();
  return Boolean(exitDate && !returnDate);
};

/** «на виході» / «на виконанні» / активний вихід у «Статус бійців». */
export const isBchsOnCombatExit = (person: BchsPersonnelAwayPerson) => {
  if (hasActiveFighterExit(person)) return true;

  const blob = [
    person.medicalPlace,
    person.status,
    person.medicalNote,
    person.direction,
    person.destination,
  ]
    .map((value) => normalizeBchsText(value))
    .join(" ");
  return /(?:на\s+)?виход[іи](?!н)|на\s+виконанн/.test(blob);
};

export const formatBchsStayPlaceLabel = (value: string) => {
  const trimmed = String(value ?? "").replace(/\s+/g, " ").trim();
  return trimmed || BCHS_EMPTY_STAY_PLACE_LABEL;
};

export const buildBchsStayPlaceStats = (
  people: BchsPersonnelAwayPerson[],
): AnalyticsMetric[] => {
  const counts = new Map<string, { label: string; value: number }>();
  for (const person of people) {
    const raw = String(person.medicalPlace ?? "").replace(/\s+/g, " ").trim();
    const key = raw ? normalizeBchsText(raw) : "";
    const label = raw || BCHS_EMPTY_STAY_PLACE_LABEL;
    const current = counts.get(key);
    if (current) current.value += 1;
    else counts.set(key, { label, value: 1 });
  }

  return Array.from(counts.values()).sort((left, right) => {
    const leftEmpty = left.label === BCHS_EMPTY_STAY_PLACE_LABEL;
    const rightEmpty = right.label === BCHS_EMPTY_STAY_PLACE_LABEL;
    if (leftEmpty !== rightEmpty) return leftEmpty ? 1 : -1;
    return right.value - left.value || left.label.localeCompare(right.label, "uk");
  });
};

export const summarizeBchsPersonnelStay = (
  people: BchsPersonnelAwayPerson[],
): BchsPersonnelStayStats => {
  const roster = filterBchsNovaPeople(people);
  return {
    total: roster.length,
    stayPlaces: buildBchsStayPlaceStats(roster),
    battleReady: roster.filter(isBchsBattleReadyPerson).length,
    notBattleReady: roster.filter((person) => !isBchsBattleReadyPerson(person))
      .length,
    onExit: roster
      .filter(isBchsOnCombatExit)
      .slice()
      .sort((left, right) =>
        left.fullName.localeCompare(right.fullName, "uk"),
      ),
  };
};

export const applyBchsAwayFromPersonnel = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  logBchsDetachedPeopleDebug(people);

  const unitRows = analytics.comparisonRows.map((row) => {
    if (row.rowNumber <= 11 || isBchsTotalUnit(row.unit)) return row;

    const stats = computeBchsUnitAwayStats(people, row.unit);
    return createBchsComparisonRow({
      ...row,
      awayOfficers: stats.officers,
      awaySergeants: stats.sergeants,
      awaySoldiers: stats.soldiers,
      awayInOtherUnits: stats.total,
      awayDestinationsText: stats.destinationText,
    });
  });

  const detailRows = unitRows.filter(
    (row) => row.rowNumber > 11 && !isBchsTotalUnit(row.unit),
  );
  const totalAwayOfficers = detailRows.reduce(
    (sum, row) => sum + row.awayOfficers,
    0,
  );
  const totalAwaySergeants = detailRows.reduce(
    (sum, row) => sum + row.awaySergeants,
    0,
  );
  const totalAwaySoldiers = detailRows.reduce(
    (sum, row) => sum + row.awaySoldiers,
    0,
  );
  const totalAway = totalAwayOfficers + totalAwaySergeants + totalAwaySoldiers;
  const totalDestinations = mergeBchsMetricMaps(
    detailRows.map((row) => parseBchsDestinationText(row.awayDestinationsText)),
  );

  const comparisonRows = unitRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;

    return createBchsComparisonRow({
      ...row,
      awayOfficers: totalAwayOfficers,
      awaySergeants: totalAwaySergeants,
      awaySoldiers: totalAwaySoldiers,
      awayInOtherUnits: totalAway,
      // Like Excel: no AO breakdown on "Усього по основним підрозділам"
      awayDestinationsText: "",
    });
  });

  const table = analytics.table
    ? {
        columns: analytics.table.columns,
        rows: analytics.table.rows.map((tableRow) => {
          const match =
            comparisonRows.find((row) => row.rowNumber === tableRow.rowNumber) ??
            null;
          if (!match) return tableRow;

          return {
            ...tableRow,
            values: {
              ...tableRow.values,
              AK: match.awayOfficers,
              AL: match.awaySergeants,
              AM: match.awaySoldiers,
              AN:
                match.awayOfficers +
                match.awaySergeants +
                match.awaySoldiers,
              AO: match.awayDestinationsText,
              AJ: match.available,
              AW: match.actualOfficers,
              AX: match.actualSergeants,
              AY: match.actualSoldiers,
              AZ: match.inRanksActually,
            },
          };
        }),
      }
    : undefined;

  const total =
    comparisonRows.find((row) => row.rowNumber === 11) ??
    comparisonRows[0] ??
    analytics.total;

  return {
    ...analytics,
    total,
    rows: comparisonRows,
    comparisonRows,
    table,
    detachedDestinations: totalDestinations,
  };
};

export const applyBchsAbsentFromPersonnel = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  const unitRows = analytics.comparisonRows.map((row) => {
    if (row.rowNumber <= 11 || isBchsTotalUnit(row.unit)) return row;

    const rankStats = computeBchsUnitRankListedAvailableStats(people, row.unit);
    const businessTrip = computeBchsUnitBusinessTripCount(people, row.unit);
    return applyBchsRankListedAvailableToComparisonRow(
      createBchsComparisonRow({
        ...row,
        businessTrip,
      }),
      rankStats,
    );
  });

  const detailRows = unitRows.filter(
    (row) => row.rowNumber > 11 && !isBchsTotalUnit(row.unit),
  );
  const totalRankStats = sumBchsComparisonRankListedAvailable(detailRows);
  const totalBusinessTrip = detailRows.reduce(
    (sum, row) => sum + row.businessTrip,
    0,
  );

  const comparisonRows = unitRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;

    return applyBchsRankListedAvailableToComparisonRow(
      createBchsComparisonRow({
        ...row,
        businessTrip: totalBusinessTrip,
      }),
      totalRankStats,
    );
  });

  const table = analytics.table
    ? {
        columns: analytics.table.columns,
        rows: analytics.table.rows.map((tableRow) => {
          const match =
            comparisonRows.find((row) => row.rowNumber === tableRow.rowNumber) ??
            null;
          if (!match) return tableRow;

          return {
            ...tableRow,
            values: {
              ...tableRow.values,
              G: match.listedOfficers,
              H: match.listedSergeants,
              I: match.listedSoldiers,
              J: match.listed,
              L: match.availableOfficers,
              M: match.availableSergeants,
              N: match.availableSoldiers,
              O: match.available,
              V: match.absentOfficers,
              W: match.absentSergeants,
              X: match.absentSoldiers,
              Y: match.absent,
              Z: match.businessTrip,
              AJ: match.available,
            },
          };
        }),
      }
    : undefined;

  const total =
    comparisonRows.find((row) => row.rowNumber === 11) ??
    comparisonRows[0] ??
    analytics.total;

  return {
    ...analytics,
    total,
    rows: comparisonRows,
    comparisonRows,
    table,
  };
};

const mergeBchsAbsenceCategoryStats = (
  row: BchsComparisonRow,
  stats: BchsUnitAbsenceCategoryStats,
): BchsComparisonRow =>
  createBchsComparisonRow({
    ...row,
    training: stats.training,
    hospitalWounded: stats.hospitalWounded,
    hospitalIllness: stats.hospitalIllness,
    vacation: stats.vacation,
    awol: stats.awol,
    missing: stats.missing,
    killed: stats.killed,
    medWounded: stats.medWounded,
    medIllness: stats.medIllness,
  });

const sumBchsAbsenceCategoryStats = (
  rows: BchsComparisonRow[],
): BchsUnitAbsenceCategoryStats =>
  rows.reduce(
    (total, row) => ({
      training: total.training + row.training,
      hospitalWounded: total.hospitalWounded + row.hospitalWounded,
      hospitalIllness: total.hospitalIllness + row.hospitalIllness,
      vacation: total.vacation + row.vacation,
      awol: total.awol + row.awol,
      missing: total.missing + row.missing,
      killed: total.killed + row.killed,
      medWounded: total.medWounded + row.medWounded,
      medIllness: total.medIllness + row.medIllness,
    }),
    emptyBchsUnitAbsenceCategoryStats(),
  );

/** Excel AA–AI з roster (як COUNTIFS на Аркуш2). */
export const applyBchsAbsenceCategoriesFromPersonnel = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  const unitRows = analytics.comparisonRows.map((row) => {
    if (row.rowNumber <= 11 || isBchsTotalUnit(row.unit)) return row;
    const stats = computeBchsUnitAbsenceCategoryStats(people, row.unit);
    return mergeBchsAbsenceCategoryStats(row, stats);
  });

  const detailRows = unitRows.filter(
    (row) => row.rowNumber > 11 && !isBchsTotalUnit(row.unit),
  );
  const totalStats = sumBchsAbsenceCategoryStats(detailRows);

  const comparisonRows = unitRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;
    return mergeBchsAbsenceCategoryStats(row, totalStats);
  });

  const table = analytics.table
    ? {
        columns: analytics.table.columns,
        rows: analytics.table.rows.map((tableRow) => {
          const match =
            comparisonRows.find((row) => row.rowNumber === tableRow.rowNumber) ??
            null;
          if (!match) return tableRow;

          return {
            ...tableRow,
            values: {
              ...tableRow.values,
              AA: match.training,
              AB: match.hospitalWounded,
              AC: match.hospitalIllness,
              AD: match.vacation,
              AE: match.awol,
              AF: match.missing,
              AG: match.killed,
              AH: match.medWounded,
              AI: match.medIllness,
            },
          };
        }),
      }
    : undefined;

  const total =
    comparisonRows.find((row) => row.rowNumber === 11) ??
    comparisonRows[0] ??
    analytics.total;

  return {
    ...analytics,
    total,
    rows: comparisonRows,
    comparisonRows,
    table,
  };
};

export const emptyBchsUnitAttachedStats = (): BchsUnitAttachedStats => ({
  officers: 0,
  sergeants: 0,
  soldiers: 0,
  total: 0,
  sources: new Map(),
  sourcesText: "",
});

export const isBchsAttachedPresentStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  // Excel AR24: U = *в строю* або *Новоприбулий* (окремі COUNTIFS, одна людина — один раз).
  return normalized.includes("в строю") || normalized.includes("новоприбулий");
};

/** Excel COUNTIFS wildcard: *текст* → includes після normalize. */
export const matchesBchsExcelWildcard = (value: string, pattern: string) => {
  const normalized = normalizeBchsText(value);
  const core = normalizeBchsText(pattern.replace(/^\*+|\*+$/g, ""));
  return Boolean(core) && normalized.includes(core);
};

/** Як Excel COUNTIFS по званню, але з нормалізацією і/и/i (капітан ≈ капитан). */
export const matchesBchsRankWildcard = (value: string, pattern: string) => {
  const normalized = normalizeBchsRankMatchText(value);
  const core = normalizeBchsRankMatchText(pattern.replace(/^\*+|\*+$/g, ""));
  return Boolean(core) && normalized.includes(core);
};

export const isBchsPrikomaniRoster = (rosterUnit: string) =>
  matchesBchsExcelWildcard(rosterUnit, "*ПРИКОМАНДИРОВАНІ*");

/** Excel AP24: M = *майор*|*капитан*|*лейтенант* (+ укр. «капітан»). */
export const isBchsAttachedOfficerRankTitle = (rankTitle: string) =>
  matchesBchsRankWildcard(rankTitle, "*майор*") ||
  matchesBchsRankWildcard(rankTitle, "*капитан*") ||
  matchesBchsRankWildcard(rankTitle, "*капітан*") ||
  matchesBchsRankWildcard(rankTitle, "*лейтенант*");

/** Excel AQ24: M = *сержант*|*старшина*. */
export const isBchsAttachedSergeantRankTitle = (rankTitle: string) =>
  matchesBchsRankWildcard(rankTitle, "*сержант*") ||
  matchesBchsRankWildcard(rankTitle, "*старшина*");

/** Excel AR24: M = *солдат*|*матрос*|*рядовий*. */
export const isBchsAttachedSoldierRankTitle = (rankTitle: string) =>
  matchesBchsRankWildcard(rankTitle, "*солдат*") ||
  matchesBchsRankWildcard(rankTitle, "*матрос*") ||
  matchesBchsRankWildcard(rankTitle, "*рядовий*");

export const isBchsBrezRoster = (rosterUnit: string) =>
  normalizeBchsText(rosterUnit) === "брез";

/** Excel AP–AR для РРЕБ: лише X (БЗВП/БРЕЗ) містить «БРЕЗ» — не колонка AA. */
export const isBchsBrezAttachedPerson = (person: BchsPersonnelAwayPerson) =>
  normalizeBchsText(person.bzvpStatus).includes("брез");

/**
 * AA «Відрядження (БРЕЗ)» часто містить «був БРЕЗ» (історія), це не поточна
 * позначка для COUNTIFS Аркуш2!X. Раніше помилково потрапляло в AP–AR.
 */
export const hasBchsFormerBrezAssignmentOnly = (
  person: BchsPersonnelAwayPerson,
) => {
  const assignment = normalizeBchsText(person.brezAssignment ?? "");
  if (!assignment.includes("брез")) return false;
  return !normalizeBchsText(person.bzvpStatus).includes("брез");
};

export const isBchsEngineerUnit = (unitName: string) =>
  /інженер|сапер/i.test(unitName);

export const isBchsRebUnit = (unitName: string) => {
  const unit = normalizeBchsText(unitName);
  return unit.includes("радіоелектрон") || unit.includes("відділення реб");
};

export const matchesBchsAttachedUnit = (
  arkush1Unit: string,
  person: BchsPersonnelAwayPerson,
) => {
  if (isBchsEngineerUnit(arkush1Unit)) {
    return isBchsPrikomaniRoster(person.rosterUnit);
  }
  if (isBchsRebUnit(arkush1Unit)) {
    return isBchsBrezAttachedPerson(person);
  }
  return false;
};

/** Excel AP/AQ/AR: звання лише з M (патерни COUNTIFS), не з I. */
export const classifyBchsAttachedRankByExcel = (
  rankTitle: string,
): "officers" | "sergeants" | "soldiers" | null => {
  if (!rankTitle.trim()) return null;
  if (isBchsAttachedOfficerRankTitle(rankTitle)) return "officers";
  if (isBchsAttachedSergeantRankTitle(rankTitle)) return "sergeants";
  if (isBchsAttachedSoldierRankTitle(rankTitle)) return "soldiers";
  return null;
};

export const classifyBchsAttachedRank = (
  _rankCategory: string,
  rankTitle: string,
): "officers" | "sergeants" | "soldiers" | null =>
  classifyBchsAttachedRankByExcel(rankTitle);

export const normalizeBchsAttachedSourceLabel = (
  battalion: string,
  rosterUnit: string,
  bzvpStatus = "",
  brezAssignment = "",
) => {
  if (
    isBchsBrezRoster(rosterUnit) ||
    normalizeBchsText(`${bzvpStatus} ${brezAssignment}`).includes("брез")
  )
    return "БРЕЗ";

  const raw = battalion.replace(/\s+/g, " ").trim();
  const normalized = normalizeBchsText(raw);
  if (!normalized) return "інші";

  if (normalized.includes("шквал")) return "ШКВАЛ";
  if (normalized === "210") return "210";
  if (normalized === "155" || /^155\s*омбр/.test(normalized)) return "155 ОМБр";
  if (normalized.includes("41") && normalized.includes("омбр")) return "41ОМБр";
  if (normalized.includes("омбр")) return raw;

  return "інші";
};

export const computeBchsUnitAttachedStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitAttachedStats => {
  if (isBchsTotalUnit(unitName)) return emptyBchsUnitAttachedStats();

  const stats = emptyBchsUnitAttachedStats();

  // Excel AP–AR: B=ПРИКОМАНДИРОВАНІ / X=БРЕЗ — без фільтра A=«нова».
  people.forEach((person) => {
    if (!isBchsAttachedPresentStatus(person.status)) return;
    if (!matchesBchsAttachedUnit(unitName, person)) return;

    const rank = classifyBchsAttachedRank(
      person.rankCategory,
      person.rankTitle,
    );
    if (rank === "officers") stats.officers += 1;
    else if (rank === "sergeants") stats.sergeants += 1;
    else if (rank === "soldiers") stats.soldiers += 1;
    else return;

    const source = normalizeBchsAttachedSourceLabel(
      person.battalion,
      person.rosterUnit,
      person.bzvpStatus,
      person.brezAssignment,
    );
    if (source) {
      stats.sources.set(source, (stats.sources.get(source) ?? 0) + 1);
    }
  });

  stats.total = stats.officers + stats.sergeants + stats.soldiers;
  stats.sourcesText = formatBchsDestinationText(stats.sources);
  return stats;
};

/** Excel AU: посада порожня або примітка про прибуття (не штатна посада). */
export const isBchsWithoutStaffPosition = (person: BchsPersonnelAwayPerson) => {
  const position = normalizeBchsText(person.position);
  if (!position) return true;
  return (
    position.includes("прибув") ||
    position.includes("чл") ||
    position.includes("без посади")
  );
};

/** Excel AU: A=«нова», U=*новоприбулий*, без штатної посади; не блоки ПРИКОМАНД./БРЕЗ. */
export const isBchsUnassignedNewcomer = (person: BchsPersonnelAwayPerson) => {
  const status = normalizeBchsText(person.status);
  if (!status.includes("новоприбулий")) return false;
  if (normalizeBchsText(person.battalion) !== "нова") return false;
  if (
    isBchsPrikomaniRoster(person.rosterUnit) ||
    isBchsBrezRoster(person.rosterUnit)
  ) {
    return false;
  }
  return isBchsWithoutStaffPosition(person);
};

export const computeBchsTotalUnassignedNewcomers = (
  people: BchsPersonnelAwayPerson[],
): number =>
  filterBchsNovaPeople(people).filter(isBchsUnassignedNewcomer).length;

/** Excel AU per unit: COUNTIFS(A=«нова», U=*новоприбулий*, B, без посади). */
export const computeBchsUnitUnassignedNewcomers = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): number => {
  if (isBchsTotalUnit(unitName)) return 0;

  let count = 0;
  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsUnassignedNewcomer(person)) return;
    if (!matchesBchsRosterUnit(unitName, person.rosterUnit)) return;
    count += 1;
  });
  return count;
};

/** Excel AV: COUNTIFS(U;"*пошук*", B, M<>), A=«нова». */
export const computeBchsUnitSearchInProgressCount = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): number => {
  if (isBchsTotalUnit(unitName)) return 0;

  let count = 0;
  filterBchsNovaPeople(people).forEach((person) => {
    if (!normalizeBchsText(person.status).includes("пошук")) return;
    if (!matchesBchsRosterUnit(unitName, person.rosterUnit)) return;
    if (!isBchsListedWithRankTitle(person)) return;
    count += 1;
  });
  return count;
};

const emptyBchsUnitCombatComponentStats = (): BchsUnitCombatComponentStats => ({
  assaultReady: 0,
  assaultRecovery: 0,
  assaultExecution: 0,
  noBzvp: 0,
  assaultTotal: 0,
  vehicleCrew: 0,
  droneCrew: 0,
  crewServedWeapons: 0,
  commandCombat: 0,
  supportCombat: 0,
  combatComponent: 0,
});

const isBchsInfantryRoleType = (roleType: string) =>
  normalizeBchsText(roleType).includes("піхот");

const isBchsDriverRoleType = (roleType: string) =>
  normalizeBchsText(roleType).includes("водій");

const isBchsPilotRoleType = (roleType: string) => {
  const role = normalizeBchsText(roleType);
  return role.includes("пілот") || role.includes("бпла");
};

const isBchsGrenadeRoleType = (roleType: string) => {
  const role = normalizeBchsText(roleType);
  return (
    role.includes("гранатомет") ||
    role.includes("міномет") ||
    role.includes("кулемет")
  );
};

const isBchsCommandRoleType = (roleType: string) => {
  const role = normalizeBchsText(roleType);
  return role.includes("упр") || role.includes("штаб");
};

const isBchsCombatAttachedRoster = (person: BchsPersonnelAwayPerson) =>
  isBchsPrikomaniRoster(person.rosterUnit) || isBchsBrezRoster(person.rosterUnit);

/** Excel U: «в строю»; для ПРИКОМАНД./БРЕЗ також «Новоприбулий». */
const isBchsCombatComponentStatus = (
  person: BchsPersonnelAwayPerson,
  allowNewcomer: boolean,
) => {
  const status = normalizeBchsText(person.status);
  if (status.includes("в строю")) return true;
  return allowNewcomer && status.includes("новоприбулий");
};

const matchesBchsCombatComponentRoster = (
  unitName: string,
  person: BchsPersonnelAwayPerson,
) => {
  if (isBchsEngineerUnit(unitName)) {
    return (
      matchesBchsRosterUnit(unitName, person.rosterUnit) ||
      isBchsPrikomaniRoster(person.rosterUnit)
    );
  }
  if (isBchsRebUnit(unitName)) {
    return (
      matchesBchsRosterUnit(unitName, person.rosterUnit) ||
      isBchsBrezRoster(person.rosterUnit)
    );
  }
  return matchesBchsRosterUnit(unitName, person.rosterUnit);
};

const isInBchsCombatComponentPool = (
  unitName: string,
  person: BchsPersonnelAwayPerson,
) => {
  if (!matchesBchsCombatComponentRoster(unitName, person)) return false;

  const attached = isBchsCombatAttachedRoster(person);
  if (!attached && normalizeBchsText(person.battalion) !== "нова") return false;

  return isBchsCombatComponentStatus(person, attached);
};

const isBchsBattleReadyCombatStatus = (combatReadiness: string) =>
  normalizeBchsText(combatReadiness) === "бг";

const isBchsOnCombatExecution = (person: BchsPersonnelAwayPerson) =>
  normalizeBchsText(person.medicalPlace).includes("на виконанні");

const isBchsNoBzvpStatus = (bzvpStatus: string) =>
  normalizeBchsText(bzvpStatus).includes("бзвп");

/** Excel BB–BL: COUNTIFS по V/W/X/AE + ПРИКОМАНД./БРЕЗ для ІСВ/РРЕБ. */
export const computeBchsUnitCombatComponentStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitCombatComponentStats => {
  if (isBchsTotalUnit(unitName)) return emptyBchsUnitCombatComponentStats();

  const pool = people.filter((person) =>
    isInBchsCombatComponentPool(unitName, person),
  );

  if (isBchsCommandUnit(unitName)) {
    const commandCombat = pool.filter((person) =>
      isBchsCommandRoleType(person.roleType),
    ).length;
    const supportCombat = Math.max(0, pool.length - commandCombat);
    return {
      assaultReady: 0,
      assaultRecovery: 0,
      assaultExecution: 0,
      noBzvp: 0,
      assaultTotal: 0,
      vehicleCrew: 0,
      droneCrew: 0,
      crewServedWeapons: 0,
      commandCombat,
      supportCombat,
      combatComponent: pool.length,
    };
  }

  const infantry = pool.filter((person) =>
    isBchsInfantryRoleType(person.roleType),
  );
  const assaultReady = infantry.filter((person) =>
    isBchsBattleReadyCombatStatus(person.combatReadiness),
  ).length;
  const noBzvp = infantry.filter((person) =>
    isBchsNoBzvpStatus(person.bzvpStatus),
  ).length;
  const assaultExecutionInf = infantry.filter((person) =>
    isBchsOnCombatExecution(person),
  ).length;
  const assaultExecutionGr = pool.filter(
    (person) =>
      isBchsGrenadeRoleType(person.roleType) &&
      isBchsOnCombatExecution(person),
  ).length;
  const assaultExecution = assaultExecutionInf + assaultExecutionGr;
  const assaultRecovery = Math.max(
    0,
    infantry.length - assaultExecutionInf - assaultReady - noBzvp,
  );
  const vehicleCrew = pool.filter((person) =>
    isBchsDriverRoleType(person.roleType),
  ).length;
  const droneCrew = pool.filter((person) =>
    isBchsPilotRoleType(person.roleType),
  ).length;
  const grenadePool = pool.filter((person) =>
    isBchsGrenadeRoleType(person.roleType),
  );
  const crewServedWeapons = Math.max(
    0,
    grenadePool.length -
      grenadePool.filter((person) => isBchsOnCombatExecution(person)).length -
      grenadePool.filter((person) => isBchsNoBzvpStatus(person.bzvpStatus))
        .length,
  );
  const commandCombat = pool.filter((person) =>
    isBchsCommandRoleType(person.roleType),
  ).length;
  const assaultTotal =
    assaultReady + assaultRecovery + assaultExecution + noBzvp;
  const supportCombat = Math.max(
    0,
    pool.length -
      assaultTotal -
      vehicleCrew -
      droneCrew -
      crewServedWeapons -
      commandCombat,
  );

  return {
    assaultReady,
    assaultRecovery,
    assaultExecution,
    noBzvp,
    assaultTotal,
    vehicleCrew,
    droneCrew,
    crewServedWeapons,
    commandCombat,
    supportCombat,
    combatComponent: pool.length,
  };
};

export const sumBchsComparisonCombatComponent = (
  rows: BchsComparisonRow[],
): BchsUnitCombatComponentStats => ({
  assaultReady: rows.reduce((sum, row) => sum + row.assaultReady, 0),
  assaultRecovery: rows.reduce((sum, row) => sum + row.assaultRecovery, 0),
  assaultExecution: rows.reduce((sum, row) => sum + row.assaultExecution, 0),
  noBzvp: rows.reduce((sum, row) => sum + row.noBzvp, 0),
  assaultTotal: rows.reduce((sum, row) => sum + row.assaultTotal, 0),
  vehicleCrew: rows.reduce((sum, row) => sum + row.vehicleCrew, 0),
  droneCrew: rows.reduce((sum, row) => sum + row.droneCrew, 0),
  crewServedWeapons: rows.reduce((sum, row) => sum + row.crewServedWeapons, 0),
  commandCombat: rows.reduce((sum, row) => sum + row.commandCombat, 0),
  supportCombat: rows.reduce((sum, row) => sum + row.supportCombat, 0),
  combatComponent: rows.reduce((sum, row) => sum + row.combatComponent, 0),
});

export const applyBchsCombatComponentFromPersonnel = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  const unitRows = analytics.comparisonRows.map((row) => {
    if (row.rowNumber <= 11 || isBchsTotalUnit(row.unit)) return row;

    const combat = computeBchsUnitCombatComponentStats(people, row.unit);
    return createBchsComparisonRow({
      ...row,
      ...combat,
    });
  });

  const detailRows = unitRows.filter(
    (row) => row.rowNumber > 11 && !isBchsTotalUnit(row.unit),
  );
  const totalCombat = sumBchsComparisonCombatComponent(detailRows);

  const comparisonRows = unitRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;
    return createBchsComparisonRow({
      ...row,
      ...totalCombat,
    });
  });

  return {
    ...analytics,
    comparisonRows,
    rows: comparisonRows,
    total: comparisonRows[0] ?? analytics.total,
  };
};

export const applyBchsNewcomersAndSearchFromPersonnel = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  const novaPeople = filterBchsNovaPeople(people);
  const unitRows = analytics.comparisonRows.map((row) => {
    if (row.rowNumber <= 11 || isBchsTotalUnit(row.unit)) return row;

    return createBchsComparisonRow({
      ...row,
      unassignedNewcomers: computeBchsUnitUnassignedNewcomers(people, row.unit),
      searchInProgress: computeBchsUnitSearchInProgressCount(
        novaPeople,
        row.unit,
      ),
    });
  });

  const detailRows = unitRows.filter(
    (row) => row.rowNumber > 11 && !isBchsTotalUnit(row.unit),
  );

  const comparisonRows = unitRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;

    return createBchsComparisonRow({
      ...row,
      unassignedNewcomers: detailRows.reduce(
        (sum, item) => sum + item.unassignedNewcomers,
        0,
      ),
      searchInProgress: detailRows.reduce(
        (sum, item) => sum + item.searchInProgress,
        0,
      ),
    });
  });

  return {
    ...analytics,
    comparisonRows,
    rows: comparisonRows,
    total: comparisonRows[0] ?? analytics.total,
  };
};

export const applyBchsAttachedFromPersonnel = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  const unitRows = analytics.comparisonRows.map((row) => {
    if (row.rowNumber <= 11 || isBchsTotalUnit(row.unit)) return row;

    const stats = computeBchsUnitAttachedStats(people, row.unit);
    return createBchsComparisonRow({
      ...row,
      attachedOfficers: stats.officers,
      attachedSergeants: stats.sergeants,
      attachedSoldiers: stats.soldiers,
      attachedFromOtherUnits: stats.total,
      attachedSourcesText: stats.sourcesText,
      // Перерахувати BA від нових AP–AR, не лишати застарілі 1100% з Excel.
      actualPercent: 0,
    });
  });

  const detailRows = unitRows.filter(
    (row) => row.rowNumber > 11 && !isBchsTotalUnit(row.unit),
  );
  const totalAttachedOfficers = detailRows.reduce(
    (sum, row) => sum + row.attachedOfficers,
    0,
  );
  const totalAttachedSergeants = detailRows.reduce(
    (sum, row) => sum + row.attachedSergeants,
    0,
  );
  const totalAttachedSoldiers = detailRows.reduce(
    (sum, row) => sum + row.attachedSoldiers,
    0,
  );
  const totalAttached =
    totalAttachedOfficers + totalAttachedSergeants + totalAttachedSoldiers;
  const totalSources = mergeBchsMetricMaps(
    detailRows.map((row) => parseBchsDestinationText(row.attachedSourcesText)),
  );

  const comparisonRows = unitRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;

    return createBchsComparisonRow({
      ...row,
      attachedOfficers: totalAttachedOfficers,
      attachedSergeants: totalAttachedSergeants,
      attachedSoldiers: totalAttachedSoldiers,
      attachedFromOtherUnits: totalAttached,
      attachedSourcesText: "",
    });
  });

  const table = analytics.table
    ? {
        columns: analytics.table.columns,
        rows: analytics.table.rows.map((tableRow) => {
          const match =
            comparisonRows.find((row) => row.rowNumber === tableRow.rowNumber) ??
            null;
          if (!match) return tableRow;

          return {
            ...tableRow,
            values: {
              ...tableRow.values,
              AP: match.attachedOfficers,
              AQ: match.attachedSergeants,
              AR: match.attachedSoldiers,
              AS: match.attachedFromOtherUnits,
              AT: match.attachedSourcesText,
              AJ: match.available,
              AW: match.actualOfficers,
              AX: match.actualSergeants,
              AY: match.actualSoldiers,
              AZ: match.inRanksActually,
            },
          };
        }),
      }
    : undefined;

  const total =
    comparisonRows.find((row) => row.rowNumber === 11) ??
    comparisonRows[0] ??
    analytics.total;

  return {
    ...analytics,
    total,
    rows: comparisonRows,
    comparisonRows,
    table,
    attachedSources: totalSources,
  };
};

export const applyBchsPersonnelDerivedColumns = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
  options: { force?: boolean } = {},
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;

  const novaPeople = filterBchsNovaPeople(people);
  const canApplyNovaDerived = novaPeople.length > 0;

  if (canApplyNovaDerived && !options.force) {
    const novaDetachedCount = novaPeople.filter((person) =>
      isBchsDetachedStatus(person.status),
    ).length;
    const storedAway = bchsToNumber(analytics.total.awayInOtherUnits);
    if (novaDetachedCount === 0 && storedAway > 0) {
      // Status column not mapped yet — do not replace «в інших» with 0.
      const attachedOnly = applyBchsAttachedFromPersonnel(analytics, people);
      return {
        ...attachedOnly,
        dataIssues: mergeBchsDataIssues(
          buildBchsBrezMiscountIssues(people),
          buildBchsComparisonAnomalyIssues(attachedOnly.comparisonRows),
          attachedOnly.dataIssues,
          analytics.dataIssues,
        ),
      };
    }
  }

  let derived = analytics;
  if (canApplyNovaDerived) {
    derived = applyBchsAwayFromPersonnel(derived, novaPeople);
    derived = applyBchsAbsentFromPersonnel(derived, novaPeople);
    derived = applyBchsAbsenceCategoriesFromPersonnel(derived, novaPeople);
  }
  derived = applyBchsAttachedFromPersonnel(derived, people);
  derived = applyBchsNewcomersAndSearchFromPersonnel(derived, people);
  derived = applyBchsCombatComponentFromPersonnel(derived, people);
  const peopleForIssues = canApplyNovaDerived
    ? people
    : novaPeople.length > 0
      ? novaPeople
      : people;
  const dataIssues = mergeBchsDataIssues(
    buildBchsGeneralListDataIssues(peopleForIssues),
    buildBchsBrezMiscountIssues(people),
    buildBchsComparisonAnomalyIssues(derived.comparisonRows),
    derived.dataIssues,
  );

  return {
    ...derived,
    dataIssues,
  };
};

/** Ensure AK–AO / AP–AT are filled from roster people before Excel write. */
export const enrichBchsAnalyticsForExport = (
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
): BchsAnalyticsSnapshot => {
  if (people.length === 0) return analytics;
  return applyBchsPersonnelDerivedColumns(analytics, people, { force: true });
};

export const getBchsDestinationCellValue = (
  anValue: unknown,
  aoValue: unknown = "",
) => {
  const aoText = valueToDisplay(aoValue as CellValue).trim();
  if (aoText && parseBchsDestinationText(aoText).size > 0) return aoText;

  const anText = valueToDisplay(anValue as CellValue).trim();
  if (anText && parseBchsDestinationText(anText).size > 0) return anText;

  return aoText || anText;
};

/** AN = Всього for away-in-other-units block. */
export const resolveBchsAwayInOtherUnits = (anValue: unknown) => bchsToNumber(anValue);

export const mergeBchsMetricMaps = (maps: Array<Map<string, number>>) => {
  const merged = new Map<string, number>();

  maps.forEach((map) => {
    map.forEach((value, label) => {
      merged.set(label, (merged.get(label) ?? 0) + value);
    });
  });

  return Array.from(merged.entries())
    .map(([label, value]) => ({ label, value }))
    .sort(
      (left, right) =>
        right.value - left.value || left.label.localeCompare(right.label, "uk"),
    );
};

export const buildBchsDetachedDestinationsFromMaps = (
  maps: Array<Map<string, number>>,
) => {
  if (maps.length === 0) return [];

  const [totalMap, ...unitMaps] = maps;
  if (totalMap && sumBchsDestinationCounts(totalMap) > 0) {
    return mergeBchsMetricMaps([totalMap]);
  }

  return mergeBchsMetricMaps(unitMaps);
};

export const buildBchsDetachedDestinationsFromSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) =>
  buildBchsDetachedDestinationsFromMaps(
    Array.from({ length: 18 }, (_, index) => {
      const rowNumber = 11 + index;
      return parseBchsDestinationText(
        getBchsDestinationCellValue(
          getSheetValue(sheet, rowNumber, "AN"),
          getSheetValue(sheet, rowNumber, "AO"),
        ),
      );
    }),
  );

export const buildBchsDetachedDestinationsFromTable = (
  table: BchsAnalyticsSnapshot["table"] | undefined,
) => {
  if (!table) return [];

  return buildBchsDetachedDestinationsFromMaps(
    table.rows.map((row) =>
      parseBchsDestinationText(
        getBchsDestinationCellValue(row.values.AN, row.values.AO),
      ),
    ),
  );
};

export const getBchsReportDate = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) => {
  const title = valueToDisplay(getSheetValue(sheet, 1, "B"));
  const match = title.match(/\b\d{1,2}[./]\d{1,2}[./]\d{2,4}\b/);

  return match?.[0] ?? "";
};

export const buildBchsAbsenceReasons = (row: BchsAnalyticsRow): AnalyticsMetric[] => {
  const hospital = row.hospitalWounded + row.hospitalIllness;
  const medrota = row.medWounded + row.medIllness;
  const other = row.businessTrip + row.training + row.vacation;

  return [
    { label: "Шпиталь", value: hospital },
    { label: "СЗЧ", value: row.awol },
    { label: "Зниклі безвісти", value: row.missing },
    { label: "Медрота", value: medrota },
    { label: "Загиблі", value: row.killed },
    { label: "Інше", value: other },
  ].filter((item) => item.value > 0);
};

export const emptyBchsSupplementRow = (
  overrides: Partial<BchsSupplementRow> = {},
): BchsSupplementRow => ({
  rowNumber: 0,
  battalion: "",
  unit: "Усього",
  staff: 0,
  listed: 0,
  available: 0,
  staffedPercent: 0,
  combatTask: 0,
  replacementReserve: 0,
  taskReserve: 0,
  commanderReserve: 0,
  absent: 0,
  businessTrip: 0,
  training: 0,
  hospitalWounded: 0,
  hospitalIllness: 0,
  vacation: 0,
  awol: 0,
  missing: 0,
  killed: 0,
  medWounded: 0,
  medIllness: 0,
  detached: 0,
  attached: 0,
  newcomers: 0,
  inRanks: 0,
  assaultReady: 0,
  assaultRecovery: 0,
  assaultExecution: 0,
  noBzvp: 0,
  assaultTotal: 0,
  vehicleCrew: 0,
  droneCrew: 0,
  crewServedWeapons: 0,
  commandCombat: 0,
  supportCombat: 0,
  bzvpBuckets: [],
  totalBzvp: 0,
  ...overrides,
});

export const getBchsCellText = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
  rowNumber: number,
  columnLetter: string,
) =>
  valueToDisplay(getSheetValue(sheet, rowNumber, columnLetter))
    .replace(/\s+/g, " ")
    .trim();

export const sumBchsSupplementRows = (
  rows: BchsSupplementRow[],
  label = "Усього",
): BchsSupplementRow => {
  const bucketMap = new Map<string, number>();

  rows.forEach((row) => {
    row.bzvpBuckets.forEach((bucket) => {
      bucketMap.set(bucket.label, (bucketMap.get(bucket.label) ?? 0) + bucket.value);
    });
  });

  const total = rows.reduce(
    (acc, row) => ({
      staff: acc.staff + row.staff,
      listed: acc.listed + row.listed,
      available: acc.available + row.available,
      combatTask: acc.combatTask + row.combatTask,
      replacementReserve: acc.replacementReserve + row.replacementReserve,
      taskReserve: acc.taskReserve + row.taskReserve,
      commanderReserve: acc.commanderReserve + row.commanderReserve,
      absent: acc.absent + row.absent,
      businessTrip: acc.businessTrip + row.businessTrip,
      training: acc.training + row.training,
      hospitalWounded: acc.hospitalWounded + row.hospitalWounded,
      hospitalIllness: acc.hospitalIllness + row.hospitalIllness,
      vacation: acc.vacation + row.vacation,
      awol: acc.awol + row.awol,
      missing: acc.missing + row.missing,
      killed: acc.killed + row.killed,
      medWounded: acc.medWounded + row.medWounded,
      medIllness: acc.medIllness + row.medIllness,
      detached: acc.detached + row.detached,
      attached: acc.attached + row.attached,
      newcomers: acc.newcomers + row.newcomers,
      inRanks: acc.inRanks + row.inRanks,
      assaultReady: acc.assaultReady + row.assaultReady,
      assaultRecovery: acc.assaultRecovery + row.assaultRecovery,
      assaultExecution: acc.assaultExecution + row.assaultExecution,
      noBzvp: acc.noBzvp + row.noBzvp,
      assaultTotal: acc.assaultTotal + row.assaultTotal,
      vehicleCrew: acc.vehicleCrew + row.vehicleCrew,
      droneCrew: acc.droneCrew + row.droneCrew,
      crewServedWeapons: acc.crewServedWeapons + row.crewServedWeapons,
      commandCombat: acc.commandCombat + row.commandCombat,
      supportCombat: acc.supportCombat + row.supportCombat,
    }),
    {
      staff: 0,
      listed: 0,
      available: 0,
      combatTask: 0,
      replacementReserve: 0,
      taskReserve: 0,
      commanderReserve: 0,
      absent: 0,
      businessTrip: 0,
      training: 0,
      hospitalWounded: 0,
      hospitalIllness: 0,
      vacation: 0,
      awol: 0,
      missing: 0,
      killed: 0,
      medWounded: 0,
      medIllness: 0,
      detached: 0,
      attached: 0,
      newcomers: 0,
      inRanks: 0,
      assaultReady: 0,
      assaultRecovery: 0,
      assaultExecution: 0,
      noBzvp: 0,
      assaultTotal: 0,
      vehicleCrew: 0,
      droneCrew: 0,
      crewServedWeapons: 0,
      commandCombat: 0,
      supportCombat: 0,
    },
  );
  const buckets = Array.from(bucketMap.entries()).map(([bucketLabel, value]) => ({
    label: bucketLabel,
    value,
  }));

  return emptyBchsSupplementRow({
    unit: label,
    ...total,
    staffedPercent: total.staff > 0 ? total.listed / total.staff : 0,
    bzvpBuckets: buckets,
    totalBzvp: buckets.reduce((sum, item) => sum + item.value, 0),
  });
};

export const buildBchsSupplementComparisonRow = (
  supplement: BchsSupplementRow,
): BchsComparisonRow =>
  createBchsComparisonRow({
    rowNumber: supplement.rowNumber,
    unit: supplement.unit,
    staff: supplement.staff,
    listed: supplement.listed,
    staffedPercent: supplement.staffedPercent,
    available: supplement.available,
    shortage: supplement.staff - supplement.listed,
    absent: supplement.absent || Math.max(0, supplement.listed - supplement.available),
    businessTrip: supplement.businessTrip,
    training: supplement.training,
    hospitalWounded: supplement.hospitalWounded,
    hospitalIllness: supplement.hospitalIllness,
    vacation: supplement.vacation,
    awol: supplement.awol,
    missing: supplement.missing,
    killed: supplement.killed,
    medWounded: supplement.medWounded,
    medIllness: supplement.medIllness,
    inRanksActually: supplement.inRanks || supplement.available,
    actualPercent: supplement.staff > 0 ? (supplement.inRanks || supplement.available) / supplement.staff : 0,
    combatComponent:
      supplement.assaultTotal +
      supplement.vehicleCrew +
      supplement.droneCrew +
      supplement.crewServedWeapons +
      supplement.commandCombat +
      supplement.supportCombat,
    awayInOtherUnits: supplement.detached,
    attachedFromOtherUnits: supplement.attached,
    unassignedNewcomers: supplement.newcomers,
    noBzvp: supplement.noBzvp || supplement.totalBzvp,
    assaultReady: supplement.assaultReady,
    assaultRecovery: supplement.assaultRecovery,
    assaultExecution: supplement.assaultExecution,
    assaultTotal: supplement.assaultTotal,
    vehicleCrew: supplement.vehicleCrew,
    droneCrew: supplement.droneCrew,
    crewServedWeapons: supplement.crewServedWeapons,
    commandCombat: supplement.commandCombat,
    supportCombat: supplement.supportCombat,
  });

export const isBchsPersonnelBzvpSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) =>
  /підрозділ/i.test(getBchsCellText(sheet, 1, "B")) &&
  /штатна кількість/i.test(getBchsCellText(sheet, 1, "C"));

export const isBchsPersonnelGeneralListSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) => {
  if (!sheet) return false;

  const sheetName = normalizeBchsText(sheet.sheetName);
  const nameMatches =
    /загальн.*спис/i.test(sheetName) || /ос.*загальн.*спис/i.test(sheetName);
  const headerMatches =
    /підрозділ/i.test(getBchsCellText(sheet, 1, "B")) &&
    /піб/i.test(getBchsCellText(sheet, 1, "N")) &&
    /статус/i.test(getBchsCellText(sheet, 1, "U")) &&
    /якому.*підрозділі/i.test(getBchsCellText(sheet, 1, "AC"));

  return nameMatches || headerMatches;
};

export const isBchsAppendixSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) =>
  /батальйон.*підрозділ/i.test(getBchsCellText(sheet, 4, "B")) &&
  /за штатом/i.test(getBchsCellText(sheet, 4, "C")) &&
  /в строю/i.test(getBchsCellText(sheet, 4, "S"));

export const buildBchsPersonnelBzvpSupplement = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
): BchsSupplementSnapshot | null => {
  if (!isBchsPersonnelBzvpSheet(sheet)) return null;

  const bucketLetters = sheet
    ? sheet.columnIndexes
        .filter((index) => index >= columnLetterToIndex("K"))
        .map(getColumnLabel)
        .filter((letter) => getBchsCellText(sheet, 3, letter))
    : [];
  const rows: BchsSupplementRow[] = [];
  const totals: BchsSupplementRow[] = [];
  let currentBattalion = "";

  sheet?.rawRows.forEach((_, index) => {
    const rowNumber = index + 1;
    if (rowNumber <= 3) return;

    const groupLabel = getBchsCellText(sheet, rowNumber, "A");
    const unitLabel = getBchsCellText(sheet, rowNumber, "B");
    const isTotal = /всього/i.test(groupLabel);
    if (groupLabel && !isTotal) currentBattalion = groupLabel;

    if (!unitLabel && !isTotal) return;

    const staff = bchsToNumber(getSheetValue(sheet, rowNumber, "C"));
    const listed = bchsToNumber(getSheetValue(sheet, rowNumber, "D"));
    const available = bchsToNumber(getSheetValue(sheet, rowNumber, "E"));
    if (!staff && !listed && !available && !isTotal) return;

    const bzvpBuckets = bucketLetters
      .map((letter) => ({
        label: getBchsCellText(sheet, 3, letter),
        value: bchsToNumber(getSheetValue(sheet, rowNumber, letter)),
      }))
      .filter((bucket) => bucket.label);
    const row = emptyBchsSupplementRow({
      rowNumber,
      battalion: currentBattalion,
      unit: isTotal ? `${currentBattalion || "Блок"} · всього` : unitLabel,
      staff,
      listed,
      available,
      staffedPercent: normalizeBchsPercentValue(
        getSheetValue(sheet, rowNumber, "F"),
        listed,
        staff,
      ),
      combatTask: bchsToNumber(getSheetValue(sheet, rowNumber, "G")),
      replacementReserve: bchsToNumber(getSheetValue(sheet, rowNumber, "H")),
      taskReserve: bchsToNumber(getSheetValue(sheet, rowNumber, "I")),
      commanderReserve: bchsToNumber(getSheetValue(sheet, rowNumber, "J")),
      absent: available > 0 ? Math.max(0, listed - available) : 0,
      inRanks: available,
      bzvpBuckets,
      totalBzvp: bzvpBuckets.reduce((sum, item) => sum + item.value, 0),
    });

    rows.push(row);
    if (isTotal) totals.push(row);
  });

  if (rows.length === 0) return null;

  const total =
    totals.find((row) =>
      row.battalion.replace(/\s+/g, "").toLowerCase().includes("1пб"),
    ) ??
    totals.find((row) => row.available > 0) ??
    sumBchsSupplementRows(totals.length > 0 ? totals : rows);

  return {
    kind: "personnel-bzvp",
    title: "Особовий склад підрозділів + БЗВП",
    reportDate: "",
    total,
    rows,
    totals,
    absenceReasons: [],
    combatCategories: [],
    reserveMetrics: [
      { label: "На виконанні БЗ", value: total.combatTask },
      { label: "Резерв заміни", value: total.replacementReserve },
      { label: "Резерв виконання БЗ", value: total.taskReserve },
      { label: "Резерв командира", value: total.commanderReserve },
    ],
    bzvpBuckets: total.bzvpBuckets,
  };
};

export const buildBchsAppendixSupplement = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
): BchsSupplementSnapshot | null => {
  if (!isBchsAppendixSheet(sheet)) return null;

  const rows = (sheet?.rawRows ?? [])
    .map((_, index) => {
      const rowNumber = index + 1;
      const unit = getBchsCellText(sheet, rowNumber, "B");
      if (rowNumber < 6 || !unit) return null;

      return emptyBchsSupplementRow({
        rowNumber,
        unit,
        staff: bchsToNumber(getSheetValue(sheet, rowNumber, "C")),
        listed: bchsToNumber(getSheetValue(sheet, rowNumber, "D")),
        absent: bchsToNumber(getSheetValue(sheet, rowNumber, "E")),
        businessTrip: bchsToNumber(getSheetValue(sheet, rowNumber, "F")),
        training: bchsToNumber(getSheetValue(sheet, rowNumber, "G")),
        hospitalWounded: bchsToNumber(getSheetValue(sheet, rowNumber, "H")),
        hospitalIllness: bchsToNumber(getSheetValue(sheet, rowNumber, "I")),
        vacation: bchsToNumber(getSheetValue(sheet, rowNumber, "J")),
        awol: bchsToNumber(getSheetValue(sheet, rowNumber, "K")),
        missing: bchsToNumber(getSheetValue(sheet, rowNumber, "L")),
        killed: bchsToNumber(getSheetValue(sheet, rowNumber, "M")),
        medWounded: bchsToNumber(getSheetValue(sheet, rowNumber, "N")),
        medIllness: bchsToNumber(getSheetValue(sheet, rowNumber, "O")),
        combatTask: bchsToNumber(getSheetValue(sheet, rowNumber, "S")),
        detached: bchsToNumber(getSheetValue(sheet, rowNumber, "P")),
        attached: bchsToNumber(getSheetValue(sheet, rowNumber, "Q")),
        newcomers: bchsToNumber(getSheetValue(sheet, rowNumber, "R")),
        inRanks: bchsToNumber(getSheetValue(sheet, rowNumber, "S")),
        assaultReady: bchsToNumber(getSheetValue(sheet, rowNumber, "T")),
        assaultRecovery: bchsToNumber(getSheetValue(sheet, rowNumber, "U")),
        assaultExecution: bchsToNumber(getSheetValue(sheet, rowNumber, "V")),
        noBzvp: bchsToNumber(getSheetValue(sheet, rowNumber, "W")),
        assaultTotal: bchsToNumber(getSheetValue(sheet, rowNumber, "X")),
        vehicleCrew: bchsToNumber(getSheetValue(sheet, rowNumber, "Y")),
        droneCrew: bchsToNumber(getSheetValue(sheet, rowNumber, "Z")),
        crewServedWeapons: bchsToNumber(getSheetValue(sheet, rowNumber, "AA")),
        commandCombat: bchsToNumber(getSheetValue(sheet, rowNumber, "AB")),
        supportCombat: bchsToNumber(getSheetValue(sheet, rowNumber, "AC")),
      });
    })
    .filter((row): row is BchsSupplementRow => Boolean(row && row.staff));

  if (rows.length === 0) return null;

  const total =
    rows.find((row) => /1\s*піхотний батальйон/i.test(row.unit)) ??
    rows.find((row) => /загалом/i.test(row.unit)) ??
    rows[0];
  const absenceReasons = [
    { label: "Відрядження", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "F")) },
    { label: "Навчання", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "G")) },
    { label: "Шпиталь", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "H")) + bchsToNumber(getSheetValue(sheet, total.rowNumber, "I")) },
    { label: "Відпустка", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "J")) },
    { label: "СЗЧ", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "K")) },
    { label: "Зниклі безвісти", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "L")) },
    { label: "Загиблі", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "M")) },
    { label: "Медрота", value: bchsToNumber(getSheetValue(sheet, total.rowNumber, "N")) + bchsToNumber(getSheetValue(sheet, total.rowNumber, "O")) },
    { label: "Відкомандировані", value: total.detached },
  ].filter((item) => item.value > 0);

  return {
    kind: "appendix",
    title: "БЧС додаток",
    reportDate: getBchsReportDate(sheet),
    total,
    rows,
    totals: rows,
    absenceReasons,
    combatCategories: [
      { label: "Штурмовики", value: total.assaultTotal },
      { label: "Екіпажі техніки", value: total.vehicleCrew },
      { label: "Екіпажі БПЛА", value: total.droneCrew },
      { label: "Колективне озброєння", value: total.crewServedWeapons },
      { label: "Управління", value: total.commandCombat },
      { label: "Забезпечення", value: total.supportCombat },
    ],
    reserveMetrics: [
      { label: "Прикомандировані", value: total.attached },
      { label: "Новоприбулі", value: total.newcomers },
      { label: "Відкомандировані", value: total.detached },
    ],
    bzvpBuckets: [
      { label: "Готові", value: total.assaultReady },
      { label: "На відновленні", value: total.assaultRecovery },
      { label: "На виконанні", value: total.assaultExecution },
      { label: "Без БЗВП", value: total.noBzvp },
    ],
  };
};

export const createBchsComparisonRow = (
  row: Partial<BchsComparisonRow>,
): BchsComparisonRow => {
  const base = { ...emptyBchsRow, ...row };
  const awayOfficers = bchsToNumber(row.awayOfficers);
  const awaySergeants = bchsToNumber(row.awaySergeants);
  const awaySoldiers = bchsToNumber(row.awaySoldiers);
  const hasAwayRankBreakdown =
    row.awayOfficers != null ||
    row.awaySergeants != null ||
    row.awaySoldiers != null;
  // AN must always equal AK+AL+AM when rank breakdown is present —
  // never trust a stale sheet total (e.g. old AO text sum).
  const awayInOtherUnits = hasAwayRankBreakdown
    ? awayOfficers + awaySergeants + awaySoldiers
    : bchsToNumber(row.awayInOtherUnits);
  const attachedOfficers = bchsToNumber(row.attachedOfficers);
  const attachedSergeants = bchsToNumber(row.attachedSergeants);
  const attachedSoldiers = bchsToNumber(row.attachedSoldiers);
  const hasAttachedRankBreakdown =
    row.attachedOfficers != null ||
    row.attachedSergeants != null ||
    row.attachedSoldiers != null;
  const attachedFromOtherUnits = hasAttachedRankBreakdown
    ? attachedOfficers + attachedSergeants + attachedSoldiers
    : bchsToNumber(row.attachedFromOtherUnits);
  const staff = bchsToNumber(base.staff);
  const listed = bchsToNumber(base.listed);
  const available = bchsToNumber(base.available);
  const absentOfficers = bchsToNumber(row.absentOfficers);
  const absentSergeants = bchsToNumber(row.absentSergeants);
  const absentSoldiers = bchsToNumber(row.absentSoldiers);
  const hasAbsentRankBreakdown =
    row.absentOfficers != null ||
    row.absentSergeants != null ||
    row.absentSoldiers != null;
  // Excel V:Y = G−L by rank; fallback Y = J−O when roster breakdown unavailable.
  const listedOfficers = bchsToNumber(row.listedOfficers);
  const listedSergeants = bchsToNumber(row.listedSergeants);
  const listedSoldiers = bchsToNumber(row.listedSoldiers);
  const absent = hasAbsentRankBreakdown
    ? absentOfficers + absentSergeants + absentSoldiers
    : listed > 0
      ? Math.max(0, listed - available)
      : bchsToNumber(base.absent);
  const availableOfficers = hasAbsentRankBreakdown
    ? Math.max(0, listedOfficers - absentOfficers)
    : bchsToNumber(row.availableOfficers);
  const availableSergeants = hasAbsentRankBreakdown
    ? Math.max(0, listedSergeants - absentSergeants)
    : bchsToNumber(row.availableSergeants);
  const availableSoldiers = hasAbsentRankBreakdown
    ? Math.max(0, listedSoldiers - absentSoldiers)
    : bchsToNumber(row.availableSoldiers);
  const availableTotal =
    availableOfficers + availableSergeants + availableSoldiers;
  // Excel: AW/AX/AY = G/H/I − V/W/X − AK/AL/AM + AP/AQ/AR
  const actualOfficers = Math.max(
    0,
    listedOfficers - absentOfficers - awayOfficers + attachedOfficers,
  );
  const actualSergeants = Math.max(
    0,
    listedSergeants - absentSergeants - awaySergeants + attachedSergeants,
  );
  const actualSoldiers = Math.max(
    0,
    listedSoldiers - absentSoldiers - awaySoldiers + attachedSoldiers,
  );
  const inRanksActually = actualOfficers + actualSergeants + actualSoldiers;

  return {
    ...base,
    staffOfficers: bchsToNumber(row.staffOfficers),
    staffSergeants: bchsToNumber(row.staffSergeants),
    staffSoldiers: bchsToNumber(row.staffSoldiers),
    listedOfficers,
    listedSergeants,
    listedSoldiers,
    availableOfficers,
    availableSergeants,
    availableSoldiers,
    available: hasAbsentRankBreakdown ? availableTotal : bchsToNumber(base.available),
    staffedPercent: normalizeBchsPercentValue(
      base.staffedPercent,
      available || listed,
      staff,
    ),
    actualPercent: normalizeBchsPercentValue(
      base.actualPercent,
      inRanksActually,
      staff,
    ),
    inRanksActually,
    actualOfficers,
    actualSergeants,
    actualSoldiers,
    awayInOtherUnits,
    awayOfficers,
    awaySergeants,
    awaySoldiers,
    awayDestinationsText: String(row.awayDestinationsText ?? "").trim(),
    attachedFromOtherUnits,
    attachedOfficers,
    attachedSergeants,
    attachedSoldiers,
    attachedSourcesText: String(row.attachedSourcesText ?? "").trim(),
    unassignedNewcomers: bchsToNumber(row.unassignedNewcomers),
    searchInProgress: bchsToNumber(row.searchInProgress),
    noBzvp: bchsToNumber(row.noBzvp),
    levelPercent: staff > 0 ? inRanksActually / staff : 0,
    balanceActual: listed - absent - awayInOtherUnits + attachedFromOtherUnits,
    assaultReady: bchsToNumber(row.assaultReady),
    assaultRecovery: bchsToNumber(row.assaultRecovery),
    assaultExecution: bchsToNumber(row.assaultExecution),
    assaultTotal: bchsToNumber(row.assaultTotal),
    droneCrew: bchsToNumber(row.droneCrew),
    vehicleCrew: bchsToNumber(row.vehicleCrew),
    crewServedWeapons: bchsToNumber(row.crewServedWeapons),
    commandCombat: bchsToNumber(row.commandCombat),
    supportCombat: bchsToNumber(row.supportCombat),
    combatComponent:
      bchsToNumber(row.combatComponent) ||
      bchsToNumber(row.assaultTotal) +
        bchsToNumber(row.vehicleCrew) +
        bchsToNumber(row.droneCrew) +
        bchsToNumber(row.crewServedWeapons) +
        bchsToNumber(row.commandCombat) +
        bchsToNumber(row.supportCombat),
  };
};

export const formatBchsTableValue = (value: unknown, letter: string) => {
  if (BCHS_PERCENT_COLUMNS.has(letter)) return toPercent(value);
  return valueToDisplay(value as CellValue);
};

export const buildBchsComparisonRowsFromTable = (
  table: BchsAnalyticsSnapshot["table"] | undefined,
) => {
  if (!table) return [];

  return table.rows.map((row) =>
    createBchsComparisonRow({
      rowNumber: row.rowNumber,
      unit:
        valueToDisplay(row.values.B as CellValue) || `Рядок ${row.rowNumber}`,
      staffOfficers: bchsToNumber(row.values.C),
      staffSergeants: bchsToNumber(row.values.D),
      staffSoldiers: bchsToNumber(row.values.E),
      staff: bchsToNumber(row.values.F),
      listedOfficers: bchsToNumber(row.values.G),
      listedSergeants: bchsToNumber(row.values.H),
      listedSoldiers: bchsToNumber(row.values.I),
      listed: bchsToNumber(row.values.J),
      staffedPercent: row.values.K,
      availableOfficers: bchsToNumber(row.values.L),
      availableSergeants: bchsToNumber(row.values.M),
      availableSoldiers: bchsToNumber(row.values.N),
      available: bchsToNumber(row.values.O),
      shortage: bchsToNumber(row.values.T),
      shortagePercent: row.values.U,
      absentOfficers: bchsToNumber(row.values.V),
      absentSergeants: bchsToNumber(row.values.W),
      absentSoldiers: bchsToNumber(row.values.X),
      absent: bchsToNumber(row.values.Y),
      businessTrip: bchsToNumber(row.values.Z),
      training: bchsToNumber(row.values.AA),
      hospitalWounded: bchsToNumber(row.values.AB),
      hospitalIllness: bchsToNumber(row.values.AC),
      vacation: bchsToNumber(row.values.AD),
      awol: bchsToNumber(row.values.AE),
      missing: bchsToNumber(row.values.AF),
      killed: bchsToNumber(row.values.AG),
      medWounded: bchsToNumber(row.values.AH),
      medIllness: bchsToNumber(row.values.AI),
      inRanksActually: bchsToNumber(row.values.AZ),
      actualPercent: row.values.BA,
      combatComponent: bchsToNumber(row.values.BL),
      actualOfficers: bchsToNumber(row.values.AW),
      actualSergeants: bchsToNumber(row.values.AX),
      actualSoldiers: bchsToNumber(row.values.AY),
      awayInOtherUnits: resolveBchsAwayInOtherUnits(row.values.AN),
      awayOfficers: bchsToNumber(row.values.AK),
      awaySergeants: bchsToNumber(row.values.AL),
      awaySoldiers: bchsToNumber(row.values.AM),
      awayDestinationsText: valueToDisplay(row.values.AO as CellValue),
      attachedOfficers: bchsToNumber(row.values.AP),
      attachedSergeants: bchsToNumber(row.values.AQ),
      attachedSoldiers: bchsToNumber(row.values.AR),
      attachedFromOtherUnits: bchsToNumber(row.values.AS),
      attachedSourcesText: valueToDisplay(row.values.AT as CellValue),
      unassignedNewcomers: bchsToNumber(row.values.AU),
      searchInProgress: bchsToNumber(row.values.AV),
      noBzvp: bchsToNumber(row.values.BE),
      assaultReady: bchsToNumber(row.values.BB),
      assaultRecovery: bchsToNumber(row.values.BC),
      assaultExecution: bchsToNumber(row.values.BD),
      assaultTotal: bchsToNumber(row.values.BF),
      vehicleCrew: bchsToNumber(row.values.BG),
      droneCrew: bchsToNumber(row.values.BH),
      crewServedWeapons: bchsToNumber(row.values.BI),
      commandCombat: bchsToNumber(row.values.BJ),
      supportCombat: bchsToNumber(row.values.BK),
    }),
  );
};

export const getBchsHeaderLabel = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
  letter: string,
) => {
  const parts = [4, 5, 7]
    .map((rowNumber) =>
      valueToDisplay(getSheetValue(sheet, rowNumber, letter))
        .replace(/\s+/g, " ")
        .trim(),
    )
    .filter(Boolean);

  return Array.from(new Set(parts)).join(" · ") || letter;
};

export const buildBchsAnalyticsTable = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) => {
  const columns = Array.from(
    { length: BCHS_ANALYTICS_END_COLUMN - BCHS_ANALYTICS_START_COLUMN + 1 },
    (_, index) => {
      const originalColumnIndex = BCHS_ANALYTICS_START_COLUMN + index;
      const letter = getColumnLabel(originalColumnIndex);

      return {
        key: letter,
        letter,
        label: getBchsHeaderLabel(sheet, letter),
        isPercent: BCHS_PERCENT_COLUMNS.has(letter),
      };
    },
  );

  const rows = Array.from({ length: 18 }, (_, index) => {
    const rowNumber = 11 + index;

    return {
      rowNumber,
      values: Object.fromEntries(
        columns.map((column) => [
          column.key,
          formatBchsTableValue(
            getSheetValue(sheet, rowNumber, column.letter),
            column.letter,
          ),
        ]),
      ),
    };
  });

  return { columns, rows };
};

export const buildBchsAnalytics = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
): BchsAnalyticsSnapshot => {
  const totalRow = 11;
  const table = buildBchsAnalyticsTable(sheet);
  const rows = Array.from({ length: 18 }, (_, index) => {
    const rowNumber = totalRow + index;

    return createBchsComparisonRow({
      rowNumber,
      unit:
        valueToDisplay(getSheetValue(sheet, rowNumber, "B")) ||
        `Рядок ${rowNumber}`,
      staffOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "C")),
      staffSergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "D")),
      staffSoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "E")),
      staff: bchsToNumber(getSheetValue(sheet, rowNumber, "F")),
      listedOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "G")),
      listedSergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "H")),
      listedSoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "I")),
      listed: bchsToNumber(getSheetValue(sheet, rowNumber, "J")),
      staffedPercent: getSheetValue(sheet, rowNumber, "K"),
      availableOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "L")),
      availableSergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "M")),
      availableSoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "N")),
      available: bchsToNumber(getSheetValue(sheet, rowNumber, "O")),
      shortage: bchsToNumber(getSheetValue(sheet, rowNumber, "T")),
      shortagePercent: getSheetValue(sheet, rowNumber, "U"),
      absentOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "V")),
      absentSergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "W")),
      absentSoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "X")),
      absent: bchsToNumber(getSheetValue(sheet, rowNumber, "Y")),
      businessTrip: bchsToNumber(getSheetValue(sheet, rowNumber, "Z")),
      training: bchsToNumber(getSheetValue(sheet, rowNumber, "AA")),
      hospitalWounded: bchsToNumber(getSheetValue(sheet, rowNumber, "AB")),
      hospitalIllness: bchsToNumber(getSheetValue(sheet, rowNumber, "AC")),
      vacation: bchsToNumber(getSheetValue(sheet, rowNumber, "AD")),
      awol: bchsToNumber(getSheetValue(sheet, rowNumber, "AE")),
      missing: bchsToNumber(getSheetValue(sheet, rowNumber, "AF")),
      killed: bchsToNumber(getSheetValue(sheet, rowNumber, "AG")),
      medWounded: bchsToNumber(getSheetValue(sheet, rowNumber, "AH")),
      medIllness: bchsToNumber(getSheetValue(sheet, rowNumber, "AI")),
      inRanksActually: bchsToNumber(getSheetValue(sheet, rowNumber, "AZ")),
      actualPercent: getSheetValue(sheet, rowNumber, "BA"),
      combatComponent: bchsToNumber(getSheetValue(sheet, rowNumber, "BL")),
      actualOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "AW")),
      actualSergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "AX")),
      actualSoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "AY")),
      awayInOtherUnits: resolveBchsAwayInOtherUnits(
        getSheetValue(sheet, rowNumber, "AN"),
      ),
      awayOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "AK")),
      awaySergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "AL")),
      awaySoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "AM")),
      awayDestinationsText: valueToDisplay(
        getSheetValue(sheet, rowNumber, "AO") as CellValue,
      ),
      attachedOfficers: bchsToNumber(getSheetValue(sheet, rowNumber, "AP")),
      attachedSergeants: bchsToNumber(getSheetValue(sheet, rowNumber, "AQ")),
      attachedSoldiers: bchsToNumber(getSheetValue(sheet, rowNumber, "AR")),
      attachedFromOtherUnits: bchsToNumber(
        getSheetValue(sheet, rowNumber, "AS"),
      ),
      attachedSourcesText: valueToDisplay(
        getSheetValue(sheet, rowNumber, "AT") as CellValue,
      ),
      unassignedNewcomers: bchsToNumber(getSheetValue(sheet, rowNumber, "AU")),
      searchInProgress: bchsToNumber(getSheetValue(sheet, rowNumber, "AV")),
      noBzvp: bchsToNumber(getSheetValue(sheet, rowNumber, "BE")),
      assaultReady: bchsToNumber(getSheetValue(sheet, rowNumber, "BB")),
      assaultRecovery: bchsToNumber(getSheetValue(sheet, rowNumber, "BC")),
      assaultExecution: bchsToNumber(getSheetValue(sheet, rowNumber, "BD")),
      assaultTotal: bchsToNumber(getSheetValue(sheet, rowNumber, "BF")),
      vehicleCrew: bchsToNumber(getSheetValue(sheet, rowNumber, "BG")),
      droneCrew: bchsToNumber(getSheetValue(sheet, rowNumber, "BH")),
      crewServedWeapons: bchsToNumber(getSheetValue(sheet, rowNumber, "BI")),
      commandCombat: bchsToNumber(getSheetValue(sheet, rowNumber, "BJ")),
      supportCombat: bchsToNumber(getSheetValue(sheet, rowNumber, "BK")),
      levelPercent: 0,
      balanceActual: 0,
    });
  }).map((row) => ({
    ...row,
    levelPercent: row.staff > 0 ? row.inRanksActually / row.staff : 0,
    balanceActual:
      row.listed -
      row.absent -
      row.awayInOtherUnits +
      row.attachedFromOtherUnits,
  }));

  return {
    reportDate: getBchsReportDate(sheet),
    total: rows[0] ?? createBchsComparisonRow(emptyBchsRow),
    rows,
    comparisonRows: rows,
    table,
    detachedDestinations: buildBchsDetachedDestinationsFromSheet(sheet),
    attachedSources: [],
    absenceReasons: buildBchsAbsenceReasons(rows[0] ?? emptyBchsRow),
    dataIssues: buildBchsComparisonAnomalyIssues(rows),
  } satisfies BchsAnalyticsSnapshot;
};

type BchsGeneralListUnitAccumulator = {
  unit: string;
  staff: number;
  staffOfficers: number;
  staffSergeants: number;
  staffSoldiers: number;
  listed: number;
  listedOfficers: number;
  listedSergeants: number;
  listedSoldiers: number;
  available: number;
  availableOfficers: number;
  availableSergeants: number;
  availableSoldiers: number;
  absent: number;
  absentOfficers: number;
  absentSergeants: number;
  absentSoldiers: number;
  businessTrip: number;
  training: number;
  hospitalWounded: number;
  hospitalIllness: number;
  medWounded: number;
  medIllness: number;
  vacation: number;
  awol: number;
  missing: number;
  killed: number;
  inRanksActually: number;
  actualOfficers: number;
  actualSergeants: number;
  actualSoldiers: number;
  unassignedNewcomers: number;
  noBzvp: number;
  assaultReady: number;
  assaultRecovery: number;
  assaultExecution: number;
  assaultTotal: number;
  vehicleCrew: number;
  droneCrew: number;
  crewServedWeapons: number;
  commandCombat: number;
  supportCombat: number;
};

const emptyBchsGeneralListUnit = (
  unit: string,
): BchsGeneralListUnitAccumulator => ({
  unit,
  staff: 0,
  staffOfficers: 0,
  staffSergeants: 0,
  staffSoldiers: 0,
  listed: 0,
  listedOfficers: 0,
  listedSergeants: 0,
  listedSoldiers: 0,
  available: 0,
  availableOfficers: 0,
  availableSergeants: 0,
  availableSoldiers: 0,
  absent: 0,
  businessTrip: 0,
  absentOfficers: 0,
  absentSergeants: 0,
  absentSoldiers: 0,
  training: 0,
  hospitalWounded: 0,
  hospitalIllness: 0,
  medWounded: 0,
  medIllness: 0,
  vacation: 0,
  awol: 0,
  missing: 0,
  killed: 0,
  inRanksActually: 0,
  actualOfficers: 0,
  actualSergeants: 0,
  actualSoldiers: 0,
  unassignedNewcomers: 0,
  noBzvp: 0,
  assaultReady: 0,
  assaultRecovery: 0,
  assaultExecution: 0,
  assaultTotal: 0,
  vehicleCrew: 0,
  droneCrew: 0,
  crewServedWeapons: 0,
  commandCombat: 0,
  supportCombat: 0,
});

export const isBchsGeneralListPresentStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized === "в строю" || normalized.includes("новоприбулий");
};

/** Excel O (наявність): «в строю» / новоприбулий + відком. за межі ПБ. */
export const isBchsGeneralListAvailableStatus = (status: string) =>
  isBchsGeneralListPresentStatus(status) || isBchsDetachedStatus(status);

export type BchsNovaRosterUnitStats = {
  staff: number;
  listed: number;
  present: number;
};

/** Count nova roster people for exact/alias unit labels (Управління=ж, Штаб, …). */
export const countBchsNovaRosterUnitStats = (
  people: BchsPersonnelAwayPerson[],
  unitAliases: string[],
): BchsNovaRosterUnitStats => {
  const aliases = unitAliases.map(normalizeBchsText).filter(Boolean);
  const stats: BchsNovaRosterUnitStats = { staff: 0, listed: 0, present: 0 };

  filterBchsNovaPeople(people).forEach((person) => {
    const roster = normalizeBchsText(person.rosterUnit);
    if (!aliases.some((alias) => roster === alias || roster.includes(alias)))
      return;

    stats.staff += 1;
    if (person.fullName.trim()) stats.listed += 1;
    if (isBchsGeneralListPresentStatus(person.status)) stats.present += 1;
  });

  return stats;
};

/** Excel U: точне «в строю» (без новоприбулих) — для 1ПБ таблиці. */
export const isBchsExactInRanksStatus = (status: string) =>
  normalizeBchsText(status) === "в строю";

/** Excel U: *в строю* — для COUNTIFS колонки G у 1ПБ таблиці. */
export const isBchsPersonnel1PbInRanksStatus = (status: string) =>
  normalizeBchsText(status).includes("в строю");

const isBchsNovaRosterExactUnitMatch = (
  rosterUnit: string,
  aliases: string[],
) => {
  const roster = normalizeBchsText(rosterUnit);
  return aliases.some((alias) => roster === normalizeBchsText(alias));
};

/** Короткі коди B (напр. «ж») — лише exact match, без substring. */
const isBchsPersonnel1PbRosterUnitMatch = (
  rosterUnit: string,
  aliases: string[],
) => {
  const roster = normalizeBchsText(rosterUnit);
  return aliases.some((alias) => {
    const normalized = normalizeBchsText(alias);
    if (normalized.length <= 2) return roster === normalized;
    return roster === normalized || roster.includes(normalized);
  });
};

const isBchsNovaRosterWildcardUnitMatch = (
  rosterUnit: string,
  pattern: string,
) => normalizeBchsText(rosterUnit).includes(normalizeBchsText(pattern));

export type BchsPersonnel1PbUnitMode =
  | "roster-command"
  | "roster-presence"
  | "zenit"
  | "arkush1";

export type BchsPersonnel1PbUnitStats = {
  staff: number;
  listed: number;
  presence: number;
  listedStaffRatio: number;
  combatTask: number;
  commanderReserve: number;
};

const BCHS_PERSONNEL_1PB_ROSTER_ALIASES: Record<string, string[]> = {
  /** Управління в цій таблиці = roster B «ж» (exact, не substring). */
  Управління: ["ж"],
  Штаб: ["штаб"],
  "Група безпілотних систем": ["група безпілотних систем"],
};

const BCHS_PERSONNEL_1PB_UNIT_MODES: Record<string, BchsPersonnel1PbUnitMode> =
  {
    Управління: "roster-command",
    Штаб: "roster-command",
    "Група безпілотних систем": "roster-command",
    "Відділення радіоелектронної боротьби": "roster-presence",
    "Інженерно-саперне відділення": "roster-presence",
    "Зенітно-ракетний взвод": "zenit",
  };

export const resolveBchsPersonnel1PbUnitAliases = (unitName: string) => {
  const fromMap = BCHS_PERSONNEL_1PB_ROSTER_ALIASES[unitName];
  const aliases = fromMap
    ? [...fromMap]
    : (() => {
        const normalized = normalizeBchsText(unitName);
        if (normalized.includes("вмтз"))
          return ["взвод матеріально-технічного забезпечення", "вмтз"];
        if (normalized === "ісв" || normalized.includes("інженерно-саперне"))
          return ["інженерно-саперне відділення", "ісв"];
        if (normalized.includes("реб"))
          return ["відділення радіоелектронної боротьби", "відділення реб"];
        return [unitName];
      })();

  const extras = aliases.flatMap((alias) => {
    const text = normalizeBchsText(alias);
    if (text.includes("піхотна рота"))
      return [text.replace("піхотна рота", "штурмова рота")];
    if (text.includes("штурмова рота"))
      return [text.replace("штурмова рота", "піхотна рота")];
    return [];
  });

  return [...aliases, ...extras];
};

const getBchsPersonnel1PbUnitMode = (unitName: string): BchsPersonnel1PbUnitMode =>
  BCHS_PERSONNEL_1PB_UNIT_MODES[unitName] ?? "arkush1";

/** Управління / Штаб / ГБС: COUNTIFS оф.+серж. за B (Управління = «ж» exact). */
export const countBchsNovaRosterCommandUnitStats = (
  people: BchsPersonnelAwayPerson[],
  unitAliases: string[],
): BchsNovaRosterUnitStats => {
  const aliases = unitAliases.map(normalizeBchsText).filter(Boolean);
  const stats: BchsNovaRosterUnitStats = { staff: 0, listed: 0, present: 0 };
  const commandRanks = new Set([BCHS_RANK_OFFICER, BCHS_RANK_SERGEANT]);

  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsPersonnel1PbRosterUnitMatch(person.rosterUnit, aliases)) return;

    const rank = normalizeBchsText(person.rankCategory);
    if (!commandRanks.has(rank)) return;

    stats.staff += 1;
    if (isBchsListedWithRankTitle(person)) stats.listed += 1;
    if (isBchsExactInRanksStatus(person.status)) stats.present += 1;
  });

  return stats;
};

/** РЕБ / ІСВ: COUNTIFS усі звання, U=«в стroю». */
export const countBchsNovaRosterPresenceAllRanks = (
  people: BchsPersonnelAwayPerson[],
  unitAliases: string[],
) => {
  const aliases = unitAliases.map(normalizeBchsText).filter(Boolean);
  let count = 0;

  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsNovaRosterExactUnitMatch(person.rosterUnit, aliases)) return;
    if (!isBchsExactInRanksStatus(person.status)) return;
    count += 1;
  });

  return count;
};

/** Excel G: COUNTIFS(B *unit*, U *в stрою*, AE «на виконанні»). */
export const countBchsNovaCombatTask = (
  people: BchsPersonnelAwayPerson[],
  unitPattern: string,
) => {
  let count = 0;

  filterBchsNovaPeople(people).forEach((person) => {
    if (!isBchsNovaRosterWildcardUnitMatch(person.rosterUnit, unitPattern))
      return;
    if (!isBchsPersonnel1PbInRanksStatus(person.status)) return;
    if (normalizeBchsText(person.medicalPlace) !== "на виконанні") return;
    count += 1;
  });

  return count;
};

/** Dark green PIB fill in roster «Ос загальний» / Аркуш2 (резерв командира полку). */
export const BCHS_COMMANDER_RESERVE_PIB_FILL_RGB = new Set([
  "FF38761D",
  "FF00B050",
  "FF548235",
  "FF70AD47",
  "FF375623",
  "FF385723",
  "FF006600",
  "FF008000",
  "FF92D050",
  "FFC6EFCE",
  "FF216E39",
]);

const rgbToHueSatLight = (r: number, g: number, b: number) => {
  const red = r / 255;
  const green = g / 255;
  const blue = b / 255;
  const max = Math.max(red, green, blue);
  const min = Math.min(red, green, blue);
  const lightness = (max + min) / 2;
  const delta = max - min;
  const saturation =
    delta === 0 ? 0 : delta / (1 - Math.abs(2 * lightness - 1));
  let hue = 0;
  if (delta !== 0) {
    if (max === red) hue = 60 * (((green - blue) / delta) % 6);
    else if (max === green) hue = 60 * ((blue - red) / delta + 2);
    else hue = 60 * ((red - green) / delta + 4);
    if (hue < 0) hue += 360;
  }
  return { hue, saturation, lightness };
};

export const isBchsCommanderReservePibHighlight = (
  rgb: string | null | undefined,
) => {
  if (!rgb) return false;
  const hex = rgb.replace(/^#/, "").toUpperCase();
  const argb = hex.length === 6 ? `FF${hex}` : hex;
  if (BCHS_COMMANDER_RESERVE_PIB_FILL_RGB.has(argb)) return true;
  if (argb.length !== 8) return false;
  const r = Number.parseInt(argb.slice(2, 4), 16);
  const g = Number.parseInt(argb.slice(4, 6), 16);
  const b = Number.parseInt(argb.slice(6, 8), 16);
  if (![r, g, b].every(Number.isFinite)) return false;
  if (g >= 45 && g > r && g > b && g - r >= 12 && g - b >= 12) return true;
  const { hue, saturation, lightness } = rgbToHueSatLight(r, g, b);
  return (
    hue >= 70 &&
    hue <= 165 &&
    saturation >= 0.18 &&
    lightness >= 0.12 &&
    lightness <= 0.78
  );
};

/** Excel L = «10» — особа в штатці (штатна посада в ШПК). */
export const isBchsInStaffTableMobilization = (
  person: BchsPersonnelAwayPerson,
) => normalizeBchsText(person.mobilizationContract) === "10";

/** Особа в штатці: є ПІБ, штатна посада, звання за ШПК (H/M). */
export const isBchsInStaffTablePerson = (
  person: BchsPersonnelAwayPerson,
) => {
  if (!person.fullName.trim()) return false;
  if (isBchsWithoutStaffPosition(person)) return false;
  return Boolean(
    normalizeBchsText(person.shpkFact) || normalizeBchsText(person.rankTitle),
  );
};

/** @deprecated alias — L=10 + штатка; для J використовуй isBchsInStaffTablePerson. */
export const isBchsStaffTableReservePerson = (
  person: BchsPersonnelAwayPerson,
) => {
  if (!isBchsInStaffTablePerson(person)) return false;
  return isBchsInStaffTableMobilization(person);
};

const isBchsCommanderReserveStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  if (normalized === "в строю") return true;
  return normalized.includes("новоприбул");
};

/**
 * Excel J — резерв командира полку:
 * «нова» / «в строю» / «новоприбулий» + ПІБ із зеленою заливкою.
 */
export const isBchsCommanderReservePerson = (
  person: BchsPersonnelAwayPerson,
) => {
  if (!person.fullName.trim()) return false;
  if (!isBchsCommanderReserveStatus(person.status)) return false;
  return isBchsCommanderReservePibHighlight(person.pibHighlightRgb);
};

export const summarizeBchsCommanderReserve = (
  people: BchsPersonnelAwayPerson[],
) => {
  const novaPeople = filterBchsNovaPeople(people);
  const pool = novaPeople.length > 0 ? novaPeople : people;
  const inRanks = pool.filter(
    (person) =>
      Boolean(person.fullName.trim()) &&
      isBchsCommanderReserveStatus(person.status),
  );
  const withFill = inRanks.filter((person) => person.pibHighlightRgb);
  const green = inRanks.filter((person) =>
    isBchsCommanderReservePibHighlight(person.pibHighlightRgb),
  );
  const fillSamples = [
    ...new Set(
      withFill
        .map((person) => person.pibHighlightRgb)
        .filter((value): value is string => Boolean(value)),
    ),
  ].slice(0, 8);

  return {
    pool: pool.length,
    inRanks: inRanks.length,
    withFill: withFill.length,
    green: green.length,
    fillSamples,
  };
};

export const countBchsNovaCommanderReserve = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
) => {
  const aliases = resolveBchsPersonnel1PbUnitAliases(unitName);
  const novaPeople = filterBchsNovaPeople(people);
  const pool = novaPeople.length > 0 ? novaPeople : people;
  let count = 0;

  pool.forEach((person) => {
    if (!isBchsPersonnel1PbRosterUnitMatch(person.rosterUnit, aliases)) return;
    if (!isBchsCommanderReservePerson(person)) return;
    count += 1;
  });

  return count;
};

/** Розрахунок одного рядка таблиці «1ПБ Особовий склад + БЗВП». */
export const computeBchsPersonnel1PbUnitStats = (
  unitName: string,
  comparisonRow: BchsComparisonRow | undefined,
  people: BchsPersonnelAwayPerson[],
): BchsPersonnel1PbUnitStats => {
  const mode = getBchsPersonnel1PbUnitMode(unitName);
  const aliases = resolveBchsPersonnel1PbUnitAliases(unitName);
  const commanderReserve = countBchsNovaCommanderReserve(people, unitName);

  if (mode === "zenit") {
    return {
      staff: 0,
      listed: 0,
      presence: 0,
      listedStaffRatio: 0,
      combatTask: 0,
      commanderReserve,
    };
  }

  let staff = 0;
  let listed = 0;
  let presence = 0;

  if (mode === "roster-command") {
    const roster = countBchsNovaRosterCommandUnitStats(people, aliases);
    staff = roster.staff;
    listed = roster.listed;
    presence = roster.present;
  } else if (mode === "roster-presence") {
    staff = comparisonRow?.staff ?? 0;
    listed = comparisonRow?.listed ?? 0;
    presence = countBchsNovaRosterPresenceAllRanks(people, aliases);
  } else {
    staff = comparisonRow?.staff ?? 0;
    listed = comparisonRow?.listed ?? 0;
    presence = comparisonRow?.combatComponent ?? 0;
  }

  const listedStaffRatio = staff > 0 ? listed / staff : 0;

  const combatTask =
    mode === "roster-command" ? 0 : countBchsNovaCombatTask(people, unitName);

  return {
    staff,
    listed,
    presence,
    listedStaffRatio,
    combatTask,
    commanderReserve,
  };
};

const explainBchsGeneralListStatusIssue = (
  person: BchsPersonnelAwayPerson,
) => {
  const status = normalizeBchsText(person.status);
  if (!person.fullName.trim()) return "";
  if (!status) return "порожній статус";
  if (status.includes("невідом")) return "статус позначений як невідомий";
  if (
    isBchsGeneralListPresentStatus(person.status) ||
    isBchsDetachedStatus(person.status) ||
    status.includes("відряд") ||
    status.includes("навч") ||
    status.includes("ліку") ||
    status.includes("відпуст") ||
    status.includes("сзч") ||
    status.includes("зник") ||
    status.includes("загиб")
  )
    return "";
  return "статус не потрапив у жодну категорію БЧС";
};

const hasBchsUnknownDetachedDestination = (person: BchsPersonnelAwayPerson) => {
  if (!person.fullName.trim()) return false;
  if (!isBchsDetachedStatus(person.status)) return false;
  const destination = normalizeBchsText(
    normalizeBchsDestinationLabel(person.destination),
  );
  return !destination || destination.includes("невідом");
};

const hasMixedScriptRankTitle = (rankTitle: string) =>
  /[A-Za-z]/.test(rankTitle) && /[А-Яа-яІіЇїЄєҐґ]/.test(rankTitle);

const explainBchsRankTitleIssue = (
  person: BchsPersonnelAwayPerson,
): string | null => {
  if (!person.fullName.trim()) return null;
  if (!isBchsListedWithRankTitle(person)) return null;

  const title = person.rankTitle.trim();

  // Лише реальні помилки написання (змішана латиниця), не коректна укр. орфографія.
  // «капітан» з українською «і» — правильно; Excel-маски *капитан* — окрема тема формул.
  if (hasMixedScriptRankTitle(title)) {
    return `у званні змішані латиниця і кирилиця: «${title}» (перевірте лат. i/a/o замість кирилиці)`;
  }

  return null;
};

const buildBchsGeneralListDataIssues = (
  people: BchsPersonnelAwayPerson[],
): BchsDataIssue[] =>
  people
    .filter((person) => normalizeBchsText(person.battalion) === "нова")
    .flatMap((person) => {
      const base = {
        fullName: person.fullName.trim() || "(без ПІБ)",
        rosterUnit: normalizeBchsRosterUnitLabel(person.rosterUnit),
        status: person.status.trim() || "(порожньо)",
        rankTitle: person.rankTitle.trim() || undefined,
        rankCategory: person.rankCategory.trim() || undefined,
      };
      const issues: BchsDataIssue[] = [];
      const statusReason = explainBchsGeneralListStatusIssue(person);
      if (statusReason) {
        issues.push({ ...base, kind: "status", reason: statusReason });
      }
      if (hasBchsUnknownDetachedDestination(person)) {
        issues.push({
          ...base,
          kind: "destination",
          destination:
            normalizeBchsDestinationLabel(person.destination) || "(порожньо)",
          reason: "невідоме місце відкомандирування",
        });
      }
      const rankReason = explainBchsRankTitleIssue(person);
      if (rankReason) {
        issues.push({ ...base, kind: "rank", reason: rankReason });
      }
      return issues;
    })
    .filter((issue): issue is BchsDataIssue => Boolean(issue));

const parseBrezCountFromSourcesText = (sourcesText: string) => {
  const match = String(sourcesText ?? "").match(/брез\s*[-–—]?\s*(\d+)/i);
  return match ? Number(match[1]) : 0;
};

/** Аномалії Аркуш1: роздутий БРЕЗ / абсурдний % фактичної наявності. */
export const buildBchsComparisonAnomalyIssues = (
  rows: BchsComparisonRow[],
): BchsDataIssue[] =>
  rows
    .filter(
      (row) =>
        row.rowNumber > 11 &&
        !isBchsTotalUnit(row.unit) &&
        row.staff > 0 &&
        row.unit.trim(),
    )
    .flatMap((row) => {
      const issues: BchsDataIssue[] = [];
      const actualPercent =
        row.staff > 0 ? row.inRanksActually / row.staff : 0;
      const storedPercent = bchsToNumber(row.actualPercent);
      const percent = Math.max(actualPercent, storedPercent);
      const brezAttached = parseBrezCountFromSourcesText(
        row.attachedSourcesText,
      );
      const attached = row.attachedFromOtherUnits;
      const inflatedBrez = brezAttached >= 40;
      const absurdPercent = percent >= 5;
      const inflatedAttached =
        isBchsRebUnit(row.unit) && attached >= 40 && attached > row.staff * 2;

      if (!inflatedBrez && !absurdPercent && !inflatedAttached) return issues;

      const pctLabel = `${Math.round(percent * 100)}%`;
      const parts: string[] = [];
      if (absurdPercent || inflatedAttached) {
        parts.push(
          `факт. укомплектованість ${pctLabel} (= ${row.inRanksActually} факт / ${row.staff} штат)`,
        );
      }
      if (brezAttached > 0) {
        parts.push(`у «Звідки прикомандировані» — БРЕЗ-${brezAttached}`);
      } else if (attached > 0) {
        parts.push(`прикомандировані загалом ${attached}`);
      }
      parts.push(
        "Ймовірна причина: у підрахунок БРЕЗ потрапили позначки «був БРЕЗ» з колонки AA (Відрядження), хоча Excel COUNTIFS для РРЕБ дивиться лише на X (БЗВП/БРЕЗ). Після виправлення мають лишитись лише актуальні позначки БРЕЗ у X (~20–30 осіб).",
      );

      issues.push({
        fullName: row.unit,
        rosterUnit: row.unit,
        status: `штат ${row.staff} · факт ${row.inRanksActually} · приком. ${attached}`,
        kind: "anomaly",
        reason: parts.join(". "),
      });

      return issues;
    });

/** Люди з «був БРЕЗ» лише в AA — не мають рахуватись у AP–AR РРЕБ. */
export const buildBchsBrezMiscountIssues = (
  people: BchsPersonnelAwayPerson[],
): BchsDataIssue[] => {
  const formerOnly = people.filter(
    (person) =>
      isBchsAttachedPresentStatus(person.status) &&
      hasBchsFormerBrezAssignmentOnly(person) &&
      Boolean(classifyBchsAttachedRank(person.rankCategory, person.rankTitle)),
  );
  if (formerOnly.length === 0) return [];

  const byRank = { officers: 0, sergeants: 0, soldiers: 0 };
  for (const person of formerOnly) {
    const rank = classifyBchsAttachedRank(
      person.rankCategory,
      person.rankTitle,
    );
    if (rank) byRank[rank] += 1;
  }

  return [
    {
      fullName: `БРЕЗ (хибний підрахунок AA)`,
      rosterUnit: "Відділення РРЕБ / БРЕЗ",
      status: "був БРЕЗ",
      kind: "brez",
      reason: `${formerOnly.length} осіб зі статусом «в строю/новоприбулий» мають «був БРЕЗ» у колонці AA, але без «БРЕЗ» у X (БЗВП/БРЕЗ). Excel AP–AR їх не рахує; раніше застосунок помилково додавав їх → БРЕЗ-${formerOnly.length} (оф.${byRank.officers} / серж.${byRank.sergeants} / солд.${byRank.soldiers}).`,
    },
    ...formerOnly.slice(0, 12).map((person) => ({
      fullName: person.fullName.trim() || "(без ПІБ)",
      rosterUnit: normalizeBchsRosterUnitLabel(person.rosterUnit),
      status: person.status.trim() || "(порожньо)",
      rankTitle: person.rankTitle.trim() || undefined,
      rankCategory: person.rankCategory.trim() || undefined,
      kind: "brez" as const,
      reason: `AA: «${person.brezAssignment.trim()}» · X: «${person.bzvpStatus.trim() || "пусто"}» — не входить у прикомандировані РРЕБ`,
    })),
  ];
};

const mergeBchsDataIssues = (
  ...groups: Array<BchsDataIssue[] | undefined>
): BchsDataIssue[] | undefined => {
  const merged = groups.flatMap((group) => group ?? []);
  return merged.length > 0 ? merged : undefined;
};

const addBchsGeneralListTotals = (
  total: BchsGeneralListUnitAccumulator,
  row: BchsGeneralListUnitAccumulator,
) => {
  total.staff += row.staff;
  total.staffOfficers += row.staffOfficers;
  total.staffSergeants += row.staffSergeants;
  total.staffSoldiers += row.staffSoldiers;
  total.listed += row.listed;
  total.listedOfficers += row.listedOfficers;
  total.listedSergeants += row.listedSergeants;
  total.listedSoldiers += row.listedSoldiers;
  total.available += row.available;
  total.availableOfficers += row.availableOfficers;
  total.availableSergeants += row.availableSergeants;
  total.availableSoldiers += row.availableSoldiers;
  total.absent += row.absent;
  total.absentOfficers += row.absentOfficers;
  total.absentSergeants += row.absentSergeants;
  total.absentSoldiers += row.absentSoldiers;
  total.businessTrip += row.businessTrip;
  total.training += row.training;
  total.hospitalWounded += row.hospitalWounded;
  total.hospitalIllness += row.hospitalIllness;
  total.medWounded += row.medWounded;
  total.medIllness += row.medIllness;
  total.vacation += row.vacation;
  total.awol += row.awol;
  total.missing += row.missing;
  total.killed += row.killed;
  total.inRanksActually += row.inRanksActually;
  total.actualOfficers += row.actualOfficers;
  total.actualSergeants += row.actualSergeants;
  total.actualSoldiers += row.actualSoldiers;
  total.unassignedNewcomers += row.unassignedNewcomers;
  total.noBzvp += row.noBzvp;
  total.assaultReady += row.assaultReady;
  total.assaultRecovery += row.assaultRecovery;
  total.assaultExecution += row.assaultExecution;
  total.assaultTotal += row.assaultTotal;
  total.vehicleCrew += row.vehicleCrew;
  total.droneCrew += row.droneCrew;
  total.crewServedWeapons += row.crewServedWeapons;
  total.commandCombat += row.commandCombat;
  total.supportCombat += row.supportCombat;
};

const bchsGeneralListUnitToRow = (
  item: BchsGeneralListUnitAccumulator,
  rowNumber: number,
) =>
  createBchsComparisonRow({
    rowNumber,
    unit: item.unit,
    staffOfficers: item.staffOfficers,
    staffSergeants: item.staffSergeants,
    staffSoldiers: item.staffSoldiers,
    staff: item.staff,
    listedOfficers: item.listedOfficers,
    listedSergeants: item.listedSergeants,
    listedSoldiers: item.listedSoldiers,
    listed: item.listed,
    staffedPercent: item.staff ? item.listed / item.staff : 0,
    availableOfficers: item.availableOfficers,
    availableSergeants: item.availableSergeants,
    availableSoldiers: item.availableSoldiers,
    available: item.available,
    shortage: Math.max(item.staff - item.listed, 0),
    shortagePercent: item.staff
      ? Math.max(item.staff - item.listed, 0) / item.staff
      : 0,
    absent: item.absent,
    absentOfficers: item.absentOfficers,
    absentSergeants: item.absentSergeants,
    absentSoldiers: item.absentSoldiers,
    businessTrip: item.businessTrip,
    training: item.training,
    hospitalWounded: item.hospitalWounded,
    hospitalIllness: item.hospitalIllness,
    medWounded: item.medWounded,
    medIllness: item.medIllness,
    vacation: item.vacation,
    awol: item.awol,
    missing: item.missing,
    killed: item.killed,
    inRanksActually: item.inRanksActually,
    actualPercent: item.staff ? item.inRanksActually / item.staff : 0,
    combatComponent: item.inRanksActually,
    actualOfficers: item.actualOfficers,
    actualSergeants: item.actualSergeants,
    actualSoldiers: item.actualSoldiers,
    unassignedNewcomers: item.unassignedNewcomers,
    noBzvp: item.noBzvp,
    assaultReady: item.assaultReady,
    assaultRecovery: item.assaultRecovery,
    assaultExecution: item.assaultExecution,
    assaultTotal: item.assaultTotal,
    vehicleCrew: item.vehicleCrew,
    droneCrew: item.droneCrew,
    crewServedWeapons: item.crewServedWeapons,
    commandCombat: item.commandCombat,
    supportCombat: item.supportCombat,
  });

const BCHS_GENERAL_LIST_UNIT_ORDER = [
  "Командування",
  "1 піхотна рота",
  "2 піхотна рота",
  "3 піхотна рота",
  "Рота безпілотних авіаційних комплексів",
  "Гранатометний взвод",
  "Мінометний взвод",
  "Кулеметний взвод",
  "Взвод протитанкових ракетних комплексів",
  "Взвод логістично-евакуаційних безпілотних наземних систем",
  "Взвод перехоплювачів безпілотних літальних апаратів",
  "Розвідувальний взвод",
  "Інженерно-саперне відділення",
  "Відділення радіоелектронної боротьби",
  "Взвод зв'язку",
  "Взвод матеріально-технічного забезпечення",
  "Медичний пункт",
];

const getBchsGeneralListUnitOrder = (unit: string) => {
  const normalized = normalizeBchsText(unit).replace(/[’ʼ`]/g, "'");
  const index = BCHS_GENERAL_LIST_UNIT_ORDER.findIndex(
    (item) => normalizeBchsText(item).replace(/[’ʼ`]/g, "'") === normalized,
  );
  return index >= 0 ? index : Number.MAX_SAFE_INTEGER;
};

export const buildBchsAnalyticsFromGeneralList = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
): BchsAnalyticsSnapshot | null => {
  if (!isBchsPersonnelGeneralListSheet(sheet)) return null;

  const allPeople = extractBchsAwayPeopleFromSheet(sheet);
  const novaPeople = filterBchsNovaPeople(allPeople);
  const units = new Map<string, BchsGeneralListUnitAccumulator>();

  const getUnit = (unitName: string) => {
    const unit = normalizeBchsRosterUnitLabel(unitName);
    const current = units.get(unit) ?? emptyBchsGeneralListUnit(unit);
    units.set(unit, current);
    return current;
  };

  novaPeople.forEach((person) => {
      const unit = getUnit(person.rosterUnit);
      const fullName = person.fullName.trim();
      const rank = normalizeBchsText(person.rankCategory);
      const roleType = normalizeBchsText(person.roleType);
      const combatReadiness = normalizeBchsText(person.combatReadiness);
      const bzvpStatus = normalizeBchsText(person.bzvpStatus);

      unit.staff += 1;
      if (rank === BCHS_RANK_OFFICER) unit.staffOfficers += 1;
      else if (rank === BCHS_RANK_SERGEANT) unit.staffSergeants += 1;
      else if (rank === BCHS_RANK_SOLDIER) unit.staffSoldiers += 1;

      if (fullName) {
        unit.listed += 1;
      }

      if (isBchsGeneralListPresentStatus(person.status)) {
        unit.inRanksActually += 1;
        if (rank === BCHS_RANK_OFFICER) unit.actualOfficers += 1;
        else if (rank === BCHS_RANK_SERGEANT) unit.actualSergeants += 1;
        else if (rank === BCHS_RANK_SOLDIER) unit.actualSoldiers += 1;
      }

      // Excel Z: COUNTIFS(U;"*Відрядження*", M<>).
      if (
        isBchsListedWithRankTitle(person) &&
        isBchsBusinessTripStatus(person.status)
      ) {
        unit.businessTrip += 1;
      }
      if (bzvpStatus.includes("без бзвп")) unit.noBzvp += 1;
      if (roleType.includes("піхот")) {
        unit.assaultTotal += 1;
        if (combatReadiness === "бг") unit.assaultReady += 1;
        else if (combatReadiness.includes("тимчасово")) unit.assaultRecovery += 1;
        else unit.assaultExecution += 1;
      }
      if (roleType.includes("пілот") || roleType.includes("бпла"))
        unit.droneCrew += 1;
      if (roleType.includes("водій"))
        unit.vehicleCrew += 1;
      if (
        roleType.includes("гранатомет") ||
        roleType.includes("міномет") ||
        roleType.includes("кулемет")
      )
        unit.crewServedWeapons += 1;
      if (roleType.includes("упр") || roleType.includes("штаб"))
        unit.commandCombat += 1;
      if (
        roleType.includes("забезпеч") ||
        roleType.includes("охорон") ||
        roleType.includes("медик") ||
        roleType.includes("кухар") ||
        roleType.includes("майстер")
      )
        unit.supportCombat += 1;
    });

  Array.from(units.values()).forEach((unit) => {
    const rankStats = computeBchsUnitRankListedAvailableStats(
      novaPeople,
      unit.unit,
    );
    unit.listedOfficers = rankStats.listed.officers;
    unit.listedSergeants = rankStats.listed.sergeants;
    unit.listedSoldiers = rankStats.listed.soldiers;
    unit.listed = rankStats.listed.total;
    unit.availableOfficers = rankStats.available.officers;
    unit.availableSergeants = rankStats.available.sergeants;
    unit.availableSoldiers = rankStats.available.soldiers;
    unit.available = rankStats.available.total;
    unit.absentOfficers = rankStats.absent.officers;
    unit.absentSergeants = rankStats.absent.sergeants;
    unit.absentSoldiers = rankStats.absent.soldiers;
    unit.absent = rankStats.absent.total;
    unit.businessTrip = computeBchsUnitBusinessTripCount(novaPeople, unit.unit);

    const absenceStats = computeBchsUnitAbsenceCategoryStats(
      novaPeople,
      unit.unit,
    );
    unit.training = absenceStats.training;
    unit.hospitalWounded = absenceStats.hospitalWounded;
    unit.hospitalIllness = absenceStats.hospitalIllness;
    unit.vacation = absenceStats.vacation;
    unit.awol = absenceStats.awol;
    unit.missing = absenceStats.missing;
    unit.killed = absenceStats.killed;
    unit.medWounded = absenceStats.medWounded;
    unit.medIllness = absenceStats.medIllness;
  });

  const detailRows = Array.from(units.values())
    .sort(
      (left, right) =>
        getBchsGeneralListUnitOrder(left.unit) -
          getBchsGeneralListUnitOrder(right.unit) ||
        left.unit.localeCompare(right.unit, "uk"),
    )
    .map((item, index) => bchsGeneralListUnitToRow(item, 12 + index));
  const totalAccumulator = emptyBchsGeneralListUnit("Усього");
  Array.from(units.values()).forEach((item) =>
    addBchsGeneralListTotals(totalAccumulator, item),
  );
  totalAccumulator.unassignedNewcomers =
    computeBchsTotalUnassignedNewcomers(allPeople);
  const total = bchsGeneralListUnitToRow(totalAccumulator, 11);
  const comparisonRows = [total, ...detailRows];

  const baseAnalytics = {
    reportDate: getBchsReportDate(sheet),
    total,
    rows: comparisonRows,
    comparisonRows,
    table: undefined,
    detachedDestinations: [],
    attachedSources: [],
    absenceReasons: buildBchsAbsenceReasons(total),
    dataIssues: buildBchsGeneralListDataIssues(allPeople),
  } satisfies BchsAnalyticsSnapshot;

  // Bake away/attached into general-list analytics (same people source as Excel Аркуш2).
  return applyBchsPersonnelDerivedColumns(baseAnalytics, allPeople);
};

export const isLegacyBchsAnalyticsSheet = (
  sheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
) => Boolean(sheet?.columnIndexes.includes(columnLetterToIndex("BL")));

export const buildBchsAnalyticsFromAppendix = (
  supplement: BchsSupplementSnapshot,
): BchsAnalyticsSnapshot => {
  const total = buildBchsSupplementComparisonRow(supplement.total);
  const rows = supplement.rows.map(buildBchsSupplementComparisonRow);

  return {
    reportDate: supplement.reportDate,
    total,
    rows,
    comparisonRows: rows,
    table: undefined,
    detachedDestinations: [],
    attachedSources: [],
    absenceReasons: supplement.absenceReasons,
    supplement,
  } satisfies BchsAnalyticsSnapshot;
};

export const buildBchsAnalyticsFromPersonnelBzvp = (
  supplement: BchsSupplementSnapshot,
): BchsAnalyticsSnapshot => {
  const total = buildBchsSupplementComparisonRow(supplement.total);
  const comparisonRows = supplement.totals.map(buildBchsSupplementComparisonRow);

  return {
    reportDate: supplement.reportDate,
    total,
    rows: comparisonRows,
    comparisonRows,
    table: undefined,
    detachedDestinations: [],
    attachedSources: [],
    absenceReasons: [],
    supplement,
  } satisfies BchsAnalyticsSnapshot;
};

export const buildBchsAnalyticsFromWorkbook = (
  workbook: ExcelWorkbookSnapshot | null,
  fallbackSheet: ExcelWorkbookSnapshot["sheets"][number] | undefined,
): BchsAnalyticsSnapshot => {
  if (!workbook) return buildBchsAnalytics(undefined);

  const legacySheet =
    workbook.sheets.find(
      (sheet) =>
        sheet.sheetName.trim().toLowerCase() === "аркуш1" &&
        isLegacyBchsAnalyticsSheet(sheet),
    ) ?? workbook.sheets.find(isLegacyBchsAnalyticsSheet);
  const appendixSupplement = workbook.sheets
    .map(buildBchsAppendixSupplement)
    .find(Boolean);
  const personnelSupplement = workbook.sheets
    .map(buildBchsPersonnelBzvpSupplement)
    .find(Boolean);

  if (legacySheet) {
    const personnelSheet =
      workbook.sheets.find(isBchsPersonnelGeneralListSheet) ??
      workbook.sheets.find(
        (sheet) => sheet.sheetName.trim().toLowerCase() === "аркуш2",
      ) ??
      workbook.sheets.find(
        (sheet) =>
          sheet !== legacySheet &&
          sheet.columnIndexes.includes(28) &&
          sheet.rawRows.length > 20,
      );
    const baseAnalytics = {
      ...buildBchsAnalytics(legacySheet),
      supplement: appendixSupplement ?? personnelSupplement ?? undefined,
    } satisfies BchsAnalyticsSnapshot;

    return applyBchsPersonnelDerivedColumns(
      baseAnalytics,
      extractBchsNovaPeopleFromSheet(personnelSheet),
    );
  }

  if (appendixSupplement) return buildBchsAnalyticsFromAppendix(appendixSupplement);
  if (personnelSupplement)
    return buildBchsAnalyticsFromPersonnelBzvp(personnelSupplement);
  const generalListSheet = workbook.sheets.find(isBchsPersonnelGeneralListSheet);
  const generalListAnalytics = buildBchsAnalyticsFromGeneralList(generalListSheet);
  if (generalListAnalytics) return generalListAnalytics;

  return buildBchsAnalytics(fallbackSheet);
};
export const parseBchsImportAnalytics = (item: BackendEjournalImport | undefined) => {
  if (!item?.notes) return null;

  try {
    const parsed = JSON.parse(item.notes) as {
      source?: string;
      analytics?:
        | Partial<typeof emptyBchsRow>
        | {
            total?: Partial<typeof emptyBchsRow>;
            rows?: Partial<typeof emptyBchsRow>[];
            table?: {
              columns?: BchsAnalyticsTableColumn[];
              rows?: BchsAnalyticsTableRow[];
            };
            detachedDestinations?: AnalyticsMetric[];
            attachedSources?: AnalyticsMetric[];
            comparisonRows?: Partial<BchsComparisonRow>[];
            reportDate?: string;
            absenceReasons?: AnalyticsMetric[];
            supplement?: BchsSupplementSnapshot;
          };
    };
    if (parsed.source !== "BCHS" || !parsed.analytics) return null;

    const maybeAnalytics = parsed.analytics;
    const total =
      "total" in maybeAnalytics ? maybeAnalytics.total : maybeAnalytics;
    const rows = "rows" in maybeAnalytics ? maybeAnalytics.rows : undefined;
    if (!total) return null;
    const totalRow = total as Partial<BchsComparisonRow>;
    const analyticsRows = rows as Partial<BchsComparisonRow>[] | undefined;

    const table =
      "table" in maybeAnalytics && maybeAnalytics.table
        ? {
            columns: maybeAnalytics.table.columns ?? [],
            rows: maybeAnalytics.table.rows ?? [],
          }
        : undefined;
    const tableComparisonRows = buildBchsComparisonRowsFromTable(table);
    const normalizedRows =
      tableComparisonRows.length > 0
        ? tableComparisonRows
        : (analyticsRows?.map(createBchsComparisonRow) ?? [
            createBchsComparisonRow(totalRow),
          ]);
    const normalizedTotal = normalizedRows[0] ?? createBchsComparisonRow(totalRow);
    const analyticsBundle =
      "total" in maybeAnalytics
        ? (maybeAnalytics as {
            detachedDestinations?: AnalyticsMetric[];
            attachedSources?: AnalyticsMetric[];
            comparisonRows?: Partial<BchsComparisonRow>[];
            reportDate?: string;
            absenceReasons?: AnalyticsMetric[];
            supplement?: BchsSupplementSnapshot;
          })
        : null;

    return {
      reportDate: analyticsBundle?.reportDate ?? "",
      total: normalizedTotal,
      rows: normalizedRows,
      comparisonRows:
        tableComparisonRows.length > 0
          ? tableComparisonRows
          : analyticsBundle?.comparisonRows
            ? analyticsBundle.comparisonRows.map(createBchsComparisonRow)
            : normalizedRows,
      table,
      detachedDestinations:
        analyticsBundle?.detachedDestinations ??
        buildBchsDetachedDestinationsFromTable(table),
      attachedSources:
        analyticsBundle?.attachedSources ??
        buildBchsDetachedDestinationsFromMaps(
          (table?.rows ?? []).map((row) =>
            parseBchsDestinationText(row.values.AT),
          ),
        ),
      absenceReasons:
        analyticsBundle?.absenceReasons ??
        buildBchsAbsenceReasons(normalizedTotal),
      supplement: analyticsBundle?.supplement,
    } satisfies BchsAnalyticsSnapshot;
  } catch {
    return null;
  }
};
