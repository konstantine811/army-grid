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
  BchsSupplementRow,
  BchsSupplementSnapshot,
  BchsUnitAttachedStats,
  BchsUnitAwayStats,
} from "./bchsTypes";

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
 * AP=Офіцери, AQ=Сержанти, AR=Солдати, AS=Всього, AT=Звідки прикомандировані
 *
 * Counted from Аркуш2:
 * - B = ПРИКОМАНДИРОВАНІ → Інженерно-саперне відділення
 * - B = БРЕЗ → Відділення радіоелектронної боротьби
 * - U = "В строю" або "Новоприбулий"
 * - rank from I, fallback to M (Звання)
 * - AT sources: for БРЕЗ → "БРЕЗ"; for ПРИКОМ. → A (№), known units kept, rest → "інші"
 */

export const normalizeBchsText = (value: unknown) =>
  valueToDisplay(value as CellValue)
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase()
    .replace(/ё/g, "е");

export const BCHS_COMMAND_ROSTER_ALIASES = [
  "штаб",
  "група безпілотних систем",
] as const;

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
    normalized === "командування"
  )
    return "Командування";
  return label;
};

export const isBchsDetachedStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  // Require "за межі" so "в межах ПБ" is not counted.
  return normalized.includes("відком") && normalized.includes("за межі");
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
  const roster = normalizeBchsText(normalizeBchsRosterUnitLabel(rosterUnit));
  if (!unit || !roster) return false;

  if (isBchsCommandUnit(unit)) {
    return BCHS_COMMAND_ROSTER_ALIASES.some((alias) => roster === alias);
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
  const detached = people.filter(
    (person) =>
      normalizeBchsText(person.battalion) === "нова" &&
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

  people.forEach((person) => {
    if (normalizeBchsText(person.battalion) !== "нова") return;
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
  const rankCol = columnIndex(8);
  const rankTitleCol = columnIndex(12);
  const fullNameCol = columnIndex(13);
  const statusCol = columnIndex(20);
  const roleTypeCol = columnIndex(21);
  const combatReadinessCol = columnIndex(22);
  const bzvpStatusCol = columnIndex(23);
  const destinationCol = columnIndex(28);
  const medicalPlaceCol = columnIndex(30);
  const medicalNoteCol = columnIndex(31);

  return sheet.rawRows
    .slice(1)
    .map((row) => ({
      battalion: valueToDisplay(row[battalionCol] as CellValue),
      rosterUnit: valueToDisplay(row[rosterCol] as CellValue),
      rankCategory: valueToDisplay(row[rankCol] as CellValue),
      rankTitle: valueToDisplay(row[rankTitleCol] as CellValue),
      fullName: valueToDisplay(row[fullNameCol] as CellValue),
      status: valueToDisplay(row[statusCol] as CellValue),
      roleType: valueToDisplay(row[roleTypeCol] as CellValue),
      combatReadiness: valueToDisplay(row[combatReadinessCol] as CellValue),
      bzvpStatus: valueToDisplay(row[bzvpStatusCol] as CellValue),
      destination: valueToDisplay(row[destinationCol] as CellValue),
      medicalPlace: valueToDisplay(row[medicalPlaceCol] as CellValue),
      medicalNote: valueToDisplay(row[medicalNoteCol] as CellValue),
    }))
    .filter(
      (person) =>
        person.rosterUnit ||
        person.status ||
        person.rankCategory ||
        person.rankTitle ||
        person.fullName ||
        person.roleType ||
        person.combatReadiness ||
        person.bzvpStatus ||
        person.destination ||
        person.medicalPlace ||
        person.medicalNote,
    );
};

export const extractBchsAwayPeopleFromDbRows = (
  rows: Array<Record<string, unknown>>,
): BchsPersonnelAwayPerson[] =>
  rows.map((row) => ({
    battalion: valueToDisplay(
      (row.column_1 ?? row["№"] ?? row.battalion ?? "") as CellValue,
    ),
    rosterUnit: valueToDisplay(
      (row.підрозділ ?? row.rosterUnit ?? row.B ?? "") as CellValue,
    ),
    rankCategory: valueToDisplay(
      (row.column_9 ?? row.rankCategory ?? row.I ?? "") as CellValue,
    ),
    rankTitle: valueToDisplay(
      (row.звання ?? row.rankTitle ?? row.M ?? "") as CellValue,
    ),
    fullName: valueToDisplay(
      (row.піб ?? row.fullName ?? row.N ?? "") as CellValue,
    ),
    status: valueToDisplay(
      (row.статус ?? row.status ?? row.U ?? "") as CellValue,
    ),
    roleType: valueToDisplay(
      (row.тип_в_с ?? row.roleType ?? row.V ?? "") as CellValue,
    ),
    combatReadiness: valueToDisplay(
      (row.статус_бг ?? row.combatReadiness ?? row.W ?? "") as CellValue,
    ),
    bzvpStatus: valueToDisplay(
      (row.бзвп_брез ?? row.bzvpStatus ?? row.X ?? "") as CellValue,
    ),
    destination: valueToDisplay(
      (row.в_якому_підрозділі ??
        row.externalUnit ??
        row.AC ??
        "") as CellValue,
    ),
    medicalPlace: valueToDisplay(
      (row.місце_перебування ?? row.medicalPlace ?? row.AE ?? "") as CellValue,
    ),
    medicalNote: valueToDisplay(
      (row.примітки ?? row.medicalNote ?? row.AF ?? "") as CellValue,
    ),
  }));

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
  return normalized === "в строю" || normalized.includes("новоприбулий");
};

export const isBchsPrikomaniRoster = (rosterUnit: string) =>
  normalizeBchsText(rosterUnit).includes("приком");

export const isBchsBrezRoster = (rosterUnit: string) =>
  normalizeBchsText(rosterUnit) === "брез";

export const isBchsEngineerUnit = (unitName: string) =>
  /інженер|сапер/i.test(unitName);

export const isBchsRebUnit = (unitName: string) => {
  const unit = normalizeBchsText(unitName);
  return unit.includes("радіоелектрон") || unit.includes("відділення реб");
};

export const matchesBchsAttachedUnit = (arkush1Unit: string, rosterUnit: string) => {
  if (isBchsEngineerUnit(arkush1Unit)) return isBchsPrikomaniRoster(rosterUnit);
  if (isBchsRebUnit(arkush1Unit)) return isBchsBrezRoster(rosterUnit);
  return false;
};

export const classifyBchsAttachedRank = (
  rankCategory: string,
  rankTitle: string,
): "officers" | "sergeants" | "soldiers" | null => {
  // Attached rows often have empty I; Excel ranks them by M (Звання).
  const title = normalizeBchsText(rankTitle);
  if (title) {
    if (
      /лейтенант|капітан|майор|підполковник|полковник|генерал/.test(title)
    )
      return "officers";
    if (/сержан|старшин/.test(title)) return "sergeants";
    return "soldiers";
  }

  const category = normalizeBchsText(rankCategory);
  if (category === BCHS_RANK_OFFICER) return "officers";
  if (category === BCHS_RANK_SERGEANT) return "sergeants";
  if (category === BCHS_RANK_SOLDIER) return "soldiers";
  return null;
};

export const normalizeBchsAttachedSourceLabel = (
  battalion: string,
  rosterUnit: string,
) => {
  if (isBchsBrezRoster(rosterUnit)) return "БРЕЗ";

  const raw = battalion.replace(/\s+/g, " ").trim();
  const normalized = normalizeBchsText(raw);
  if (!normalized) return "інші";

  if (normalized.includes("шквал")) return "ШКВАЛ";
  if (normalized === "210") return "210";
  if (normalized === "155" || /^155\s*омбр/.test(normalized)) return "155 ОМБр";
  if (normalized.includes("омбр")) return raw;

  return "інші";
};

export const computeBchsUnitAttachedStats = (
  people: BchsPersonnelAwayPerson[],
  unitName: string,
): BchsUnitAttachedStats => {
  if (isBchsTotalUnit(unitName)) return emptyBchsUnitAttachedStats();

  const stats = emptyBchsUnitAttachedStats();

  people.forEach((person) => {
    if (!isBchsAttachedPresentStatus(person.status)) return;
    if (!matchesBchsAttachedUnit(unitName, person.rosterUnit)) return;

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
    );
    if (source) {
      stats.sources.set(source, (stats.sources.get(source) ?? 0) + 1);
    }
  });

  stats.total = stats.officers + stats.sergeants + stats.soldiers;
  stats.sourcesText = formatBchsDestinationText(stats.sources);
  return stats;
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
): BchsAnalyticsSnapshot => {
  const derived = applyBchsAttachedFromPersonnel(
    applyBchsAwayFromPersonnel(analytics, people),
    people,
  );
  const dataIssues = people.length > 0
    ? buildBchsGeneralListDataIssues(people)
    : [];

  return {
    ...derived,
    dataIssues: dataIssues.length > 0
      ? dataIssues
      : derived.dataIssues,
  };
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
  const inRanksActually = bchsToNumber(base.inRanksActually);
  const staff = bchsToNumber(base.staff);
  const listed = bchsToNumber(base.listed);
  const absent = bchsToNumber(base.absent);
  const available = bchsToNumber(base.available);

  return {
    ...base,
    staffOfficers: bchsToNumber(row.staffOfficers),
    staffSergeants: bchsToNumber(row.staffSergeants),
    staffSoldiers: bchsToNumber(row.staffSoldiers),
    listedOfficers: bchsToNumber(row.listedOfficers),
    listedSergeants: bchsToNumber(row.listedSergeants),
    listedSoldiers: bchsToNumber(row.listedSoldiers),
    availableOfficers: bchsToNumber(row.availableOfficers),
    availableSergeants: bchsToNumber(row.availableSergeants),
    availableSoldiers: bchsToNumber(row.availableSoldiers),
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
    actualOfficers: bchsToNumber(row.actualOfficers),
    actualSergeants: bchsToNumber(row.actualSergeants),
    actualSoldiers: bchsToNumber(row.actualSoldiers),
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

const isBchsGeneralListPresentStatus = (status: string) => {
  const normalized = normalizeBchsText(status);
  return normalized === "в строю" || normalized.includes("новоприбулий");
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
      };
      const issues: BchsDataIssue[] = [];
      const statusReason = explainBchsGeneralListStatusIssue(person);
      if (statusReason) issues.push({ ...base, reason: statusReason });
      if (hasBchsUnknownDetachedDestination(person)) {
        issues.push({
          ...base,
          destination:
            normalizeBchsDestinationLabel(person.destination) || "(порожньо)",
          reason: "невідоме місце відкомандирування",
        });
      }
      return issues;
    })
    .filter((issue): issue is BchsDataIssue => Boolean(issue));

const isBchsWoundedMedicalNote = (value: string) => {
  const normalized = normalizeBchsText(value);
  return /поран|ранен|травм/.test(normalized);
};

const isBchsIllnessMedicalNote = (value: string) => {
  const normalized = normalizeBchsText(value);
  return /хвор|болез|захвор/.test(normalized);
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
  "Взвод зв'язку",
  "Взвод логістично-евакуаційних безпілотних наземних систем",
  "Взвод матеріально-технічного забезпечення",
  "Взвод перехоплювачів безпілотних літальних апаратів",
  "Взвод протитанкових ракетних комплексів",
  "Відділення радіоелектронної боротьби",
  "Гранатометний взвод",
  "Група безпілотних систем",
  "Зенітно-ракетний взвод",
  "Інженерно-саперне відділення",
  "Кулеметний взвод",
  "Мінометний взвод",
  "Розвідувальний взвод",
  "Рота безпілотних авіаційних комплексів",
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

  const people = extractBchsAwayPeopleFromSheet(sheet);
  const units = new Map<string, BchsGeneralListUnitAccumulator>();

  const getUnit = (unitName: string) => {
    const unit = normalizeBchsRosterUnitLabel(unitName);
    const current = units.get(unit) ?? emptyBchsGeneralListUnit(unit);
    units.set(unit, current);
    return current;
  };

  people
    .filter((person) => normalizeBchsText(person.battalion) === "нова")
    .forEach((person) => {
      const unit = getUnit(person.rosterUnit);
      const status = normalizeBchsText(person.status);
      const fullName = person.fullName.trim();
      const rank = normalizeBchsText(person.rankCategory);
      const roleType = normalizeBchsText(person.roleType);
      const combatReadiness = normalizeBchsText(person.combatReadiness);
      const bzvpStatus = normalizeBchsText(person.bzvpStatus);
      const medicalPlace = normalizeBchsText(person.medicalPlace);
      const medicalText = `${person.medicalPlace} ${person.medicalNote}`;

      unit.staff += 1;
      if (rank === BCHS_RANK_OFFICER) unit.staffOfficers += 1;
      else if (rank === BCHS_RANK_SERGEANT) unit.staffSergeants += 1;
      else if (rank === BCHS_RANK_SOLDIER) unit.staffSoldiers += 1;

      if (fullName) {
        unit.listed += 1;
        if (rank === BCHS_RANK_OFFICER) unit.listedOfficers += 1;
        else if (rank === BCHS_RANK_SERGEANT) unit.listedSergeants += 1;
        else if (rank === BCHS_RANK_SOLDIER) unit.listedSoldiers += 1;
      }

      if (isBchsGeneralListPresentStatus(person.status)) {
        unit.available += 1;
        unit.inRanksActually += 1;
        if (rank === BCHS_RANK_OFFICER) {
          unit.availableOfficers += 1;
          unit.actualOfficers += 1;
        } else if (rank === BCHS_RANK_SERGEANT) {
          unit.availableSergeants += 1;
          unit.actualSergeants += 1;
        } else if (rank === BCHS_RANK_SOLDIER) {
          unit.availableSoldiers += 1;
          unit.actualSoldiers += 1;
        }
      } else if (fullName) {
        unit.absent += 1;
      }

      if (status.includes("відряд") || isBchsDetachedStatus(person.status))
        unit.businessTrip += 1;
      if (status.includes("навч")) unit.training += 1;
      if (status.includes("ліку") || medicalPlace.includes("шпитал")) {
        const wounded = isBchsWoundedMedicalNote(medicalText);
        const illness =
          isBchsIllnessMedicalNote(medicalText) || (!wounded && fullName);

        if (medicalPlace.includes("мед")) {
          if (wounded) unit.medWounded += 1;
          else if (illness) unit.medIllness += 1;
        } else {
          if (wounded) unit.hospitalWounded += 1;
          else if (illness) unit.hospitalIllness += 1;
        }
      }
      if (status.includes("відпуст")) unit.vacation += 1;
      if (status.includes("сзч")) unit.awol += 1;
      if (status.includes("зник")) unit.missing += 1;
      if (status.includes("загиб")) unit.killed += 1;
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
  totalAccumulator.unassignedNewcomers = people.filter(
    (person) =>
      normalizeBchsText(person.status).includes("новоприбулий") &&
      normalizeBchsText(person.battalion) !== "нова",
  ).length;
  const total = bchsGeneralListUnitToRow(totalAccumulator, 11);
  const comparisonRows = [total, ...detailRows];

  return {
    reportDate: getBchsReportDate(sheet),
    total,
    rows: comparisonRows,
    comparisonRows,
    table: undefined,
    detachedDestinations: [],
    attachedSources: [],
    absenceReasons: buildBchsAbsenceReasons(total),
    dataIssues: buildBchsGeneralListDataIssues(people),
  } satisfies BchsAnalyticsSnapshot;
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
      extractBchsAwayPeopleFromSheet(personnelSheet),
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
