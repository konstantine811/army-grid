import type {
  CellValue,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  extractBchsAwayPeopleFromSheet,
  filterBchsNovaPeople,
  isBchsAvailableForAbsentFormula,
  isBchsListedWithRankTitle,
  isBchsPersonnelGeneralListSheet,
} from "../bchs/bchsCalc";
import {
  classifyRankGroup,
  looksLikePersonName,
  normalizeLooseText,
  normalizePersonName,
  shortPersonName,
} from "./socPassportFields";
import type { RankGroup } from "./socPassportTypes";

export type HousingIdpPerson = {
  excelRow: number;
  unit: string;
  rank: string;
  rankGroup: RankGroup;
  fullName: string;
  address: string;
  note: string;
};

export type HousingIdpMatchRow = HousingIdpPerson & {
  inMorning: boolean;
  morningStatus: string;
  /** Є в ранковому списку (лишився в обліку). */
  remaining: boolean;
  /** У ранковому і в наявності (в строю / відком.). */
  remainingPresent: boolean;
};

export type HousingIdpStatsResult = {
  fileName: string;
  sheetName: string;
  people: HousingIdpMatchRow[];
  totals: {
    inFile: number;
    remaining: number;
    remainingPresent: number;
    leftNotInMorning: number;
    byRankRemaining: {
      officer: number;
      sergeant: number;
      soldier: number;
    };
  };
  remainingNames: string[];
  leftNames: string[];
};

const cellText = (value: CellValue | undefined): string => {
  if (value == null) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
};

const normalizeHeader = (value: CellValue | undefined) =>
  normalizeLooseText(cellText(value)).replace(/[.]/g, " ");

const columnIndex = (
  headerRow: CellValue[],
  parts: string[],
  exclude: string[] = [],
) =>
  headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    if (!header) return false;
    if (exclude.some((part) => header.includes(part))) return false;
    return parts.every((part) => header.includes(part));
  });

const findHousingSheet = (workbook: ExcelWorkbookSnapshot) => {
  for (const sheet of workbook.sheets) {
    for (let i = 0; i < Math.min(12, sheet.rawRows.length); i += 1) {
      const headers = (sheet.rawRows[i] ?? []).map((cell) =>
        normalizeHeader(cell),
      );
      const hasName =
        headers.some((h) => h === "піб") ||
        headers.some((h) => h.includes("прізвище"));
      const hasAddress = headers.some(
        (h) => h.includes("адрес") || h.includes("проживан") || h.includes("тот"),
      );
      if (hasName && (hasAddress || headers.some((h) => h.includes("зван")))) {
        return { sheet, headerIndex: i };
      }
    }
  }
  return null;
};

export const parseHousingIdpWorkbook = (
  workbook: ExcelWorkbookSnapshot,
): HousingIdpPerson[] => {
  const found = findHousingSheet(workbook);
  if (!found) {
    throw new Error(
      'У файлі «Додаток про житло» не знайдено таблицю з колонкою ПІБ.',
    );
  }
  const { sheet, headerIndex } = found;
  const headerRow = sheet.rawRows[headerIndex] ?? [];
  const nameCol =
    columnIndex(headerRow, ["піб"]) >= 0
      ? columnIndex(headerRow, ["піб"])
      : columnIndex(headerRow, ["прізвище"]);
  const rankCol = columnIndex(headerRow, ["зван"]);
  const unitCol =
    columnIndex(headerRow, ["найменуван"]) >= 0
      ? columnIndex(headerRow, ["найменуван"])
      : columnIndex(headerRow, ["в", "ч"]);
  const addressCol =
    columnIndex(headerRow, ["адрес"]) >= 0
      ? columnIndex(headerRow, ["адрес"])
      : columnIndex(headerRow, ["проживан"]);
  const noteCol = columnIndex(headerRow, ["приміт"]);

  if (nameCol < 0) {
    throw new Error('Не знайдено колонку «ПІБ» у додатку про житло.');
  }

  const people: HousingIdpPerson[] = [];
  const seen = new Set<string>();
  for (let i = headerIndex + 1; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const fullName = cellText(row[nameCol]);
    if (!looksLikePersonName(fullName)) continue;
    const key = normalizePersonName(fullName);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    const rank = cellText(row[rankCol >= 0 ? rankCol : 2]);
    people.push({
      excelRow: i + 1,
      unit: cellText(row[unitCol >= 0 ? unitCol : 1]),
      rank,
      rankGroup: classifyRankGroup(rank),
      fullName,
      address: cellText(row[addressCol >= 0 ? addressCol : 4]),
      note: cellText(row[noteCol >= 0 ? noteCol : 5]),
    });
  }
  return people;
};

const buildMorningNameIndex = (morning: ExcelWorkbookSnapshot) => {
  const sheet =
    morning.sheets.find(isBchsPersonnelGeneralListSheet) ??
    morning.sheets.find((candidate) =>
      /ос.*загальн|загальн.*спис/i.test(candidate.sheetName),
    );
  if (!sheet) {
    throw new Error(
      "У ранковому звіті не знайдено аркуш «1.ОС Загальний список».",
    );
  }
  const byName = new Map<
    string,
    { status: string; rank: string; fullName: string }
  >();
  const nova = filterBchsNovaPeople(extractBchsAwayPeopleFromSheet(sheet));
  for (const person of nova) {
    if (!isBchsListedWithRankTitle(person)) continue;
    const fullName = person.fullName.trim();
    if (!looksLikePersonName(fullName)) continue;
    const key = normalizePersonName(fullName);
    if (!key || byName.has(key)) continue;
    byName.set(key, {
      status: person.status,
      rank: person.rankTitle || person.rankCategory,
      fullName,
    });
    const short = shortPersonName(fullName);
    if (short && !byName.has(short)) {
      byName.set(short, {
        status: person.status,
        rank: person.rankTitle || person.rankCategory,
        fullName,
      });
    }
  }
  return { sheetName: sheet.sheetName, byName };
};

export const buildHousingIdpStats = (
  housing: ExcelWorkbookSnapshot,
  morning: ExcelWorkbookSnapshot,
): HousingIdpStatsResult => {
  const parsed = parseHousingIdpWorkbook(housing);
  const { sheetName, byName } = buildMorningNameIndex(morning);

  const byRankRemaining = { officer: 0, sergeant: 0, soldier: 0 };
  const remainingNames: string[] = [];
  const leftNames: string[] = [];

  const people: HousingIdpMatchRow[] = parsed.map((person) => {
    const key = normalizePersonName(person.fullName);
    const morningHit =
      byName.get(key) || byName.get(shortPersonName(person.fullName));
    const inMorning = Boolean(morningHit);
    const morningStatus = morningHit?.status ?? "";
    const remaining = inMorning;
    const remainingPresent =
      remaining && isBchsAvailableForAbsentFormula(morningStatus);
    if (remaining) {
      byRankRemaining[person.rankGroup] += 1;
      remainingNames.push(person.fullName);
    } else {
      leftNames.push(person.fullName);
    }
    return {
      ...person,
      inMorning,
      morningStatus,
      remaining,
      remainingPresent,
    };
  });

  return {
    fileName: housing.fileName,
    sheetName,
    people,
    totals: {
      inFile: people.length,
      remaining: remainingNames.length,
      remainingPresent: people.filter((row) => row.remainingPresent).length,
      leftNotInMorning: leftNames.length,
      byRankRemaining,
    },
    remainingNames,
    leftNames,
  };
};

/** Імена ВПО з додатку, які лишились у ранковому (для isIdp у соцпакеті). */
export const housingIdpRemainingNameSet = (
  stats: HousingIdpStatsResult | null | undefined,
): Set<string> => {
  const set = new Set<string>();
  if (!stats) return set;
  for (const name of stats.remainingNames) {
    const key = normalizePersonName(name);
    if (key) set.add(key);
    const short = shortPersonName(name);
    if (short) set.add(short);
  }
  return set;
};
