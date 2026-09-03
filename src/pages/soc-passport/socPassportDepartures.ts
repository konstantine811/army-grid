import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { tryParseExcelSerialDate, formatUkDate } from "../../shared/format";
import {
  extractBchsAwayPeopleFromSheet,
  filterBchsNovaPeople,
  isBchsBrezAttachedPerson,
  isBchsBrezRoster,
  isBchsListedWithRankTitle,
  isBchsPersonnelGeneralListSheet,
} from "../bchs/bchsCalc";
import {
  parsePbArchive,
  parsePbMovements,
  parsePbShPeople,
} from "../ejournal/ejoosSyncPlan";
import {
  ARRIVAL_SOURCE_LABELS,
  classifyOosArrivedFrom,
  classifyRankGroup,
  combineMorningBrezFields,
  isBrezMarkerText,
  looksLikePersonName,
  normalizeLooseText,
  normalizePersonName,
} from "./socPassportFields";
import type { ArrivalSource, RankGroup } from "./socPassportTypes";

/** Категорії «Вибули» для соцпакета (звільнення). Переведені — окремо. */
export type DepartureCategory =
  | "discharged_unit"
  | "demobilization"
  | "health"
  | "unsuitable_court"
  | "other"
  | "transfer";

export type DeparturePerson = {
  excelRow: number;
  rank: string;
  rankGroup: RankGroup;
  rankGroupLabel: string;
  fullName: string;
  personId: string;
  positionIndex: string;
  ground: string;
  type: string;
  destination: string;
  orderNumber: string;
  orderDate: string;
  excludeDate: string;
  category: DepartureCategory;
  categoryLabel: string;
  matchNote: string;
};

export type RankBreakdown = {
  officer: number;
  sergeant: number;
  soldier: number;
};

export type DepartureSummaryRow = {
  category: DepartureCategory;
  label: string;
  count: number;
  byRank: RankBreakdown;
};

export type EnrollmentArrivalPerson = {
  excelRow: number;
  rank: string;
  fullName: string;
  personId: string;
  positionIndex: string;
  arrivedFrom: string;
  enrollDate: string;
  rankGroup: RankGroup;
  rankGroupLabel: string;
  arrivalSource: ArrivalSource;
  arrivalSourceLabel: string;
};

export type ArrivalsMonthResult = {
  sourceSheet: string;
  month: number;
  year: number;
  monthLabel: string;
  people: EnrollmentArrivalPerson[];
  byRank: {
    officer: number;
    sergeant: number;
    soldier: number;
  };
  bySource: Record<ArrivalSource, number>;
  sourceSummary: Array<{
    source: ArrivalSource;
    label: string;
    count: number;
    byRank: RankBreakdown;
  }>;
  total: number;
};

export type DeparturesResult = {
  sourceSheet: string;
  people: DeparturePerson[];
  summary: DepartureSummaryRow[];
  totals: {
    all: number;
    discharges: number;
    transfers: number;
    byRank: RankBreakdown;
    dischargesByRank: RankBreakdown;
    transfersByRank: RankBreakdown;
  };
  /** Фільтр: дата виключення / наказу ≥ цієї дати (включно). */
  periodFrom: string | null;
  periodFromLabel: string | null;
  /** Усі рядки аркуша до фільтра за датою. */
  totalUnfiltered: number;
  skippedNoDate: number;
  skippedBeforePeriod: number;
  /** Прибули за датою зарахування до списків частини (ООС), місяць серпень. */
  arrivalsAugust: ArrivalsMonthResult | null;
  /**
   * Штатка (ранковий): БРЕЗ з ранку; решта — «Звідки прибув» з ООС
   * (порожнє = ТЦК). Розбивка за званням.
   */
  arrivalsFromMorning: ArrivalsMonthResult | null;
  /** Прибули з 1ПБ · Рух тип ПРИБУВ за датою наказу (той самий місяць). */
  arrivalsAugustPb: ArrivalsMonthResult | null;
  /** Виведені у розпорядження з 1ПБ archive. */
  dispositionFromArchive: DispositionArchiveResult | null;
  /** СЗЧ з 1ПБ · Рух. */
  szchFromRuh: SzchRuhResult | null;
  /** Втрати (загиблі / безвісти / у бою / інші) з 1ПБ Рух (+ archive). */
  combatLossesFromPb: CombatLossesResult | null;
};

export type CombatLossReason =
  | "killed"
  | "missing"
  | "inCombat"
  | "otherCircumstances";

export type CombatLossPerson = {
  excelRow: number;
  rank: string;
  rankGroup: RankGroup;
  rankGroupLabel: string;
  fullName: string;
  personId: string;
  sourceSheet: string;
  typeOrAbsence: string;
  status: string;
  placeOrNote: string;
  orderDate: string;
  reason: CombatLossReason;
  reasonLabel: string;
  matchNote: string;
};

export type CombatLossSummaryRow = {
  reason: CombatLossReason;
  label: string;
  count: number;
  byRank: RankBreakdown;
};

export type CombatLossesResult = {
  sourceSheet: string;
  people: CombatLossPerson[];
  summary: CombatLossSummaryRow[];
  totals: {
    all: number;
    byRank: RankBreakdown;
  };
  periodFrom: string | null;
  periodFromLabel: string | null;
  skippedBeforePeriod: number;
  skippedNoDate: number;
};

export type SzchPerson = {
  excelRow: number;
  rank: string;
  rankGroup: RankGroup;
  rankGroupLabel: string;
  fullName: string;
  personId: string;
  movementNumber: string;
  type: string;
  status: string;
  destination: string;
  note: string;
  orderDate: string;
  matchNote: string;
};

export type SzchRuhResult = {
  sourceSheet: string;
  people: SzchPerson[];
  totals: {
    all: number;
    byRank: RankBreakdown;
  };
  periodFrom: string | null;
  periodFromLabel: string | null;
  totalMovements: number;
  skippedBeforePeriod: number;
  skippedAfterPeriod: number;
  skippedNoDate: number;
};

export type DispositionReason =
  | "treatment"
  | "organizational"
  | "other";

export type DispositionPerson = {
  excelRow: number;
  rank: string;
  rankGroup: RankGroup;
  rankGroupLabel: string;
  fullName: string;
  personId: string;
  absenceType: string;
  place: string;
  departDate: string;
  orderDate: string;
  returnDate: string;
  reason: DispositionReason;
  reasonLabel: string;
  matchNote: string;
};

export type DispositionSummaryRow = {
  reason: DispositionReason;
  label: string;
  count: number;
  byRank: RankBreakdown;
};

export type DispositionArchiveResult = {
  sourceSheet: string;
  people: DispositionPerson[];
  summary: DispositionSummaryRow[];
  totals: {
    all: number;
    byRank: RankBreakdown;
  };
  periodFrom: string | null;
  periodFromLabel: string | null;
  totalArchiveRows: number;
  skippedNotDisposition: number;
  skippedBeforePeriod: number;
  skippedNoDate: number;
  skippedReturned: number;
};

const cellText = (value: CellValue | undefined): string => {
  if (value == null || value === "") return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return formatUkDate(value);
  }
  const fromSerial = tryParseExcelSerialDate(value);
  if (fromSerial) return formatUkDate(fromSerial);
  if (typeof value === "object") {
    const maybe = value as { text?: () => string; value?: () => unknown };
    if (typeof maybe.text === "function") return cellText(maybe.text() as CellValue);
    if (typeof maybe.value === "function") return cellText(maybe.value() as CellValue);
  }
  return String(value).replace(/\s+/g, " ").trim();
};

const findSheet = (workbook: ExcelWorkbookSnapshot, patterns: RegExp[]) => {
  for (const pattern of patterns) {
    const match = workbook.sheets.find((sheet) =>
      pattern.test(normalizeLooseText(sheet.sheetName)),
    );
    if (match) return match;
  }
  return undefined;
};

export const findExcludedSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/3\.?\s*виключ/, /виключен/]);

export const findOosSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/2\.?\s*оос/, /^оос$/, /облік\s*особов/]);

const RANK_GROUP_LABELS: Record<RankGroup, string> = {
  officer: "Офіцери",
  sergeant: "Сержанти",
  soldier: "Солдати",
};

const emptyRankBreakdown = (): RankBreakdown => ({
  officer: 0,
  sergeant: 0,
  soldier: 0,
});

const bumpRank = (target: RankBreakdown, rankGroup: RankGroup) => {
  target[rankGroup] += 1;
};

const rankBreakdownOf = (
  people: Array<{ rankGroup: RankGroup }>,
): RankBreakdown => {
  const byRank = emptyRankBreakdown();
  people.forEach((person) => bumpRank(byRank, person.rankGroup));
  return byRank;
};

const MONTH_LABELS_UK = [
  "",
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
];

const parseCellDate = (value: CellValue | undefined): Date | null => {
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
  if (Number.isNaN(date.getTime())) return null;
  return date;
};

const normalizeHeader = (value: CellValue | undefined) =>
  normalizeLooseText(cellText(value)).replace(/[.]/g, " ");

const columnIndex = (headerRow: CellValue[], parts: string[], exclude: string[] = []) =>
  headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    if (!header) return false;
    if (exclude.some((part) => header.includes(part))) return false;
    return parts.every((part) => header.includes(part));
  });

const CATEGORY_LABELS: Record<DepartureCategory, string> = {
  discharged_unit: "Звільнені з військової частини",
  demobilization: "У звʼязку з оголошенням демобілізації",
  health: "За станом здоровʼя",
  unsuitable_court: "За службовою невідповідністю / вироком суду",
  other: "Інше (звільнення / некласифіковано)",
  transfer: "Переведені / розпорядження (не звільнення)",
};

export const DEPARTURE_CATEGORY_ORDER: DepartureCategory[] = [
  "discharged_unit",
  "demobilization",
  "health",
  "unsuitable_court",
  "other",
  "transfer",
];

export const classifyDepartureReason = (
  ground: string,
  type: string,
): { category: DepartureCategory; matchNote: string } => {
  const text = normalizeLooseText(`${ground} ${type}`);
  if (!text) {
    return { category: "other", matchNote: "порожня підстава" };
  }

  const isTransferOnly =
    /(перевед|розпоряд)/.test(text) &&
    !/(звільн|демобіл|здоров|медич|непридат|невідповід|вирок|суду|кримінал|засудж)/.test(
      text,
    );
  if (isTransferOnly) {
    return { category: "transfer", matchNote: "переведення/розпорядження" };
  }

  if (/демобіл/.test(text)) {
    return { category: "demobilization", matchNote: "демобілізація" };
  }
  if (/(здоров|медич|непридат|лкк|влк|хвор)/.test(text)) {
    return { category: "health", matchNote: "стан здоровʼя" };
  }
  if (/(невідповід|вирок|суду|кримінал|засудж|обвинуваль)/.test(text)) {
    return {
      category: "unsuitable_court",
      matchNote: "невідповідність / вирок суду",
    };
  }
  if (/звільн/.test(text)) {
    return { category: "discharged_unit", matchNote: "звільнення з в/ч" };
  }

  return { category: "other", matchNote: "інше / без чіткого маркера" };
};

const parseExcludedByHeaders = (sheet: ExcelSheetSnapshot): DeparturePerson[] => {
  let headerIndex = -1;
  for (let i = 0; i < Math.min(12, sheet.rawRows.length); i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const headers = row.map((cell) => normalizeHeader(cell));
    const hasName = headers.some((h) => h === "піб" || h.includes("прізвище"));
    if (hasName) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex < 0) return [];

  const headerRow = sheet.rawRows[headerIndex] ?? [];
  const rankCol = columnIndex(headerRow, ["зван"]);
  const nameCol = columnIndex(headerRow, ["піб"]);
  const nameColAlt = columnIndex(headerRow, ["прізвище"]);
  const idCol = columnIndex(headerRow, ["id"]);
  const indexCol = columnIndex(headerRow, ["індекс"]);
  const groundCol = columnIndex(headerRow, ["підстав"]);
  const typeCol = columnIndex(headerRow, ["тип"], ["підстав"]);
  const typeColAlt = columnIndex(headerRow, ["вид"]);
  const destCol = columnIndex(headerRow, ["куди"]);
  const orderNumCol = columnIndex(headerRow, ["наказ"]);
  const orderDateCol = columnIndex(headerRow, ["дата"], ["виключ", "наказ"]);
  const excludeDateCol = columnIndex(headerRow, ["виключ"]);

  const people: DeparturePerson[] = [];
  for (let i = headerIndex + 1; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const fullName =
      cellText(row[nameCol >= 0 ? nameCol : nameColAlt]) ||
      cellText(row[1]) ||
      cellText(row[6]);
    if (!fullName || fullName.length < 3) continue;
    const ground =
      cellText(row[groundCol]) ||
      cellText(row[typeCol >= 0 ? typeCol : typeColAlt]) ||
      cellText(row[31]); // AF fallback
    const type =
      cellText(row[typeCol >= 0 ? typeCol : typeColAlt]) || cellText(row[31]);
    const classified = classifyDepartureReason(ground, type);
    const rank = cellText(row[rankCol >= 0 ? rankCol : 0]) || cellText(row[5]);
    const rankGroup = classifyRankGroup(rank);
    people.push({
      excelRow: i + 1,
      rank,
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      fullName,
      personId: cellText(row[idCol >= 0 ? idCol : 2]) || cellText(row[7]),
      positionIndex:
        cellText(row[indexCol >= 0 ? indexCol : 3]) || cellText(row[3]),
      ground,
      type,
      destination: cellText(row[destCol]) || cellText(row[30]), // AE
      orderNumber: cellText(row[orderNumCol]) || cellText(row[29]),
      orderDate: cellText(row[orderDateCol]) || cellText(row[27]),
      excludeDate: cellText(row[excludeDateCol]) || cellText(row[27]),
      category: classified.category,
      categoryLabel: CATEGORY_LABELS[classified.category],
      matchNote: classified.matchNote,
    });
  }
  return people;
};

/** Fallback layout як у шаблоні ЕЖООС (без надійних заголовків). */
const parseExcludedFixed = (sheet: ExcelSheetSnapshot): DeparturePerson[] => {
  const people: DeparturePerson[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const fullName = cellText(row[1]) || cellText(row[6]);
    const personId = cellText(row[2]) || cellText(row[7]);
    if (!fullName && !personId) continue;
    if (!fullName) continue;
    const ground = cellText(row[4]) || cellText(row[31]);
    const type = cellText(row[31]) || cellText(row[4]);
    const classified = classifyDepartureReason(ground, type);
    const rank = cellText(row[0]) || cellText(row[5]);
    const rankGroup = classifyRankGroup(rank);
    people.push({
      excelRow: i + 1,
      rank,
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      fullName,
      personId,
      positionIndex: cellText(row[3]),
      ground,
      type,
      destination: cellText(row[30]) || cellText(row[8]),
      orderNumber: cellText(row[29]) || cellText(row[11]),
      orderDate: cellText(row[27]) || cellText(row[12]),
      excludeDate: cellText(row[27]) || cellText(row[28]),
      category: classified.category,
      categoryLabel: CATEGORY_LABELS[classified.category],
      matchNote: classified.matchNote,
    });
  }
  return people;
};

const personEventDate = (person: DeparturePerson): Date | null =>
  parseCellDate(person.excludeDate) || parseCellDate(person.orderDate);

const filterDeparturesSince = (
  people: DeparturePerson[],
  since: Date,
): {
  people: DeparturePerson[];
  skippedNoDate: number;
  skippedBeforePeriod: number;
} => {
  let skippedNoDate = 0;
  let skippedBeforePeriod = 0;
  const filtered: DeparturePerson[] = [];
  for (const person of people) {
    const eventDate = personEventDate(person);
    if (!eventDate) {
      skippedNoDate += 1;
      continue;
    }
    if (eventDate.getTime() < since.getTime()) {
      skippedBeforePeriod += 1;
      continue;
    }
    filtered.push(person);
  }
  return { people: filtered, skippedNoDate, skippedBeforePeriod };
};

/** Прибули до в/ч за «Дата зарахування до списків частини» (ООС) за місяць. */
export const buildArrivalsForMonth = (
  workbook: ExcelWorkbookSnapshot,
  month: number,
  year: number,
): ArrivalsMonthResult => {
  const sheet = findOosSheet(workbook);
  if (!sheet) {
    throw new Error('У файлі ЕЖООС не знайдено аркуш «2. ООС».');
  }

  let headerIndex = -1;
  for (let i = 0; i < Math.min(12, sheet.rawRows.length); i += 1) {
    const headers = (sheet.rawRows[i] ?? []).map((cell) => normalizeHeader(cell));
    const hasName = headers.some((h) => h.includes("прізвище") || h === "піб");
    const hasEnroll = headers.some(
      (h) => h.includes("зарахуван") && h.includes("списк"),
    );
    if (hasName && (hasEnroll || headers.some((h) => h.includes("звідки")))) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex < 0) {
    throw new Error('На аркуші «2. ООС» не знайдено рядок заголовків.');
  }

  const headerRow = sheet.rawRows[headerIndex] ?? [];
  const rankCol = columnIndex(headerRow, ["зван"]);
  const nameCol = columnIndex(headerRow, ["піб"]);
  const nameColAlt = columnIndex(headerRow, ["прізвище"]);
  const idCol = columnIndex(headerRow, ["id"]);
  const indexCol = columnIndex(headerRow, ["індекс"]);
  let fromCol = columnIndex(headerRow, ["звідки", "прибув"]);
  if (fromCol < 0) fromCol = columnIndex(headerRow, ["звідки"], ["куди"]);
  // Типовий шаблон ЕЖООС: G = «Звідки прибув» (індекс 6).
  if (fromCol < 0) fromCol = 6;
  let enrollCol = columnIndex(headerRow, ["зарахуван", "списк"]);
  if (enrollCol < 0) {
    enrollCol = columnIndex(headerRow, ["дата", "зарахуван"]);
  }
  // Типовий шаблон ЕЖООС 4.x: колонка H (індекс 7).
  if (enrollCol < 0) enrollCol = 7;

  const people: EnrollmentArrivalPerson[] = [];
  for (let i = headerIndex + 1; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const fullName =
      cellText(row[nameCol >= 0 ? nameCol : nameColAlt]) || cellText(row[1]);
    if (!fullName || fullName.length < 3) continue;

    const enrollRaw = row[enrollCol];
    const enrollDate = parseCellDate(enrollRaw);
    if (!enrollDate) continue;
    if (
      enrollDate.getUTCMonth() + 1 !== month ||
      enrollDate.getUTCFullYear() !== year
    ) {
      continue;
    }

    const rank = cellText(row[rankCol >= 0 ? rankCol : 0]);
    const rankGroup = classifyRankGroup(rank);
    const arrivedFrom = cellText(row[fromCol]);
    const arrivalSource = classifyOosArrivedFrom(arrivedFrom);
    people.push({
      excelRow: i + 1,
      rank,
      fullName,
      personId: cellText(row[idCol >= 0 ? idCol : 2]),
      positionIndex: cellText(row[indexCol >= 0 ? indexCol : 3]),
      arrivedFrom: arrivedFrom || "(порожньо → ТЦК)",
      enrollDate: formatUkDate(enrollDate),
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      arrivalSource,
      arrivalSourceLabel: ARRIVAL_SOURCE_LABELS[arrivalSource],
    });
  }

  return buildArrivalsSummary(people, {
    sourceSheet: sheet.sheetName,
    month,
    year,
  });
};

const buildArrivalsSummary = (
  people: EnrollmentArrivalPerson[],
  meta: {
    sourceSheet: string;
    month: number;
    year: number;
  },
): ArrivalsMonthResult => {
  const byRank = emptyRankBreakdown();
  const bySource: Record<ArrivalSource, number> = {
    tck: 0,
    trainingCenter: 0,
    recruiting: 0,
    brez: 0,
    unitTransfer: 0,
    other: 0,
    unknown: 0,
  };
  const bySourceRank: Record<ArrivalSource, RankBreakdown> = {
    tck: emptyRankBreakdown(),
    trainingCenter: emptyRankBreakdown(),
    recruiting: emptyRankBreakdown(),
    brez: emptyRankBreakdown(),
    unitTransfer: emptyRankBreakdown(),
    other: emptyRankBreakdown(),
    unknown: emptyRankBreakdown(),
  };
  people.forEach((person) => {
    bumpRank(byRank, person.rankGroup);
    bySource[person.arrivalSource] += 1;
    bumpRank(bySourceRank[person.arrivalSource], person.rankGroup);
  });
  const sourceOrder: ArrivalSource[] = [
    "brez",
    "tck",
    "trainingCenter",
    "unitTransfer",
    "recruiting",
    "other",
    "unknown",
  ];
  const sourceSummary = sourceOrder
    .filter(
      (source) =>
        bySource[source] > 0 ||
        source === "tck" ||
        source === "trainingCenter" ||
        source === "unitTransfer" ||
        source === "brez",
    )
    .map((source) => ({
      source,
      label: ARRIVAL_SOURCE_LABELS[source],
      count: bySource[source],
      byRank: bySourceRank[source],
    }));
  const monthName = MONTH_LABELS_UK[meta.month] || String(meta.month);
  return {
    sourceSheet: meta.sourceSheet,
    month: meta.month,
    year: meta.year,
    monthLabel: `${monthName} ${meta.year}`,
    people,
    byRank,
    bySource,
    sourceSummary,
    total: people.length,
  };
};

/**
 * Штатка (ранковий звіт) як база:
 * 1) хто з БРЕЗ у ранковому → «З БРЕЗ»
 * 2) решта → «Звідки прибув» з ООС (порожнє = ТЦК)
 * Розбивка: офіцери / сержанти / солдати.
 */
export const buildArrivalSourcesFromMorning = (
  ejoos: ExcelWorkbookSnapshot,
  morning: ExcelWorkbookSnapshot,
): ArrivalsMonthResult => {
  const morningSheet =
    morning.sheets.find(isBchsPersonnelGeneralListSheet) ??
    findSheet(morning, [/ос.*загальн/, /загальн.*спис/]);
  if (!morningSheet) {
    throw new Error(
      "У ранковому звіті (Штатка) не знайдено аркуш «1.ОС Загальний список».",
    );
  }

  const oosFrom = buildOosArrivedFromLookup(ejoos);
  const novaPeople = filterBchsNovaPeople(
    extractBchsAwayPeopleFromSheet(morningSheet),
  );

  const people: EnrollmentArrivalPerson[] = [];
  const seen = new Set<string>();
  let brezCount = 0;
  let emptyAsTck = 0;
  let fromOos = 0;

  for (const roster of novaPeople) {
    if (!isBchsListedWithRankTitle(roster)) continue;
    const fullName = roster.fullName.trim();
    if (!looksLikePersonName(fullName)) continue;
    const nameKey = normalizePersonName(fullName);
    if (!nameKey || seen.has(nameKey)) continue;
    seen.add(nameKey);

    const rank =
      roster.rankTitle || roster.rankCategory || roster.shpkFact || "";
    // Категорія ШПК (оф./серж./солд.) надійніша для розбивки, ніж лише звання.
    const rankGroup = classifyRankGroup(
      roster.rankCategory || roster.rankTitle || roster.shpkFact || "",
      roster.position,
    );
    const morningBrez = combineMorningBrezFields(
      roster.bzvpStatus ?? "",
      roster.brezAssignment ?? "",
    );
    const isBrez =
      isBchsBrezAttachedPerson(roster) ||
      isBchsBrezRoster(roster.rosterUnit ?? "") ||
      isBrezMarkerText(morningBrez);

    let arrivalSource: ArrivalSource;
    let arrivedFrom: string;

    if (isBrez) {
      arrivalSource = "brez";
      arrivedFrom = morningBrez || roster.rosterUnit || "БРЕЗ";
      brezCount += 1;
    } else {
      const arrivedFromRaw =
        oosFrom.byName.get(normalizeLooseText(fullName)) ||
        oosFrom.byName.get(nameKey) ||
        "";
      arrivalSource = classifyOosArrivedFrom(arrivedFromRaw);
      arrivedFrom = arrivedFromRaw || "(порожньо → ТЦК)";
      if (!arrivedFromRaw) emptyAsTck += 1;
      else fromOos += 1;
    }

    people.push({
      excelRow: people.length + 1,
      rank,
      fullName,
      personId: "",
      positionIndex: "",
      arrivedFrom,
      enrollDate: "",
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      arrivalSource,
      arrivalSourceLabel: ARRIVAL_SOURCE_LABELS[arrivalSource],
    });
  }

  const summary = buildArrivalsSummary(people, {
    sourceSheet: `Штатка (${morningSheet.sheetName}) + ООС «Звідки прибув»`,
    month: 0,
    year: new Date().getFullYear(),
  });
  summary.monthLabel = `ранковий · ${people.length} осіб`;
  console.info(
    `[Звідки / Штатка] усього ${people.length}, БРЕЗ ${brezCount}, з ООС тексту ${fromOos}, порожнє→ТЦК ${emptyAsTck}`,
  );
  return summary;
};

export const logArrivalSourcesFromMorning = (
  result: ArrivalsMonthResult,
  fileName?: string,
) => {
  console.group(
    `[Звідки / Штатка] ${fileName ?? ""} · ${result.monthLabel} · ${result.total}`,
  );
  console.table(
    result.sourceSummary.map((row) => ({
      джерело: row.label,
      кількість: row.count,
    })),
  );
  console.info("за званням", result.byRank);
  for (const source of result.sourceSummary) {
    if (source.count <= 0) continue;
    const list = result.people.filter((p) => p.arrivalSource === source.source);
    console.group(`${source.label} (${list.length})`);
    console.table(
      list.map((p) => ({
        fullName: p.fullName,
        rank: p.rank,
        rankGroup: p.rankGroupLabel,
        arrivedFrom: p.arrivedFrom,
      })),
    );
    console.groupEnd();
  }
  console.groupEnd();
};

/** Lookup «Звідки прибув» з ООС за ID / нормалізованим ПІБ. */
export const buildOosArrivedFromLookup = (
  workbook: ExcelWorkbookSnapshot,
): { byId: Map<string, string>; byName: Map<string, string> } => {
  const byId = new Map<string, string>();
  const byName = new Map<string, string>();
  const sheet = findOosSheet(workbook);
  if (!sheet) return { byId, byName };

  let headerIndex = -1;
  for (let i = 0; i < Math.min(12, sheet.rawRows.length); i += 1) {
    const headers = (sheet.rawRows[i] ?? []).map((cell) => normalizeHeader(cell));
    const hasName = headers.some((h) => h.includes("прізвище") || h === "піб");
    const hasFrom = headers.some((h) => h.includes("звідки"));
    if (hasName && hasFrom) {
      headerIndex = i;
      break;
    }
  }
  if (headerIndex < 0) return { byId, byName };

  const headerRow = sheet.rawRows[headerIndex] ?? [];
  const nameCol = columnIndex(headerRow, ["піб"]);
  const nameColAlt = columnIndex(headerRow, ["прізвище"]);
  const idCol = columnIndex(headerRow, ["id"]);
  let fromCol = columnIndex(headerRow, ["звідки", "прибув"]);
  if (fromCol < 0) fromCol = columnIndex(headerRow, ["звідки"], ["куди"]);
  if (fromCol < 0) fromCol = 6;

  for (let i = headerIndex + 1; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i] ?? [];
    const fullName =
      cellText(row[nameCol >= 0 ? nameCol : nameColAlt]) || cellText(row[1]);
    if (!fullName || fullName.length < 3) continue;
    const arrivedFrom = cellText(row[fromCol]);
    if (!arrivedFrom) continue;
    const personId = cellText(row[idCol >= 0 ? idCol : 2]);
    if (personId) byId.set(personId, arrivedFrom);
    byName.set(normalizeLooseText(fullName), arrivedFrom);
    const personKey = normalizePersonName(fullName);
    if (personKey) byName.set(personKey, arrivedFrom);
  }
  return { byId, byName };
};

/** Прибули до в/ч з 1ПБ · аркуш Рух · тип ПРИБУВ за датою наказу. */
export const buildArrivalsFromPb = (
  workbook: ExcelWorkbookSnapshot,
  month: number,
  year: number,
  options?: { ejoos?: ExcelWorkbookSnapshot | null },
): ArrivalsMonthResult => {
  const hasSh = workbook.sheets.some((sheet) => /^sh$/i.test(sheet.sheetName));
  const hasRuh = workbook.sheets.some((sheet) => /^рух$/i.test(sheet.sheetName));
  if (!hasSh || !hasRuh) {
    throw new Error(
      "У файлі 1ПБ потрібні аркуші sh і Рух (зазвичай ще archive).",
    );
  }

  const movements = parsePbMovements(workbook);
  const shPeople = parsePbShPeople(workbook);
  const shById = new Map(
    shPeople
      .filter((person) => person.personId)
      .map((person) => [person.personId, person]),
  );
  const shByName = new Map(
    shPeople
      .filter((person) => person.fullName)
      .map((person) => [normalizeLooseText(person.fullName), person]),
  );
  const oosFrom = options?.ejoos
    ? buildOosArrivedFromLookup(options.ejoos)
    : { byId: new Map<string, string>(), byName: new Map<string, string>() };

  const people: EnrollmentArrivalPerson[] = [];
  const seen = new Set<string>();

  for (const event of movements) {
    if (event.type !== "ПРИБУВ") continue;
    const enrollDate = parseCellDate(event.orderDate);
    if (!enrollDate) continue;
    if (
      enrollDate.getUTCMonth() + 1 !== month ||
      enrollDate.getUTCFullYear() !== year
    ) {
      continue;
    }

    const dedupeKey = [
      event.personId || normalizeLooseText(event.fullName),
      formatUkDate(enrollDate),
      event.movementNumber || String(event.excelRow),
    ].join("|");
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    const fromSh =
      (event.personId ? shById.get(event.personId) : undefined) ||
      shByName.get(normalizeLooseText(event.fullName));
    const rank = event.rank || fromSh?.rank || "";
    const rankGroup = classifyRankGroup(rank);
    // Класифікація лише з «Звідки прибув» (ООС / sh / Рух), не з «куди»/примітки.
    const arrivedFrom =
      event.arrivedFrom ||
      fromSh?.arrivedFrom ||
      (event.personId ? oosFrom.byId.get(event.personId) : undefined) ||
      oosFrom.byName.get(normalizeLooseText(event.fullName)) ||
      "";
    // Порожнє «Звідки прибув» (ООС) = ТЦК.
    const arrivalSource = classifyOosArrivedFrom(arrivedFrom);

    people.push({
      excelRow: event.excelRow,
      rank,
      fullName: event.fullName || fromSh?.fullName || "",
      personId: event.personId || fromSh?.personId || "",
      positionIndex: event.nextIndex || fromSh?.positionIndex || "",
      arrivedFrom: arrivedFrom || "(порожньо → ТЦК)",
      enrollDate: formatUkDate(enrollDate),
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      arrivalSource,
      arrivalSourceLabel: ARRIVAL_SOURCE_LABELS[arrivalSource],
    });
  }

  people.sort((a, b) => {
    const byDate = a.enrollDate.localeCompare(b.enrollDate, "uk");
    if (byDate !== 0) return byDate;
    return a.fullName.localeCompare(b.fullName, "uk");
  });

  return buildArrivalsSummary(people, {
    sourceSheet: "Рух (ПРИБУВ) · Звідки прибув",
    month,
    year,
  });
};

export const DISPOSITION_REASON_ORDER: DispositionReason[] = [
  "treatment",
  "organizational",
  "other",
];

export const DISPOSITION_REASON_LABELS: Record<DispositionReason, string> = {
  treatment: "У звʼязку з тривалим лікуванням",
  organizational: "У звʼязку з організаційними заходами",
  other: "Інші причини",
};

const isDispositionArchiveRow = (absenceType: string, place: string) => {
  const text = normalizeLooseText(`${absenceType} ${place}`);
  if (!text) return false;
  // Звільнення / переведення в іншу в/ч — не «у розпорядження».
  if (
    /(звільн|демобіл|виключ)/.test(text) ||
    (/(перевед)/.test(text) && !/розпор/.test(text))
  ) {
    return false;
  }
  // Відпустка / відрядження без розпорядження — не беремо.
  if (
    /(відпустк|відряджен|відком)/.test(text) &&
    !/(розпор|лікуван|організац)/.test(text)
  ) {
    return false;
  }
  return /(розпор|виведен|в\s*розпор|лікуван|тривал.*лік|шпитал|організац|орг\.?\s*зах)/.test(
    text,
  );
};

export const classifyDispositionReason = (
  absenceType: string,
  place: string,
): { reason: DispositionReason; matchNote: string } => {
  const text = normalizeLooseText(`${absenceType} ${place}`);
  if (/(лікуван|тривал.*лік|шпитал|мед\.?закл|влк|хвороб|поранен|реабіліт)/.test(text)) {
    return { reason: "treatment", matchNote: "лікування" };
  }
  if (/(організац|орг\.?\s*зах|штатн|реорган|скорочен)/.test(text)) {
    return { reason: "organizational", matchNote: "організаційні заходи" };
  }
  return { reason: "other", matchNote: "інше / некласифіковано" };
};

/**
 * Виведені у розпорядження з 1ПБ archive.
 * За замовчуванням: з 01.01.2026, лише відкриті періоди (без дати повернення).
 */
export const buildDispositionFromArchive = (
  workbook: ExcelWorkbookSnapshot,
  options?: {
    sinceDate?: string | null;
    /** Якщо true — лише періоди без дати повернення. */
    openOnly?: boolean;
  },
): DispositionArchiveResult => {
  const hasArchive = workbook.sheets.some((sheet) =>
    /^archive$/i.test(sheet.sheetName),
  );
  if (!hasArchive) {
    throw new Error('У файлі 1ПБ не знайдено аркуш «archive».');
  }

  const periods = parsePbArchive(workbook);
  const sinceRaw =
    options?.sinceDate === undefined
      ? DEFAULT_DEPARTURES_SINCE
      : options.sinceDate;
  const since = sinceRaw ? parseCellDate(sinceRaw) : null;
  const openOnly = options?.openOnly !== false;

  let skippedNotDisposition = 0;
  let skippedBeforePeriod = 0;
  let skippedNoDate = 0;
  let skippedReturned = 0;
  const people: DispositionPerson[] = [];
  const seen = new Set<string>();

  for (const period of periods) {
    if (!isDispositionArchiveRow(period.absenceType, period.place)) {
      skippedNotDisposition += 1;
      continue;
    }
    if (openOnly && period.returnDate) {
      skippedReturned += 1;
      continue;
    }
    const eventDate =
      parseCellDate(period.departDate) || parseCellDate(period.orderDate);
    if (since) {
      if (!eventDate) {
        skippedNoDate += 1;
        continue;
      }
      if (eventDate.getTime() < since.getTime()) {
        skippedBeforePeriod += 1;
        continue;
      }
    }

    const dedupeKey =
      period.personId ||
      normalizeLooseText(period.fullName) ||
      String(period.excelRow);
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    const rankGroup = classifyRankGroup(period.rank);
    const classified = classifyDispositionReason(
      period.absenceType,
      period.place,
    );
    people.push({
      excelRow: period.excelRow,
      rank: period.rank,
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      fullName: period.fullName,
      personId: period.personId,
      absenceType: period.absenceType,
      place: period.place,
      departDate: period.departDate,
      orderDate: period.orderDate,
      returnDate: period.returnDate,
      reason: classified.reason,
      reasonLabel: DISPOSITION_REASON_LABELS[classified.reason],
      matchNote: classified.matchNote,
    });
  }

  people.sort((a, b) => a.fullName.localeCompare(b.fullName, "uk"));

  const byReason = Object.fromEntries(
    DISPOSITION_REASON_ORDER.map((reason) => [reason, [] as DispositionPerson[]]),
  ) as Record<DispositionReason, DispositionPerson[]>;
  people.forEach((person) => byReason[person.reason].push(person));

  const summary: DispositionSummaryRow[] = DISPOSITION_REASON_ORDER.map(
    (reason) => ({
      reason,
      label: DISPOSITION_REASON_LABELS[reason],
      count: byReason[reason].length,
      byRank: rankBreakdownOf(byReason[reason]),
    }),
  );

  const result: DispositionArchiveResult = {
    sourceSheet: "archive",
    people,
    summary,
    totals: {
      all: people.length,
      byRank: rankBreakdownOf(people),
    },
    periodFrom: since ? formatUkDate(since) : null,
    periodFromLabel: since
      ? `з ${formatUkDate(since)} (дата вибуття / наказу)`
      : null,
    totalArchiveRows: periods.length,
    skippedNotDisposition,
    skippedBeforePeriod,
    skippedNoDate,
    skippedReturned,
  };
  logDispositionByReason(result, workbook.fileName);
  return result;
};

export const logDispositionByReason = (
  result: DispositionArchiveResult,
  fileName?: string,
) => {
  console.group(
    `[Розпорядження / archive] ${fileName || result.sourceSheet}${
      result.periodFrom ? ` · з ${result.periodFrom}` : ""
    } · ${result.totals.all}`,
  );
  console.log("Підсумок:", {
    усього: result.totals.all,
    офіцери: result.totals.byRank.officer,
    сержанти: result.totals.byRank.sergeant,
    солдати: result.totals.byRank.soldier,
    неРозпорядження: result.skippedNotDisposition,
    доПеріоду: result.skippedBeforePeriod,
    безДати: result.skippedNoDate,
    ужеПовернулись: result.skippedReturned,
  });
  for (const row of result.summary) {
    const list = result.people.filter((person) => person.reason === row.reason);
    console.groupCollapsed(
      `${row.label} · ${list.length} (оф. ${row.byRank.officer}, серж. ${row.byRank.sergeant}, солд. ${row.byRank.soldier})`,
    );
    console.table(
      list.map((person) => ({
        fullName: person.fullName,
        rank: person.rank,
        rankGroup: person.rankGroupLabel,
        absenceType: person.absenceType,
        place: person.place,
        departDate: person.departDate,
        matchNote: person.matchNote,
        excelRow: person.excelRow,
      })),
    );
    console.log("ПІБ за званням:", {
      офіцери: list
        .filter((p) => p.rankGroup === "officer")
        .map((p) => p.fullName),
      сержанти: list
        .filter((p) => p.rankGroup === "sergeant")
        .map((p) => p.fullName),
      солдати: list
        .filter((p) => p.rankGroup === "soldier")
        .map((p) => p.fullName),
    });
    console.groupEnd();
  }
  console.groupEnd();
};

const isSzchMovementText = (...parts: string[]) => {
  const text = normalizeLooseText(parts.filter(Boolean).join(" "));
  if (!text) return false;
  return /(сзч|самовільн|самовольн|дезерт)/.test(text);
};

/**
 * СЗЧ з 1ПБ · аркуш Рух (тип / статус / примітка / куди / яка зміна).
 * За замовчуванням з 01.01.2026 за датою наказу.
 */
export const buildSzchFromRuh = (
  workbook: ExcelWorkbookSnapshot,
  options?: { sinceDate?: string | null; untilDate?: string | null },
): SzchRuhResult => {
  const hasRuh = workbook.sheets.some((sheet) => /^рух$/i.test(sheet.sheetName));
  if (!hasRuh) {
    throw new Error('У файлі 1ПБ не знайдено аркуш «Рух».');
  }

  const movements = parsePbMovements(workbook);
  const shPeople = parsePbShPeople(workbook);
  const shById = new Map(
    shPeople
      .filter((person) => person.personId)
      .map((person) => [person.personId, person]),
  );
  const shByName = new Map(
    shPeople
      .filter((person) => person.fullName)
      .map((person) => [normalizeLooseText(person.fullName), person]),
  );

  const sinceRaw =
    options?.sinceDate === undefined
      ? DEFAULT_DEPARTURES_SINCE
      : options.sinceDate;
  const since = sinceRaw ? parseCellDate(sinceRaw) : null;
  const untilRaw = options?.untilDate ?? null;
  const until = untilRaw ? parseCellDate(untilRaw) : null;

  let skippedBeforePeriod = 0;
  let skippedAfterPeriod = 0;
  let skippedNoDate = 0;
  const people: SzchPerson[] = [];
  const seen = new Set<string>();

  for (const event of movements) {
    if (
      !isSzchMovementText(
        event.type,
        event.status,
        event.note,
        event.destination,
        event.changeText,
      )
    ) {
      continue;
    }

    const eventDate = parseCellDate(event.orderDate);
    if (since) {
      if (!eventDate) {
        skippedNoDate += 1;
        continue;
      }
      if (eventDate.getTime() < since.getTime()) {
        skippedBeforePeriod += 1;
        continue;
      }
    }
    if (until) {
      if (!eventDate) {
        skippedNoDate += 1;
        continue;
      }
      if (eventDate.getTime() > until.getTime()) {
        skippedAfterPeriod += 1;
        continue;
      }
    }

    const dedupeKey =
      event.personId ||
      normalizeLooseText(event.fullName) ||
      String(event.excelRow);
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    const fromSh =
      (event.personId ? shById.get(event.personId) : undefined) ||
      shByName.get(normalizeLooseText(event.fullName));
    const rank = event.rank || fromSh?.rank || "";
    const rankGroup = classifyRankGroup(rank);
    const matchSource = [
      event.type,
      event.status,
      event.note,
      event.destination,
      event.changeText,
    ]
      .filter((part) => isSzchMovementText(part))
      .join(" · ");

    people.push({
      excelRow: event.excelRow,
      rank,
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      fullName: event.fullName || fromSh?.fullName || "",
      personId: event.personId || fromSh?.personId || "",
      movementNumber: event.movementNumber,
      type: event.type,
      status: event.status,
      destination: event.destination,
      note: event.note,
      orderDate: event.orderDate,
      matchNote: matchSource || "СЗЧ",
    });
  }

  people.sort((a, b) => {
    const byDate = a.orderDate.localeCompare(b.orderDate, "uk");
    if (byDate !== 0) return byDate;
    return a.fullName.localeCompare(b.fullName, "uk");
  });

  const result: SzchRuhResult = {
    sourceSheet: "Рух",
    people,
    totals: {
      all: people.length,
      byRank: rankBreakdownOf(people),
    },
    periodFrom: since ? formatUkDate(since) : null,
    periodFromLabel: since
      ? `з ${formatUkDate(since)} (дата наказу в Рух)`
      : null,
    totalMovements: movements.length,
    skippedBeforePeriod,
    skippedAfterPeriod,
    skippedNoDate,
  };
  logSzchFromRuh(result, workbook.fileName);
  return result;
};

const mergeSzchPeople = (primary: SzchPerson[], extra: SzchPerson[]) => {
  const merged = [...primary];
  const seen = new Set(
    primary.map(
      (person) =>
        person.personId ||
        normalizeLooseText(person.fullName) ||
        String(person.excelRow),
    ),
  );
  for (const person of extra) {
    const key =
      person.personId ||
      normalizeLooseText(person.fullName) ||
      String(person.excelRow);
    if (seen.has(key)) continue;
    seen.add(key);
    merged.push(person);
  }
  merged.sort((a, b) => {
    const byDate = a.orderDate.localeCompare(b.orderDate, "uk");
    if (byDate !== 0) return byDate;
    return a.fullName.localeCompare(b.fullName, "uk");
  });
  return merged;
};

export const buildSzchFromArchive = (
  workbook: ExcelWorkbookSnapshot,
  options?: { sinceDate?: string | null; untilDate?: string | null },
): Pick<SzchRuhResult, "people" | "skippedBeforePeriod" | "skippedAfterPeriod" | "skippedNoDate"> => {
  const hasArchive = workbook.sheets.some((sheet) =>
    /^archive$/i.test(sheet.sheetName),
  );
  if (!hasArchive) {
    return {
      people: [],
      skippedBeforePeriod: 0,
      skippedAfterPeriod: 0,
      skippedNoDate: 0,
    };
  }

  const periods = parsePbArchive(workbook);
  const sinceRaw =
    options?.sinceDate === undefined
      ? DEFAULT_DEPARTURES_SINCE
      : options.sinceDate;
  const since = sinceRaw ? parseCellDate(sinceRaw) : null;
  const untilRaw = options?.untilDate ?? null;
  const until = untilRaw ? parseCellDate(untilRaw) : null;

  let skippedBeforePeriod = 0;
  let skippedAfterPeriod = 0;
  let skippedNoDate = 0;
  const people: SzchPerson[] = [];
  const seen = new Set<string>();

  for (const period of periods) {
    if (!isSzchMovementText(period.absenceType, period.place)) continue;

    const eventDate =
      parseCellDate(period.orderDate) || parseCellDate(period.departDate);
    if (since) {
      if (!eventDate) {
        skippedNoDate += 1;
        continue;
      }
      if (eventDate.getTime() < since.getTime()) {
        skippedBeforePeriod += 1;
        continue;
      }
    }
    if (until) {
      if (!eventDate) {
        skippedNoDate += 1;
        continue;
      }
      if (eventDate.getTime() > until.getTime()) {
        skippedAfterPeriod += 1;
        continue;
      }
    }

    const dedupeKey =
      period.personId ||
      normalizeLooseText(period.fullName) ||
      String(period.excelRow);
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    const rankGroup = classifyRankGroup(period.rank);
    people.push({
      excelRow: period.excelRow,
      rank: period.rank,
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      fullName: period.fullName,
      personId: period.personId,
      movementNumber: period.periodNumber,
      type: period.absenceType,
      status: "",
      destination: period.place,
      note: "",
      orderDate: period.orderDate || period.departDate,
      matchNote: "archive",
    });
  }

  return {
    people,
    skippedBeforePeriod,
    skippedAfterPeriod,
    skippedNoDate,
  };
};

/** Унікальні особи з СЗЧ за календарний рік (1ПБ · Рух + archive). */
export const buildSzchYearSummary = (
  workbook: ExcelWorkbookSnapshot,
  options?: { year?: number },
): SzchRuhResult => {
  const year = options?.year ?? new Date().getFullYear();
  const sinceDate = `01.01.${year}`;
  const untilDate = `31.12.${year}`;
  const rangeOptions = { sinceDate, untilDate };

  let ruh: SzchRuhResult | null = null;
  try {
    ruh = buildSzchFromRuh(workbook, rangeOptions);
  } catch {
    ruh = null;
  }

  const archive = buildSzchFromArchive(workbook, rangeOptions);
  const people = mergeSzchPeople(ruh?.people ?? [], archive.people);
  const sourceSheet = ruh ? "Рух + archive" : "archive";

  return {
    sourceSheet,
    people,
    totals: {
      all: people.length,
      byRank: rankBreakdownOf(people),
    },
    periodFrom: sinceDate,
    periodFromLabel: `за ${year} рік · унікальні особи (дата наказу в Рух / archive)`,
    totalMovements: (ruh?.totalMovements ?? 0) + archive.people.length,
    skippedBeforePeriod:
      (ruh?.skippedBeforePeriod ?? 0) + archive.skippedBeforePeriod,
    skippedAfterPeriod:
      (ruh?.skippedAfterPeriod ?? 0) + archive.skippedAfterPeriod,
    skippedNoDate: (ruh?.skippedNoDate ?? 0) + archive.skippedNoDate,
  };
};

export const logSzchFromRuh = (result: SzchRuhResult, fileName?: string) => {
  console.group(
    `[СЗЧ / Рух] ${fileName || result.sourceSheet}${
      result.periodFrom ? ` · з ${result.periodFrom}` : ""
    } · ${result.totals.all}`,
  );
  console.log("Підсумок:", {
    усього: result.totals.all,
    офіцери: result.totals.byRank.officer,
    сержанти: result.totals.byRank.sergeant,
    солдати: result.totals.byRank.soldier,
    доПеріоду: result.skippedBeforePeriod,
    безДати: result.skippedNoDate,
    усьогоРядківРух: result.totalMovements,
  });
  console.table(
    result.people.map((person) => ({
      fullName: person.fullName,
      rank: person.rank,
      rankGroup: person.rankGroupLabel,
      type: person.type,
      status: person.status,
      orderDate: person.orderDate,
      matchNote: person.matchNote,
      excelRow: person.excelRow,
    })),
  );
  console.log("ПІБ за званням:", {
    офіцери: result.people
      .filter((p) => p.rankGroup === "officer")
      .map((p) => p.fullName),
    сержанти: result.people
      .filter((p) => p.rankGroup === "sergeant")
      .map((p) => p.fullName),
    солдати: result.people
      .filter((p) => p.rankGroup === "soldier")
      .map((p) => p.fullName),
  });
  console.groupEnd();
};

export const COMBAT_LOSS_REASON_ORDER: CombatLossReason[] = [
  "killed",
  "missing",
  "inCombat",
  "otherCircumstances",
];

export const COMBAT_LOSS_REASON_LABELS: Record<CombatLossReason, string> = {
  killed: "Загиблі",
  missing: "Безвісти",
  inCombat: "У бою",
  otherCircumstances: "Інші обставини",
};

const isCombatLossText = (...parts: string[]) => {
  const text = normalizeLooseText(parts.filter(Boolean).join(" "));
  if (!text) return false;
  // СЗЧ уже окремо.
  if (/(сзч|самовільн)/.test(text) && !/(загибл|безвіст|у\s*бою|200\b|500\b)/.test(text)) {
    return false;
  }
  return /(загибл|загинул|смерт|убит|\b200\b|безвіст|зникл|\bзб\b|\b500\b|у\s*бою|бойов|інш[іi]\s*обстав|нещасн|аварі)/.test(
    text,
  );
};

export const classifyCombatLossReason = (
  ...parts: string[]
): { reason: CombatLossReason; matchNote: string } => {
  const text = normalizeLooseText(parts.filter(Boolean).join(" "));
  if (/(безвіст|зникл.*безв|\bзб\b|\b500\b)/.test(text)) {
    return { reason: "missing", matchNote: "безвісти" };
  }
  if (/(загибл|загинул|смерт|убит|\b200\b)/.test(text)) {
    return { reason: "killed", matchNote: "загиблий" };
  }
  if (/(у\s*бою|бойов|під\s*час\s*бою|внаслідок\s*бою)/.test(text)) {
    return { reason: "inCombat", matchNote: "у бою" };
  }
  if (/(інш[іi]\s*обстав|нещасн|аварі|обставин)/.test(text)) {
    return { reason: "otherCircumstances", matchNote: "інші обставини" };
  }
  return { reason: "otherCircumstances", matchNote: "інше / некласифіковано" };
};

/**
 * Втрати з 1ПБ: Рух + archive.
 * Категорії: загиблі / безвісти / у бою / інші обставини.
 * За замовчуванням з 01.01.2026.
 */
export const buildCombatLossesFromPb = (
  workbook: ExcelWorkbookSnapshot,
  options?: { sinceDate?: string | null },
): CombatLossesResult => {
  const hasRuh = workbook.sheets.some((sheet) => /^рух$/i.test(sheet.sheetName));
  if (!hasRuh) {
    throw new Error('У файлі 1ПБ не знайдено аркуш «Рух».');
  }

  const movements = parsePbMovements(workbook);
  const archive = parsePbArchive(workbook);
  const shPeople = parsePbShPeople(workbook);
  const shById = new Map(
    shPeople
      .filter((person) => person.personId)
      .map((person) => [person.personId, person]),
  );
  const shByName = new Map(
    shPeople
      .filter((person) => person.fullName)
      .map((person) => [normalizeLooseText(person.fullName), person]),
  );

  const sinceRaw =
    options?.sinceDate === undefined
      ? DEFAULT_DEPARTURES_SINCE
      : options.sinceDate;
  const since = sinceRaw ? parseCellDate(sinceRaw) : null;

  let skippedBeforePeriod = 0;
  let skippedNoDate = 0;
  const people: CombatLossPerson[] = [];
  const seen = new Set<string>();

  const pushPerson = (input: {
    excelRow: number;
    sourceSheet: string;
    personId: string;
    fullName: string;
    rank: string;
    typeOrAbsence: string;
    status: string;
    placeOrNote: string;
    orderDate: string;
    matchParts: string[];
  }) => {
    const eventDate = parseCellDate(input.orderDate);
    if (since) {
      if (!eventDate) {
        skippedNoDate += 1;
        return;
      }
      if (eventDate.getTime() < since.getTime()) {
        skippedBeforePeriod += 1;
        return;
      }
    }

    const dedupeKey =
      input.personId ||
      normalizeLooseText(input.fullName) ||
      `${input.sourceSheet}:${input.excelRow}`;
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    const fromSh =
      (input.personId ? shById.get(input.personId) : undefined) ||
      shByName.get(normalizeLooseText(input.fullName));
    const rank = input.rank || fromSh?.rank || "";
    const rankGroup = classifyRankGroup(rank);
    const classified = classifyCombatLossReason(...input.matchParts);

    people.push({
      excelRow: input.excelRow,
      rank,
      rankGroup,
      rankGroupLabel: RANK_GROUP_LABELS[rankGroup],
      fullName: input.fullName || fromSh?.fullName || "",
      personId: input.personId || fromSh?.personId || "",
      sourceSheet: input.sourceSheet,
      typeOrAbsence: input.typeOrAbsence,
      status: input.status,
      placeOrNote: input.placeOrNote,
      orderDate: input.orderDate,
      reason: classified.reason,
      reasonLabel: COMBAT_LOSS_REASON_LABELS[classified.reason],
      matchNote: classified.matchNote,
    });
  };

  for (const event of movements) {
    const parts = [
      event.type,
      event.status,
      event.note,
      event.destination,
      event.changeText,
    ];
    if (!isCombatLossText(...parts)) continue;
    pushPerson({
      excelRow: event.excelRow,
      sourceSheet: "Рух",
      personId: event.personId,
      fullName: event.fullName,
      rank: event.rank,
      typeOrAbsence: event.type,
      status: event.status,
      placeOrNote: [event.note, event.destination, event.changeText]
        .filter(Boolean)
        .join(" · "),
      orderDate: event.orderDate,
      matchParts: parts,
    });
  }

  for (const period of archive) {
    const parts = [period.absenceType, period.place];
    if (!isCombatLossText(...parts)) continue;
    pushPerson({
      excelRow: period.excelRow,
      sourceSheet: "archive",
      personId: period.personId,
      fullName: period.fullName,
      rank: period.rank,
      typeOrAbsence: period.absenceType,
      status: "",
      placeOrNote: period.place,
      orderDate: period.departDate || period.orderDate,
      matchParts: parts,
    });
  }

  people.sort((a, b) => {
    const byDate = a.orderDate.localeCompare(b.orderDate, "uk");
    if (byDate !== 0) return byDate;
    return a.fullName.localeCompare(b.fullName, "uk");
  });

  const byReason = Object.fromEntries(
    COMBAT_LOSS_REASON_ORDER.map((reason) => [reason, [] as CombatLossPerson[]]),
  ) as Record<CombatLossReason, CombatLossPerson[]>;
  people.forEach((person) => byReason[person.reason].push(person));

  const summary: CombatLossSummaryRow[] = COMBAT_LOSS_REASON_ORDER.map(
    (reason) => ({
      reason,
      label: COMBAT_LOSS_REASON_LABELS[reason],
      count: byReason[reason].length,
      byRank: rankBreakdownOf(byReason[reason]),
    }),
  );

  const result: CombatLossesResult = {
    sourceSheet: "Рух + archive",
    people,
    summary,
    totals: {
      all: people.length,
      byRank: rankBreakdownOf(people),
    },
    periodFrom: since ? formatUkDate(since) : null,
    periodFromLabel: since
      ? `з ${formatUkDate(since)} (дата наказу / вибуття)`
      : null,
    skippedBeforePeriod,
    skippedNoDate,
  };
  logCombatLosses(result, workbook.fileName);
  return result;
};

export const logCombatLosses = (
  result: CombatLossesResult,
  fileName?: string,
) => {
  console.group(
    `[Втрати] ${fileName || result.sourceSheet}${
      result.periodFrom ? ` · з ${result.periodFrom}` : ""
    } · ${result.totals.all}`,
  );
  console.log("Підсумок:", {
    усього: result.totals.all,
    офіцери: result.totals.byRank.officer,
    сержанти: result.totals.byRank.sergeant,
    солдати: result.totals.byRank.soldier,
    доПеріоду: result.skippedBeforePeriod,
    безДати: result.skippedNoDate,
  });
  for (const row of result.summary) {
    const list = result.people.filter((person) => person.reason === row.reason);
    console.groupCollapsed(
      `${row.label} · ${list.length} (оф. ${row.byRank.officer}, серж. ${row.byRank.sergeant}, солд. ${row.byRank.soldier})`,
    );
    console.table(
      list.map((person) => ({
        fullName: person.fullName,
        rank: person.rank,
        rankGroup: person.rankGroupLabel,
        source: person.sourceSheet,
        typeOrAbsence: person.typeOrAbsence,
        placeOrNote: person.placeOrNote,
        orderDate: person.orderDate,
        matchNote: person.matchNote,
        excelRow: person.excelRow,
      })),
    );
    console.log("ПІБ за званням:", {
      офіцери: list
        .filter((p) => p.rankGroup === "officer")
        .map((p) => p.fullName),
      сержанти: list
        .filter((p) => p.rankGroup === "sergeant")
        .map((p) => p.fullName),
      солдати: list
        .filter((p) => p.rankGroup === "soldier")
        .map((p) => p.fullName),
    });
    console.groupEnd();
  }
  console.groupEnd();
};

export const DEFAULT_DEPARTURES_SINCE = "01.01.2026";

export const buildDeparturesFromEjoos = (
  workbook: ExcelWorkbookSnapshot,
  options?: {
    arrivalsMonth?: number;
    arrivalsYear?: number;
    /** Дата з якої рахувати вибули (включно). Формат DD.MM.YYYY або ISO. null = усі. */
    sinceDate?: string | null;
  },
): DeparturesResult => {
  const sheet = findExcludedSheet(workbook);
  if (!sheet) {
    throw new Error('У файлі ЕЖООС не знайдено аркуш «3. Виключені».');
  }

  let allPeople = parseExcludedByHeaders(sheet);
  if (allPeople.length < 3) {
    const fixed = parseExcludedFixed(sheet);
    if (fixed.length > allPeople.length) allPeople = fixed;
  }

  const sinceRaw =
    options?.sinceDate === undefined
      ? DEFAULT_DEPARTURES_SINCE
      : options.sinceDate;
  const since = sinceRaw ? parseCellDate(sinceRaw) : null;
  const filtered = since
    ? filterDeparturesSince(allPeople, since)
    : {
        people: allPeople,
        skippedNoDate: 0,
        skippedBeforePeriod: 0,
      };
  const people = filtered.people;

  const byCategoryPeople = Object.fromEntries(
    DEPARTURE_CATEGORY_ORDER.map((category) => [category, [] as DeparturePerson[]]),
  ) as Record<DepartureCategory, DeparturePerson[]>;
  people.forEach((person) => {
    byCategoryPeople[person.category].push(person);
  });

  const summary: DepartureSummaryRow[] = DEPARTURE_CATEGORY_ORDER.map(
    (category) => {
      const group = byCategoryPeople[category];
      return {
        category,
        label: CATEGORY_LABELS[category],
        count: group.length,
        byRank: rankBreakdownOf(group),
      };
    },
  );

  const transferPeople = byCategoryPeople.transfer;
  const dischargePeople = people.filter((person) => person.category !== "transfer");
  const transfers = transferPeople.length;
  const discharges = dischargePeople.length;

  const arrivalsMonth = options?.arrivalsMonth ?? 8;
  const arrivalsYear = options?.arrivalsYear ?? new Date().getFullYear();
  let arrivalsAugust: ArrivalsMonthResult | null = null;
  try {
    arrivalsAugust = buildArrivalsForMonth(workbook, arrivalsMonth, arrivalsYear);
  } catch {
    arrivalsAugust = null;
  }

  let arrivalsFromMorning: ArrivalsMonthResult | null = null;

  const result: DeparturesResult = {
    sourceSheet: sheet.sheetName,
    people,
    summary,
    totals: {
      all: people.length,
      discharges,
      transfers,
      byRank: rankBreakdownOf(people),
      dischargesByRank: rankBreakdownOf(dischargePeople),
      transfersByRank: rankBreakdownOf(transferPeople),
    },
    periodFrom: since ? formatUkDate(since) : null,
    periodFromLabel: since
      ? `з ${formatUkDate(since)} (дата виключення / наказу)`
      : null,
    totalUnfiltered: allPeople.length,
    skippedNoDate: filtered.skippedNoDate,
    skippedBeforePeriod: filtered.skippedBeforePeriod,
    arrivalsAugust,
    arrivalsFromMorning,
    arrivalsAugustPb: null,
    dispositionFromArchive: null,
    szchFromRuh: null,
    combatLossesFromPb: null,
  };
  logDeparturesByCategory(result, workbook.fileName);
  return result;
};

/** Додати «Звідки» зі Штатки (ранковий) + ООС до вже порахованих Вибули. */
export const withMorningArrivalSources = (
  departures: DeparturesResult,
  ejoos: ExcelWorkbookSnapshot,
  morning: ExcelWorkbookSnapshot,
): DeparturesResult => {
  try {
    const arrivalsFromMorning = buildArrivalSourcesFromMorning(ejoos, morning);
    logArrivalSourcesFromMorning(
      arrivalsFromMorning,
      `${morning.fileName} + ${ejoos.fileName}`,
    );
    return { ...departures, arrivalsFromMorning };
  } catch (error) {
    console.warn("[Звідки / Штатка]", error);
    return { ...departures, arrivalsFromMorning: null };
  }
};

/** Імена вибулих у консоль, згруповані за категоріями. */
export const logDeparturesByCategory = (
  result: DeparturesResult,
  fileName?: string,
) => {
  const byCategory: Record<string, Array<Record<string, string>>> = {};
  for (const category of DEPARTURE_CATEGORY_ORDER) {
    byCategory[CATEGORY_LABELS[category]] = [];
  }
  for (const person of result.people) {
    const bucket = byCategory[person.categoryLabel] ?? (byCategory[person.categoryLabel] = []);
    bucket.push({
      fullName: person.fullName,
      rank: person.rank,
      rankGroup: person.rankGroupLabel,
      personId: person.personId,
      excludeDate: person.excludeDate,
      orderDate: person.orderDate,
      ground: person.ground || person.type,
      destination: person.destination,
      matchNote: person.matchNote,
      excelRow: String(person.excelRow),
    });
  }

  console.group(
    `[Вибули] ${fileName || result.sourceSheet}${
      result.periodFrom ? ` · з ${result.periodFrom}` : ""
    } · ${result.totals.all} з ${result.totalUnfiltered}`,
  );
  console.log("Підсумок:", {
    усьогоУПеріоді: result.totals.all,
    офіцери: result.totals.byRank.officer,
    сержанти: result.totals.byRank.sergeant,
    солдати: result.totals.byRank.soldier,
    звільнення: result.totals.discharges,
    переведені: result.totals.transfers,
    доПеріоду: result.skippedBeforePeriod,
    безДати: result.skippedNoDate,
  });
  for (const row of result.summary) {
    const list = byCategory[row.label] ?? [];
    console.groupCollapsed(
      `${row.label} · ${list.length} (оф. ${row.byRank.officer}, серж. ${row.byRank.sergeant}, солд. ${row.byRank.soldier})`,
    );
    console.table(list);
    console.log("ПІБ за званням:", {
      офіцери: list
        .filter((item) => item.rankGroup === "Офіцери")
        .map((item) => item.fullName),
      сержанти: list
        .filter((item) => item.rankGroup === "Сержанти")
        .map((item) => item.fullName),
      солдати: list
        .filter((item) => item.rankGroup === "Солдати")
        .map((item) => item.fullName),
    });
    console.groupEnd();
  }
  console.groupEnd();
};
