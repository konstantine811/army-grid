import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { valueToDisplay } from "../../excelRoundTrip";
import {
  extractBchsAwayPeopleFromSheet,
  filterBchsNovaPeople,
  isBchsAvailableForAbsentFormula,
  isBchsListedWithRankTitle,
  isBchsPersonnelGeneralListSheet,
  isBchsWithoutStaffPosition,
  normalizeBchsText,
} from "../bchs/bchsCalc";
import type { BchsPersonnelAwayPerson } from "../bchs/bchsTypes";
import {
  ageBandOf,
  classifyArrivalSource,
  classifyNationality,
  classifyRankGroup,
  classifyRegion,
  classifyServiceType,
  classifySex,
  exitBandOf,
  hasIdpFlag,
  hasUbdFlag,
  looksLikePersonName,
  normalizeLooseText,
  normalizePersonName,
  parseAgeYears,
  parseRelatives,
  REGION_LABELS,
  shortPersonName,
} from "./socPassportFields";
import type {
  RankGroup,
  ServiceType,
  SocPerson,
  SocStaffSlot,
} from "./socPassportTypes";

const isBrezText = (value: string) => /бре[зc]|brez/i.test(normalizeLooseText(value));

const cellText = (value: CellValue | undefined) =>
  valueToDisplay(value ?? null).replace(/\s+/g, " ").trim();

const normalizeHeader = (value: CellValue | undefined) =>
  normalizeLooseText(cellText(value)).replace(/[.]/g, " ");

const findSheet = (
  workbook: ExcelWorkbookSnapshot | null | undefined,
  patterns: RegExp[],
) => {
  if (!workbook) return undefined;
  for (const pattern of patterns) {
    const match = workbook.sheets.find((sheet) =>
      pattern.test(normalizeLooseText(sheet.sheetName)),
    );
    if (match) return match;
  }
  return undefined;
};

const headerLooksLike = (row: CellValue[], partsGroups: string[][]) => {
  const headers = row.map((cell) => normalizeHeader(cell));
  return partsGroups.every((parts) =>
    headers.some((header) => parts.every((part) => header.includes(part))),
  );
};

const findHeaderRow = (sheet: ExcelSheetSnapshot, partsGroups: string[][]) => {
  const limit = Math.min(sheet.rawRows.length, 12);
  for (let index = 0; index < limit; index += 1) {
    const row = sheet.rawRows[index] ?? [];
    if (headerLooksLike(row, partsGroups)) return index;
  }
  return -1;
};

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

const readCell = (row: CellValue[] | undefined, index: number) =>
  index >= 0 ? cellText(row?.[index]) : "";

type NamedRow = {
  excelRowNumber: number;
  values: CellValue[];
};

const sheetDataRows = (sheet: ExcelSheetSnapshot, headerRowIndex: number): NamedRow[] =>
  sheet.rawRows.slice(headerRowIndex + 1).map((values, offset) => ({
    excelRowNumber: headerRowIndex + 2 + offset,
    values,
  }));

export const findShpoSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/1\.?\s*шпо/, /штатно.?посад/, /\bшпо\b/]);

export const findOosSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/2\.?\s*оос/, /\bоос\b/, /обл[іi]к особового/]);

export const findMorningRosterSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/ос.*загальн/, /загальн.*спис/]);

export const findFighterStatusSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/статус.?б[іi]йц/]);

type ShpoColumns = {
  index: number;
  position: number;
  staffRank: number;
  rank: number;
  name: number;
};

const resolveShpoColumns = (headerRow: CellValue[]): ShpoColumns => ({
  index: columnIndex(headerRow, ["індекс", "посад"]),
  position: columnIndex(headerRow, ["посад"], ["індекс"]),
  staffRank: columnIndex(headerRow, ["шпк"]),
  rank: columnIndex(headerRow, ["звання"]),
  name:
    columnIndex(headerRow, ["прізвищ"]) >= 0
      ? columnIndex(headerRow, ["прізвищ"])
      : columnIndex(headerRow, ["піб"]),
});

type OosColumns = {
  rank: number;
  name: number;
  id: number;
  positionIndex: number;
  arrivedFrom: number;
  serviceType: number;
  birthDate: number;
  birthPlace: number;
  sex: number;
  calledBy: number;
  relatives: number;
  extra: number;
};

const resolveOosColumns = (headerRow: CellValue[]): OosColumns => ({
  rank: columnIndex(headerRow, ["звання"]),
  name:
    columnIndex(headerRow, ["прізвищ"]) >= 0
      ? columnIndex(headerRow, ["прізвищ"])
      : columnIndex(headerRow, ["піб"]),
  id: columnIndex(headerRow, ["id"]),
  positionIndex: columnIndex(headerRow, ["індекс", "посад"]),
  arrivedFrom: columnIndex(headerRow, ["звідки", "прибув"]),
  serviceType: columnIndex(headerRow, ["вид", "служб"]),
  birthDate: columnIndex(headerRow, ["дата", "народж"]),
  birthPlace: columnIndex(headerRow, ["місце", "народж"]),
  sex: columnIndex(headerRow, ["стать"]),
  calledBy: columnIndex(headerRow, ["ким", "призван"]),
  relatives: columnIndex(headerRow, ["дані", "родич"]),
  extra: columnIndex(headerRow, ["додатков", "інформ"]),
});

type MorningColumns = {
  category: number;
  service: number;
  rank: number;
  name: number;
  years: number;
  status: number;
  brez: number;
};

const resolveMorningColumns = (headerRow: CellValue[]): MorningColumns => ({
  category:
    columnIndex(headerRow, ["шпк", "факт"]) >= 0
      ? columnIndex(headerRow, ["шпк", "факт"])
      : columnIndex(headerRow, ["оф"]),
  service: columnIndex(headerRow, ["мобіліз"]),
  rank: columnIndex(headerRow, ["звання"]),
  name:
    columnIndex(headerRow, ["піб"]) >= 0
      ? columnIndex(headerRow, ["піб"])
      : columnIndex(headerRow, ["прізвищ"]),
  years:
    columnIndex(headerRow, ["повних", "рок"]) >= 0
      ? columnIndex(headerRow, ["повних", "рок"])
      : columnIndex(headerRow, ["років"]),
  status: columnIndex(headerRow, ["статус"]),
  brez: columnIndex(headerRow, ["бре"]),
});

type StatusColumns = {
  name: number;
  exitDate: number;
  direction: number;
};

const resolveStatusColumns = (headerRow: CellValue[]): StatusColumns => ({
  name:
    columnIndex(headerRow, ["прізвищ"]) >= 0
      ? columnIndex(headerRow, ["прізвищ"])
      : columnIndex(headerRow, ["піб"]),
  exitDate: columnIndex(headerRow, ["дата", "вих"]),
  direction: columnIndex(headerRow, ["напрям"]),
});

type OosRecord = {
  name: string;
  rank: string;
  id: string;
  positionIndex: string;
  arrivedFrom: string;
  serviceType: ServiceType;
  birthDate: string;
  birthPlace: string;
  sex: ReturnType<typeof classifySex>;
  calledBy: string;
  relatives: string;
  extra: string;
};

type MorningRecord = {
  name: string;
  rank: string;
  category: string;
  years: string;
  status: string;
  brez: string;
};

const indexByName = <T extends { name: string }>(rows: T[]) => {
  const full = new Map<string, T[]>();
  const short = new Map<string, T[]>();
  for (const row of rows) {
    const fullKey = normalizePersonName(row.name);
    const shortKey = shortPersonName(row.name);
    if (fullKey) {
      const list = full.get(fullKey) ?? [];
      list.push(row);
      full.set(fullKey, list);
    }
    if (shortKey) {
      const list = short.get(shortKey) ?? [];
      list.push(row);
      short.set(shortKey, list);
    }
  }
  return { full, short };
};

const pickMatch = <T,>(
  name: string,
  indexes: { full: Map<string, T[]>; short: Map<string, T[]> },
) => {
  const fullKey = normalizePersonName(name);
  const shortKey = shortPersonName(name);
  const exact = indexes.full.get(fullKey);
  if (exact?.length) return exact[0];
  const shortened = indexes.short.get(shortKey);
  if (shortened?.length === 1) return shortened[0];
  return undefined;
};

const parseOosRows = (sheet: ExcelSheetSnapshot): OosRecord[] => {
  const headerRowIndex = findHeaderRow(sheet, [["звання"], ["прізвищ"]]);
  if (headerRowIndex < 0) return [];
  const columns = resolveOosColumns(sheet.rawRows[headerRowIndex] ?? []);
  if (columns.name < 0) return [];

  return sheetDataRows(sheet, headerRowIndex).flatMap((row) => {
    const name = readCell(row.values, columns.name);
    if (!looksLikePersonName(name)) return [];
    return [
      {
        name,
        rank: readCell(row.values, columns.rank),
        id: readCell(row.values, columns.id),
        positionIndex: readCell(row.values, columns.positionIndex),
        arrivedFrom: readCell(row.values, columns.arrivedFrom),
        serviceType: classifyServiceType(readCell(row.values, columns.serviceType)),
        birthDate: readCell(row.values, columns.birthDate),
        birthPlace: readCell(row.values, columns.birthPlace),
        sex: classifySex(readCell(row.values, columns.sex)),
        calledBy: readCell(row.values, columns.calledBy),
        relatives: readCell(row.values, columns.relatives),
        extra: readCell(row.values, columns.extra),
      },
    ];
  });
};

const parseMorningRows = (sheet: ExcelSheetSnapshot): MorningRecord[] => {
  const headerRowIndex = findHeaderRow(sheet, [["піб"], ["статус"]]);
  if (headerRowIndex < 0) return [];
  const columns = resolveMorningColumns(sheet.rawRows[headerRowIndex] ?? []);
  if (columns.name < 0) return [];

  return sheetDataRows(sheet, headerRowIndex).flatMap((row) => {
    const name = readCell(row.values, columns.name);
    if (!looksLikePersonName(name)) return [];
    return [
      {
        name,
        rank: readCell(row.values, columns.rank),
        category: readCell(row.values, columns.category),
        years: readCell(row.values, columns.years),
        status: readCell(row.values, columns.status),
        brez: readCell(row.values, columns.brez),
      },
    ];
  });
};

const parseExitCounts = (sheet: ExcelSheetSnapshot) => {
  const headerRowIndex = findHeaderRow(sheet, [["прізвищ"], ["дата"]]);
  const counts = new Map<string, Set<string>>();
  if (headerRowIndex < 0) return counts;
  const columns = resolveStatusColumns(sheet.rawRows[headerRowIndex] ?? []);
  if (columns.name < 0) return counts;

  for (const row of sheetDataRows(sheet, headerRowIndex)) {
    const name = readCell(row.values, columns.name);
    if (!looksLikePersonName(name)) continue;
    const key = normalizePersonName(name);
    const shortKey = shortPersonName(name);
    const stamp =
      readCell(row.values, columns.exitDate) ||
      readCell(row.values, columns.direction) ||
      `row-${row.excelRowNumber}`;
    for (const mapKey of [key, shortKey]) {
      if (!mapKey) continue;
      const set = counts.get(mapKey) ?? new Set<string>();
      set.add(stamp);
      counts.set(mapKey, set);
    }
  }
  return counts;
};

const rankGroupOf = (
  personRank: string,
  staffRank: string,
  position: string,
  morningCategory: string,
): RankGroup => {
  if (morningCategory) return classifyRankGroup(morningCategory, position);
  if (personRank) return classifyRankGroup(personRank, position);
  if (staffRank) return classifyRankGroup(staffRank, position);
  return classifyRankGroup("", position);
};

const isBchsStaffSlot = (person: BchsPersonnelAwayPerson) => {
  const rank = normalizeBchsText(person.rankCategory);
  if (rank === "оф." || rank === "серж." || rank.includes("солд")) return true;
  return Boolean(person.position.trim()) && !isBchsWithoutStaffPosition(person);
};

const parseShpoNameIndex = (sheet: ExcelSheetSnapshot | undefined) => {
  const map = new Map<string, { positionIndex: string; position: string; rank: string; staffRank: string }>();
  if (!sheet) return map;
  const headerRowIndex = findHeaderRow(sheet, [["посад"], ["прізвищ"]]);
  if (headerRowIndex < 0) return map;
  const columns = resolveShpoColumns(sheet.rawRows[headerRowIndex] ?? []);
  if (columns.name < 0) return map;

  for (const row of sheetDataRows(sheet, headerRowIndex)) {
    const name = readCell(row.values, columns.name);
    if (!looksLikePersonName(name)) continue;
    const key = normalizePersonName(name);
    if (!key || map.has(key)) continue;
    map.set(key, {
      positionIndex: readCell(row.values, columns.index),
      position: readCell(row.values, columns.position),
      rank: readCell(row.values, columns.rank),
      staffRank: readCell(row.values, columns.staffRank),
    });
  }
  return map;
};

export const parseSocPassportSources = ({
  ejoos,
  morning,
  asOf = new Date(),
}: {
  ejoos: ExcelWorkbookSnapshot;
  morning?: ExcelWorkbookSnapshot | null;
  asOf?: Date;
}) => {
  if (!morning) {
    throw new Error("Завантажте ранковий звіт — штат, список і наявність рахуються як у БЧС.");
  }

  const shpoSheet = findShpoSheet(ejoos);
  const oosSheet = findOosSheet(ejoos);
  const morningSheet =
    morning.sheets.find(isBchsPersonnelGeneralListSheet) ??
    findMorningRosterSheet(morning);
  const statusSheet = findFighterStatusSheet(morning);

  if (!oosSheet) {
    throw new Error("У файлі ЕЖООС не знайдено аркуш «2. ООС».");
  }
  if (!morningSheet) {
    throw new Error("У ранковому звіті не знайдено аркуш «1.ОС Загальний список».");
  }

  const oosRows = parseOosRows(oosSheet);
  const morningYearRows = parseMorningRows(morningSheet);
  const exitCounts = statusSheet ? parseExitCounts(statusSheet) : new Map<string, Set<string>>();
  const oosIndex = indexByName(oosRows);
  const oosByPosition = new Map<string, OosRecord>();
  for (const row of oosRows) {
    const key = row.positionIndex.replace(/\s+/g, "");
    if (key && !oosByPosition.has(key)) oosByPosition.set(key, row);
  }
  const yearIndex = indexByName(morningYearRows);
  const shpoByName = parseShpoNameIndex(shpoSheet);

  const novaPeople = filterBchsNovaPeople(extractBchsAwayPeopleFromSheet(morningSheet));
  const staffSlots: SocStaffSlot[] = novaPeople.filter(isBchsStaffSlot).map((person) => ({
    position: person.position,
    positionIndex: "",
    staffRank: person.shpkFact || person.rankCategory,
    rankGroup: rankGroupOf(person.rankTitle, person.shpkFact, person.position, person.rankCategory),
    occupied: looksLikePersonName(person.fullName),
    name: looksLikePersonName(person.fullName) ? person.fullName : "",
  }));

  const people: SocPerson[] = [];
  const seenNames = new Set<string>();
  const matchedOosBrezKeys = new Set<string>();
  const matchedOosUbdKeys = new Set<string>();

  const oosBrezAll = oosRows
    .filter((row) => isBrezText(row.arrivedFrom))
    .map((row) => ({
      name: row.name,
      normalizedName: normalizePersonName(row.name),
      shortName: shortPersonName(row.name),
      arrivedFrom: row.arrivedFrom,
      rank: row.rank,
      serviceType: row.serviceType,
      positionIndex: row.positionIndex,
    }));

  const oosUbdAll = oosRows
    .filter((row) => hasUbdFlag(row.extra, row.relatives))
    .map((row) => ({
      name: row.name,
      normalizedName: normalizePersonName(row.name),
      shortName: shortPersonName(row.name),
      rank: row.rank,
      serviceType: row.serviceType,
      positionIndex: row.positionIndex,
      relatives: row.relatives,
      extra: row.extra,
      ubdHitIn:
        /убд|удб/i.test(row.extra) && /убд|удб/i.test(row.relatives)
          ? "extra+relatives"
          : /убд|удб/i.test(row.extra)
            ? "extra"
            : "relatives",
    }));

  for (const roster of novaPeople) {
    if (!isBchsListedWithRankTitle(roster)) continue;
    const name = roster.fullName.trim();
    if (!looksLikePersonName(name)) continue;
    const normalizedName = normalizePersonName(name);
    if (seenNames.has(normalizedName)) continue;
    seenNames.add(normalizedName);

    const shpo = shpoByName.get(normalizedName);
    const oos =
      (shpo?.positionIndex && oosByPosition.get(shpo.positionIndex.replace(/\s+/g, ""))) ||
      pickMatch(name, oosIndex);
    if (oos && isBrezText(oos.arrivedFrom)) {
      matchedOosBrezKeys.add(normalizePersonName(oos.name));
    }
    if (oos && hasUbdFlag(oos.extra, oos.relatives)) {
      matchedOosUbdKeys.add(normalizePersonName(oos.name));
    }
    const years = pickMatch(name, yearIndex)?.years ?? "";
    const relatives = parseRelatives(oos?.relatives ?? "", asOf);
    const region = classifyRegion(oos?.birthPlace ?? "", oos?.calledBy ?? "");
    const nationality = classifyNationality(
      oos?.relatives ?? "",
      oos?.extra ?? "",
      oos?.birthPlace ?? "",
    );
    const age = parseAgeYears(oos?.birthDate ?? "", years, asOf);
    const present = isBchsAvailableForAbsentFormula(roster.status);
    const arrivalSource = classifyArrivalSource(
      oos?.arrivedFrom ?? "",
      "",
      roster.bzvpStatus,
    );
    const regionOccupied = region.occupied || /Occupied|crimea/.test(region.key);
    const parseNotes = [
      ...relatives.notes,
      ...region.notes,
      ...nationality.notes,
    ];
    if (!oos) parseNotes.push("немає рядка в ООС");
    const exitSet =
      exitCounts.get(normalizedName) ??
      exitCounts.get(shortPersonName(name));
    const exitCount = exitSet?.size ?? 0;

    people.push({
      id: `${normalizedName}:${roster.position}`,
      name,
      normalizedName,
      shortName: shortPersonName(name),
      position: roster.position || shpo?.position || "",
      positionIndex: shpo?.positionIndex ?? "",
      rank: roster.rankTitle || oos?.rank || shpo?.rank || "",
      staffRank: roster.shpkFact || shpo?.staffRank || "",
      rankGroup: rankGroupOf(
        roster.rankTitle || oos?.rank || "",
        roster.shpkFact,
        roster.position,
        roster.rankCategory,
      ),
      serviceType: oos?.serviceType ?? "mobilized",
      sex: oos?.sex && oos.sex !== "unknown" ? oos.sex : "male",
      birthDate: oos?.birthDate ?? "",
      age,
      ageBand: ageBandOf(age),
      birthPlace: oos?.birthPlace ?? "",
      region: region.key,
      regionLabel: REGION_LABELS[region.key],
      regionOccupied,
      nationality: nationality.key,
      marital:
        relatives.marital === "unknown" ? "unmarried" : relatives.marital,
      childrenUnder18: relatives.childrenUnder18,
      children3plus: relatives.childrenUnder18 >= 3,
      relativesServing: relatives.relativesServing,
      relativesAbroad: relatives.relativesAbroad,
      relativesHostile: relatives.relativesHostile,
      relativesRaw: oos?.relatives ?? "",
      extraRaw: oos?.extra ?? "",
      arrivedFrom: oos?.arrivedFrom ?? "",
      calledBy: oos?.calledBy ?? "",
      arrivalSource,
      hasUbd: hasUbdFlag(oos?.extra ?? "", oos?.relatives ?? ""),
      isIdp: hasIdpFlag(oos?.extra ?? "", oos?.relatives ?? "", oos?.birthPlace ?? ""),
      morningStatus: roster.status,
      bzvpStatus: roster.bzvpStatus ?? "",
      onStaff: !isBchsWithoutStaffPosition(roster),
      onList: true,
      present,
      inDisposition: !present,
      exitCount,
      exitBand: exitBandOf(exitCount),
      match: {
        oos: Boolean(oos),
        morning: true,
        exits: Boolean(exitSet?.size),
      },
      parseNotes,
    });
  }

  const oosBrezKeptInMorning = oosBrezAll.filter((row) =>
    matchedOosBrezKeys.has(row.normalizedName),
  );
  const oosBrezDroppedByMorning = oosBrezAll.filter(
    (row) => !matchedOosBrezKeys.has(row.normalizedName),
  );

  const countedBrezFromOos = people
    .filter((person) => person.arrivalSource === "brez" && isBrezText(person.arrivedFrom))
    .map((person) => ({
      name: person.name,
      shortName: person.shortName,
      arrivedFrom: person.arrivedFrom,
      bzvpStatus: person.bzvpStatus,
      present: person.present,
      inDisposition: person.inDisposition,
    }));

  const countedBrezFromMorningOnly = people
    .filter((person) => person.arrivalSource === "brez" && !isBrezText(person.arrivedFrom))
    .map((person) => ({
      name: person.name,
      shortName: person.shortName,
      arrivedFrom: person.arrivedFrom,
      bzvpStatus: person.bzvpStatus,
      present: person.present,
      inDisposition: person.inDisposition,
    }));

  const countedBrezAll = people
    .filter((person) => person.arrivalSource === "brez")
    .map((person) => ({
      name: person.name,
      shortName: person.shortName,
      arrivedFrom: person.arrivedFrom,
      bzvpStatus: person.bzvpStatus,
      source: isBrezText(person.arrivedFrom)
        ? "oos"
        : isBrezText(person.bzvpStatus)
          ? "morning"
          : "both/unknown",
      present: person.present,
      inDisposition: person.inDisposition,
    }));

  const brezDebug = {
    oosBrezAll,
    oosBrezKeptInMorning,
    oosBrezDroppedByMorning,
    countedBrezFromOos,
    countedBrezFromMorningOnly,
    countedBrezAll,
    counts: {
      oosBrezAll: oosBrezAll.length,
      oosBrezKeptInMorning: oosBrezKeptInMorning.length,
      oosBrezDroppedByMorning: oosBrezDroppedByMorning.length,
      countedBrezFromOos: countedBrezFromOos.length,
      countedBrezFromMorningOnly: countedBrezFromMorningOnly.length,
      countedBrezAll: countedBrezAll.length,
    },
  };

  const oosUbdKeptInMorning = oosUbdAll.filter((row) =>
    matchedOosUbdKeys.has(row.normalizedName),
  );
  const oosUbdDroppedByMorning = oosUbdAll.filter(
    (row) => !matchedOosUbdKeys.has(row.normalizedName),
  );

  const countedUbdInMorning = people
    .filter((person) => person.hasUbd)
    .map((person) => ({
      name: person.name,
      shortName: person.shortName,
      rank: person.rank,
      rankGroup: person.rankGroup,
      serviceType: person.serviceType,
      present: person.present,
      inDisposition: person.inDisposition,
      relativesRaw: person.relativesRaw,
      extraRaw: person.extraRaw,
    }));

  const ubdDebug = {
    oosUbdAll,
    oosUbdKeptInMorning,
    oosUbdDroppedByMorning,
    countedUbdInMorning,
    counts: {
      oosUbdAll: oosUbdAll.length,
      oosUbdKeptInMorning: oosUbdKeptInMorning.length,
      oosUbdDroppedByMorning: oosUbdDroppedByMorning.length,
      countedUbdInMorning: countedUbdInMorning.length,
    },
  };

  return {
    people,
    staffSlots,
    sheets: {
      shpo: shpoSheet?.sheetName ?? "",
      oos: oosSheet.sheetName,
      morning: morningSheet.sheetName,
      fighterStatus: statusSheet?.sheetName,
    },
    oosRowCount: oosRows.length,
    morningRowCount: novaPeople.length,
    brezDebug,
    ubdDebug,
  };
};
