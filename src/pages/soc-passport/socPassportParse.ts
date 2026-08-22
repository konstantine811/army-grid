import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { valueToDisplay } from "../../excelRoundTrip";
import {
  extractBchsAwayPeopleFromSheet,
  filterBchsNovaPeople,
  isBchsAvailableForAbsentFormula,
  isBchsListedWithRankTitle,
  isBchsMissingStatus,
  isBchsPersonnelGeneralListSheet,
  isBchsWithoutStaffPosition,
  normalizeBchsText,
} from "../bchs/bchsCalc";
import type { BchsPersonnelAwayPerson } from "../bchs/bchsTypes";
import {
  ageBandOf,
  buildUbdRosterLookup,
  classifyArrivalSource,
  combineMorningBrezFields,
  isBrezMarkerText,
  classifyNationality,
  classifyRankGroup,
  classifyRegion,
  classifyServiceType,
  classifySex,
  exitBandOf,
  hasIdpFlag,
  hasUbdFlag,
  isCombatZoneDislocation,
  isTransiterDestination,
  isRrebDestination,
  isPpoDestination,
  lookupExitStampSet,
  looksLikePersonName,
  matchUbdRosterRecord,
  mergeExitCountMaps,
  namesLikelySamePerson,
  normalizeExitDateKey,
  normalizeLooseText,
  normalizePersonName,
  normalizeRnokpp,
  parseAgeYears,
  parseRelatives,
  parseUbdInfo,
  REGION_LABELS,
  resolveCombatExitCount,
  shortPersonName,
  type UbdRosterMatchRecord,
  type UbdRosterStatus,
} from "./socPassportFields";
import type {
  RankGroup,
  ServiceType,
  SocPerson,
  SocStaffSlot,
} from "./socPassportTypes";

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

export const findTempArrivedSheet = (workbook: ExcelWorkbookSnapshot) =>
  findSheet(workbook, [/4\.?\s*тимчасово\s*прибу/, /тимчасово\s*прибулі/]);

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
  location: number;
};

const resolveOosColumns = (headerRow: CellValue[]): OosColumns => ({
  rank: columnIndex(headerRow, ["звання"]),
  name:
    columnIndex(headerRow, ["прізвищ"]) >= 0
      ? columnIndex(headerRow, ["прізвищ"])
      : columnIndex(headerRow, ["піб"]),
  id: columnIndex(headerRow, ["id"]),
  rnokpp: columnIndex(headerRow, ["рнокпп"]),
  positionIndex: columnIndex(headerRow, ["індекс", "посад"]),
  arrivedFrom: columnIndex(headerRow, ["звідки", "прибув"]),
  serviceType: columnIndex(headerRow, ["вид", "служб"]),
  birthDate: columnIndex(headerRow, ["дата", "народж"]),
  birthPlace: columnIndex(headerRow, ["місце", "народж"]),
  sex: columnIndex(headerRow, ["стать"]),
  calledBy: columnIndex(headerRow, ["ким", "призван"]),
  relatives: columnIndex(headerRow, ["дані", "родич"]),
  extra: columnIndex(headerRow, ["додатков", "інформ"]),
  location: columnIndex(headerRow, ["місце", "дислок"]),
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
  rnokpp: string;
  positionIndex: string;
  arrivedFrom: string;
  serviceType: ServiceType;
  birthDate: string;
  birthPlace: string;
  sex: ReturnType<typeof classifySex>;
  calledBy: string;
  relatives: string;
  extra: string;
  location: string;
};

type TempArrivalRecord = {
  name: string;
  rank: string;
  arrivedFrom: string;
  arrivalDate: string;
  dislocation: string;
  departureDate: string;
  extra: string;
  inCombatZone: boolean;
};

type UbdRosterRecord = UbdRosterMatchRecord;

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
        rnokpp: normalizeRnokpp(readCell(row.values, columns.rnokpp)),
        positionIndex: readCell(row.values, columns.positionIndex),
        arrivedFrom: readCell(row.values, columns.arrivedFrom),
        serviceType: classifyServiceType(readCell(row.values, columns.serviceType)),
        birthDate: readCell(row.values, columns.birthDate),
        birthPlace: readCell(row.values, columns.birthPlace),
        sex: classifySex(readCell(row.values, columns.sex)),
        calledBy: readCell(row.values, columns.calledBy),
        relatives: readCell(row.values, columns.relatives),
        extra: readCell(row.values, columns.extra),
        location: readCell(row.values, columns.location),
      },
    ];
  });
};

type TempArrivalColumns = {
  rank: number;
  name: number;
  arrivedFrom: number;
  arrivalDate: number;
  dislocation: number;
  departureDate: number;
  extra: number;
};

const resolveTempArrivalColumns = (headerRow: CellValue[]): TempArrivalColumns => ({
  rank: columnIndex(headerRow, ["звання"]),
  name:
    columnIndex(headerRow, ["прізвищ"]) >= 0
      ? columnIndex(headerRow, ["прізвищ"])
      : columnIndex(headerRow, ["піб"]),
  arrivedFrom: columnIndex(headerRow, ["звідки", "прибув"]),
  arrivalDate: columnIndex(headerRow, ["дата", "прибут"]),
  dislocation: columnIndex(headerRow, ["місце", "дислок"]),
  departureDate: columnIndex(headerRow, ["дата", "вибут"]),
  extra: columnIndex(headerRow, ["додатков", "інформ"]),
});

const parseTempArrivalRows = (sheet: ExcelSheetSnapshot): TempArrivalRecord[] => {
  const headerRowIndex = findHeaderRow(sheet, [["піб"], ["звання"], ["прибув"]]);
  if (headerRowIndex < 0) return [];
  const columns = resolveTempArrivalColumns(sheet.rawRows[headerRowIndex] ?? []);
  if (columns.name < 0) return [];

  return sheetDataRows(sheet, headerRowIndex).flatMap((row) => {
    const name = readCell(row.values, columns.name);
    if (!looksLikePersonName(name)) return [];
    const dislocation = readCell(row.values, columns.dislocation);
    return [
      {
        name,
        rank: readCell(row.values, columns.rank),
        arrivedFrom: readCell(row.values, columns.arrivedFrom),
        arrivalDate: readCell(row.values, columns.arrivalDate),
        dislocation,
        departureDate: readCell(row.values, columns.departureDate),
        extra: readCell(row.values, columns.extra),
        inCombatZone: isCombatZoneDislocation(dislocation),
      },
    ];
  });
};

const collectCombatDutyEvidence = ({
  ubdNumber,
  hasUbd,
  ubdRosterStatus,
  oosDislocation,
  tempArrivals,
  bplaPerformed,
  morningMissing,
  rrebDestination,
  ppoDestination,
}: {
  ubdNumber: string;
  hasUbd: boolean;
  ubdRosterStatus: UbdRosterStatus | null;
  oosDislocation: string;
  tempArrivals: TempArrivalRecord[];
  bplaPerformed: boolean;
  morningMissing: boolean;
  rrebDestination: boolean;
  ppoDestination: boolean;
}) => {
  const evidence: string[] = [];
  if (ubdNumber) evidence.push(`УБД №${ubdNumber}`);
  else if (hasUbd) evidence.push("УБД (анкета ООС)");
  if (ubdRosterStatus === "submitted") evidence.push("УБД реєстр: подавалися");
  else if (ubdRosterStatus === "notSubmitted") {
    evidence.push("УБД реєстр: не подавалися");
  }
  if (bplaPerformed) evidence.push("БПЛА: виконував БЗ");
  if (morningMissing) evidence.push("Ранковий: Зниклі безвісти");
  if (rrebDestination) evidence.push("Ранковий: підрозділ РРЕБ");
  if (ppoDestination) evidence.push("Ранковий: підрозділ ППО");
  if (isCombatZoneDislocation(oosDislocation)) {
    evidence.push(`ООС: ${oosDislocation}`);
  }
  for (const row of tempArrivals) {
    if (!row.inCombatZone) continue;
    const period = [row.arrivalDate, row.departureDate].filter(Boolean).join("–");
    evidence.push(
      `Тимч. прибуття${period ? ` (${period})` : ""}: ${row.dislocation}`,
    );
  }
  return evidence;
};

const isUbdSubmittedSheetName = (sheetName: string) => {
  const text = normalizeLooseText(sheetName);
  return text.includes("подавали") && !text.includes("не подавали");
};

const isUbdNotSubmittedSheetName = (sheetName: string) =>
  normalizeLooseText(sheetName).includes("не подавали");

type UbdRosterNameColumns = {
  combined: number;
  last: number;
  first: number;
  patronymic: number;
  rnokpp: number;
};

const resolveUbdSubmittedColumns = (headerRow: CellValue[]): UbdRosterNameColumns => {
  const combined = headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    return header.includes("прізвищ") && header.includes("ім");
  });
  return {
    combined:
      combined >= 0
        ? combined
        : columnIndex(headerRow, ["прізвищ"], ["позивн"]),
    last: columnIndex(headerRow, ["прізвищ"], ["позивн"]),
    first: columnIndex(headerRow, ["ім"], ["прізвищ", "батьк"]),
    patronymic: columnIndex(headerRow, ["батьк"]),
    rnokpp: columnIndex(headerRow, ["рнокпп", "обліков"]),
  };
};

const resolveUbdNotSubmittedColumns = (headerRow: CellValue[]): UbdRosterNameColumns => {
  const combined = headerRow.findIndex((cell, index) => {
    const header = normalizeHeader(cell);
    return index <= 3 && header.includes("прізвищ") && header.includes("ім");
  });
  return {
    combined: combined >= 0 ? combined : 1,
    last: columnIndex(headerRow, ["прізвищ"], ["позивн", "ім"]),
    first: columnIndex(headerRow, ["ім"], ["прізвищ", "батьк"]),
    patronymic: columnIndex(headerRow, ["батьк"]),
    rnokpp: columnIndex(headerRow, ["рнокпп", "обліков"]),
  };
};

const readUbdRosterName = (row: CellValue[], columns: UbdRosterNameColumns) => {
  const combined = readCell(row, columns.combined);
  if (looksLikePersonName(combined)) return combined;
  const assembled = [
    readCell(row, columns.last),
    readCell(row, columns.first),
    readCell(row, columns.patronymic),
  ]
    .map((part) => part.trim())
    .filter(Boolean)
    .join(" ");
  return looksLikePersonName(assembled) ? assembled : "";
};

const parseUbdRosterSheet = (
  sheet: ExcelSheetSnapshot,
  status: UbdRosterStatus,
): UbdRosterRecord[] => {
  const headerRowIndex = findHeaderRow(sheet, [["прізвищ"]]);
  if (headerRowIndex < 0) return [];
  const headerRow = sheet.rawRows[headerRowIndex] ?? [];
  const columns =
    status === "submitted"
      ? resolveUbdSubmittedColumns(headerRow)
      : resolveUbdNotSubmittedColumns(headerRow);

  return sheetDataRows(sheet, headerRowIndex).flatMap((row) => {
    const name = readUbdRosterName(row.values, columns);
    if (!looksLikePersonName(name)) return [];
    return [
      {
        name,
        rnokpp: normalizeRnokpp(readCell(row.values, columns.rnokpp)),
        status,
        sheetName: sheet.sheetName,
      },
    ];
  });
};

const parseUbdRosterRecords = (workbook: ExcelWorkbookSnapshot) => {
  const records: UbdRosterRecord[] = [];
  for (const sheet of workbook.sheets) {
    if (isUbdSubmittedSheetName(sheet.sheetName)) {
      records.push(...parseUbdRosterSheet(sheet, "submitted"));
    } else if (isUbdNotSubmittedSheetName(sheet.sheetName)) {
      records.push(...parseUbdRosterSheet(sheet, "notSubmitted"));
    }
  }

  const byName = new Map<string, UbdRosterRecord>();
  for (const record of records) {
    const key = normalizePersonName(record.name);
    if (!key) continue;
    const existing = byName.get(key);
    if (!existing || record.status === "submitted") {
      byName.set(key, record);
    }
  }
  return [...byName.values()];
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
    const exitDateRaw = readCell(row.values, columns.exitDate);
    const stamp =
      normalizeExitDateKey(exitDateRaw) ??
      normalizeExitDateKey(readCell(row.values, columns.direction)) ??
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

type JbdExitColumns = {
  name: number;
  exitDate: number;
};

const resolveJbdExitColumns = (headerRow: CellValue[]): JbdExitColumns => ({
  name:
    columnIndex(headerRow, ["фіо"]) >= 0
      ? columnIndex(headerRow, ["фіо"])
      : columnIndex(headerRow, ["піб"]),
  exitDate: columnIndex(headerRow, ["дата", "вих"]),
});

const parseJbdExitSheet = (sheet: ExcelSheetSnapshot) => {
  const headerRowIndex = findHeaderRow(sheet, [["фіо"], ["дата", "вих"]]);
  if (headerRowIndex < 0) return new Map<string, Set<string>>();
  const columns = resolveJbdExitColumns(sheet.rawRows[headerRowIndex] ?? []);
  if (columns.name < 0 || columns.exitDate < 0) return new Map<string, Set<string>>();

  const counts = new Map<string, Set<string>>();
  for (const row of sheetDataRows(sheet, headerRowIndex)) {
    const name = readCell(row.values, columns.name);
    if (!looksLikePersonName(name)) continue;
    const exitDateRaw = readCell(row.values, columns.exitDate);
    const stamp =
      normalizeExitDateKey(exitDateRaw) ??
      normalizeExitDateKey(row.values[columns.exitDate]) ??
      null;
    if (!stamp) continue;
    const key = normalizePersonName(name);
    const shortKey = shortPersonName(name);
    for (const mapKey of [key, shortKey]) {
      if (!mapKey) continue;
      const set = counts.get(mapKey) ?? new Set<string>();
      set.add(stamp);
      counts.set(mapKey, set);
    }
  }
  return counts;
};

const parseJbdExitCounts = (workbook: ExcelWorkbookSnapshot) => {
  const merged = new Map<string, Set<string>>();
  for (const sheet of workbook.sheets) {
    const partial = parseJbdExitSheet(sheet);
    for (const [key, stamps] of partial) {
      const set = merged.get(key) ?? new Set<string>();
      for (const stamp of stamps) set.add(stamp);
      merged.set(key, set);
    }
  }
  return merged;
};

type BplaPerformerRecord = {
  name: string;
  callsign: string;
  rank: string;
};

type BplaPerformerLookup = {
  keys: Set<string>;
  records: Map<string, BplaPerformerRecord>;
};

const isBplaPerformedStatus = (status: string) => {
  const text = normalizeLooseText(status);
  return text.includes("виконував") && text.includes("бз") && !text.includes("не виконував");
};

const isBplaStatusCell = (value: string) => {
  const text = normalizeLooseText(value);
  return (
    (text.includes("виконував") && text.includes("бз")) ||
    (text.includes("не") && text.includes("виконував") && text.includes("бз"))
  );
};

/** ПІБ + статус БЗ у рядку (колонки зміщуються після normalize у readWorkbookSnapshot). */
const resolveBplaColumnsFromRow = (row: CellValue[]) => {
  let nameCol = -1;
  let statusCol = -1;
  for (let index = 0; index < row.length; index += 1) {
    const text = readCell(row, index);
    if (!text) continue;
    if (nameCol < 0 && looksLikePersonName(text)) {
      nameCol = index;
      continue;
    }
    if (statusCol < 0 && isBplaStatusCell(text)) {
      statusCol = index;
    }
  }
  if (nameCol < 0 || statusCol < 0) return null;
  const callsignCol =
    nameCol + 1 < statusCol && nameCol + 1 >= 0 ? nameCol + 1 : -1;
  const rankCol = nameCol > 0 ? nameCol - 1 : -1;
  return { nameCol, statusCol, callsignCol, rankCol };
};

const addBplaPerformer = (
  performers: Map<string, BplaPerformerRecord>,
  row: CellValue[],
  nameCol: number,
  statusCol: number,
  callsignCol: number,
  rankCol: number,
) => {
  const name = readCell(row, nameCol);
  const status = readCell(row, statusCol);
  if (!looksLikePersonName(name) || !isBplaPerformedStatus(status)) return;
  const key = normalizePersonName(name);
  if (!key || performers.has(key)) return;
  performers.set(key, {
    name,
    callsign: callsignCol >= 0 ? readCell(row, callsignCol) : "",
    rank: rankCol >= 0 ? readCell(row, rankCol) : "",
  });
};

const parseBplaExitSheet = (sheet: ExcelSheetSnapshot) => {
  const performers = new Map<string, BplaPerformerRecord>();
  const headerRowIndex = findHeaderRow(sheet, [["піб"], ["фіо"], ["прізвищ"]]);
  if (headerRowIndex >= 0) {
    const headerRow = sheet.rawRows[headerRowIndex] ?? [];
    const nameCol =
      columnIndex(headerRow, ["піб"]) >= 0
        ? columnIndex(headerRow, ["піб"])
        : columnIndex(headerRow, ["фіо"]) >= 0
          ? columnIndex(headerRow, ["фіо"])
          : columnIndex(headerRow, ["прізвищ"]);
    const statusCol =
      columnIndex(headerRow, ["бз"]) >= 0
        ? columnIndex(headerRow, ["бз"])
        : columnIndex(headerRow, ["статус"]);
    if (nameCol >= 0 && statusCol >= 0) {
      const callsignCol = columnIndex(headerRow, ["позивн"]);
      const rankCol = columnIndex(headerRow, ["звання"]);
      for (const row of sheetDataRows(sheet, headerRowIndex)) {
        addBplaPerformer(
          performers,
          row.values,
          nameCol,
          statusCol,
          callsignCol,
          rankCol,
        );
      }
      return performers;
    }
  }

  for (const row of sheet.rawRows) {
    const columns = resolveBplaColumnsFromRow(row);
    if (!columns) continue;
    addBplaPerformer(
      performers,
      row,
      columns.nameCol,
      columns.statusCol,
      columns.callsignCol,
      columns.rankCol,
    );
  }
  return performers;
};

const parseBplaPerformers = (workbook: ExcelWorkbookSnapshot): BplaPerformerLookup => {
  const records = new Map<string, BplaPerformerRecord>();
  for (const sheet of workbook.sheets) {
    const partial = parseBplaExitSheet(sheet);
    for (const [key, record] of partial) {
      if (!records.has(key)) records.set(key, record);
    }
  }
  const keys = new Set<string>();
  for (const [key, record] of records) {
    keys.add(key);
    const shortKey = shortPersonName(record.name);
    if (shortKey) keys.add(shortKey);
  }
  return { keys, records };
};

const findBplaPerformer = (
  lookup: BplaPerformerLookup | null,
  name: string,
  normalizedName: string,
  shortName: string,
): BplaPerformerRecord | null => {
  if (!lookup) return null;
  const byFull = lookup.records.get(normalizedName);
  if (byFull) return byFull;
  for (const record of lookup.records.values()) {
    if (shortPersonName(record.name) === shortName) return record;
    if (namesLikelySamePerson(name, record.name)) return record;
  }
  return null;
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
  jbdExits,
  bplaExits,
  ubdRoster,
  asOf = new Date(),
}: {
  ejoos: ExcelWorkbookSnapshot;
  morning?: ExcelWorkbookSnapshot | null;
  jbdExits?: ExcelWorkbookSnapshot | null;
  bplaExits?: ExcelWorkbookSnapshot | null;
  ubdRoster?: ExcelWorkbookSnapshot | null;
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
  const tempArrivedSheet = findTempArrivedSheet(ejoos);

  if (!oosSheet) {
    throw new Error("У файлі ЕЖООС не знайдено аркуш «2. ООС».");
  }
  if (!morningSheet) {
    throw new Error("У ранковому звіті не знайдено аркуш «1.ОС Загальний список».");
  }

  const oosRows = parseOosRows(oosSheet);
  const tempArrivalRows = tempArrivedSheet
    ? parseTempArrivalRows(tempArrivedSheet)
    : [];
  const tempArrivalIndex = indexByName(tempArrivalRows);
  const morningYearRows = parseMorningRows(morningSheet);
  const morningExitCounts = statusSheet ? parseExitCounts(statusSheet) : new Map();
  const jbdExitCounts = jbdExits ? parseJbdExitCounts(jbdExits) : new Map();
  const bplaPerformers = bplaExits ? parseBplaPerformers(bplaExits) : null;
  const exitCounts = mergeExitCountMaps(morningExitCounts, jbdExitCounts);
  const ubdRosterRows = ubdRoster ? parseUbdRosterRecords(ubdRoster) : [];
  const ubdRosterLookup = buildUbdRosterLookup(ubdRosterRows);
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
    .filter((row) => isBrezMarkerText(row.arrivedFrom))
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
    if (oos && isBrezMarkerText(oos.arrivedFrom)) {
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
    const morningBrez = combineMorningBrezFields(
      roster.bzvpStatus,
      roster.brezAssignment,
    );
    const arrivalSource = classifyArrivalSource(
      oos?.arrivedFrom ?? "",
      "",
      morningBrez,
    );
    const regionOccupied = region.occupied || /Occupied|crimea/.test(region.key);
    const parseNotes = [
      ...relatives.notes,
      ...region.notes,
      ...nationality.notes,
    ];
    if (!oos) parseNotes.push("немає рядка в ООС");
    const shortName = shortPersonName(name);
    const morningExitSet = lookupExitStampSet(
      morningExitCounts,
      normalizedName,
      shortName,
    );
    const jbdExitSet = lookupExitStampSet(jbdExitCounts, normalizedName, shortName);
    const mergedExitSet = lookupExitStampSet(exitCounts, normalizedName, shortName);
    const ubdInfo = parseUbdInfo(oos?.extra ?? "", oos?.relatives ?? "");
    const ubdRosterRecord = matchUbdRosterRecord(
      name,
      oos?.rnokpp ?? "",
      ubdRosterLookup,
      [oos?.name ?? ""],
    );
    const ubdRosterStatus = ubdRosterRecord?.status ?? null;
    const tempArrivals = [
      ...(tempArrivalIndex.full.get(normalizedName) ?? []),
      ...(tempArrivalIndex.short.get(shortPersonName(name)) ?? []),
    ].filter(
      (row, index, list) =>
        list.findIndex(
          (other) =>
            other.name === row.name &&
            other.arrivalDate === row.arrivalDate &&
            other.dislocation === row.dislocation,
        ) === index,
    );
    const oosDislocation = oos?.location ?? "";
    const isTransiter = isTransiterDestination(roster.destination);
    const bplaMatch = findBplaPerformer(
      bplaPerformers,
      name,
      normalizedName,
      shortName,
    );
    const bplaPerformed = !isTransiter && Boolean(bplaMatch);
    const callsign = (roster.callsign ?? "").trim() || bplaMatch?.callsign || "";
    const morningMissing = !isTransiter && isBchsMissingStatus(roster.status);
    const rrebDestination = !isTransiter && isRrebDestination(roster.destination);
    const ppoDestination = !isTransiter && isPpoDestination(roster.destination);
    const combatDutyEvidence = collectCombatDutyEvidence({
      ubdNumber: ubdInfo.ubdNumber,
      hasUbd: ubdInfo.hasUbd,
      ubdRosterStatus,
      oosDislocation,
      tempArrivals,
      bplaPerformed,
      morningMissing,
      rrebDestination,
      ppoDestination,
    });
    const rankGroup = rankGroupOf(
      roster.rankTitle || oos?.rank || "",
      roster.shpkFact,
      roster.position,
      roster.rankCategory,
    );
    const morningExitCount = isTransiter ? 0 : morningExitSet.size;
    const jbdExitCount = isTransiter
      ? 0
      : [...jbdExitSet].filter((stamp) => !morningExitSet.has(stamp)).length;
    const mergedExitCount = isTransiter ? 0 : mergedExitSet.size;
    const exitCount = resolveCombatExitCount(
      mergedExitCount,
      ubdRosterStatus,
      combatDutyEvidence,
      isTransiter,
    );
    if (isTransiter) {
      parseNotes.push(
        `транзитер (${roster.destination.trim() || "ТРАНЗИТЕР"}): не входить у «Не виконували» / виходи`,
      );
    } else if (jbdExitCount > 0) {
      parseNotes.push(`ЖБД: +${jbdExitCount} виход(ів) без дублікатів з «Статус бійців»`);
    } else if (ubdRosterStatus) {
      parseNotes.push(
        ubdRosterStatus === "submitted"
          ? "УБД реєстр: подавалися"
          : "УБД реєстр: не подавалися",
      );
    } else if (mergedExitCount === 0 && combatDutyEvidence.length > 0) {
      parseNotes.push(
        `бойове завдання без «Статус бійців»: ${combatDutyEvidence.join("; ")}`,
      );
    }

    people.push({
      id: `${normalizedName}:${roster.position}`,
      name,
      normalizedName,
      shortName: shortPersonName(name),
      callsign,
      position: roster.position || shpo?.position || "",
      positionIndex: shpo?.positionIndex ?? "",
      rank: roster.rankTitle || oos?.rank || shpo?.rank || "",
      staffRank: roster.shpkFact || shpo?.staffRank || "",
      rankGroup,
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
      hasUbd: ubdInfo.hasUbd,
      ubdNumber: ubdInfo.ubdNumber,
      ubdRosterStatus,
      oosDislocation,
      combatDutyEvidence,
      isIdp: hasIdpFlag(oos?.extra ?? "", oos?.relatives ?? "", oos?.birthPlace ?? ""),
      morningStatus: roster.status,
      morningAbsenceNotes: [roster.treatmentNote, roster.medicalNote]
        .filter(Boolean)
        .join(" "),
      morningDestination: roster.destination ?? "",
      isTransiter,
      bzvpStatus: roster.bzvpStatus ?? "",
      brezAssignment: roster.brezAssignment ?? "",
      onStaff: !isBchsWithoutStaffPosition(roster),
      onList: true,
      present,
      inDisposition: !present,
      morningExitCount,
      jbdExitCount,
      exitCount,
      exitBand: exitBandOf(exitCount),
      match: {
        oos: Boolean(oos),
        morning: true,
        exits: morningExitCount > 0,
        jbdExits: jbdExitCount > 0 || (jbdExitSet.size > 0 && morningExitCount === 0),
        bplaExits: bplaPerformed,
        ubdRoster: Boolean(ubdRosterStatus),
        tempArrival: tempArrivals.some((row) => row.inCombatZone),
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
    .filter(
      (person) => person.arrivalSource === "brez" && isBrezMarkerText(person.arrivedFrom),
    )
    .map((person) => ({
      name: person.name,
      shortName: person.shortName,
      arrivedFrom: person.arrivedFrom,
      bzvpStatus: person.bzvpStatus,
      brezAssignment: person.brezAssignment,
      present: person.present,
      inDisposition: person.inDisposition,
    }));

  const countedBrezFromMorningOnly = people
    .filter(
      (person) =>
        person.arrivalSource === "brez" && !isBrezMarkerText(person.arrivedFrom),
    )
    .map((person) => ({
      name: person.name,
      shortName: person.shortName,
      arrivedFrom: person.arrivedFrom,
      bzvpStatus: person.bzvpStatus,
      brezAssignment: person.brezAssignment,
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
      brezAssignment: person.brezAssignment,
      source: isBrezMarkerText(person.arrivedFrom)
        ? "oos"
        : isBrezMarkerText(
              combineMorningBrezFields(person.bzvpStatus, person.brezAssignment),
            )
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

  const ubdRosterDebug = {
    loaded: Boolean(ubdRoster),
    fileName: ubdRoster?.fileName ?? "",
    submittedAll: ubdRosterRows.filter((row) => row.status === "submitted"),
    notSubmittedAll: ubdRosterRows.filter((row) => row.status === "notSubmitted"),
    matchedInMorning: people
      .filter((person) => person.ubdRosterStatus)
      .map((person) => ({
        name: person.name,
        status: person.ubdRosterStatus,
        exitCount: person.exitCount,
        morningExitCount: person.morningExitCount,
      })),
    counts: {
      submittedInFile: ubdRosterRows.filter((row) => row.status === "submitted").length,
      notSubmittedInFile: ubdRosterRows.filter((row) => row.status === "notSubmitted")
        .length,
      matchedSubmitted: people.filter((person) => person.ubdRosterStatus === "submitted")
        .length,
      matchedNotSubmitted: people.filter(
        (person) => person.ubdRosterStatus === "notSubmitted",
      ).length,
    },
  };

  const jbdExitsDebug = {
    loaded: Boolean(jbdExits),
    fileName: jbdExits?.fileName ?? "",
    totalStamps: [...jbdExitCounts.values()].reduce((sum, set) => sum + set.size, 0),
    addedToMorning: people
      .filter((person) => person.jbdExitCount > 0)
      .map((person) => ({
        name: person.name,
        jbdExitCount: person.jbdExitCount,
        morningExitCount: person.morningExitCount,
        exitCount: person.exitCount,
      })),
    counts: {
      peopleWithJbdOnly: people.filter(
        (person) => person.jbdExitCount > 0 && person.morningExitCount === 0,
      ).length,
      peopleWithJbdMerged: people.filter((person) => person.jbdExitCount > 0).length,
      jbdStampsTotal: [...jbdExitCounts.values()].reduce((sum, set) => sum + set.size, 0),
    },
  };

  const bplaExitsDebug = {
    loaded: Boolean(bplaExits),
    fileName: bplaExits?.fileName ?? "",
    performersInFile: bplaPerformers?.records.size ?? 0,
    matchedInMorning: people
      .filter((person) => person.match.bplaExits)
      .map((person) => ({
        name: person.name,
        exitCount: person.exitCount,
        morningExitCount: person.morningExitCount,
      })),
    unmatchedInFile: [...(bplaPerformers?.records.values() ?? [])]
      .filter(
        (record) =>
          !people.some((person) =>
            namesLikelySamePerson(record.name, person.name),
          ),
      )
      .map((record) => record.name),
    counts: {
      performersInFile: bplaPerformers?.records.size ?? 0,
      matchedPerformers: people.filter((person) => person.match.bplaExits).length,
      liftedFromNoExits: people.filter(
        (person) =>
          person.match.bplaExits &&
          person.morningExitCount === 0 &&
          person.jbdExitCount === 0 &&
          person.exitCount > 0,
      ).length,
    },
  };

  const combatDutyDebug = {
    tempArrivalAll: tempArrivalRows.filter((row) => row.inCombatZone),
    inferredFromEjoosOnly: people
      .filter(
        (person) =>
          person.morningExitCount === 0 &&
          person.jbdExitCount === 0 &&
          person.combatDutyEvidence.length > 0,
      )
      .map((person) => ({
        name: person.name,
        evidence: person.combatDutyEvidence,
        ubdNumber: person.ubdNumber,
      })),
    counts: {
      tempArrivalCombatZone: tempArrivalRows.filter((row) => row.inCombatZone).length,
      inferredExitFloor: people.filter(
        (person) =>
          person.morningExitCount === 0 &&
          person.jbdExitCount === 0 &&
          person.combatDutyEvidence.length > 0,
      ).length,
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
      tempArrived: tempArrivedSheet?.sheetName,
      jbdExits: jbdExits?.fileName,
      bplaExits: bplaExits?.fileName,
      ubdRoster: ubdRoster?.fileName,
    },
    oosRowCount: oosRows.length,
    morningRowCount: novaPeople.length,
    brezDebug,
    ubdDebug,
    combatDutyDebug,
    jbdExitsDebug,
    bplaExitsDebug,
    ubdRosterDebug,
  };
};
