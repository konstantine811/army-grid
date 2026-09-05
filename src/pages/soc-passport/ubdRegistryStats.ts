import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { extractCellFillRgb, valueToDisplay } from "../../excelRoundTrip";
import { tryParseExcelSerialDate } from "../../shared/format";
import {
  looksLikePersonName,
  normalizeLooseText,
  normalizePersonName,
  normalizeRnokpp,
} from "./socPassportFields";

export type UbdRegistryBucket =
  | "submitted"
  | "notSubmitted"
  | "batch"
  | "problems"
  | "onHand"
  | "other";

export type UbdYearStat = {
  year: number | "unknown";
  /** Заяви з аркуша «Подавалися» з прив’язкою до року. */
  submitted: number;
  /** З них є ознака видачі / отримання посвідчення (+ пакети за кольором для 2026). */
  issuedOrReceived: number;
  /** Скільки з «Отримано / видано» додано з кольорових пакетів (лише 2026). */
  issuedFromBatches: number;
  /** Рік взято з дати видачі / отримання / відправки. */
  fromProcessDate: number;
  /** Рік взято з періоду участі в БД. */
  fromCombatPeriod: number;
  people: string[];
};

export type UbdSheetStat = {
  sheetName: string;
  bucket: UbdRegistryBucket;
  peopleCount: number;
  withCombatPeriod: number;
  withIssueDate: number;
  withReceived: number;
  withIssued: number;
  statusTak: number;
};

export type UbdRegistryPerson = {
  name: string;
  callsign: string;
  rank: string;
  rnokpp: string;
  position: string;
  sheetName: string;
  bucket: UbdRegistryBucket;
  combatPeriod: string;
  combatYears: number[];
  processYear: number | null;
  year: number | "unknown";
  yearSource: "issue" | "received" | "sent" | "combat" | "none";
  ubdStatus: string;
  documents: string;
  raport: string;
  sent: string;
  received: string;
  issued: string;
  personnelStatus: string;
  issuedOrReceived: boolean;
};

export type UbdBatchColorRow = {
  name: string;
  callsign: string;
  excelRow: number;
  fillRgb: string | null;
  isGreen: boolean;
  received2026: boolean;
  mark: string;
  note: string;
};

/** @deprecated alias — use UbdBatchColorRow */
export type UbdJulyAugustRow = UbdBatchColorRow;

export type UbdBatchColorRule = "nonGreenReceived" | "greenReceived";

export type UbdBatchColorSheet = {
  sheetName: string;
  rule: UbdBatchColorRule;
  ruleLabel: string;
  total: number;
  green: number;
  received2026: number;
  rows: UbdBatchColorRow[];
};

export type UbdRegistryStatsResult = {
  fileName: string;
  people: UbdRegistryPerson[];
  byYear: UbdYearStat[];
  bySheet: UbdSheetStat[];
  /** Аркуші з правилом кольору → «Отримали в 2026». */
  colorBatches: UbdBatchColorSheet[];
  /** Зручний alias на «Июль-август», якщо є. */
  julyAugust: UbdBatchColorSheet | null;
  totals: {
    uniquePeople: number;
    submitted: number;
    notSubmitted: number;
    batchRows: number;
    problems: number;
    onHand: number;
    issuedOrReceived: number;
    statusTak: number;
    withCombatPeriod: number;
    yearUnknown: number;
    julyAugustReceived2026: number;
    batchReceived2026: number;
  };
  warnings: string[];
};


const cellText = (value: CellValue | undefined) =>
  valueToDisplay(value ?? null).replace(/\s+/g, " ").trim();

const normalizeHeader = (value: CellValue | undefined) =>
  normalizeLooseText(cellText(value)).replace(/[.]/g, " ");

const headerLooksLike = (row: CellValue[], partsGroups: string[][]) => {
  const headers = row.map((cell) => normalizeHeader(cell));
  return partsGroups.every((parts) =>
    headers.some((header) => parts.every((part) => header.includes(part))),
  );
};

const findHeaderRow = (sheet: ExcelSheetSnapshot, partsGroups: string[][]) => {
  const limit = Math.min(sheet.rawRows.length, 12);
  for (let index = 0; index < limit; index += 1) {
    if (headerLooksLike(sheet.rawRows[index] ?? [], partsGroups)) return index;
  }
  return -1;
};

const columnIndex = (
  headerRow: CellValue[],
  parts: string[],
  exclude: string[] = [],
) => {
  const index = headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    if (!header) return false;
    if (exclude.some((part) => header.includes(part))) return false;
    return parts.every((part) => header.includes(part));
  });
  return index;
};

const classifySheetBucket = (sheetName: string): UbdRegistryBucket => {
  const text = normalizeLooseText(sheetName);
  if (text.includes("не подавали")) return "notSubmitted";
  if (text.includes("подавали")) return "submitted";
  if (text.includes("проблен") || text.includes("проблем")) return "problems";
  if (text.includes("у мене") || text.includes("у меня") || text.includes("убд у мене")) {
    return "onHand";
  }
  if (
    /(январ|феврал|март|апрел|ма[йи]|июн|июл|август|сентябр|октябр|ноябр|декабр|січ|лют|берез|квіт|трав|черв|лип|серп|верес|жовт|листопад|груд)/.test(
      text,
    ) ||
    text.includes("сдал доки") ||
    text.includes("забрать")
  ) {
    return "batch";
  }
  return "other";
};

const isJulyAugustSheet = (sheetName: string) => {
  const text = normalizeLooseText(sheetName);
  return (
    (/июл|лип/.test(text) && /август|серп/.test(text)) ||
    text.includes("июль-август") ||
    text.includes("липень-серпень")
  );
};

/** Июнь / май-июнь / апрель-май — зелені = отримали в 2026. */
const isGreenReceivedBatchSheet = (sheetName: string) => {
  const text = normalizeLooseText(sheetName);
  if (isJulyAugustSheet(sheetName)) return false;
  return (
    text.includes("сдал доки") ||
    text.includes("сдав доки") ||
    (text.includes("май") && text.includes("июн")) ||
    (text.includes("трав") && text.includes("черв")) ||
    (text.includes("апрел") && text.includes("май")) ||
    (text.includes("квіт") && text.includes("трав")) ||
    /^май-июн/.test(text) ||
    /^апрел/.test(text)
  );
};

const batchColorRuleForSheet = (
  sheetName: string,
): UbdBatchColorRule | null => {
  if (isJulyAugustSheet(sheetName)) return "nonGreenReceived";
  if (isGreenReceivedBatchSheet(sheetName)) return "greenReceived";
  return null;
};

/** Типові зелені заливки Excel (акцент / «готово»). */
const isExcelGreenFill = (rgb: string | null | undefined) => {
  if (!rgb) return false;
  const hex = rgb.replace(/^#/, "").replace(/^FF/i, "").toUpperCase();
  if (
    /^(92D050|00B050|70AD47|C6EFCE|A9D08E|548235|375623|00FF00|39B54A)$/.test(
      hex,
    )
  ) {
    return true;
  }
  if (hex.length !== 6) return false;
  const r = Number.parseInt(hex.slice(0, 2), 16);
  const g = Number.parseInt(hex.slice(2, 4), 16);
  const b = Number.parseInt(hex.slice(4, 6), 16);
  return g >= 160 && g > r + 30 && g > b + 30;
};

const loadXlsxPopulate = async () => {
  const module =
    await import("xlsx-populate/browser/xlsx-populate-no-encryption");
  return module.default;
};

const readSheetCellFills = async (
  file: File,
  sheetName: string,
  maxRow: number,
  columns: number[],
): Promise<Record<number, string>> => {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(file);
  const sheet =
    workbook.sheets().find((item: { name: () => string }) => item.name() === sheetName) ??
    workbook.sheet(sheetName as unknown as number);
  if (!sheet) return {};
  const fills: Record<number, string> = {};
  for (let row = 1; row <= maxRow; row += 1) {
    for (const column of columns) {
      try {
        const rgb = extractCellFillRgb(sheet.cell(row, column).style("fill"));
        if (!rgb) continue;
        const normalized = rgb.replace(/^#/, "").toUpperCase();
        if (
          normalized === "FFFFFFFF" ||
          normalized === "FFFFFF" ||
          normalized === "00000000" ||
          normalized === "000000"
        ) {
          continue;
        }
        fills[row] = normalized.startsWith("FF") && normalized.length === 8
          ? normalized
          : normalized.length === 6
            ? `FF${normalized}`
            : normalized;
        break;
      } catch {
        // ignore broken style cells
      }
    }
  }
  return fills;
};


const extractYearsFromText = (value: string): number[] => {
  const years = [...value.matchAll(/\b(20\d{2})\b/g)].map((match) =>
    Number(match[1]),
  );
  return [...new Set(years)].filter((year) => year >= 2014 && year <= 2035);
};

const parseDateLike = (value: CellValue | undefined): Date | null => {
  if (value == null || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const serial = tryParseExcelSerialDate(value);
  if (serial) return serial;
  const text = cellText(value);
  if (!text) return null;
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    const date = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const ua = text.match(/(\d{1,2})[./](\d{1,2})[./](\d{2,4})/);
  if (ua) {
    const yearRaw = Number(ua[3]);
    const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
    const date = new Date(year, Number(ua[2]) - 1, Number(ua[1]));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  return null;
};

const yearFromValue = (value: CellValue | undefined): number | null => {
  const date = parseDateLike(value);
  if (date) {
    const year = date.getFullYear();
    if (year >= 2014 && year <= 2035) return year;
  }
  const years = extractYearsFromText(cellText(value));
  return years[0] ?? null;
};

const combatYearsFromPeriod = (period: string): number[] =>
  extractYearsFromText(period);

type RegistryColumns = {
  callsign: number;
  combined: number;
  last: number;
  first: number;
  patronymic: number;
  rank: number;
  rnokpp: number;
  position: number;
  combatPeriod: number;
  issueDate: number;
  ubdStatus: number;
  documents: number;
  raport: number;
  sent: number;
  received: number;
  issued: number;
  personnelStatus: number;
};

const resolveRegistryColumns = (headerRow: CellValue[]): RegistryColumns => {
  const combined = headerRow.findIndex((cell) => {
    const header = normalizeHeader(cell);
    return header.includes("прізвищ") && header.includes("ім");
  });
  return {
    callsign: columnIndex(headerRow, ["позивн"]),
    combined:
      combined >= 0
        ? combined
        : columnIndex(headerRow, ["прізвищ"], ["позивн"]),
    last: columnIndex(headerRow, ["прізвищ"], ["позивн", "ім"]),
    first: columnIndex(headerRow, ["ім"], ["прізвищ", "батьк"]),
    patronymic: columnIndex(headerRow, ["батьк"]),
    rank: columnIndex(headerRow, ["званн"]),
    rnokpp: columnIndex(headerRow, ["обліков"]),
    position: columnIndex(headerRow, ["посад"]),
    combatPeriod: columnIndex(headerRow, ["період", "учас"]),
    issueDate: columnIndex(headerRow, ["дата", "видаг"]),
    ubdStatus: columnIndex(headerRow, ["статус", "убд"]),
    documents: columnIndex(headerRow, ["документ"]),
    raport: columnIndex(headerRow, ["рапорт"]),
    sent: columnIndex(headerRow, ["відправл"]),
    received: columnIndex(headerRow, ["отриман"]),
    issued: columnIndex(headerRow, ["видан"]),
    personnelStatus: columnIndex(headerRow, ["статус"], ["убд"]),
  };
};

const readName = (row: CellValue[], columns: RegistryColumns) => {
  const combined = cellText(row[columns.combined]);
  if (looksLikePersonName(combined)) return combined.replace(/\s+/g, " ").trim();
  const assembled = [
    cellText(row[columns.last]),
    cellText(row[columns.first]),
    cellText(row[columns.patronymic]),
  ]
    .filter(Boolean)
    .join(" ");
  return looksLikePersonName(assembled) ? assembled : "";
};

const resolvePersonYear = (input: {
  issueDate: CellValue | undefined;
  received: string;
  sent: string;
  issued: string;
  combatYears: number[];
}): Pick<UbdRegistryPerson, "year" | "yearSource" | "processYear"> => {
  const issueYear = yearFromValue(input.issueDate);
  if (issueYear != null) {
    return { year: issueYear, yearSource: "issue", processYear: issueYear };
  }
  const receivedYear = yearFromValue(input.received);
  if (receivedYear != null) {
    return {
      year: receivedYear,
      yearSource: "received",
      processYear: receivedYear,
    };
  }
  const issuedYear = yearFromValue(input.issued);
  if (issuedYear != null) {
    return { year: issuedYear, yearSource: "received", processYear: issuedYear };
  }
  const sentYear = yearFromValue(input.sent);
  if (sentYear != null) {
    return { year: sentYear, yearSource: "sent", processYear: sentYear };
  }
  if (input.combatYears.length) {
    return {
      year: Math.min(...input.combatYears),
      yearSource: "combat",
      processYear: null,
    };
  }
  return { year: "unknown", yearSource: "none", processYear: null };
};

const pickLooseName = (row: CellValue[]) => {
  for (let index = 0; index < Math.min(row.length, 6); index += 1) {
    const text = cellText(row[index]).replace(/\([^)]*\)/g, " ").trim();
    if (looksLikePersonName(text)) return text;
  }
  const candidate = cellText(row[2]) || cellText(row[1]) || cellText(row[0]);
  const cleaned = candidate.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
  return looksLikePersonName(cleaned) ? cleaned : "";
};

const parseRegistrySheet = (
  sheet: ExcelSheetSnapshot,
  bucket: UbdRegistryBucket,
): UbdRegistryPerson[] => {
  const headerRowIndex = findHeaderRow(sheet, [["прізвищ"]]);
  if (headerRowIndex < 0) return [];
  const headerRow = sheet.rawRows[headerRowIndex] ?? [];
  const columns = resolveRegistryColumns(headerRow);
  const people: UbdRegistryPerson[] = [];

  for (let rowIndex = headerRowIndex + 1; rowIndex < sheet.rawRows.length; rowIndex += 1) {
    const row = sheet.rawRows[rowIndex] ?? [];
    const name = readName(row, columns);
    if (!name) continue;
    const combatPeriod = cellText(row[columns.combatPeriod]);
    const combatYears = combatYearsFromPeriod(combatPeriod);
    const received = cellText(row[columns.received]);
    const sent = cellText(row[columns.sent]);
    const issued = cellText(row[columns.issued]);
    const yearInfo = resolvePersonYear({
      issueDate: row[columns.issueDate],
      received,
      sent,
      issued,
      combatYears,
    });
    const ubdStatus = cellText(row[columns.ubdStatus]);
    people.push({
      name,
      callsign: cellText(row[columns.callsign]),
      rank: cellText(row[columns.rank]),
      rnokpp: normalizeRnokpp(cellText(row[columns.rnokpp])),
      position: cellText(row[columns.position]),
      sheetName: sheet.sheetName,
      bucket,
      combatPeriod,
      combatYears,
      ...yearInfo,
      ubdStatus,
      documents: cellText(row[columns.documents]),
      raport: cellText(row[columns.raport]),
      sent,
      received,
      issued,
      personnelStatus: cellText(row[columns.personnelStatus]),
      issuedOrReceived: Boolean(received.trim() || issued.trim()),
    });
  }
  return people;
};

const parseLoosePeopleSheet = (
  sheet: ExcelSheetSnapshot,
  bucket: UbdRegistryBucket,
): UbdRegistryPerson[] => {
  const people: UbdRegistryPerson[] = [];
  const seen = new Set<string>();
  for (const row of sheet.rawRows) {
    const name = pickLooseName(row);
    if (!name) continue;
    const key = normalizePersonName(name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    const note = row
      .map((cell) => cellText(cell))
      .filter(Boolean)
      .slice(1, 6)
      .join(" · ");
    const years = extractYearsFromText(note);
    people.push({
      name,
      callsign: "",
      rank: "",
      rnokpp: "",
      position: "",
      sheetName: sheet.sheetName,
      bucket,
      combatPeriod: "",
      combatYears: years,
      processYear: years[0] ?? null,
      year: years[0] ?? "unknown",
      yearSource: years[0] ? "sent" : "none",
      ubdStatus: "",
      documents: "",
      raport: "",
      sent: note,
      received: "",
      issued: "",
      personnelStatus: "",
      issuedOrReceived: /отриман|видан|вруч|убд/.test(normalizeLooseText(note)),
    });
  }
  return people;
};

const parseBatchColorSheet = (
  sheet: ExcelSheetSnapshot,
  fillsByRow: Record<number, string>,
  rule: UbdBatchColorRule,
): UbdBatchColorRow[] => {
  if (rule === "greenReceived") {
    const seen = new Set<string>();
    const result: UbdBatchColorRow[] = [];
    for (let index = 0; index < sheet.rawRows.length; index += 1) {
      const row = sheet.rawRows[index] ?? [];
      const excelRow = index + 1;
      const name = pickLooseName(row);
      if (!name) continue;
      if (/сдал\s*документ|рапорт[оі]?в?\b/i.test(name)) continue;
      const fillRgb = fillsByRow[excelRow] ?? null;
      const isGreen = isExcelGreenFill(fillRgb);
      if (!isGreen) continue;
      const key = normalizePersonName(name);
      if (!key || seen.has(key)) continue;
      seen.add(key);
      const col1 = cellText(row[1]);
      const callsign =
        col1 && !looksLikePersonName(col1) && col1.length <= 24 ? col1 : "";
      const mark = cellText(row[1]) || cellText(row[2]) || cellText(row[3]);
      const note = row
        .slice(1, 8)
        .map((cell) => cellText(cell))
        .filter(Boolean)
        .join(" · ");
      result.push({
        name,
        callsign,
        excelRow,
        fillRgb,
        isGreen: true,
        received2026: true,
        mark,
        note,
      });
    }
    return result;
  }

  // nonGreenReceived — «Июль-август»
  type Candidate = {
    excelRow: number;
    name: string;
    callsign: string;
    mark: string;
    note: string;
    anchored: boolean;
  };

  const candidates: Candidate[] = [];
  for (let index = 0; index < sheet.rawRows.length; index += 1) {
    const row = sheet.rawRows[index] ?? [];
    const excelRow = index + 1;
    const name = pickLooseName(row);
    if (!name) continue;
    if (/сдал\s*документ|рапорт[оі]?в?\b/i.test(name)) continue;

    const col0 = cellText(row[0]);
    const col1 = cellText(row[1]);
    const mark = cellText(row[3]);
    const note = row
      .slice(4, 12)
      .map((cell) => cellText(cell))
      .filter(Boolean)
      .join(" · ");
    const hasNumber = /^\d+$/.test(col0);
    const callsign =
      col1 && !looksLikePersonName(col1) && col1.length <= 24 ? col1 : "";
    const anchored =
      hasNumber ||
      mark === "+" ||
      mark === "-" ||
      Boolean(callsign) ||
      /рапорт|отдал|отрим|готов\s+на\s+отдач/i.test(`${mark} ${note}`);
    candidates.push({
      excelRow,
      name,
      callsign,
      mark,
      note,
      anchored,
    });
  }

  const anchoredRows = candidates.filter((row) => row.anchored);
  if (!anchoredRows.length) return [];
  const minRow = Math.min(...anchoredRows.map((row) => row.excelRow));
  const maxRow = Math.max(...anchoredRows.map((row) => row.excelRow));

  const seen = new Set<string>();
  const result: UbdBatchColorRow[] = [];
  for (const row of candidates) {
    if (row.excelRow < minRow || row.excelRow > maxRow) continue;
    if (!row.anchored) {
      const hasNeighborAnchor = anchoredRows.some(
        (anchor) => Math.abs(anchor.excelRow - row.excelRow) <= 2,
      );
      if (!hasNeighborAnchor) continue;
    }
    const key = normalizePersonName(row.name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    const fillRgb = fillsByRow[row.excelRow] ?? null;
    const isGreen = isExcelGreenFill(fillRgb);
    result.push({
      name: row.name,
      callsign: row.callsign,
      excelRow: row.excelRow,
      fillRgb,
      isGreen,
      received2026: !isGreen,
      mark: row.mark,
      note: row.note,
    });
  }
  return result;
};

const emptySheetStat = (
  sheetName: string,
  bucket: UbdRegistryBucket,
): UbdSheetStat => ({
  sheetName,
  bucket,
  peopleCount: 0,
  withCombatPeriod: 0,
  withIssueDate: 0,
  withReceived: 0,
  withIssued: 0,
  statusTak: 0,
});

const bumpYear = (
  map: Map<number | "unknown", UbdYearStat>,
  year: number | "unknown",
  person: UbdRegistryPerson,
) => {
  const current =
    map.get(year) ??
    ({
      year,
      submitted: 0,
      issuedOrReceived: 0,
      issuedFromBatches: 0,
      fromProcessDate: 0,
      fromCombatPeriod: 0,
      people: [],
    } satisfies UbdYearStat);
  current.submitted += 1;
  if (person.issuedOrReceived) current.issuedOrReceived += 1;
  if (
    person.yearSource === "issue" ||
    person.yearSource === "received" ||
    person.yearSource === "sent"
  ) {
    current.fromProcessDate += 1;
  } else if (person.yearSource === "combat") {
    current.fromCombatPeriod += 1;
  }
  if (current.people.length < 40) current.people.push(person.name);
  map.set(year, current);
};

export const parseUbdRegistryStats = async (
  workbook: ExcelWorkbookSnapshot,
): Promise<UbdRegistryStatsResult> => {
  const warnings: string[] = [];
  const bySheet: UbdSheetStat[] = [];
  const submittedPeople: UbdRegistryPerson[] = [];
  const notSubmittedPeople: UbdRegistryPerson[] = [];
  const otherPeople: UbdRegistryPerson[] = [];
  const colorBatches: UbdBatchColorSheet[] = [];

  for (const sheet of workbook.sheets) {
    const bucket = classifySheetBucket(sheet.sheetName);
    const sheetStat = emptySheetStat(sheet.sheetName, bucket);
    const colorRule = batchColorRuleForSheet(sheet.sheetName);

    if (bucket === "submitted" || bucket === "notSubmitted") {
      const parsed = parseRegistrySheet(sheet, bucket);
      sheetStat.peopleCount = parsed.length;
      for (const person of parsed) {
        if (person.combatPeriod) sheetStat.withCombatPeriod += 1;
        if (person.yearSource === "issue") sheetStat.withIssueDate += 1;
        if (person.received) sheetStat.withReceived += 1;
        if (person.issued) sheetStat.withIssued += 1;
        if (normalizeLooseText(person.ubdStatus) === "так") {
          sheetStat.statusTak += 1;
        }
      }
      if (bucket === "submitted") submittedPeople.push(...parsed);
      else notSubmittedPeople.push(...parsed);
      if (!parsed.length) {
        warnings.push(
          `Аркуш «${sheet.sheetName}»: не знайдено рядків з ПІБ (перевірте заголовок).`,
        );
      }
    } else if (colorRule) {
      let fillsByRow: Record<number, string> = {};
      try {
        fillsByRow = await readSheetCellFills(
          workbook.file,
          sheet.sheetName,
          Math.min(sheet.rawRows.length, 400),
          [1, 2, 3],
        );
      } catch {
        warnings.push(
          `Аркуш «${sheet.sheetName}»: не вдалося прочитати кольори рядків.`,
        );
      }
      const batchRows = parseBatchColorSheet(sheet, fillsByRow, colorRule);
      const receivedRows = batchRows.filter((row) => row.received2026);
      sheetStat.peopleCount =
        colorRule === "greenReceived" ? receivedRows.length : batchRows.length;
      sheetStat.withReceived = receivedRows.length;
      const batch: UbdBatchColorSheet = {
        sheetName: sheet.sheetName,
        rule: colorRule,
        ruleLabel:
          colorRule === "greenReceived"
            ? "зелені = отримали в 2026"
            : "не зелені = отримали в 2026",
        total: batchRows.length,
        green: batchRows.filter((row) => row.isGreen).length,
        received2026: receivedRows.length,
        rows: batchRows,
      };
      colorBatches.push(batch);
      otherPeople.push(
        ...receivedRows.map((row) => ({
          name: row.name,
          callsign: row.callsign,
          rank: "",
          rnokpp: "",
          position: "",
          sheetName: sheet.sheetName,
          bucket,
          combatPeriod: "",
          combatYears: [2026],
          processYear: 2026,
          year: 2026 as const,
          yearSource: "received" as const,
          ubdStatus: "",
          documents: "",
          raport: "",
          sent: row.note,
          received: `${sheet.sheetName} (${batch.ruleLabel})`,
          issued: "",
          personnelStatus: row.mark,
          issuedOrReceived: true,
        })),
      );
    } else {
      const parsed = parseLoosePeopleSheet(sheet, bucket);
      sheetStat.peopleCount = parsed.length;
      otherPeople.push(...parsed);
    }

    bySheet.push(sheetStat);
  }

  const yearMap = new Map<number | "unknown", UbdYearStat>();
  for (const person of submittedPeople) {
    bumpYear(yearMap, person.year, person);
  }

  // Пакети за кольором → «Отримано / видано» за 2026 (без дублікатів ПІБ).
  const issuedKeys2026 = new Set<string>();
  for (const person of submittedPeople) {
    if (person.year === 2026 && person.issuedOrReceived) {
      const key = normalizePersonName(person.name);
      if (key) issuedKeys2026.add(key);
    }
  }
  const issuedBeforeBatches = issuedKeys2026.size;
  for (const batch of colorBatches) {
    for (const row of batch.rows) {
      if (!row.received2026) continue;
      const key = normalizePersonName(row.name);
      if (!key) continue;
      issuedKeys2026.add(key);
    }
  }
  const year2026 =
    yearMap.get(2026) ??
    ({
      year: 2026,
      submitted: 0,
      issuedOrReceived: 0,
      issuedFromBatches: 0,
      fromProcessDate: 0,
      fromCombatPeriod: 0,
      people: [],
    } satisfies UbdYearStat);
  year2026.issuedOrReceived = issuedKeys2026.size;
  year2026.issuedFromBatches = Math.max(
    0,
    issuedKeys2026.size - issuedBeforeBatches,
  );
  yearMap.set(2026, year2026);

  const byYear = [...yearMap.values()].sort((left, right) => {
    if (left.year === "unknown") return 1;
    if (right.year === "unknown") return -1;
    return left.year - right.year;
  });

  const unique = new Map<string, UbdRegistryPerson>();
  for (const person of [
    ...submittedPeople,
    ...notSubmittedPeople,
    ...otherPeople,
  ]) {
    const key = normalizePersonName(person.name);
    if (!key) continue;
    const existing = unique.get(key);
    if (!existing || person.bucket === "submitted") {
      unique.set(key, person);
    }
  }

  if (!submittedPeople.length) {
    warnings.push(
      "Не знайдено аркуш «Подавалися» з даними — річна статистика «зроблено УБД» порожня.",
    );
  }

  const julyAugust =
    colorBatches.find((batch) => isJulyAugustSheet(batch.sheetName)) ?? null;
  const batchReceived2026 = colorBatches.reduce(
    (sum, batch) => sum + batch.received2026,
    0,
  );

  return {
    fileName: workbook.fileName,
    people: submittedPeople,
    byYear,
    bySheet,
    colorBatches,
    julyAugust,
    totals: {
      uniquePeople: unique.size,
      submitted: submittedPeople.length,
      notSubmitted: notSubmittedPeople.length,
      batchRows: otherPeople.filter((person) => person.bucket === "batch").length,
      problems: otherPeople.filter((person) => person.bucket === "problems")
        .length,
      onHand: otherPeople.filter((person) => person.bucket === "onHand").length,
      issuedOrReceived: byYear.reduce(
        (sum, row) => sum + row.issuedOrReceived,
        0,
      ),
      statusTak: submittedPeople.filter(
        (person) => normalizeLooseText(person.ubdStatus) === "так",
      ).length,
      withCombatPeriod: submittedPeople.filter((person) => person.combatPeriod)
        .length,
      yearUnknown: submittedPeople.filter((person) => person.year === "unknown")
        .length,
      julyAugustReceived2026: julyAugust?.received2026 ?? 0,
      batchReceived2026,
    },
    warnings,
  };
};
