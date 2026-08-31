import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { mapPbStatusToEjoosWithRules, readOperatorSettings } from "./ejoosStatusMap";
import type { EjoosTimesheetCode } from "./ejoosRules";
import {
  findLatestPriorOwnUnitStaffMove,
  isAmbiguousStaffTransfer,
  isDispositionToStaffPlacement,
  isOutboundStaffMove,
  isOwnFirstPbDestination,
  isOwnUnitStaffMove,
  resolveOutboundTransferDestination,
} from "./ejoosMovementRules";
import {
  isStaleVacatedAbsence,
  isUnrecordedSameMonthTransit,
  ownUnitMoveSupersededByOutbound,
  skipExternalIfAlreadyProcessed,
} from "./ejoosExcludePolicy";
import {
  absenceSpansBeforeEpisode,
  archivePeriodTouchesJournalMonth,
  buildTimesheetAbsenceSpans,
  clipAbsenceSpansToActiveEpisode,
  currentStatusConfirmsOpenAbsence,
  encodeTimesheetAbsenceSpans,
  archiveReturnContradictsCurrentSh,
  findTimesheetMonthHeaderCell,
  formatTimesheetMonthHeader,
  isTimesheetDepartureMark,
  journalDayFromDateMs,
  sameTimesheetDayMark,
  timesheetMarkFromArchive,
  extractTimesheetDestinationFromPosition,
} from "./ejoosTimesheetText";
import { findDuplicateTimesheetExtras } from "./ejoosTimesheetDuplicates";

export { extractTimesheetDestinationFromPosition };

export type EjoosOpClass = "ready" | "needs_input" | "conflict";

export type EjoosOpKind =
  | "timesheet_day"
  | "shpo_occupant"
  | "absent_upsert"
  | "absent_close"
  | "exclude_transfer"
  | "move_to_disposition"
  | "data_mismatch"
  | "position_change"
  | "rank_change"
  | "arrival"
  | "other_manual";

export type EjoosSyncOp = {
  id: string;
  kind: EjoosOpKind;
  class: EjoosOpClass;
  sheet: string;
  personId: string;
  fullName: string;
  positionIndex: string;
  rank: string;
  before: string;
  after: string;
  sourceRef: string;
  why: string;
  confidence: "high" | "review" | "manual";
  /** Extra payload for apply */
  payload: Record<string, string>;
  /** Stable source-event key persisted after a successful DB version write. */
  movementKey?: string;
  checkedDefault: boolean;
};

export type EjoosSyncPlan = {
  ejoosName: string;
  pbName: string;
  timesheetDay: number;
  timesheetDayLabel: string;
  ops: EjoosSyncOp[];
  summary: {
    ready: number;
    needsInput: number;
    conflict: number;
  };
  limitsNote?: string;
};

const norm = (value: CellValue | unknown) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const day = String(value.getDate()).padStart(2, "0");
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${day}.${month}.${value.getFullYear()}`;
  }
  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) {
    // Excel serial date
    const utc = Date.UTC(1899, 11, 30) + Math.floor(value) * 86400000;
    const date = new Date(utc);
    const day = String(date.getUTCDate()).padStart(2, "0");
    const month = String(date.getUTCMonth() + 1).padStart(2, "0");
    return `${day}.${month}.${date.getUTCFullYear()}`;
  }
  if (value && typeof value === "object" && !(value instanceof Date)) {
    const maybe = value as { text?: () => string; value?: () => unknown };
    if (typeof maybe.text === "function") return norm(maybe.text());
    if (typeof maybe.value === "function") return norm(maybe.value());
    return "";
  }
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
};

const normKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

/** Уточнення на кшталт «(08.02.1985 р.н.)» — не інше ПІБ, а лише позначка. */
const canonicalName = (value: string) =>
  normKey(
    value
      .replace(/\([^)]*\)/g, " ")
      .replace(/[.,;]/g, " ")
      .replace(/\s+/g, " "),
  );

const dateMs = (value: string) => {
  const match = String(value ?? "").match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (!match) return 0;
  const year = Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
  const result = Date.UTC(year, Number(match[2]) - 1, Number(match[1]));
  return Number.isFinite(result) ? result : 0;
};

export const findEjoosSheet = (
  workbook: ExcelWorkbookSnapshot,
  matcher: RegExp,
) => workbook.sheets.find((sheet) => matcher.test(sheet.sheetName));

const findSheet = findEjoosSheet;

const headerMap = (row: CellValue[]) => {
  const map = new Map<string, number>();
  row.forEach((cell, index) => {
    const key = normKey(norm(cell));
    if (key && !map.has(key)) map.set(key, index);
  });
  return map;
};

const findCol = (map: Map<string, number>, ...needles: RegExp[]) => {
  for (const [key, index] of map.entries()) {
    if (needles.some((re) => re.test(key))) return index;
  }
  return -1;
};

const cell = (row: CellValue[] | undefined, index: number) =>
  index >= 0 ? norm(row?.[index]) : "";

/** Штатний ID 1ПБ/ЕЖООС — коротке число, не РНОКПП/ІПН з 8+ цифр. */
const isJournalPersonId = (value: string) => /^\d{1,7}$/.test(value.trim());

/** ID is never an Excel date. Reject date-shaped fallback values instead of showing them as IDs. */
const normId = (value: CellValue | unknown) => {
  if (value instanceof Date) return "";
  const text =
    typeof value === "number" && Number.isFinite(value)
      ? String(Math.trunc(value))
      : String(value ?? "")
          .replace(/\s+/g, " ")
          .trim();
  if (/^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text)) return "";
  return text === "0" || text === "0.0" ? "" : text;
};

const journalIdCell = (row: CellValue[] | undefined, index: number) => {
  const value = idCell(row, index);
  return isJournalPersonId(value) ? value : "";
};

const idCell = (row: CellValue[] | undefined, index: number) =>
  index >= 0 ? normId(row?.[index]) : "";

export type PbShPerson = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  positionTitle: string;
  status: string;
  /** Якщо є в sh — «Звідки прибув». */
  arrivedFrom: string;
};

export type PbArchivePeriod = {
  excelRow: number;
  periodNumber: string;
  personId: string;
  fullName: string;
  rank: string;
  positionTitle: string;
  absenceType: string;
  departDate: string;
  place: string;
  orderNumber: string;
  orderDate: string;
  plannedReturn: string;
  /** Фактична / дата повернення, якщо є. */
  returnDate: string;
  returnOrderNumber: string;
  returnOrderDate: string;
};

export type PbMovement = {
  excelRow: number;
  movementNumber: string;
  type: string;
  personId: string;
  fullName: string;
  rank: string;
  previousIndex: string;
  nextIndex: string;
  destination: string;
  orderNumber: string;
  orderDate: string;
  basisNumber: string;
  basisDate: string;
  changeText: string;
  status: string;
  note: string;
  /** Колонка «Звідки» / «Звідки прибув», якщо є в Рух. */
  arrivedFrom: string;
};

export const createMovementKey = (event: PbMovement) => {
  const identity = event.personId
    ? `id:${normKey(event.personId)}`
    : `name:${normKey(event.fullName)}`;
  if (identity.endsWith(":")) return "";
  return [
    identity,
    normKey(event.type),
    normKey(event.orderDate),
    normKey(event.orderNumber),
    normKey(event.previousIndex),
    normKey(event.nextIndex),
  ].join("|");
};

export const collectProcessedMovementKeys = (
  versions: Array<{ changeProtocol?: unknown }> | null | undefined,
) => {
  const keys = new Set<string>();
  for (const version of versions ?? []) {
    const protocol = version.changeProtocol;
    if (!protocol || typeof protocol !== "object") continue;
    const ops = (protocol as { ops?: unknown }).ops;
    if (!Array.isArray(ops)) continue;
    for (const op of ops) {
      if (!op || typeof op !== "object") continue;
      const key = (op as { movementKey?: unknown }).movementKey;
      if (typeof key === "string" && key.trim()) keys.add(key.trim());
    }
  }
  return keys;
};

/** «Куди» у Рух часто = «x» / порожньо; тоді військова частина є в примітці. */
const resolveMovementDestination = (rawDest: string, note: string) => {
  const clean = (value: string) => {
    const text = value.replace(/\s+/g, " ").trim();
    if (!text) return "";
    const upper = text.toUpperCase();
    if (
      upper === "X" ||
      upper === "Х" ||
      upper === "0" ||
      upper === "0.0" ||
      upper === "-" ||
      upper === "—" ||
      upper === "." ||
      upper.includes("РОЗПОР") ||
      upper.includes("ПЕРЕВ") ||
      upper.includes("ПОСАД") ||
      upper.includes("ПРИБ") ||
      upper.includes("ЗВІЛ")
    ) {
      return "";
    }
    return text;
  };
  const destination = clean(rawDest);
  if (destination) return destination;
  const noteText = note.replace(/\s+/g, " ").trim();
  return /(?:[АA]\s*\d{4}(?!\d)|військов(?:ої|а)\s+частин)/iu.test(noteText)
    ? noteText
    : "";
};

const formatTransferDestinationForTimesheet = (value: string) => {
  const codes = [...value.matchAll(/[АA]\s*(\d{4})(?!\d)/giu)].map(
    (match) => `А${match[1]}`,
  );
  const unique = [...new Set(codes)];
  return unique.length ? unique.join(" / ") : value;
};

/**
 * «2103791 Старший кухар → 2103179 Стрілець 3 піхотного відділення …» —
 * для «куди вибув» потрібна лише нова посада без службового індексу.
 */
const positionChangeDestination = (event: PbMovement) => {
  const text = norm(event.changeText);
  if (!text) return norm(event.destination);
  const tail = text.split(/→|->|=>/).pop()?.trim() || text;
  return tail.replace(/^\d{5,}[\s.:;-]*/, "").trim() || tail;
};

/**
 * У колонці повернення archive часто стоїть «до окремого розпорядження», «0»
 * або «-». Період вважаємо закритим лише за фактичною датою.
 */
const hasActualReturn = (value: string) => Boolean(dateMs(value));

/** У `sh` замість індексу бувають маркери «ВИВЕДЕНО», «#N/A» тощо. */
const isPositionIndex = (value: string) => /^\d{5,}$/.test(value.trim());

const normalizeRankLabel = (value: string) =>
  value
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^(?:звання|військов(?:е|ое)\s+звання)\s+/i, "")
    .toLocaleLowerCase("uk-UA");

/** «солдат → СТАРШИЙ СОЛДАТ» або лише нове звання в «Яка зміна». */
export const parseRankPromotion = (event: PbMovement) => {
  const text = [event.changeText, event.note].filter(Boolean).join(" ");
  const match = text.match(
    /([А-ЯІЇЄҐа-яіїєґ'’.\-\s]{3,}?)\s*(?:→|->|=>|—)\s*([А-ЯІЇЄҐа-яіїєґ'’.\-\s]{3,})/u,
  );
  if (match) {
    return {
      previousRank: normalizeRankLabel(match[1]),
      nextRank: normalizeRankLabel(match[2]),
    };
  }
  const nextRank =
    normalizeRankLabel(event.changeText) || normalizeRankLabel(event.rank);
  const previousRank = normalizeRankLabel(event.rank);
  return {
    previousRank: previousRank === nextRank ? "" : previousRank,
    nextRank,
  };
};

const isRankAssignmentEvent = (event: PbMovement) => {
  if (event.type === "ЗВАННЯ" || event.type.startsWith("ЗВАН")) return true;
  return false;
};

const unitCodeFromMovement = (event: PbMovement) =>
  formatTransferDestinationForTimesheet(
    [event.destination, event.changeText, event.note].join(" "),
  );

const isSzchCancellation = (event: PbMovement) =>
  /СКАС(?:УВАННЯ|ОВАНО|УВАТИ)?.*СЗЧ|СЗЧ.*СКАС/iu.test(
    [event.type, event.status, event.note, event.changeText].join(" "),
  );

const isDispositionAbsenceStatus = (value: string) =>
  /СЗЧ|САМОВІЛ|БЕЗВІСТ|(?:^|[^А-ЯІЇЄҐ])ЗБ(?:$|[^А-ЯІЇЄҐ])/iu.test(value);

const movementBlob = (event: PbMovement) =>
  [event.type, event.status, event.note, event.changeText, event.destination].join(
    " ",
  );

/** Окремий рядок «СКАСУВАННЯ переведення», а не скасований рядок ПЕРЕВ. */
const isTransferCancellation = (event: PbMovement) => {
  if (isSzchCancellation(event)) return false;
  if (
    event.type === "ПЕРЕВ" ||
    event.type === "ПОСАДА" ||
    event.type === "ЗВАННЯ"
  ) {
    return false;
  }
  if (event.type === "СКАСУВАННЯ") return true;
  const text = movementBlob(event);
  return /скас(?:уванн|овано|увано|увати)/iu.test(text) && /перев/iu.test(text);
};

/** Анульований рядок РУХ: «скасовано» у статусі, примітці або «куди». */
const isCancelledMovementRecord = (event: PbMovement) => {
  if (isSzchCancellation(event) || isTransferCancellation(event)) return false;
  const fields = [event.status, event.note, event.destination, event.changeText];
  return fields.some((value) => {
    const text = String(value ?? "").trim();
    if (!text) return false;
    if (/^скас(?:овано|увано)$/iu.test(text)) return true;
    return /(?:^|[\s,;./(])скас(?:овано|увано)(?:$|[\s,;./)])/iu.test(text);
  });
};

export const parsePbShPeople = (workbook: ExcelWorkbookSnapshot): PbShPerson[] => {
  const sheet = findSheet(workbook, /^sh$/i);
  if (!sheet) return [];
  const headers = headerMap(sheet.rawRows[0] ?? []);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /^піб$/, /прізвище/);
  const rankCol = findCol(headers, /зван/);
  const indexCol = findCol(headers, /індекс\s*посад/);
  const posCol = findCol(headers, /^посада$/);
  const statusCol = findCol(headers, /^статус$/, /^перебуван/);
  const fromCol = findCol(headers, /звідки.*прибув|^звідки$/);
  const people: PbShPerson[] = [];

  sheet.rawRows.slice(1).forEach((row, offset) => {
    const fullName = cell(row, nameCol);
    const positionIndex = cell(row, indexCol);
    const status = cell(row, statusCol);
    const personId = idCell(row, idCol);
    if (!fullName && !personId && !positionIndex) return;
    if (!fullName && (!status || status === "0")) return;
    people.push({
      excelRow: offset + 2,
      personId: personId && personId !== "0" ? personId : "",
      fullName,
      rank: cell(row, rankCol),
      positionIndex,
      positionTitle: cell(row, posCol),
      status,
      arrivedFrom: cell(row, fromCol),
    });
  });
  return people;
};

export const parsePbArchive = (workbook: ExcelWorkbookSnapshot): PbArchivePeriod[] => {
  const sheet = findSheet(workbook, /^archive$/i);
  if (!sheet) return [];
  const headerRowIndex = sheet.rawRows.findIndex((row) =>
    /вид\s*вибут|прізвище/i.test(row.map(norm).join(" ")),
  );
  if (headerRowIndex < 0) return [];
  const headers = headerMap(sheet.rawRows[headerRowIndex] ?? []);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /прізвище|піб/);
  const rankCol = findCol(headers, /зван/);
  const typeCol = findCol(headers, /вид\s*вибут/);
  const dateCol = findCol(headers, /з якої дати|дата вибут/);
  const placeCol = findCol(headers, /куди виб/);
  const orderNumCol = findCol(headers, /номер наказу вибут/);
  const orderDateCol = findCol(headers, /дата наказу вибут/);
  const plannedCol = findCol(headers, /планова дата/);
  // Фактичне повернення — колонки 29–31 («Дата прибуття» / «Дата наказу» /
  // «Номер наказу»), не «Планова дата прибуття» і не наказ вибуття.
  const returnCol = findCol(
    headers,
    /^дата прибуття$/,
    /фактичн.*(?:поверн|прибут)/,
    /^дата поверн/,
  );
  const returnOrderNumCol = findCol(
    headers,
    /^номер наказу$/,
    /номер наказу\s*(?:прибут|поверн)/,
    /наказ\s*(?:на\s+)?(?:прибут|поверн)/,
  );
  const returnOrderDateCol = findCol(
    headers,
    /^дата наказу$/,
    /дата наказу\s*(?:прибут|поверн)/,
  );
  const posCol = findCol(headers, /займана посад|посада/);
  const numCol = findCol(headers, /№\s*з\/п|^№$/);
  const periods: PbArchivePeriod[] = [];

  sheet.rawRows.slice(headerRowIndex + 1).forEach((row, offset) => {
    const fullName = cell(row, nameCol);
    if (!fullName) return;
    periods.push({
      excelRow: headerRowIndex + offset + 2,
      periodNumber: cell(row, numCol),
      personId: idCell(row, idCol),
      fullName,
      rank: cell(row, rankCol),
      positionTitle: cell(row, posCol),
      absenceType: cell(row, typeCol),
      departDate: cell(row, dateCol),
      place: cell(row, placeCol),
      orderNumber: cell(row, orderNumCol),
      orderDate: cell(row, orderDateCol),
      plannedReturn: cell(row, plannedCol),
      returnDate: cell(row, returnCol),
      returnOrderNumber: cell(row, returnOrderNumCol),
      returnOrderDate: cell(row, returnOrderDateCol),
    });
  });
  return periods;
};

export const parsePbMovements = (workbook: ExcelWorkbookSnapshot): PbMovement[] => {
  const sheet = findSheet(workbook, /^рух$/i);
  if (!sheet) return [];
  const headers = headerMap(sheet.rawRows[0] ?? []);
  const numCol = findCol(headers, /^№$/);
  const typeCol = findCol(headers, /^тип$/);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /^піб$/);
  const rankCol = findCol(headers, /зван/);
  const statusCol = findCol(headers, /^статус$/);
  const prevIdxCol = findCol(headers, /індекс.*попер|попер.*індекс/);
  const nextIdxCol = findCol(headers, /індекс.*як|яка зміна.*індекс|індекси посад \(яка/);
  // Fallback columns when layout differs (older file had ID in H, type in E)
  const changeCol = findCol(headers, /яка зміна/, /попер/);
  const destCol = findCol(
    headers,
    /^куди(?:\s|$)/,
    /куди.*(?:перев|виб)/,
  );
  const fromCol = findCol(headers, /звідки.*прибув|^звідки$/);
  const noteCol = findCol(headers, /^примітка$/);
  const orderNumCol = findCol(headers, /^наказ$/);
  const orderDateCol = findCol(headers, /^дата$/);
  const movements: PbMovement[] = [];

  const normalizeType = (value: string) => {
    const text = value.toUpperCase();
    if (text.includes("ПОСАД")) return "ПОСАДА";
    if (text.includes("ПРИБ")) return "ПРИБУВ";
    if (text.includes("ЗМІНИШТАТ") || text.includes("ЗМІНИ ШТАТ")) {
      return "ПОСАДА";
    }
    if (text.includes("РОЗПОР")) return "РОЗПОРЯДЖ";
    if (text.includes("СКАС")) {
      if (text.includes("СЗЧ") || text.includes("САМОВІЛ")) return "СЗЧ";
      return "СКАСУВАННЯ";
    }
    if (text.includes("ПЕРЕВ")) return "ПЕРЕВ";
    if (text.includes("ЗВАН")) return "ЗВАННЯ";
    if (text.includes("ЗВІЛ")) return "ЗВІЛЬН";
    if (text.includes("СЗЧ") || text.includes("САМОВІЛ")) return "СЗЧ";
    return text || "—";
  };

  sheet.rawRows.slice(1).forEach((row, offset) => {
    const type = normalizeType(cell(row, typeCol >= 0 ? typeCol : 4));
    const fullName = cell(row, nameCol >= 0 ? nameCol : 6);
    const movementNumber = cell(row, numCol >= 0 ? numCol : 1);
    if (!fullName && !movementNumber) return;
    if (type === "—" && !fullName) return;
    const note = cell(row, noteCol >= 0 ? noteCol : 18);
    const rawDest = cell(row, destCol);
    movements.push({
      excelRow: offset + 2,
      movementNumber,
      type,
      // ID беремо лише з явно знайденої колонки. Позиційний fallback міг
      // прочитати дату народження як ID, якщо структура РУХ змінилася.
      personId: idCell(row, idCol),
      fullName,
      // Звання читаємо лише з колонки, яку однозначно визначено заголовком.
      // Позиційний fallback F міг підставити стороннє значення з іншого макета.
      rank: cell(row, rankCol),
      previousIndex: cell(row, prevIdxCol),
      nextIndex: cell(row, nextIdxCol),
      destination:
        type === "ПЕРЕВ"
          ? resolveOutboundTransferDestination(rawDest, note)
          : resolveMovementDestination(rawDest, note),
      orderNumber: cell(row, orderNumCol),
      orderDate: cell(row, orderDateCol),
      // У поточному форматі РУХ: O+P — документ-підстава, I+J — стройовий наказ.
      basisNumber: cell(row, 14),
      basisDate: cell(row, 15),
      changeText: cell(row, changeCol),
      status: cell(row, statusCol >= 0 ? statusCol : 2),
      note,
      arrivedFrom: cell(row, fromCol),
    });
  });
  return movements;
};

export type EjoosAbsentRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  ground: string;
  place: string;
  departDate: string;
  actualReturn: string;
};

export type EjoosTimesheetRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  dayValue: string;
};

/** Повний рядок Табеля: штатний або історичний, з усіма днями місяця. */
export type EjoosTimesheetPersonScan = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  hasDepartureText: boolean;
  firstDepartureDay: number;
  plusDays: number[];
  /** 1-based day → нормалізована позначка (або «вибув»). */
  dayCodes: string[];
};

export type EjoosShpoRow = {
  excelRow: number;
  positionIndex: string;
  personId: string;
  fullName: string;
  rank: string;
};

export type EjoosArrivalRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  arriveDate: string;
  fromUnit: string;
};

export type EjoosOosRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
};

export type EjoosExcludedRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  orderNumber: string;
  orderDate: string;
};

export const parseEjoosAbsents = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosAbsentRow[] => {
  if (!sheet) return [];
  const rows: EjoosAbsentRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]);
    const positionIndex = norm(row?.[3]);
    if (!fullName && !positionIndex) continue;
    rows.push({
      excelRow: i + 1,
      personId: normId(row?.[2]),
      fullName,
      rank: norm(row?.[0]),
      positionIndex,
      ground: norm(row?.[4]),
      place: norm(row?.[5]),
      departDate: norm(row?.[6]),
      actualReturn: norm(row?.[12]),
    });
  }
  return rows;
};

export const parseEjoosTimesheetDay = (
  sheet: ExcelSheetSnapshot | undefined,
  day: number,
): EjoosTimesheetRow[] => {
  if (!sheet || day < 1 || day > 31) return [];
  const dayCol = 8 + (day - 1);
  const rows: EjoosTimesheetRow[] = [];
  for (let i = 6; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const positionIndex = norm(row?.[1]);
    const fullName = norm(row?.[6]);
    if (!positionIndex || !/^\d/.test(positionIndex)) continue;
    rows.push({
      excelRow: i + 1,
      personId: normId(row?.[7]),
      fullName,
      rank: norm(row?.[5]),
      positionIndex,
      dayValue: norm(row?.[dayCol]),
    });
  }
  return rows;
};

export const parseEjoosTimesheetPeople = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosTimesheetPersonScan[] => {
  if (!sheet) return [];
  const rows: EjoosTimesheetPersonScan[] = [];
  for (let i = 6; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const positionIndex = norm(row?.[1]);
    const fullName = norm(row?.[6]);
    const personId = normId(row?.[7]);
    if (!positionIndex || !/^\d/.test(positionIndex)) continue;
    if (!fullName && !personId) continue;
    const plusDays: number[] = [];
    const dayCodes: string[] = [];
    let hasDepartureText = false;
    let firstDepartureDay = 0;
    for (let day = 1; day <= 31; day += 1) {
      const value = row?.[8 + (day - 1)];
      if (isTimesheetDepartureMark(value)) {
        hasDepartureText = true;
        if (!firstDepartureDay) firstDepartureDay = day;
        dayCodes[day] = "вибув";
        continue;
      }
      const code = norm(value);
      dayCodes[day] = code;
      if (code === "+") plusDays.push(day);
    }
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[5]),
      positionIndex,
      hasDepartureText,
      firstDepartureDay,
      plusDays,
      dayCodes,
    });
  }
  return rows;
};

export const usablePersonId = (...ids: Array<string | undefined | null>) => {
  for (const id of ids) {
    const value = String(id || "").trim();
    if (value && value !== "0") return value;
  }
  return "";
};

/** Якщо в рядку немає ID — беремо з ШПО за ПІБ або індексом посади. */
export const personIdFromShpo = (
  rows: EjoosShpoRow[],
  input: { fullName?: string; positionIndex?: string; personId?: string },
) => {
  const known = usablePersonId(input.personId);
  if (known) return known;
  const name = canonicalName(input.fullName || "");
  const index = String(input.positionIndex || "").trim();
  const byName = name
    ? rows.find(
        (row) => usablePersonId(row.personId) && canonicalName(row.fullName) === name,
      )
    : undefined;
  if (byName) return byName.personId;
  const byIndex = index
    ? rows.find(
        (row) =>
          usablePersonId(row.personId) && row.positionIndex === index,
      )
    : undefined;
  return usablePersonId(byIndex?.personId);
};

export const parseEjoosShpo = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosShpoRow[] => {
  if (!sheet) return [];
  const rows: EjoosShpoRow[] = [];
  for (let i = 6; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const positionIndex = norm(row?.[0]);
    if (!positionIndex || !/^\d/.test(positionIndex)) continue;
    rows.push({
      excelRow: i + 1,
      positionIndex,
      rank: norm(row?.[5]),
      fullName: norm(row?.[6]),
      personId: normId(row?.[7]),
    });
  }
  return rows;
};

export const parseEjoosArrivals = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosArrivalRow[] => {
  if (!sheet) return [];
  const rows: EjoosArrivalRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]) || norm(row?.[6]);
    // Колонка C у «Тимчасово прибулі» — часто РНОКПП, не штатний ID.
    const personId = journalIdCell(row, 2);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]) || norm(row?.[5]),
      positionIndex: norm(row?.[4]) || norm(row?.[3]),
      fromUnit: norm(row?.[5]) || norm(row?.[8]),
      arriveDate: norm(row?.[7]) || norm(row?.[10]),
    });
  }
  return rows;
};

export const parseEjoosOos = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosOosRow[] => {
  if (!sheet) return [];
  const rows: EjoosOosRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]);
    const personId = normId(row?.[2]);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]),
      positionIndex: norm(row?.[3]),
    });
  }
  return rows;
};

export const parseEjoosExcluded = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosExcludedRow[] => {
  if (!sheet) return [];
  const rows: EjoosExcludedRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]) || norm(row?.[6]);
    const personId = normId(row?.[2]) || normId(row?.[7]);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      orderDate: norm(row?.[28]) || norm(row?.[10]),
      orderNumber: norm(row?.[29]) || norm(row?.[11]),
    });
  }
  return rows;
};

export const parseTimesheetDayFromPbName = (
  fileName: string,
  fallback = new Date(),
) => {
  const match = fileName.match(/(\d{2})[._-](\d{2})[._-](\d{2,4})/);
  if (match) {
    const day = Number(match[1]);
    if (day >= 1 && day <= 31) {
      const month = Number(match[2]);
      const yearRaw = Number(match[3]);
      const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
      return {
        day,
        label: `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`,
      };
    }
  }
  // Also 1ПБ_25082026
  const compact = fileName.match(/_(\d{2})(\d{2})(\d{4})/);
  if (compact) {
    const day = Number(compact[1]);
    const month = Number(compact[2]);
    const year = Number(compact[3]);
    return {
      day,
      label: `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`,
    };
  }
  return {
    day: fallback.getDate(),
    label: fallback.toLocaleDateString("uk-UA"),
  };
};

/** День табеля з назви файлу 1ПБ / ЄЖООС або сьогодні. */
export const resolveJournalTimesheetDay = parseTimesheetDayFromPbName;

/** Якщо план зібрано раніше в цьому місяці — на застосуванні тягнемо табель до сьогодні. */
export const refreshPlanTimesheetHorizon = (
  plan: Pick<EjoosSyncPlan, "timesheetDay" | "timesheetDayLabel">,
  now = new Date(),
): Pick<EjoosSyncPlan, "timesheetDay" | "timesheetDayLabel"> => {
  const today = now.getDate();
  const todayLabel = now.toLocaleDateString("uk-UA");
  const match = String(plan.timesheetDayLabel || "").match(
    /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
  );
  if (match) {
    const planMonth = Number(match[2]);
    const planYear =
      Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
    const sameMonth =
      planMonth === now.getMonth() + 1 && planYear === now.getFullYear();
    if (sameMonth && today > plan.timesheetDay) {
      return { timesheetDay: today, timesheetDayLabel: todayLabel };
    }
    return {
      timesheetDay: plan.timesheetDay,
      timesheetDayLabel: plan.timesheetDayLabel,
    };
  }
  if (today > plan.timesheetDay) {
    return { timesheetDay: today, timesheetDayLabel: todayLabel };
  }
  return {
    timesheetDay: plan.timesheetDay,
    timesheetDayLabel: plan.timesheetDayLabel,
  };
};

const opId = (parts: string[]) =>
  parts
    .map((part) => part.replace(/\s+/g, "_").slice(0, 40))
    .filter(Boolean)
    .join("__");

export const buildEjoosSyncPlan = (
  ejoos: ExcelWorkbookSnapshot,
  pb: ExcelWorkbookSnapshot,
  options?: {
    statusRules?: import("./ejoosRules").EjoosStatusRule[];
    processedMovementKeys?: Iterable<string>;
  },
): EjoosSyncPlan => {
  const statusRules =
    options?.statusRules ?? readOperatorSettings().statusRules;
  const mapStatus = (raw: string) => mapPbStatusToEjoosWithRules(raw, statusRules);
  const { day: timesheetDay, label: timesheetDayLabel } = parseTimesheetDayFromPbName(
    pb.fileName,
  );
  const shPeople = parsePbShPeople(pb);
  const archiveAll = parsePbArchive(pb);
  const movementsAll = parsePbMovements(pb);
  const activeMovementsAll = movementsAll.filter(
    (event) => !isCancelledMovementRecord(event),
  );
  const processedMovementKeys = new Set(options?.processedMovementKeys ?? []);
  const wasMovementProcessed = (event: PbMovement) => {
    const key = createMovementKey(event);
    return Boolean(key && processedMovementKeys.has(key));
  };

  const absentSheet = findSheet(ejoos, /тимчасов.*відсут/i);
  const arrivalSheet = findSheet(ejoos, /тимчасов.*прибул/i);
  const excludedSheet = findSheet(ejoos, /виключен/i);
  const timesheetSheet = findSheet(ejoos, /табель/i);
  const shpoSheet = findSheet(ejoos, /шпо|штатно.?посад/i);
  const oosSheet = findSheet(ejoos, /(^|[.\s])оос($|[\s])/i);
  const ejoosAbsents = parseEjoosAbsents(absentSheet);
  const ejoosArrivals = parseEjoosArrivals(arrivalSheet);
  const ejoosExcluded = parseEjoosExcluded(excludedSheet);
  const ejoosDays = parseEjoosTimesheetDay(timesheetSheet, timesheetDay);
  const timesheetPeople = parseEjoosTimesheetPeople(timesheetSheet);
  const ejoosShpo = parseEjoosShpo(shpoSheet);
  const ejoosOos = parseEjoosOos(oosSheet);

  const shIds = new Set(
    shPeople.map((person) => person.personId).filter(Boolean),
  );
  const shNames = new Set(
    shPeople.map((person) => canonicalName(person.fullName)).filter(Boolean),
  );
  const shPersonById = new Map(
    shPeople
      .filter((person) => person.personId)
      .map((person) => [person.personId, person]),
  );
  const shPersonByName = new Map(
    shPeople
      .filter((person) => person.fullName)
      .map((person) => [canonicalName(person.fullName), person]),
  );
  const shpoPersonById = new Map(
    ejoosShpo
      .filter((row) => row.personId)
      .map((row) => [row.personId, row]),
  );
  const shpoPersonByName = new Map(
    ejoosShpo
      .filter((row) => row.fullName)
      .map((row) => [canonicalName(row.fullName), row]),
  );
  const oosPersonById = new Map(
    ejoosOos
      .filter((row) => row.personId)
      .map((row) => [row.personId, row]),
  );
  const oosPersonByName = new Map(
    ejoosOos
      .filter((row) => row.fullName)
      .map((row) => [canonicalName(row.fullName), row]),
  );
  // Поточний облік — лише ШПО + ООС. «Виключені» і старі рядки Табеля —
  // історія завершених вибуттів, а не доказ, що людина зараз вибула.
  const ejoosOccupied = [
    ...ejoosShpo.filter((row) => row.fullName || row.personId),
    ...ejoosOos.filter((row) => row.fullName || row.personId),
  ];
  const ejoosIds = new Set(
    ejoosOccupied.map((row) => row.personId).filter(Boolean),
  );
  const ejoosNames = new Set(
    ejoosOccupied.map((row) => canonicalName(row.fullName)).filter(Boolean),
  );
  const activeArrivalIds = new Set(
    ejoosArrivals.map((row) => row.personId).filter(Boolean),
  );
  const activeArrivalNames = new Set(
    ejoosArrivals.map((row) => canonicalName(row.fullName)).filter(Boolean),
  );
  const openAbsentRows = ejoosAbsents.filter((row) => !row.actualReturn);
  const openAbsentIds = new Set(
    openAbsentRows.map((row) => row.personId).filter(Boolean),
  );
  const openAbsentNames = new Set(
    openAbsentRows.map((row) => canonicalName(row.fullName)).filter(Boolean),
  );

  // Один ID — одна людина. Написання ПІБ у джерелах різне (Миколайович /
  // Михайлович), а частина рядків ЕЖООС узагалі без ID, тому для звірки
  // тримаємо всі варіанти ПІБ, зв'язані спільним ID.
  const nameVariantsById = new Map<
    string,
    Map<string, { display: string; sources: Set<string> }>
  >();
  const addNameVariant = (
    personId: string,
    fullName: string,
    source: string,
  ) => {
    const id = normId(personId);
    const name = norm(fullName);
    const key = canonicalName(name);
    if (!id || !key) return;
    const variants =
      nameVariantsById.get(id) ??
      new Map<string, { display: string; sources: Set<string> }>();
    const variant = variants.get(key) ?? { display: name, sources: new Set() };
    variant.sources.add(source);
    variants.set(key, variant);
    nameVariantsById.set(id, variants);
  };
  for (const event of activeMovementsAll) {
    addNameVariant(event.personId, event.fullName, "Рух");
  }
  for (const period of archiveAll) {
    addNameVariant(period.personId, period.fullName, "archive");
  }
  for (const person of shPeople) {
    addNameVariant(person.personId, person.fullName, "sh");
  }
  for (const row of ejoosShpo) addNameVariant(row.personId, row.fullName, "ШПО");
  for (const row of ejoosOos) addNameVariant(row.personId, row.fullName, "ООС");
  for (const row of ejoosAbsents) {
    addNameVariant(row.personId, row.fullName, "Тимч. відсутні");
  }
  for (const row of ejoosArrivals) {
    addNameVariant(row.personId, row.fullName, "Тимч. прибулі");
  }

  const idsByName = new Map<string, Set<string>>();
  for (const [personId, variants] of nameVariantsById) {
    for (const key of variants.keys()) {
      const ids = idsByName.get(key) ?? new Set<string>();
      ids.add(personId);
      idsByName.set(key, ids);
    }
  }
  /** Усі написання ПІБ цієї особи: пряме, і всі зв'язані спільним ID. */
  const personNameKeys = (personId: string, fullName: string) => {
    const keys = new Set<string>();
    const name = canonicalName(norm(fullName));
    if (name) keys.add(name);
    const ids = new Set<string>();
    const id = normId(personId);
    if (id) ids.add(id);
    if (!id && name) {
      const linked = idsByName.get(name);
      // Однакове ПІБ у різних ID — не вгадуємо, працюємо лише з прямим ПІБ.
      if (linked?.size === 1) ids.add([...linked][0]);
    }
    for (const candidate of ids) {
      for (const key of nameVariantsById.get(candidate)?.keys() ?? []) {
        keys.add(key);
      }
    }
    return keys;
  };
  /** Пошук рядка за будь-яким написанням ПІБ цієї особи. */
  const byPersonName = <T,>(
    map: Map<string, T>,
    personId: string,
    fullName: string,
  ) => {
    for (const key of personNameKeys(personId, fullName)) {
      const found = map.get(key);
      if (found) return found;
    }
    return null;
  };
  /** Рядки ЕЖООС часто без ID, тому порівнюємо ще й за псевдонімами ПІБ. */
  const isSamePerson = (
    left: { personId: string; fullName: string },
    right: { personId: string; fullName: string },
  ) => {
    const leftId = normId(left.personId);
    const rightId = normId(right.personId);
    if (leftId && rightId) return leftId === rightId;
    const rightName = canonicalName(norm(right.fullName));
    if (!rightName) return false;
    return personNameKeys(left.personId, left.fullName).has(rightName);
  };

  const personStillInEjoos = (personId: string, fullName: string) =>
    Boolean(
      (personId && ejoosIds.has(personId)) ||
        [...personNameKeys(personId, fullName)].some((key) =>
          ejoosNames.has(key),
        ),
    );
  const personStillInSh = (personId: string, fullName: string) =>
    Boolean(
      (personId && shIds.has(personId)) ||
        [...personNameKeys(personId, fullName)].some((key) => shNames.has(key)),
    );
  const onStaffShpo = (personId: string, fullName: string) =>
    Boolean(
      (personId && shpoPersonById.has(personId)) ||
        byPersonName(shpoPersonByName, personId, fullName),
    );
  const hasOpenAbsence = (personId: string, fullName: string) =>
    Boolean(
      (personId && openAbsentIds.has(personId)) ||
        [...personNameKeys(personId, fullName)].some((key) =>
          openAbsentNames.has(key),
        ),
    );
  /**
   * Особу вже знято з штатної посади (розпорядження / безвісти), а відкрита
   * відсутність у ЕЖООС є. Повторно ставити на вакантний індекс або писати
   * новий рядок БЕЗВІСТИ не треба — це вже проведений стан.
   */
  const alreadyVacatedForAbsence = (personId: string, fullName: string) =>
    isStaleVacatedAbsence({
      onStaffShpo: onStaffShpo(personId, fullName),
      hasOpenAbsence: hasOpenAbsence(personId, fullName),
      stillInSh: personStillInSh(personId, fullName),
    });
  const findMovementExcludedRow = (event: PbMovement) =>
    ejoosExcluded.find((row) => {
      if (!isSamePerson(event, row)) return false;
      const sameOrder = Boolean(
        event.orderNumber &&
          row.orderNumber &&
          normKey(row.orderNumber) === normKey(event.orderNumber),
      );
      const sameDate = Boolean(
        event.orderDate &&
          row.orderDate &&
          normKey(row.orderDate) === normKey(event.orderDate),
      );
      return event.orderNumber ? sameOrder : sameDate;
    }) ?? null;
  const findLatestExcludedRow = (personId: string, fullName: string) =>
    [...ejoosExcluded]
      .filter((row) => isSamePerson({ personId, fullName }, row))
      .sort((left, right) => right.excelRow - left.excelRow)[0] ?? null;
  const movementPersonKey = (event: {
    personId: string;
    fullName: string;
  }) =>
    event.personId
      ? `id:${event.personId}`
      : event.fullName
        ? `name:${normKey(event.fullName)}`
        : "";
  const movementEventTime = (event: PbMovement) =>
    dateMs(event.orderDate || event.basisDate) || event.excelRow;

  /**
   * ПЕРЕВ + СКАСУВАННЯ в тому ж місяці: переведення не чинне.
   * Один рядок Табеля; запис у «Виключені» прибираємо.
   */
  const latestTransferCancelByPerson = new Map<string, PbMovement>();
  const cancelledExternalTransferRows = new Set<number>();
  for (const event of movementsAll) {
    if (!isTransferCancellation(event)) continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    const previous = latestTransferCancelByPerson.get(key);
    if (
      !previous ||
      movementEventTime(event) > movementEventTime(previous) ||
      (movementEventTime(event) === movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      latestTransferCancelByPerson.set(key, event);
    }
  }
  for (const event of movementsAll) {
    if (event.type !== "ПЕРЕВ" || isOwnUnitStaffMove(event)) continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    if (isCancelledMovementRecord(event) || isTransferCancellation(event)) {
      cancelledExternalTransferRows.add(event.excelRow);
      continue;
    }
    const cancel = latestTransferCancelByPerson.get(key);
    if (!cancel) continue;
    if (
      movementEventTime(cancel) > movementEventTime(event) ||
      (movementEventTime(cancel) === movementEventTime(event) &&
        cancel.excelRow > event.excelRow)
    ) {
      cancelledExternalTransferRows.add(event.excelRow);
    }
  }
  const isLaterMovement = (later: PbMovement, earlier: PbMovement) =>
    movementEventTime(later) > movementEventTime(earlier) ||
    (movementEventTime(later) === movementEventTime(earlier) &&
      later.excelRow > earlier.excelRow);
  const laterOutboundStaffDeparture = (event: PbMovement): PbMovement | null => {
    if (personStillInSh(event.personId, event.fullName)) return null;
    let found: PbMovement | null = null;
    for (const other of activeMovementsAll) {
      if (other.excelRow === event.excelRow) continue;
      if (!isSamePerson(event, other)) continue;
      if (cancelledExternalTransferRows.has(other.excelRow)) continue;
      if (!isOutboundStaffMove(other)) continue;
      if (!isLaterMovement(other, event)) continue;
      if (!found || isLaterMovement(other, found)) found = other;
    }
    return found;
  };
  const ownUnitMoveSuperseded = (event: PbMovement) =>
    ownUnitMoveSupersededByOutbound({
      stillInSh: personStillInSh(event.personId, event.fullName),
      hasLaterOutbound: Boolean(laterOutboundStaffDeparture(event)),
    });
  const transferCancelForPerson = (personId: string, fullName: string) => {
    if (personId) {
      const found = latestTransferCancelByPerson.get(`id:${personId}`);
      if (found) return found;
    }
    for (const event of latestTransferCancelByPerson.values()) {
      if (isSamePerson({ personId, fullName }, event)) return event;
    }
    if (fullName) {
      return latestTransferCancelByPerson.get(`name:${normKey(fullName)}`) ?? null;
    }
    return null;
  };
  const transferCancelOf = (event: PbMovement) =>
    transferCancelForPerson(event.personId, event.fullName);
  const cancelledTransferOf = (event: { personId: string; fullName: string }) => {
    const cancel = transferCancelForPerson(event.personId, event.fullName);
    if (!cancel) return null;
    let found: PbMovement | null = null;
    for (const movement of movementsAll) {
      if (movement.type !== "ПЕРЕВ" || isOwnUnitStaffMove(movement)) {
        continue;
      }
      if (!cancelledExternalTransferRows.has(movement.excelRow)) continue;
      if (!isSamePerson(event, movement)) continue;
      if (movementEventTime(movement) > movementEventTime(cancel)) continue;
      if (
        !found ||
        movementEventTime(movement) > movementEventTime(found) ||
        (movementEventTime(movement) === movementEventTime(found) &&
          movement.excelRow > found.excelRow)
      ) {
        found = movement;
      }
    }
    return found;
  };

  const shpoByIndex = new Map(ejoosShpo.map((row) => [row.positionIndex, row]));
  const dayByIndex = new Map<string, EjoosTimesheetRow>();
  const dayRowScore = (row: EjoosTimesheetRow) => {
    const shpo = shpoByIndex.get(row.positionIndex);
    if (!shpo) return !row.personId && !row.fullName ? 2 : 1;
    if (shpo.personId && row.personId === shpo.personId) return 4;
    if (
      shpo.fullName &&
      normKey(row.fullName) === normKey(shpo.fullName)
    ) {
      return 4;
    }
    if (!shpo.personId && !shpo.fullName && !row.personId && !row.fullName) {
      return 3;
    }
    return 1;
  };
  for (const row of ejoosDays) {
    const current = dayByIndex.get(row.positionIndex);
    if (!current || dayRowScore(row) > dayRowScore(current)) {
      dayByIndex.set(row.positionIndex, row);
    }
  }
  const timesheetScanByRow = new Map(
    timesheetPeople.map((row) => [row.excelRow, row]),
  );
  const timesheetActivityScore = (row: EjoosTimesheetRow) => {
    const scan = timesheetScanByRow.get(row.excelRow);
    if (scan?.hasDepartureText && !scan.plusDays.length) return 0;
    if (scan?.hasDepartureText) return 1 + dayRowScore(row);
    return 10 + dayRowScore(row);
  };
  const dayById = new Map<string, EjoosTimesheetRow>();
  for (const row of ejoosDays) {
    if (!row.personId) continue;
    const current = dayById.get(row.personId);
    if (!current || timesheetActivityScore(row) > timesheetActivityScore(current)) {
      dayById.set(row.personId, row);
    }
  }
  const activeTimesheetRowOf = (
    personId: string,
    fullName: string,
    staffIndex = "",
  ) => {
    const byIndex = staffIndex ? dayByIndex.get(staffIndex) : undefined;
    if (
      byIndex &&
      isSamePerson({ personId, fullName }, byIndex) &&
      timesheetActivityScore(byIndex) > 0
    ) {
      return byIndex;
    }
    if (personId) {
      const byId = dayById.get(personId);
      if (byId && timesheetActivityScore(byId) > 0) return byId;
    }
    return (
      timesheetRowsOf(personId, fullName)
        .map((scan) =>
          ejoosDays.find((row) => row.excelRow === scan.excelRow),
        )
        .filter((row): row is EjoosTimesheetRow => Boolean(row))
        .sort((left, right) => timesheetActivityScore(right) - timesheetActivityScore(left))[0] ??
      null
    );
  };
  const timesheetRowsOf = (personId: string, fullName: string) =>
    timesheetPeople.filter((row) => isSamePerson({ personId, fullName }, row));
  const priorEpisodeTimesheetOf = (
    personId: string,
    fullName: string,
    activeExcelRow: number,
    staffIndex: string,
  ) =>
    timesheetRowsOf(personId, fullName).find(
      (row) =>
        row.excelRow !== activeExcelRow &&
        (row.hasDepartureText ||
          Boolean(
            staffIndex &&
              row.positionIndex &&
              row.positionIndex !== staffIndex,
          )),
    ) ?? null;
  const isVacantStaffRow = (row: { personId: string; fullName: string } | null | undefined) =>
    Boolean(row) && !row!.personId && !row!.fullName;
  /** Рядок Табеля особи на штатному індексі — не чужий рядок dayByIndex. */
  const staffIndexTimesheetForPerson = (
    personId: string,
    fullName: string,
    positionIndex: string,
  ) => {
    const personRows = timesheetRowsOf(personId, fullName).filter(
      (row) => row.positionIndex === positionIndex,
    );
    if (personRows.length) {
      return [...personRows].sort(
        (left, right) => right.plusDays.length - left.plusDays.length,
      )[0];
    }
    const indexRow = positionIndex ? dayByIndex.get(positionIndex) : undefined;
    if (indexRow && isSamePerson({ personId, fullName }, indexRow)) return indexRow;
    if (isVacantStaffRow(indexRow)) return indexRow ?? null;
    return null;
  };
  const shOccupantByIndex = new Map(
    shPeople
      .filter(
        (person) =>
          person.positionIndex && isPositionIndex(person.positionIndex),
      )
      .map((person) => [person.positionIndex, person]),
  );
  /**
   * Скасоване переведення — не розбиваємо Табель на історію + новий рядок.
   * Фактичного вибуття не було; лишається один активний рядок з дати постановки.
   */
  const timesheetNeedsTransferCancelSplit = (
    _personId: string,
    _fullName: string,
    _staffIndex: string,
    _cancelDate: string,
  ) => false;

  /** Дата постановки на поточний штатний індекс (ПОСАДА / внутрішній ПЕРЕВ). */
  const staffAppointmentDateFor = (
    personId: string,
    fullName: string,
    positionIndex: string,
  ) => {
    let found: PbMovement | null = null;
    for (const event of activeMovementsAll) {
      if (!isSamePerson({ personId, fullName }, event)) continue;
      if (!eventInLeadWindow(event)) continue;
      const isInbound =
        isOwnUnitStaffMove(event);
      if (!isInbound) continue;
      const targetIndex = event.nextIndex || "";
      if (
        positionIndex &&
        targetIndex &&
        isPositionIndex(targetIndex) &&
        targetIndex !== positionIndex
      ) {
        continue;
      }
      if (
        !found ||
        movementEventTime(event) > movementEventTime(found) ||
        (movementEventTime(event) === movementEventTime(found) &&
          event.excelRow > found.excelRow)
      ) {
        found = event;
      }
    }
    return found?.orderDate || found?.basisDate || "";
  };

  /**
   * ПОСАДА також може бути далеко за межами хвоста РУХ. Підтягуємо останню
   * подію, яка підтверджує поточну особу та індекс у `sh`; інакше різниця
   * sh ↔ старий ШПО помилково виглядає як звичайний конфлікт зайнятості.
   */
  const currentPositionMovements = new Map<
    string,
    (typeof movementsAll)[number]
  >();
  for (const event of activeMovementsAll) {
    if (
      !isOwnUnitStaffMove(event) ||
      !event.fullName
    ) {
      continue;
    }
    const person =
      (event.personId && shPersonById.get(event.personId)) ||
      byPersonName(shPersonByName, event.personId, event.fullName) ||
      shPeople.find((candidate) => isSamePerson(candidate, event)) ||
      null;
    const currentEjoosPerson =
      ((person?.personId || event.personId) &&
        shpoPersonById.get(person?.personId || event.personId)) ||
      byPersonName(shpoPersonByName, event.personId, event.fullName) ||
      ((person?.personId || event.personId) &&
        oosPersonById.get(person?.personId || event.personId)) ||
      byPersonName(oosPersonByName, event.personId, event.fullName) ||
      null;
    if (!person?.positionIndex || !isPositionIndex(person.positionIndex)) {
      // Особу вже прибрали з поточного sh, а її стару позицію зайняв хтось
      // інший. Не губимо підтверджений внутрішній ПЕРЕВ — показуємо його
      // окремою операцією для ручної перевірки/застосування.
      // Якщо далі є зовнішнє вибуття, внутрішню постановку не проводимо:
      // кінцевий стан визначає ПЕРЕВ / sh, а не проміжна ПОСАДА.
      if (
        currentEjoosPerson &&
        isOwnUnitStaffMove(event) &&
        !ownUnitMoveSuperseded(event)
      ) {
        currentPositionMovements.set(canonicalName(event.fullName), event);
      }
      continue;
    }
    const confirmsCurrentIndex =
      event.nextIndex === person.positionIndex ||
      String(event.changeText || "").includes(person.positionIndex);
    const currentOosPerson =
      ((person.personId || event.personId) &&
        oosPersonById.get(person.personId || event.personId)) ||
      byPersonName(oosPersonByName, event.personId, event.fullName) ||
      null;
    const oosPositionIndexes = (currentOosPerson?.positionIndex || "")
      .split(/[^0-9]+/)
      .filter(Boolean);
    const oosNeedsPositionUpdate =
      !currentOosPerson ||
      !oosPositionIndexes.includes(person.positionIndex);
    const indexActuallyChanged = Boolean(
      currentEjoosPerson?.positionIndex &&
        currentEjoosPerson.positionIndex !== person.positionIndex,
    );
    const isNewPlacement = !currentEjoosPerson && confirmsCurrentIndex;
    const cancel = transferCancelForPerson(
      person.personId || event.personId,
      person.fullName || event.fullName,
    );
    const needsCancelSplit = Boolean(
      cancel?.orderDate &&
        timesheetNeedsTransferCancelSplit(
          person.personId || event.personId,
          person.fullName || event.fullName,
          person.positionIndex,
          cancel.orderDate,
        ),
    );
    // Якщо ШПО вже містить цю людину на поточному індексі sh, рух проведений
    // раніше й повторно показувати/застосовувати його не можна — окрім
    // скасованого переведення, де Табель ще не розкладений на історію + новий рядок.
    if (
      !indexActuallyChanged &&
      !isNewPlacement &&
      !oosNeedsPositionUpdate &&
      !needsCancelSplit
    ) {
      continue;
    }
    currentPositionMovements.set(canonicalName(person.fullName), event);
  }

  const currentPositionMovementRows = new Set(
    [...currentPositionMovements.values()].map((event) => event.excelRow),
  );
  const internalPositionMovementRows = new Set(
    [...currentPositionMovements.values()]
      .filter((event) => event.type === "ПЕРЕВ")
      .map((event) => event.excelRow),
  );

  /**
   * По одній особі може бути кілька непроведених подій. Історію зміни посади
   * втрачати не можна, тому ланцюг ПОСАДА / внутрішній ПЕРЕВ будуємо від
   * індексу, який ще стоїть у ЕЖООС, і проводимо кроки послідовно за датою.
   */
  const positionChainByPerson = new Map<string, PbMovement[]>();
  const chainedPositionRows = new Set<number>();
  const chainStepByRow = new Map<number, { step: number; total: number }>();
  {
    const positionEventsByPerson = new Map<string, PbMovement[]>();
    for (const event of activeMovementsAll) {
      // Ланцюг ведемо лише по переходах усередині 1ПБ: вибуття в іншу частину
      // проводиться через виключення, а не через зміну штатної посади.
      const isPositionEvent =
        isOwnUnitStaffMove(event);
      if (!isPositionEvent) continue;
      if (!event.previousIndex || !event.nextIndex) continue;
      if (event.previousIndex === event.nextIndex) continue;
      const key = movementPersonKey(event);
      if (!key) continue;
      const events = positionEventsByPerson.get(key) ?? [];
      events.push(event);
      positionEventsByPerson.set(key, events);
    }
    const eventTime = (event: PbMovement) =>
      dateMs(event.orderDate || event.basisDate);
    const isDispositionToken = (value: string) => /розпорядж/iu.test(value);
    const indexesConnect = (from: string, to: string) =>
      Boolean(from) &&
      Boolean(to) &&
      (from === to || (isDispositionToken(from) && isDispositionToken(to)));
    // Індекс у ШПО буває порожній (рядок лише з ПІБ), тому додатково
    // орієнтуємось на штатний рядок Табеля.
    const timesheetIndexById = new Map<string, string>();
    const timesheetIndexByName = new Map<string, string>();
    for (const row of ejoosDays) {
      if (!row.positionIndex) continue;
      if (/РОЗПОРЯДЖ/iu.test(row.dayValue)) continue;
      if (row.personId && !timesheetIndexById.has(row.personId)) {
        timesheetIndexById.set(row.personId, row.positionIndex);
      }
      const key = canonicalName(row.fullName);
      if (key && !timesheetIndexByName.has(key)) {
        timesheetIndexByName.set(key, row.positionIndex);
      }
    }
    for (const [key, events] of positionEventsByPerson) {
      events.sort(
        (left, right) =>
          eventTime(left) - eventTime(right) || left.excelRow - right.excelRow,
      );
      const personId = key.startsWith("id:") ? key.slice(3) : "";
      const fullName = events[0].fullName;
      const ejoosRow =
        (personId && shpoPersonById.get(personId)) ||
        byPersonName(shpoPersonByName, personId, fullName) ||
        null;
      // Точка відліку — те, що реально стоїть у ЕЖООС. Якщо індекс уже новий,
      // подія проведена раніше й у ланцюг не потрапляє.
      let currentIndex = ejoosRow?.positionIndex || "";
      // Якщо штатного рядка в ШПО вже немає, індекс у Табелі часто лишається
      // від історії вибуття в розпорядження. З нього ланцюг посад не будуємо —
      // інакше вже проведений РОЗПОРЯДЖ знову виглядає як зміна посади.
      if (!currentIndex) {
        const leftStaff =
          Boolean(
            (personId && oosPersonById.has(personId)) ||
              byPersonName(oosPersonByName, personId, fullName) ||
              (personId && openAbsentIds.has(personId)) ||
              [...personNameKeys(personId, fullName)].some((key) =>
                openAbsentNames.has(key),
              ),
          );
        if (!leftStaff) {
          currentIndex =
            (personId && timesheetIndexById.get(personId)) ||
            byPersonName(timesheetIndexByName, personId, fullName) ||
            "";
        }
      }
      // Повернення з розпорядження / СЗЧ: у ШПО штатного індексу ще немає,
      // ланцюг починаємо з ПОСАДИ «розпорядження → 2103…».
      if (!currentIndex) {
        const fromDisposition =
          events.find((event) => isDispositionToStaffPlacement(event)) ||
          events.find((event) => isDispositionToken(event.previousIndex));
        if (
          fromDisposition &&
          personStillInSh(personId, fullName)
        ) {
          currentIndex = fromDisposition.previousIndex;
        }
      }
      if (!currentIndex) continue;
      const pending: PbMovement[] = [];
      const visited = [currentIndex];
      for (const event of events) {
        if (!indexesConnect(event.previousIndex, currentIndex)) continue;
        currentIndex = event.nextIndex;
        // Повернення на індекс, який уже був у маршруті (А → Б → А):
        // проміжні кроки взаємно погашені, історію переносити нікуди.
        const loopAt = visited.indexOf(event.nextIndex);
        if (loopAt >= 0) {
          pending.length = loopAt;
          visited.length = loopAt + 1;
          continue;
        }
        pending.push(event);
        visited.push(event.nextIndex);
      }
      if (!pending.length) continue;
      // `sh` — джерело істини про поточну посаду. Якщо ланцюг веде не туди,
      // це стара або суперечлива подія: обробляємо звичайним шляхом.
      const shIndex =
        (personId && shPersonById.get(personId)?.positionIndex) ||
        byPersonName(shPersonByName, personId, fullName)?.positionIndex ||
        "";
      if (isPositionIndex(shIndex) && shIndex !== currentIndex) continue;
      const lastPending = pending[pending.length - 1];
      if (ownUnitMoveSuperseded(lastPending)) continue;
      const lastRow = lastPending.excelRow;
      const hasFollowUpEvent = activeMovementsAll.some(
        (event) =>
          movementPersonKey(event) === key &&
          event.excelRow > lastRow &&
          !cancelledExternalTransferRows.has(event.excelRow) &&
          (event.type === "РОЗПОРЯДЖ" ||
            event.type === "ПЕРЕВ" ||
            event.type === "ЗВІЛЬН"),
      );
      const total = pending.length + (hasFollowUpEvent ? 1 : 0);
      pending.forEach((event, index) => {
        chainedPositionRows.add(event.excelRow);
        chainStepByRow.set(event.excelRow, { step: index + 1, total });
      });
      positionChainByPerson.set(key, pending);
    }
  }
  const isPositionChangeRow = (excelRow: number) =>
    currentPositionMovementRows.has(excelRow) ||
    chainedPositionRows.has(excelRow);

  const operationalTypes = new Set([
    "ПОСАДА",
    "ПЕРЕВ",
    "РОЗПОРЯДЖ",
    "ПРИБУВ",
    "ЗВІЛЬН",
    "СЗЧ",
  ]);
  const latestOperationalRowByPerson = new Map<string, number>();
  for (const event of activeMovementsAll) {
    if (!operationalTypes.has(event.type)) continue;
    if (cancelledExternalTransferRows.has(event.excelRow)) continue;
    if (event.type === "СКАСУВАННЯ") continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    latestOperationalRowByPerson.set(
      key,
      Math.max(latestOperationalRowByPerson.get(key) ?? 0, event.excelRow),
    );
  }
  const effectiveMovements = activeMovementsAll.filter((event) => {
    if (event.type === "СКАСУВАННЯ") return false;
    if (cancelledExternalTransferRows.has(event.excelRow)) return false;
    if (!operationalTypes.has(event.type)) return false;
    // Крім останньої події беремо непроведені кроки зміни посади і ту
    // ПОСАДУ, яка підтверджує поточний індекс у sh. Інакше ланцюг
    // «РОЗПОРЯДЖ → ПОСАДА → ПЕРЕВ → СКАСУВАННЯ → СЗЧ» губить постановку.
    if (chainedPositionRows.has(event.excelRow)) return true;
    if (currentPositionMovementRows.has(event.excelRow)) return true;
    const key = movementPersonKey(event);
    return Boolean(
      key && latestOperationalRowByPerson.get(key) === event.excelRow,
    );
  });

  const openAbsents = ejoosAbsents.filter((row) => !row.actualReturn);
  const openById = new Map(
    openAbsents.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const openByName = new Map(
    openAbsents.map((row) => [normKey(row.fullName), row]),
  );
  const arrivalById = new Map(
    ejoosArrivals.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const arrivalByName = new Map(
    ejoosArrivals.map((row) => [normKey(row.fullName), row]),
  );
  const oosById = new Map(
    ejoosOos.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const oosByName = new Map(
    ejoosOos.map((row) => [normKey(row.fullName), row]),
  );
  const latestPositionByName = new Map<string, PbMovement>();
  for (const event of effectiveMovements) {
    if (!isPositionChangeRow(event.excelRow) || !event.fullName) {
      continue;
    }
    latestPositionByName.set(normKey(event.fullName), event);
  }
  const positionEventForShPerson = (person: PbShPerson) => {
    const event = latestPositionByName.get(normKey(person.fullName));
    if (!event) return null;
    if (
      event.nextIndex &&
      person.positionIndex &&
      event.nextIndex !== person.positionIndex
    ) {
      return null;
    }
    return event;
  };

  // Ведуться лише зміни місяця 1ПБ (для 25.08.2026 — серпень).
  // Давніші періоди не підтягуємо і не дописуємо як нові статуси.
  const reportDateMs = dateMs(timesheetDayLabel);
  const monthStartMs = (() => {
    const match = String(timesheetDayLabel || "").match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
    if (!match) return 0;
    const year =
      Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
    return Date.UTC(year, Number(match[2]) - 1, 1);
  })();
  const journalMonthStartLabel = (() => {
    const match = String(timesheetDayLabel || "").match(
      /(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/,
    );
    if (!match) return "";
    const year =
      Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
    return `01.${String(match[2]).padStart(2, "0")}.${year}`;
  })();
  const timesheetMonthHeader = formatTimesheetMonthHeader(timesheetDayLabel);
  const timesheetHeaderCell = findTimesheetMonthHeaderCell(
    timesheetSheet?.rawRows ?? [],
  );
  const journalMonth = monthStartMs ? new Date(monthStartMs).getUTCMonth() + 1 : 0;
  const journalYear = monthStartMs ? new Date(monthStartMs).getUTCFullYear() : 0;
  const timesheetMonthHeaderMismatch = Boolean(
    timesheetMonthHeader &&
      timesheetHeaderCell &&
      (timesheetHeaderCell.month !== journalMonth ||
        timesheetHeaderCell.year !== journalYear),
  );
  const leadWindowStart = monthStartMs;
  const leadWindowEnd = reportDateMs || monthStartMs;
  const dateInLeadWindow = (value: string) => {
    const ms = dateMs(value);
    return Boolean(
      ms && leadWindowStart && ms >= leadWindowStart && ms <= leadWindowEnd,
    );
  };
  const eventInLeadWindow = (event: {
    orderDate: string;
    basisDate: string;
  }) => {
    const ms = dateMs(event.orderDate || event.basisDate);
    if (!ms || !leadWindowStart) return true;
    return ms >= leadWindowStart && ms <= leadWindowEnd;
  };
  const inTempArrivals = (personId: string, fullName: string) =>
    Boolean(
      (personId && arrivalById.get(personId)) ||
        byPersonName(arrivalByName, personId, fullName),
    );
  // Постановка цього місяця: відкритий СЗЧ з травня треба внести разом із
  // ШПО/Табелем, а не наступним прогоном після застосування ПОСАДИ.
  const inboundStaffPlacementThisMonth = (personId: string, fullName: string) => {
    const event =
      latestPositionByName.get(normKey(fullName)) ||
      [...latestPositionByName.values()].find((item) =>
        isSamePerson({ personId, fullName }, item),
      ) ||
      null;
    if (event && eventInLeadWindow(event)) return true;
    return (
      inTempArrivals(personId, fullName) &&
      Boolean(staffAppointmentDateFor(personId, fullName, ""))
    );
  };
  const shConfirmsOpenArchive = (period: {
    personId: string;
    fullName: string;
    absenceType: string;
  }) => {
    const sh =
      (period.personId && shPersonById.get(period.personId)) ||
      byPersonName(shPersonByName, period.personId, period.fullName) ||
      null;
    if (!sh) return false;
    return currentStatusConfirmsOpenAbsence(
      mapStatus(sh.status).timesheetCode,
      mapStatus(period.absenceType).timesheetCode,
    );
  };
  const archive = archiveAll.filter((period) => {
    if (!personStillInSh(period.personId, period.fullName)) return false;
    const departMs = dateMs(period.departDate);
    const returnMs = hasActualReturn(period.returnDate)
      ? dateMs(period.returnDate)
      : null;
    if (
      archivePeriodTouchesJournalMonth(
        departMs,
        returnMs,
        leadWindowStart,
        leadWindowEnd,
        { carryOpen: shConfirmsOpenArchive(period) },
      )
    ) {
      return true;
    }
    // Відкритий період до місяця: особа саме зараз заходить на штат
    // (sh уже «В СТРОЮ», тож confirmOpenCarry не спрацює).
    if (
      !returnMs &&
      departMs &&
      departMs < leadWindowStart &&
      inboundStaffPlacementThisMonth(period.personId, period.fullName)
    ) {
      return true;
    }
    return false;
  });
  const augustAbsenceSpansFor = (personId: string, fullName: string) =>
    buildTimesheetAbsenceSpans(
      archiveAll
        .filter((period) => isSamePerson({ personId, fullName }, period))
        .map((period) => ({
          departDate: period.departDate,
          returnDate: period.returnDate,
          absenceType: period.absenceType,
          excelRow: period.excelRow,
          personId: period.personId,
          fullName: period.fullName,
          departMs: dateMs(period.departDate),
          returnMs: hasActualReturn(period.returnDate)
            ? dateMs(period.returnDate)
            : null,
        })),
      {
        timesheetDay,
        monthStartMs: leadWindowStart,
        monthEndMs: leadWindowEnd,
        reportDayMs: reportDateMs,
        mapCode: (absenceType) => mapStatus(absenceType).timesheetCode || "",
        hasReturn: hasActualReturn,
        confirmOpenCarry: (period) =>
          shConfirmsOpenArchive({
            personId: period.personId || personId,
            fullName: period.fullName || fullName,
            absenceType: period.absenceType,
          }),
      },
    );

  const inboundStaffPlacementBefore = (event: PbMovement) =>
    findLatestPriorOwnUnitStaffMove(activeMovementsAll, event, {
      samePerson: isSamePerson,
      inWindow: eventInLeadWindow,
      eventTime: movementEventTime,
    });

  const inboundStaffDateFor = (personId: string, fullName: string) => {
    let inbound: PbMovement | null = null;
    let outboundBefore: PbMovement | null = null;
    for (const event of activeMovementsAll) {
      if (!isSamePerson({ personId, fullName }, event)) continue;
      if (!eventInLeadWindow(event)) continue;
      if (isOwnUnitStaffMove(event)) {
        inbound = event;
      } else if (isOutboundStaffMove(event)) {
        outboundBefore = event;
      }
    }
    if (
      inbound &&
      outboundBefore &&
      movementEventTime(inbound) > movementEventTime(outboundBefore)
    ) {
      return inbound.orderDate || inbound.basisDate;
    }
    return "";
  };

  const staffEpisodePaintPayload = (
    personId: string,
    fullName: string,
    staffIndex: string,
    activeExcelRow: number,
  ) => {
    const appointment =
      staffAppointmentDateFor(personId, fullName, staffIndex) ||
      inboundStaffDateFor(personId, fullName) ||
      "";
    const allSpans = augustAbsenceSpansFor(personId, fullName);
    const carryFromMonthStart = allSpans.some((span) => span.fromDay === 1);
    const activeFromDay = carryFromMonthStart
      ? 1
      : journalDayFromDateMs(dateMs(appointment), leadWindowStart);
    const activeSpans =
      activeFromDay > 1
        ? clipAbsenceSpansToActiveEpisode(allSpans, activeFromDay)
        : allSpans;
    const historySpans =
      !carryFromMonthStart && activeFromDay > 1
        ? absenceSpansBeforeEpisode(allSpans, activeFromDay)
        : [];
    const history = priorEpisodeTimesheetOf(
      personId,
      fullName,
      activeExcelRow,
      staffIndex,
    );
    return {
      timesheetActiveFrom: carryFromMonthStart
        ? journalMonthStartLabel
        : appointment,
      timesheetPreserveHistory:
        !carryFromMonthStart && activeFromDay > 1 ? "1" : "",
      timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(activeSpans),
      historyTimesheetExcelRow:
        history && historySpans.length ? String(history.excelRow) : "",
      historyTimesheetAbsenceSpans: encodeTimesheetAbsenceSpans(historySpans),
    };
  };

  const ops: EjoosSyncOp[] = [];
  const absenceRowsClosedByMovement = new Set<number>();
  const latestSzchCancellationByPerson = new Map<string, PbMovement>();
  for (const event of activeMovementsAll) {
    if (!isSzchCancellation(event)) continue;
    const key = event.personId
      ? `id:${event.personId}`
      : `name:${normKey(event.fullName)}`;
    if (!key.endsWith(":")) latestSzchCancellationByPerson.set(key, event);
  }
  for (const event of latestSzchCancellationByPerson.values()) {
    if (!eventInLeadWindow(event)) continue;
    const open =
      (event.personId && openById.get(event.personId)) ||
      openByName.get(normKey(event.fullName)) ||
      null;
    const person = shPeople.find(
      (candidate) =>
        (event.personId &&
          candidate.personId &&
          event.personId === candidate.personId) ||
        normKey(candidate.fullName) === normKey(event.fullName),
    );
    const timesheetRow =
      activeTimesheetRowOf(
        event.personId,
        event.fullName,
        person?.positionIndex || "",
      ) ||
      (person?.positionIndex && dayByIndex.get(person.positionIndex)) ||
      ejoosDays.find(
        (row) => normKey(row.fullName) === normKey(event.fullName),
      ) ||
      null;
    // Закритий запис і вже виставлений «+» означають, що скасування СЗЧ
    // проведене повністю; повторну операцію не створюємо.
    if (!open && timesheetRow?.dayValue === "+") continue;
    const returnDate = event.orderDate || timesheetDayLabel;
    if (open) absenceRowsClosedByMovement.add(open.excelRow);
    ops.push({
      id: opId([
        "szch_cancel",
        event.personId || event.fullName,
        String(event.excelRow),
      ]),
      kind: "absent_close",
      class: timesheetRow ? "ready" : "needs_input",
      sheet: open
        ? "5. Тимчасово відсутні / 6. Табель"
        : "6. Табель",
      personId: event.personId || open?.personId || "",
      fullName: event.fullName || open?.fullName || "",
      positionIndex: person?.positionIndex || open?.positionIndex || "",
      rank: person?.rank || event.rank || open?.rank || "",
      before: open
        ? `СЗЧ відкрито з ${open.departDate || "?"}`
        : "відкритого періоду СЗЧ у «Тимч. відсутні» немає",
      after: `скасування СЗЧ / повернення ${returnDate}`,
      sourceRef: `Рух!R${event.excelRow} СТАТУС=«${event.status}»`,
      why: open
        ? "Скасування СЗЧ закриває період тимчасової відсутності та відновлює «+» у Табелі"
        : "Відкритого рядка СЗЧ немає: ШПО/ООС не змінюємо, відновлюємо «+» у Табелі",
      confidence: timesheetRow ? "high" : "manual",
      payload: {
        excelRow: open ? String(open.excelRow) : "",
        returnDate,
        timesheetExcelRow: timesheetRow
          ? String(timesheetRow.excelRow)
          : "",
        returnDay: returnDate,
        movementNumber: event.movementNumber,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
      },
      movementKey: createMovementKey(event),
      checkedDefault: Boolean(timesheetRow),
    });
  }

  // SHPO / Табель occupant identity from sh (by position index)
  shPeople.forEach((person) => {
    if (!person.positionIndex || !person.fullName) return;
    const pendingPositionOp = ops.some(
      (op) =>
        op.kind === "position_change" &&
        isSamePerson(person, op) &&
        op.positionIndex === person.positionIndex &&
        op.class === "ready",
    );
    const positionEvent = positionEventForShPerson(person);
    const shpo = shpoByIndex.get(person.positionIndex);
    const ts =
      staffIndexTimesheetForPerson(
        person.personId,
        person.fullName,
        person.positionIndex,
      ) || dayByIndex.get(person.positionIndex);
    if (!shpo && !ts) return;

    if (positionEvent || pendingPositionOp) return;
    const inTempArrivals = Boolean(
      (person.personId && arrivalById.get(person.personId)) ||
        byPersonName(arrivalByName, person.personId, person.fullName),
    );
    if (
      alreadyVacatedForAbsence(person.personId, person.fullName) &&
      !inTempArrivals
    ) {
      return;
    }

    const beforeName = shpo?.fullName || (!shpo ? ts?.fullName : "") || "—";
    const beforeRank = shpo?.rank || (!shpo ? ts?.rank : "") || "—";
    const beforeId = shpo?.personId || (!shpo ? ts?.personId : "") || "—";
    const afterName = person.fullName;
    const afterRank = person.rank || "—";
    const afterId = person.personId || "—";
    const occupantOos =
      (person.personId && oosPersonById.get(person.personId)) ||
      byPersonName(oosPersonByName, person.personId, person.fullName) ||
      null;
    const occupantExcluded = occupantOos
      ? null
      : findLatestExcludedRow(person.personId, person.fullName);

    const nameChanged = normKey(beforeName) !== normKey(afterName);
    const rankChanged = normKey(beforeRank) !== normKey(afterRank);
    const idChanged =
      Boolean(person.personId) &&
      beforeId !== "—" &&
      beforeId !== person.personId;
    const needsRestoreFromExcluded =
      !personStillInEjoos(person.personId, person.fullName) &&
      Boolean(
        occupantExcluded ||
          transferCancelForPerson(person.personId, person.fullName),
      );

    if (
      !nameChanged &&
      !rankChanged &&
      !idChanged &&
      !needsRestoreFromExcluded
    ) {
      return;
    }
    // Сама різниця у званні — не кадрова зміна зайнятості. Її показуємо
    // як DATA_MISMATCH і не підставляємо звання з sh автоматично.
    if (rankChanged && !nameChanged && !idChanged) return;

    // Same index, different person ID → conflict (do not auto-apply)
    if (idChanged && nameChanged) {
      ops.push({
        id: opId(["shpo_conflict", person.positionIndex, person.personId || afterName]),
        kind: "shpo_occupant",
        class: "conflict",
        sheet: "1. ШПО / 6. Табель",
        personId: person.personId,
        fullName: person.fullName,
        positionIndex: person.positionIndex,
        rank: person.rank,
        before: `${beforeRank} ${beforeName} (ID ${beforeId})`,
        after: `${afterRank} ${afterName} (ID ${afterId})`,
        sourceRef: `sh!R${person.excelRow} індекс ${person.positionIndex}`,
        why: "На тому ж індексі посади інший ID/ПІБ — потрібна ручна перевірка перед зміною зайнятості",
        confidence: "review",
        payload: {
          shpoExcelRow: shpo ? String(shpo.excelRow) : "",
          timesheetExcelRow: ts ? String(ts.excelRow) : "",
          nextName: afterName,
          nextRank: afterRank,
          nextPersonId: afterId,
          excludedSourceExcelRow: occupantExcluded
            ? String(occupantExcluded.excelRow)
            : "",
        },
        checkedDefault: false,
      });
      return;
    }

    ops.push({
      id: opId(["shpo", person.positionIndex, person.personId || afterName]),
      kind: "shpo_occupant",
      class: "ready",
      sheet: "1. ШПО / 6. Табель",
      personId: person.personId,
      fullName: person.fullName,
      positionIndex: person.positionIndex,
      rank: person.rank,
      before: `${beforeRank} ${beforeName} (ID ${beforeId})`,
      after: `${afterRank} ${afterName} (ID ${afterId})`,
      sourceRef: `sh!R${person.excelRow} індекс ${person.positionIndex}`,
        why: needsRestoreFromExcluded
          ? `Особа є в sh на ${person.positionIndex}, але в активних ШПО/ООС її немає. Відновити зайнятість${occupantExcluded ? ` і картку ООС з «Виключені» R${occupantExcluded.excelRow}` : ""}`
          : "Зайнятість посади в ЕЖООС відрізняється від sh — оновити ПІБ/звання/ID",
      confidence: "high",
      payload: {
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
        nextName: afterName,
        nextRank: afterRank,
        nextPersonId: person.personId,
        excludedSourceExcelRow: occupantExcluded
          ? String(occupantExcluded.excelRow)
          : "",
        timesheetActiveFrom: needsRestoreFromExcluded
          ? transferCancelForPerson(person.personId, person.fullName)
              ?.orderDate || ""
          : "",
        timesheetPreserveHistory: needsRestoreFromExcluded ? "1" : "",
        timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
          augustAbsenceSpansFor(person.personId, person.fullName),
        ),
      },
      checkedDefault: true,
    });
  });

  shPeople.forEach((person) => {
    if (!person.status || person.status === "0") return;
    const hasPositionEvent = Boolean(positionEventForShPerson(person));
    const inTempArrivals = Boolean(
      (person.personId && arrivalById.get(person.personId)) ||
        byPersonName(arrivalByName, person.personId, person.fullName),
    );
    if (
      alreadyVacatedForAbsence(person.personId, person.fullName) &&
      !inTempArrivals
    ) {
      return;
    }
    const mapped = mapStatus(person.status);

    if (mapped.timesheetCode === "+") {
      const open =
        (person.personId && openById.get(person.personId)) ||
        byPersonName(openByName, person.personId, person.fullName) ||
        openByName.get(normKey(person.fullName)) ||
        null;
      const archiveReturn = archive.find(
        (period) =>
          isSamePerson(person, period) &&
          hasActualReturn(period.returnDate) &&
          (canonicalName(period.absenceType) === canonicalName(open?.ground || "") ||
            !open?.ground),
      );
      if (
        open &&
        !absenceRowsClosedByMovement.has(open.excelRow) &&
        (dateInLeadWindow(open.departDate) ||
          (archiveReturn && dateInLeadWindow(archiveReturn.returnDate)))
      ) {
        const returnDate = archiveReturn?.returnDate || timesheetDayLabel;
        const returnOrder = [
          archiveReturn?.returnOrderDate,
          archiveReturn?.returnOrderNumber,
        ]
          .filter(Boolean)
          .join(" ");
        ops.push({
          id: opId(["absent_close", open.personId || open.fullName, String(open.excelRow)]),
          kind: "absent_close",
          class: "ready",
          sheet: "5. Тимчасово відсутні",
          personId: person.personId || open.personId,
          fullName: person.fullName || open.fullName,
          positionIndex: person.positionIndex || open.positionIndex,
          rank: person.rank,
          before: `відкрито: ${open.ground || "—"} з ${open.departDate || "?"} → ${open.place || "?"}`,
          after: returnOrder
            ? `фактичне прибуття: ${returnDate} · ${returnOrder}`
            : `фактичне прибуття: ${returnDate}`,
          sourceRef: archiveReturn
            ? `archive!R${archiveReturn.excelRow} → повернення ${returnDate}`
            : `sh!R${person.excelRow} → В СТРОЮ; ЕЖООС sheet5 R${open.excelRow}`,
          why: archiveReturn
            ? "У archive є фактичне повернення — закриваємо «Тимч. відсутні» цією датою і наказом, не днем зрізу sh"
            : "У 1ПБ знову «в строю», а в «Тимч. відсутні» період ще відкритий",
          confidence: "high",
          payload: {
            excelRow: String(open.excelRow),
            returnDate,
            returnOrderNumber: archiveReturn?.returnOrderNumber || "",
            returnOrderDate: archiveReturn?.returnOrderDate || "",
            timesheetExcelRow: String(
              activeTimesheetRowOf(
                person.personId,
                person.fullName,
                person.positionIndex,
              )?.excelRow || "",
            ),
            returnDay: returnDate,
            timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
              augustAbsenceSpansFor(person.personId, person.fullName),
            ),
            timesheetActiveFrom: "",
            timesheetSkipHistory: "1",
          },
          checkedDefault: true,
        });
        absenceRowsClosedByMovement.add(open.excelRow);
      }
    }

    // Постановку на штат фарбує position_change; окремий timesheet_day тут
    // лише дублює і може зіпсувати коди СЗЧ на початку місяця.
    if (hasPositionEvent) return;
    const personTimesheet =
      (person.personId && dayById.get(person.personId)) ||
      ejoosDays.find((row) =>
        isSamePerson(person, { personId: row.personId, fullName: row.fullName }),
      ) ||
      null;
    const indexTimesheet = person.positionIndex
      ? dayByIndex.get(person.positionIndex) ?? null
      : null;
    const timesheetRow = personTimesheet || indexTimesheet;
    const before = personTimesheet?.dayValue || "—";
    const afterCode = mapped.timesheetCode;
    const onMatchingShpo = Boolean(
      person.positionIndex &&
        shpoByIndex.get(person.positionIndex) &&
        isSamePerson(person, shpoByIndex.get(person.positionIndex)!),
    );
    const indexTakenByOther = Boolean(
      indexTimesheet &&
        (indexTimesheet.personId || indexTimesheet.fullName) &&
        !isSamePerson(person, indexTimesheet),
    );
    const missingFromTimesheet =
      !personTimesheet && onMatchingShpo && mapped.timesheetCode === "+";
    if (missingFromTimesheet) {
      ops.push({
        id: opId(["ts_restore", person.personId || person.positionIndex]),
        kind: "timesheet_day",
        class: "needs_input",
        sheet: "6. Табель",
        personId: person.personId,
        fullName: person.fullName,
        positionIndex: person.positionIndex,
        rank: person.rank,
        before: "немає в Табелі",
        after: `відновити на ${person.positionIndex}${
          indexTakenByOther ? " — рядок зайнятий іншою особою" : ""
        }`,
        sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}» · ШПО інд. ${person.positionIndex}`,
        why: indexTakenByOther
          ? "У ШПО особа є і в строю, але рядок Табеля на цьому індексі зайнятий іншою людиною — не перезаписуємо автоматично"
          : "У ШПО особа вже стоїть на актуальній посаді і в строю, а в Табелі її немає. Кадровий рух не проводимо — лише відновити рядок Табеля після перевірки",
        confidence: "review",
        payload: {
          timesheetCode: afterCode || "+",
          day: String(timesheetDay),
          excelRow:
            !indexTakenByOther && indexTimesheet
              ? String(indexTimesheet.excelRow)
              : "",
          statusRaw: person.status,
          restorePerson: "1",
          nextName: person.fullName,
          nextPersonId: person.personId,
          nextRank: shpoByIndex.get(person.positionIndex)?.rank || "",
        },
        checkedDefault: false,
      });
    }

    if (!missingFromTimesheet && afterCode) {
      const after = afterCode;
      if (before !== after) {
        const isReady = mapped.confidence === "high" && Boolean(timesheetRow);
        ops.push({
          id: opId(["ts", person.personId || person.positionIndex, String(timesheetDay), after]),
          kind: "timesheet_day",
          class: !timesheetRow
            ? "needs_input"
            : mapped.confidence === "manual"
              ? "needs_input"
              : mapped.confidence === "review"
                ? "conflict"
                : "ready",
          sheet: "6. Табель",
          personId: person.personId,
          fullName: person.fullName,
          positionIndex: person.positionIndex,
          rank: person.rank,
          before,
          after,
          sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
          why: mapped.reason,
          confidence: mapped.confidence,
          payload: {
            timesheetCode: after,
            day: String(timesheetDay),
            excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
            statusRaw: person.status,
          },
          checkedDefault: isReady,
        });
      }
    } else if (
      !missingFromTimesheet &&
      mapped.confidence !== "high" &&
      mapped.ruleId !== "absent_archive"
    ) {
      const alreadyPresent = before === "+";
      const defaultsToPresent =
        Boolean(timesheetRow) && (before === "—" || alreadyPresent);
      if (!alreadyPresent) {
        ops.push({
          id: opId(["ts_manual", person.personId || person.fullName, person.status]),
          kind: "timesheet_day",
          class: defaultsToPresent ? "ready" : "needs_input",
          sheet: "6. Табель",
          personId: person.personId,
          fullName: person.fullName,
          positionIndex: person.positionIndex,
          rank: person.rank,
          before,
          after: defaultsToPresent ? "+" : "(оберіть код)",
          sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
          why: defaultsToPresent
            ? "У Табелі за поточний день порожньо — за замовчуванням ставимо «+»"
            : mapped.reason,
          confidence: defaultsToPresent ? "high" : "manual",
          payload: {
            day: String(timesheetDay),
            excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
            statusRaw: person.status,
            timesheetCode: defaultsToPresent ? "+" : "",
          },
          checkedDefault: defaultsToPresent,
        });
      }
    }
  });

  archive.forEach((period) => {
    const open =
      (period.personId && openById.get(period.personId)) ||
      byPersonName(openByName, period.personId, period.fullName) ||
      openByName.get(normKey(period.fullName)) ||
      null;
    const sameAbsentPeriod = (row: EjoosAbsentRow) => {
      if (!isSamePerson(period, row)) return false;
      const sameKind =
        canonicalName(row.ground) === canonicalName(period.absenceType) ||
        (isDispositionAbsenceStatus(row.ground) &&
          isDispositionAbsenceStatus(period.absenceType));
      const sameDepart =
        Boolean(dateMs(row.departDate)) &&
        dateMs(row.departDate) === dateMs(period.departDate);
      return sameKind && (sameDepart || !row.departDate);
    };
    const recorded = ejoosAbsents.find(sameAbsentPeriod) ?? null;
    const reuseAbsent =
      recorded || (open && sameAbsentPeriod(open) ? open : null);
    if (
      recorded?.actualReturn &&
      dateMs(recorded.actualReturn) === dateMs(period.returnDate)
    ) {
      return;
    }
    const complete = Boolean(period.absenceType && (period.departDate || period.plannedReturn));
    const before = reuseAbsent
      ? `${reuseAbsent.ground} / ${reuseAbsent.place} / ${reuseAbsent.departDate}`
      : "(немає запису в «Тимч. відсутні»)";
    const returnOrderText = [
      period.returnOrderDate,
      period.returnOrderNumber
        ? `№${period.returnOrderNumber.replace(/^№/i, "")}`
        : "",
    ]
      .filter(Boolean)
      .join(" ");
    const after = [
      `${period.absenceType || "?"} → ${period.place || "?"} з ${period.departDate || "?"}`,
      period.returnDate
        ? `повернення ${period.returnDate}${returnOrderText ? ` · ${returnOrderText}` : ""}`
        : "",
    ]
      .filter(Boolean)
      .join(" · ");
    const needsClose =
      hasActualReturn(period.returnDate) &&
      Boolean(reuseAbsent) &&
      dateMs(reuseAbsent?.actualReturn || "") !== dateMs(period.returnDate);
    // Відкритий період тієї ж особи (по ID) уже є — різниця в по батькові
    // або в форматі рядка не є новою кадровою подією.
    const sameAbsenceKind =
      open &&
      (canonicalName(open.ground) === canonicalName(period.absenceType) ||
        (isDispositionAbsenceStatus(open.ground) &&
          isDispositionAbsenceStatus(period.absenceType)));
    if (sameAbsenceKind && !hasActualReturn(period.returnDate)) return;
    if (
      reuseAbsent &&
      absenceRowsClosedByMovement.has(reuseAbsent.excelRow)
    ) {
      return;
    }
    if (
      alreadyVacatedForAbsence(period.personId, period.fullName) &&
      open &&
      !inboundStaffPlacementThisMonth(period.personId, period.fullName)
    ) {
      return;
    }
    if (open && before === after) return;
    if (recorded && !hasActualReturn(period.returnDate)) return;
    const shPerson =
      (period.personId && shPersonById.get(period.personId)) ||
      byPersonName(shPersonByName, period.personId, period.fullName) ||
      null;
    const archiveReturnVsSh = archiveReturnContradictsCurrentSh(
      shPerson ? mapStatus(shPerson.status).timesheetCode : "",
      mapStatus(period.absenceType).timesheetCode,
      hasActualReturn(period.returnDate),
    );
    if (archiveReturnVsSh) {
      ops.push({
        id: opId([
          "archive_sh_return",
          period.personId || period.fullName,
          String(period.excelRow),
        ]),
        kind: "absent_upsert",
        class: "needs_input",
        sheet: "5. Тимч. відсутні / archive vs sh",
        personId: shPerson?.personId || period.personId,
        fullName: shPerson?.fullName || period.fullName,
        positionIndex: shPerson?.positionIndex || "",
        rank: period.rank,
        before,
        after: `archive: повернення ${period.returnDate}; sh досі ${shPerson?.status || period.absenceType}`,
        sourceRef: `archive!R${period.excelRow} · sh`,
        why: `NEEDS_REVIEW: archive вже має повернення ${period.returnDate}, а поточний sh досі ${shPerson?.status || "відсутній"}. Не закриваємо період і не тягнемо СЗЧ/ЗБ до дня звіту — перевірте, що саме застаріло.`,
        confidence: "manual",
        payload: {
          mismatchKind: "ARCHIVE_RETURN_SH_STILL_ABSENT",
          absenceType: period.absenceType,
          place: period.place,
          departDate: period.departDate,
          returnDate: period.returnDate,
          statusRaw: shPerson?.status || "",
          existingExcelRow: reuseAbsent ? String(reuseAbsent.excelRow) : "",
        },
        checkedDefault: false,
      });
      return;
    }
    const activeTs = activeTimesheetRowOf(
      period.personId,
      period.fullName,
      shPerson?.positionIndex || "",
    );
    const episodePaint = staffEpisodePaintPayload(
      period.personId,
      period.fullName,
      shPerson?.positionIndex || "",
      activeTs?.excelRow || 0,
    );
    const returnOrderNo = period.returnOrderNumber
      ? `№${period.returnOrderNumber.replace(/^№/i, "")}`
      : "";

    ops.push({
      id: opId(["absent_up", period.personId || period.fullName, period.periodNumber || String(period.excelRow)]),
      kind: "absent_upsert",
      class: complete ? "ready" : "needs_input",
      sheet: "5. Тимчасово відсутні",
      personId: shPerson?.personId || period.personId,
      fullName: shPerson?.fullName || period.fullName,
      positionIndex: shPerson?.positionIndex || "",
      rank: period.rank,
      before,
      after,
      sourceRef: `archive!R${period.excelRow} №${period.periodNumber || "—"}`,
      why: complete
        ? period.returnDate
          ? `Закрити період ${period.absenceType || "відсутності"} фактичним поверненням ${period.returnDate}${returnOrderNo ? ` ${returnOrderNo}` : ""}. Коди відсутності — у «Тимч. відсутні»${episodePaint.historyTimesheetExcelRow ? " і на історичному рядку Табеля" : ""}, не на новому штатному епізоді.`
          : "Відкритий період цього місяця з archive — внести у «Тимч. відсутні»"
        : "В archive неповні поля (дата/підстава) — дозаповніть перед застосуванням",
      confidence: complete ? "high" : "manual",
      payload: {
        absenceType: period.absenceType,
        place: period.place,
        departDate: period.departDate,
        orderNumber: period.orderNumber,
        orderDate: period.orderDate,
        plannedReturn: period.plannedReturn || "?",
        returnDate: period.returnDate,
        returnOrderNumber: period.returnOrderNumber,
        returnOrderDate: period.returnOrderDate,
        periodNumber: period.periodNumber,
        positionTitle: period.positionTitle,
        timesheetExcelRow: activeTs ? String(activeTs.excelRow) : "",
        timesheetSkipHistory: "1",
        ...episodePaint,
        historyDepartDate: "",
        historyOrderNumber: "",
        historyDepartDest: "",
        timesheetCode:
          dateMs(period.departDate) &&
          dateMs(period.returnDate) &&
          dateMs(period.returnDate) <= dateMs(period.departDate)
            ? ""
            : mapStatus(period.absenceType).timesheetCode || "",
        existingExcelRow: reuseAbsent ? String(reuseAbsent.excelRow) : "",
      },
      checkedDefault: complete && (!reuseAbsent || needsClose),
    });
  });

  const timesheetAlreadyPainted = (personId: string, fullName: string) =>
    ops.some(
      (op) =>
        isSamePerson({ personId, fullName }, op) &&
        Boolean(
          op.payload.timesheetAbsenceSpans || op.payload.timesheetActiveFrom,
        ),
    );
  for (const person of shPeople) {
    if (!personStillInSh(person.personId, person.fullName)) continue;
    if (positionEventForShPerson(person)) continue;
    if (timesheetAlreadyPainted(person.personId, person.fullName)) continue;
    if (
      ops.some(
        (op) =>
          op.payload.mismatchKind === "ARCHIVE_RETURN_SH_STILL_ABSENT" &&
          isSamePerson(person, op),
      )
    ) {
      continue;
    }
    const spans = augustAbsenceSpansFor(person.personId, person.fullName);
    const appointmentDate = staffAppointmentDateFor(
      person.personId,
      person.fullName,
      person.positionIndex,
    );
    const inboundDate = inboundStaffDateFor(person.personId, person.fullName);
    const active = activeTimesheetRowOf(
      person.personId,
      person.fullName,
      person.positionIndex,
    );
    if (!active) continue;
    const episodePaint = staffEpisodePaintPayload(
      person.personId,
      person.fullName,
      person.positionIndex,
      active.excelRow,
    );
    const activeFromLabel =
      episodePaint.timesheetActiveFrom || appointmentDate || inboundDate || "";
    let activeFromDay =
      journalDayFromDateMs(dateMs(activeFromLabel), leadWindowStart) || 1;
    const scan = timesheetScanByRow.get(active.excelRow);
    if (
      scan?.hasDepartureText &&
      !scan.plusDays.some((day) => day >= Math.max(1, activeFromDay))
    ) {
      continue;
    }
    // Якщо дату постановки з РУХ не взяли, а рядок уже «−» потім «+», і СЗЧ
    // закінчився до першого «+» — це той самий штатний епізод, не фарбуємо СЗЧ.
    if (activeFromDay <= 1 && scan?.plusDays.length) {
      const firstPlus = Math.min(...scan.plusDays);
      const absenceEndedBeforePlus =
        firstPlus > 1 && spans.every((span) => span.toDay < firstPlus);
      const prefixInactive =
        firstPlus > 1 &&
        Array.from({ length: firstPlus - 1 }, (_, index) => index + 1).every(
          (day) => {
            const actual = (scan.dayCodes[day] || "").trim();
            return (
              !actual ||
              actual === "вибув" ||
              sameTimesheetDayMark(actual, "-")
            );
          },
        );
      if (absenceEndedBeforePlus && prefixInactive) {
        activeFromDay = firstPlus;
      }
    }
    const episodeSpans =
      activeFromDay > 1
        ? clipAbsenceSpansToActiveEpisode(spans, activeFromDay)
        : spans;
    if (!episodeSpans.length && activeFromDay <= 1 && !spans.length) continue;
    let mismatch = false;
    for (let day = 1; day <= timesheetDay; day += 1) {
      const expected = timesheetMarkFromArchive(day, {
        activeFromDay,
        lastDay: timesheetDay,
        spans: episodeSpans,
        fillBeforeActive: activeFromDay > 1,
      });
      if (!expected) continue;
      const actual = (scan?.dayCodes[day] || "").trim();
      if (actual === "вибув") continue;
      if (sameTimesheetDayMark(actual, expected)) continue;
      mismatch = true;
      break;
    }
    if (!mismatch) continue;
    ops.push({
      id: opId(["ts_archive", person.personId || person.fullName]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: person.personId,
      fullName: person.fullName,
      positionIndex: person.positionIndex,
      rank: person.rank,
      before: "позначки днів не збігаються з archive",
      after:
        activeFromDay > 1
          ? `штатний рядок з ${activeFromLabel}: до постановки «-», далі «+»`
          : episodeSpans.length
            ? `фактичні коди з archive (${episodeSpans.map((span) => `${span.fromDay}–${span.toDay}:${span.code}`).join(", ")})`
            : `активний рядок з ${activeFromLabel}`,
      sourceRef: `archive + sh · Табель R${active.excelRow}`,
      why:
        activeFromDay > 1
          ? "Новий штатний епізод: коди відсутності на цей рядок не переносимо. СЗЧ лишається в «Тимч. відсутні» та на історичному рядку Табеля, якщо він є."
          : "Табель ведемо за фактичною хронологією archive, не з РУХ. Історичний рядок з вибуттям не перераховуємо.",
      confidence: "high",
      payload: {
        type: "PAINT_ARCHIVE",
        excelRow: String(active.excelRow),
        ...episodePaint,
        timesheetActiveFrom: episodePaint.timesheetActiveFrom || activeFromLabel,
        timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(episodeSpans),
        timesheetPreserveHistory: activeFromDay > 1 ? "1" : episodePaint.timesheetPreserveHistory,
        historyDepartDate: "",
        historyOrderNumber: "",
        historyDepartDest: "",
      },
      checkedDefault: true,
    });
  }

  const seenTransferCancelReview = new Set<string>();
  for (const event of movementsAll) {
    if (!isTransferCancellation(event) || !eventInLeadWindow(event)) continue;
    const key = movementPersonKey(event);
    if (!key || seenTransferCancelReview.has(key)) continue;
    seenTransferCancelReview.add(key);
    const cancelled = cancelledTransferOf(event);
    const excludedRow = cancelled
      ? findMovementExcludedRow(cancelled)
      : findMovementExcludedRow(event);
    const timesheetHasDepart = timesheetRowsOf(
      event.personId,
      event.fullName,
    ).some((row) => row.hasDepartureText);
    const evidence = Boolean(excludedRow || timesheetHasDepart);
    const shPerson =
      (event.personId && shPersonById.get(event.personId)) ||
      byPersonName(shPersonByName, event.personId, event.fullName) ||
      null;
    const inCurrentSh = personStillInSh(event.personId, event.fullName);
    const missingFromCurrentSh = !inCurrentSh;
    const dest = cancelled
      ? formatTransferDestinationForTimesheet(
          [cancelled.destination, cancelled.changeText].filter(Boolean).join(" "),
        ) || cancelled.destination
      : "";
    ops.push({
      id: opId([
        "transfer-cancel",
        event.personId || event.fullName,
        event.orderNumber || String(event.excelRow),
      ]),
      kind: "other_manual",
      class: missingFromCurrentSh ? "needs_input" : "ready",
      sheet: missingFromCurrentSh
        ? "Дані джерел / 3. Виключені"
        : "3. Виключені",
      personId: shPerson?.personId || event.personId,
      fullName: shPerson?.fullName || event.fullName,
      positionIndex:
        shPerson?.positionIndex || event.nextIndex || event.previousIndex,
      rank: shPerson?.rank || event.rank,
      before: missingFromCurrentSh
        ? `РУХ: ПЕРЕВ №${cancelled?.orderNumber || "?"} → ${dest || "?"} скасовано №${event.orderNumber || "?"}; sh: немає; ЕЖООС: Виключені${excludedRow ? ` R${excludedRow.excelRow}` : ""} / Табель закритий`
        : excludedRow
          ? `Виключені R${excludedRow.excelRow}: №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"}`
          : "рядка виключення за скасованим ПЕРЕВ немає",
      after: missingFromCurrentSh
        ? "NEEDS_REVIEW — скасування є, але в актуальній sh людини немає; ШПО / ООС / Табель автоматично не відновлюємо"
        : evidence
          ? `REMOVE_CANCELLED_EXCLUSION — прибрати запис №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"} (скасовано №${event.orderNumber || "?"})`
          : "REMOVE_CANCELLED_EXCLUSION — нового рядка у Виключених не створюємо",
      sourceRef: `Рух!R${event.excelRow} №${event.orderNumber || "?"} · скасування`,
      why: missingFromCurrentSh
        ? `Скасування переведення №${event.orderNumber || "?"} від ${event.orderDate || "?"} (ПЕРЕВ №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"} → ${dest || "?"}). Рух очікує залишення в 1 ПБ, але в актуальній sh особи немає${/нема в sh/i.test(event.status) ? " (у рядку скасування теж «Нема в sh»)" : ""}. Це не AUTO_RESTORE і не NO_ACTION: поки sh не підтвердить перебування, ШПО/ООС/новий Табель не відновлюємо.`
        : evidence
          ? `Переведення №${cancelled?.orderNumber || "?"} від ${cancelled?.orderDate || "?"} скасовано №${event.orderNumber || "?"} від ${event.orderDate || "?"}. Фактичного виключення не було — запис у «Виключені» прибираємо, Табель лишається одним рядком.`
          : `Скасування №${event.orderNumber || "?"} від ${event.orderDate || "?"}: рядка у Виключених і «вибув» у Табелі немає. Новий рядок не створюємо, особу виключеною не вважаємо.`,
      confidence: missingFromCurrentSh ? "manual" : "high",
      payload: {
        type: "TRANSFER_CANCELLED",
        reviewReason: missingFromCurrentSh
          ? "CANCEL_TRANSFER_BUT_NOT_IN_CURRENT_SH"
          : "",
        excludedExcelRow: excludedRow ? String(excludedRow.excelRow) : "",
        cancelledTransferOrder: cancelled?.orderNumber || "",
        cancelledTransferDate: cancelled?.orderDate || "",
        cancelledDestination: dest,
        transferCancelOrder: event.orderNumber,
        transferCancelDate: event.orderDate,
        actualExclusionEvidence: evidence ? "1" : "",
        inCurrentSh: inCurrentSh ? "1" : "",
        previousIndex: cancelled?.previousIndex || event.nextIndex || "",
      },
      movementKey: createMovementKey(event),
      checkedDefault: !missingFromCurrentSh,
    });
  }

  /**
   * `norm` віддає числа з діапазону Excel-дат як дату, тому ID (напр. 22814)
   * у тексті рядка губиться. ID збираємо окремо через `normId`.
   */
  type SheetScanRow = { excelRow: number; text: string; ids: Set<string> };
  const scanRows = (
    sheet: ExcelSheetSnapshot | undefined,
    pattern: RegExp,
  ): SheetScanRow[] =>
    (sheet?.rawRows ?? [])
      .map((row, index) => ({
        excelRow: index + 1,
        text: canonicalName(row.map(norm).join(" ")),
        ids: new Set(row.map((value) => normId(value)).filter(Boolean)),
      }))
      .filter((row) => pattern.test(row.text));
  const DISPOSITION_RE = /РОЗПОРЯДЖ/iu;
  const ABSENCE_STATUS_RE =
    /СЗЧ|САМОВІЛ|БЕЗВІСТ|(?:^|[^А-ЯІЇЄҐа-яіїєґ])ЗБ(?:$|[^А-ЯІЇЄҐа-яіїєґ])/iu;
  const shpoDispositionRows = scanRows(shpoSheet, DISPOSITION_RE);
  const timesheetDispositionRows = scanRows(timesheetSheet, DISPOSITION_RE);
  const shpoSzchRows = scanRows(shpoSheet, ABSENCE_STATUS_RE);
  const timesheetSzchRows = scanRows(timesheetSheet, ABSENCE_STATUS_RE);
  const personInTextRows = (event: PbMovement, rows: SheetScanRow[]) => {
    const id = normId(event.personId);
    if (id && rows.some((row) => row.ids.has(id))) return true;
    // Рядки ЕЖООС часто без ID, а написання ПІБ у джерелах різне.
    const names = [...personNameKeys(event.personId, event.fullName)];
    return names.some((name) =>
      rows.some((row) => row.text.includes(name)),
    );
  };
  const findNonStaffOccupantExcelRow = (
    sheet: ExcelSheetSnapshot | undefined,
    personId: string,
    fullName: string,
    cols: { index: number; name: number; id: number },
  ) => {
    if (!sheet) return 0;
    const names = [...personNameKeys(personId, fullName)];
    const id = normId(personId);
    for (let i = 6; i < sheet.rawRows.length; i += 1) {
      const row = sheet.rawRows[i];
      const index = norm(row?.[cols.index]);
      if (index && /^\d/.test(index)) continue;
      const rowId = normId(row?.[cols.id]);
      const rowName = canonicalName(norm(row?.[cols.name]));
      if (id && rowId && id === rowId) return i + 1;
      if (
        rowName &&
        names.some((name) => name === rowName || rowName.includes(name))
      ) {
        return i + 1;
      }
    }
    return 0;
  };
  const openAbsenceOf = (personId: string, fullName: string) =>
    (personId && openById.get(personId)) ||
    byPersonName(openByName, personId, fullName) ||
    openByName.get(normKey(fullName)) ||
    null;
  const arrivalOf = (personId: string, fullName: string) =>
    ejoosArrivals.find((row) => isSamePerson({ personId, fullName }, row)) ||
    (personId && arrivalById.get(personId)) ||
    byPersonName(arrivalByName, personId, fullName) ||
    null;
  const ownUnitIndexHistoryOf = (personId: string, fullName: string) => {
    const seen = new Set<string>();
    const entries: Array<{ index: string; date: string; order: string }> = [];
    const events = activeMovementsAll
      .filter(
        (item) =>
          isSamePerson({ personId, fullName }, item) &&
          eventInLeadWindow(item) &&
          isOwnUnitStaffMove(item) &&
          isPositionIndex(item.nextIndex),
      )
      .sort(
        (left, right) =>
          movementEventTime(left) - movementEventTime(right) ||
          left.excelRow - right.excelRow,
      );
    for (const item of events) {
      if (seen.has(item.nextIndex)) continue;
      seen.add(item.nextIndex);
      entries.push({
        index: item.nextIndex,
        date: item.orderDate || item.basisDate || "",
        order: item.orderNumber || "",
      });
    }
    return entries;
  };
  const shpoByPositionIndex = new Map(
    ejoosShpo.map((row) => [row.positionIndex, row]),
  );
  /**
   * У ШПО трапляються рядки без індексу — лише з ПІБ (залишок після
   * попередніх правок). Вони не є ні штатною посадою, ні блоком
   * розпорядження, але їх треба чистити разом із виведенням особи.
   */
  const shpoStrayRowByName = new Map<string, EjoosShpoRow>();
  {
    const dispositionExcelRows = new Set(
      shpoDispositionRows.map((row) => row.excelRow),
    );
    (shpoSheet?.rawRows ?? []).forEach((row, index) => {
      const excelRow = index + 1;
      if (excelRow < 7 || dispositionExcelRows.has(excelRow)) return;
      if (norm(row?.[0])) return;
      const fullName = norm(row?.[6]);
      if (!fullName) return;
      const key = canonicalName(fullName);
      if (!key || shpoStrayRowByName.has(key)) return;
      shpoStrayRowByName.set(key, {
        excelRow,
        positionIndex: "",
        rank: norm(row?.[5]),
        fullName,
        personId: normId(row?.[7]),
      });
    });
  }
  const activeSzchPersonIds = new Set<string>();
  const activeSzchPersonNames = new Set<string>();
  for (const period of archiveAll) {
    if (
      !isDispositionAbsenceStatus(period.absenceType) ||
      hasActualReturn(period.returnDate)
    ) {
      continue;
    }
    if (period.personId) activeSzchPersonIds.add(period.personId);
    if (period.fullName) {
      activeSzchPersonNames.add(canonicalName(period.fullName));
    }
  }
  const reportDate = dateMs(timesheetDayLabel);
  const staleMovementCutoff = reportDate
    ? reportDate - 90 * 86400000
    : 0;

  const pendingRankByPerson = new Map<string, EjoosSyncOp>();
  const latestRankEventByPerson = new Map<string, PbMovement>();
  for (const event of activeMovementsAll) {
    if (!isRankAssignmentEvent(event) || !eventInLeadWindow(event)) continue;
    const key = movementPersonKey(event);
    if (!key) continue;
    const previous = latestRankEventByPerson.get(key);
    if (
      !previous ||
      movementEventTime(event) > movementEventTime(previous) ||
      (movementEventTime(event) === movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      latestRankEventByPerson.set(key, event);
    }
  }
  for (const event of latestRankEventByPerson.values()) {
    if (!personStillInEjoos(event.personId, event.fullName)) continue;
    const { previousRank, nextRank } = parseRankPromotion(event);
    if (!nextRank) continue;
    const shpo =
      (event.personId && shpoPersonById.get(event.personId)) ||
      byPersonName(shpoPersonByName, event.personId, event.fullName) ||
      null;
    const oos =
      (event.personId && oosPersonById.get(event.personId)) ||
      byPersonName(oosPersonByName, event.personId, event.fullName) ||
      null;
    const ts =
      (event.personId && dayById.get(event.personId)) ||
      (shpo?.positionIndex && dayByIndex.get(shpo.positionIndex)) ||
      ejoosDays.find((row) => isSamePerson(event, row)) ||
      null;
    const currentRank = oos?.rank || shpo?.rank || ts?.rank || "";
    if (currentRank && normKey(currentRank) === normKey(nextRank)) continue;
    const canApply = Boolean(
      event.orderNumber && event.orderDate && (shpo || oos || ts),
    );
    const rankOp: EjoosSyncOp = {
      id: opId(["rank", event.personId || event.fullName, event.orderNumber || String(event.excelRow)]),
      kind: "rank_change",
      class: canApply ? "ready" : "needs_input",
      sheet: "1. ШПО / 2. ООС / 6. Табель",
      personId: shpo?.personId || oos?.personId || event.personId,
      fullName: shpo?.fullName || oos?.fullName || event.fullName,
      positionIndex:
        shpo?.positionIndex ||
        oos?.positionIndex ||
        ts?.positionIndex ||
        event.previousIndex ||
        event.nextIndex,
      rank: nextRank,
      before: currentRank || previousRank || "—",
      after: `${nextRank} · наказ №${event.orderNumber || "?"} від ${event.orderDate || "?"}`,
      sourceRef: `Рух!R${event.excelRow} №${event.movementNumber} · ЗВАННЯ`,
      why: canApply
        ? "Серпневе присвоєння звання треба провести до виключення, інакше в «Виключені» піде старе звання"
        : "У РУХ не вистачає номера/дати наказу про присвоєння звання",
      confidence: canApply ? "high" : "manual",
      payload: {
        previousRank: currentRank || previousRank,
        nextRank,
        orderNumber: event.orderNumber,
        orderDate: event.orderDate,
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        oosExcelRow: oos ? String(oos.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
      },
      movementKey: createMovementKey(event),
      checkedDefault: canApply,
    };
    ops.push(rankOp);
    const key = movementPersonKey(event);
    if (key) pendingRankByPerson.set(key, rankOp);
    if (event.personId) pendingRankByPerson.set(`id:${event.personId}`, rankOp);
  }

  const considerMovement = (event: PbMovement) => {
    const movementDate = dateMs(event.orderDate || event.basisDate);
    if (leadWindowStart && movementDate && movementDate < leadWindowStart) {
      return;
    }
    if (isAmbiguousStaffTransfer(event)) {
      ops.push({
        id: opId(["scope", event.personId || event.fullName, String(event.excelRow)]),
        kind: "other_manual",
        class: "needs_input",
        sheet: "Рух / Виключені або ШПО",
        personId: event.personId,
        fullName: event.fullName,
        positionIndex: event.nextIndex || event.previousIndex,
        rank: event.rank,
        before: event.destination || event.changeText || "ПЕРЕВ",
        after: "уточнити: внутрішня зміна посади чи вибуття до іншої в/ч",
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: "Не визначено, чи це переведення всередині 1ПБ, чи вибуття до іншої військової частини. Без «куди вибув» / в/ч А#### масово не застосовуємо.",
        confidence: "manual",
        payload: {
          type: "TRANSFER_SCOPE_UNCLEAR",
          transferScope: "unclear",
          destination: event.destination,
          note: event.note,
          changeText: event.changeText,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
          previousIndex: event.previousIndex,
          nextIndex: event.nextIndex,
        },
        movementKey: createMovementKey(event),
        checkedDefault: false,
      });
      return;
    }
    const hasCurrentTrace = Boolean(
      personStillInSh(event.personId, event.fullName) ||
        personStillInEjoos(event.personId, event.fullName) ||
        (event.personId &&
          (activeArrivalIds.has(event.personId) ||
            openAbsentIds.has(event.personId))) ||
        (event.fullName &&
          (activeArrivalNames.has(normKey(event.fullName)) ||
            openAbsentNames.has(normKey(event.fullName)))),
    );
    if (
      staleMovementCutoff &&
      movementDate &&
      movementDate < staleMovementCutoff &&
      !hasCurrentTrace
    ) {
      return;
    }

    if (event.type === "РОЗПОРЯДЖ") {
      const samePerson = (personId: string, fullName: string) =>
        isSamePerson(event, { personId, fullName });
      const inCurrentSh = personStillInSh(event.personId, event.fullName);
      const inActiveShpo = onStaffShpo(event.personId, event.fullName);
      const hasOpenAbsence = ejoosAbsents.some(
        (row) =>
          samePerson(row.personId, row.fullName) &&
          !row.actualReturn,
      );
      const inActiveOos = Boolean(
        (event.personId && oosPersonById.has(event.personId)) ||
          byPersonName(oosPersonByName, event.personId, event.fullName),
      );

      // Стара подія РУХ не є завданням сама по собі. Якщо людини вже немає
      // в актуальних джерелах і немає відкритої відсутності, стан вважаємо
      // відпрацьованим раніше (ALREADY_PROCESSED / NO_ACTION).
      if (
        !inCurrentSh &&
        !inActiveShpo &&
        !inActiveOos &&
        !hasOpenAbsence
      ) {
        return;
      }

      const oldPosition = shpoByPositionIndex.get(event.previousIndex);
      const oldPositionFreed =
        !oldPosition ||
        !samePerson(oldPosition.personId, oldPosition.fullName);
      const remainsInOos = inActiveOos;
      const hasSzchContext = Boolean(
        (event.personId && activeSzchPersonIds.has(event.personId)) ||
          [...personNameKeys(event.personId, event.fullName)].some((key) =>
            activeSzchPersonNames.has(key),
          ),
      );
      const szchRemains = ejoosAbsents.some(
        (row) =>
          samePerson(row.personId, row.fullName) &&
          isDispositionAbsenceStatus(row.ground) &&
          !row.actualReturn,
      );
      const dispositionInShpo = personInTextRows(event, shpoDispositionRows);
      const dispositionInTimesheet = personInTextRows(
        event,
        timesheetDispositionRows,
      );
      const szchReflectedElsewhere =
        personInTextRows(event, shpoSzchRows) ||
        personInTextRows(event, timesheetSzchRows);
      const absenceStateReflected =
        !hasSzchContext || szchRemains || szchReflectedElsewhere;
      const dispositionStateReflected =
        dispositionInShpo || dispositionInTimesheet;
      const openArchivePeriods = archiveAll
        .filter(
          (period) =>
            samePerson(period.personId, period.fullName) &&
            !hasActualReturn(period.returnDate),
        )
        .sort((a, b) => b.excelRow - a.excelRow);
      // Стан у РУХ (СЗЧ / БЕЗВІСТИ) головніший за випадковий останній період
      // архіву: інакше в блок розпорядження потрапляє «ЛІКУВАННЯ» чи службовий
      // рядок «ВНЕСЕННЯ ДАНИХ».
      const activeArchivePeriod = isDispositionAbsenceStatus(event.status)
        ? (openArchivePeriods.find((period) =>
            isDispositionAbsenceStatus(period.absenceType),
          ) ?? openArchivePeriods[0])
        : openArchivePeriods[0];
      const timesheetRow = ejoosDays.find((row) =>
        samePerson(row.personId, row.fullName),
      );
      // Рядок особи в ШПО/Табелі може бути вже в блоці розпорядження — тоді
      // його не чіпаємо. Закривати треба лише штатний рядок.
      const shpoDispositionExcelRows = new Set(
        shpoDispositionRows.map((row) => row.excelRow),
      );
      const timesheetDispositionExcelRows = new Set(
        timesheetDispositionRows.map((row) => row.excelRow),
      );
      const personShpoRow =
        (event.personId && shpoPersonById.get(event.personId)) ||
        byPersonName(shpoPersonByName, event.personId, event.fullName) ||
        byPersonName(shpoStrayRowByName, event.personId, event.fullName) ||
        null;
      const staffShpoRow =
        oldPosition && samePerson(oldPosition.personId, oldPosition.fullName)
          ? oldPosition
          : personShpoRow &&
              !shpoDispositionExcelRows.has(personShpoRow.excelRow)
            ? personShpoRow
            : null;
      const staffTimesheetRow =
        timesheetRow &&
        !timesheetDispositionExcelRows.has(timesheetRow.excelRow)
          ? timesheetRow
          : null;

      // РОЗПОРЯДЖ не є виключенням зі списків частини: ООС та СЗЧ
      // зберігаються, а до «3. Виключені» особа не переноситься.
      if (
        oldPositionFreed &&
        !staffShpoRow &&
        !staffTimesheetRow &&
        ((absenceStateReflected && dispositionStateReflected) ||
          (remainsInOos && hasOpenAbsence) ||
          alreadyVacatedForAbsence(event.personId, event.fullName))
      ) {
        return;
      }

      // Розпорядження проводимо лише після непроведених змін посади:
      // спочатку історія 2103791 → 2103179, і лише потім вивід у розпорядження.
      const pendingPositionSteps = (
        positionChainByPerson.get(movementPersonKey(event)) ?? []
      ).filter((pending) => pending.excelRow < event.excelRow);
      if (pendingPositionSteps.length) {
        const first = pendingPositionSteps[0];
        const last = pendingPositionSteps[pendingPositionSteps.length - 1];
        const total = pendingPositionSteps.length + 1;
        ops.push({
          id: opId([
            "disposition-after-position",
            event.movementNumber || String(event.excelRow),
          ]),
          kind: "other_manual",
          class: "needs_input",
          sheet: "ШПО → розпорядження / Тимчасово відсутні / Табель",
          personId: event.personId,
          fullName: event.fullName,
          rank: event.rank,
          positionIndex: event.previousIndex,
          before: `у ЕЖООС ще штатна посада ${first.previousIndex}`,
          after: `спочатку зміна посади ${first.previousIndex} → ${last.nextIndex}, потім розпорядження з ${event.previousIndex || "—"}`,
          sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
          why: `Крок ${total} з ${total}: розпорядження проводимо після зміни посади, інакше зникне історія переходу ${first.previousIndex} → ${last.nextIndex}.`,
          confidence: "manual",
          payload: {
            type: event.type,
            chainWaiting: "1",
            chainStep: String(total),
            chainTotal: String(total),
            awaitFromIndex: first.previousIndex,
            awaitToIndex: last.nextIndex,
            awaitOrderNumber: first.orderNumber,
            awaitOrderDate: first.orderDate,
            previousIndex: event.previousIndex,
            destination: event.destination || event.changeText,
            orderNumber: event.orderNumber,
            orderDate: event.orderDate,
          },
          checkedDefault: false,
        });
        return;
      }

      // Стан буває напів-проведений: у Виключених запис уже є, а в Табелі
      // особа досі «+» на штатній посаді й немає запису відсутності.
      // Єдине правильне місце для СЗЧ/БЕЗВІСТИ — «5. Тимчасово відсутні»:
      // згадка в тексті Табеля запису не заміняє.
      const needsAbsenceRecord = hasSzchContext && !szchRemains;
      const needsTimesheetClose = Boolean(staffTimesheetRow);
      const canMoveToDisposition = Boolean(
        staffShpoRow || needsTimesheetClose || needsAbsenceRecord,
      );
      if (canMoveToDisposition) {
        const absenceStatus = isDispositionAbsenceStatus(event.status)
          ? norm(event.status)
          : activeArchivePeriod?.absenceType || event.status || "РОЗПОРЯДЖЕННЯ";
        const mappedAbsence = mapStatus(absenceStatus);
        ops.push({
          id: opId([
            "move-to-disposition",
            event.movementNumber || String(event.excelRow),
          ]),
          kind: "move_to_disposition",
          // Без даних архіву запис відсутності заповнити нічим.
          class:
            needsAbsenceRecord && !activeArchivePeriod ? "needs_input" : "ready",
          sheet: "ШПО → розпорядження / Тимчасово відсутні / Табель",
          personId: event.personId || staffShpoRow?.personId || "",
          fullName: event.fullName || staffShpoRow?.fullName || "",
          rank: event.rank || staffShpoRow?.rank || "",
          positionIndex: event.previousIndex,
          before: `штатна посада ${event.previousIndex}`,
          after: [
            event.destination || event.changeText || "у розпорядження",
            absenceStatus,
            needsAbsenceRecord && !activeArchivePeriod
              ? "немає даних архіву для запису відсутності — перевірити"
              : "",
            !staffTimesheetRow && !dispositionInTimesheet
              ? "штатний рядок Табеля не знайдено — перевірити"
              : "",
            !remainsInOos ? "в ООС активного запису немає" : "",
          ]
            .filter(Boolean)
            .join(" · "),
          sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
          why: "РОЗПОРЯДЖ звільняє стару штатну посаду, але залишає особу в ООС. Виключені не змінюються.",
          confidence: activeArchivePeriod ? "high" : "review",
          payload: {
            type: event.type,
            previousIndex: event.previousIndex,
            destination: event.destination || event.changeText,
            orderNumber: event.orderNumber,
            orderDate: event.orderDate,
            basisNumber: event.basisNumber,
            basisDate: event.basisDate,
            shpoExcelRow: String(staffShpoRow?.excelRow || ""),
            timesheetExcelRow: String(staffTimesheetRow?.excelRow || ""),
            skipShpoDisposition: dispositionInShpo ? "1" : "",
            absenceExcelRow: String(
              ejoosAbsents.find(
                (row) =>
                  samePerson(row.personId, row.fullName) && !row.actualReturn,
              )?.excelRow || "",
            ),
            needsAbsenceRecord: needsAbsenceRecord ? "1" : "",
            absenceType: absenceStatus,
            absenceCode: mappedAbsence.timesheetCode || absenceStatus,
            absenceDate: activeArchivePeriod?.departDate || "",
            absencePlace: activeArchivePeriod?.place || "",
            absenceOrderNumber: activeArchivePeriod?.orderNumber || "",
            absenceOrderDate: activeArchivePeriod?.orderDate || "",
            plannedReturn: activeArchivePeriod?.plannedReturn || "",
            remainsInOos: String(remainsInOos),
            timesheetFound: String(Boolean(staffTimesheetRow)),
            hasSzchContext: String(hasSzchContext),
            szchRemains: String(szchRemains),
            szchReflectedElsewhere: String(szchReflectedElsewhere),
            dispositionInShpo: String(dispositionInShpo),
            dispositionInTimesheet: String(dispositionInTimesheet),
            timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
              augustAbsenceSpansFor(
                event.personId || staffShpoRow?.personId || "",
                event.fullName || staffShpoRow?.fullName || "",
              ),
            ),
          },
          movementKey: createMovementKey(event),
          checkedDefault: !(needsAbsenceRecord && !activeArchivePeriod),
        });
        return;
      }

      const missing = [
        !dispositionStateReflected &&
          "перевірити відображення розпорядження у ШПО або Табелі",
        hasSzchContext &&
          !absenceStateReflected &&
          "перевірити чинний запис відсутності у Тимчасово відсутніх",
      ].filter(Boolean);
      ops.push({
        id: opId([
          "disposition-review",
          event.movementNumber || String(event.excelRow),
        ]),
        kind: "other_manual",
        class: "conflict",
        sheet: "ШПО / ООС / Тимчасово відсутні / Табель",
        personId: event.personId,
        fullName: event.fullName,
        rank: event.rank,
        positionIndex: event.previousIndex,
        before: `РОЗПОРЯДЖ зі штатної посади ${event.previousIndex || "—"}`,
        after: missing.join("; "),
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: "РОЗПОРЯДЖ не видаляє особу з ООС і не додає її до Виключених без окремої події виключення зі списків частини.",
        confidence: "manual",
        payload: {
          type: event.type,
          previousIndex: event.previousIndex,
          destination: event.destination || event.changeText,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
          oldPositionFreed: String(oldPositionFreed),
          remainsInOos: String(remainsInOos),
          hasSzchContext: String(hasSzchContext),
          szchRemains: String(szchRemains),
          szchReflectedElsewhere: String(szchReflectedElsewhere),
          dispositionInShpo: String(dispositionInShpo),
          dispositionInTimesheet: String(dispositionInTimesheet),
          dispositionStateReflected: String(dispositionStateReflected),
        },
        checkedDefault: false,
      });
      return;
    }

    const isExternalUnitDeparture =
      isOutboundStaffMove(event) &&
      !(
        event.type === "ПЕРЕВ" &&
        internalPositionMovementRows.has(event.excelRow)
      ) &&
      !(event.type === "ПОСАДА" && isPositionChangeRow(event.excelRow));
    if (isExternalUnitDeparture) {
      const inboundPlacement = inboundStaffPlacementBefore(event);
      const existingExcludedEarly = findMovementExcludedRow(event);
      const unrecordedSameMonthTransit = isUnrecordedSameMonthTransit({
        hasInboundPlacement: Boolean(inboundPlacement),
        alreadyExcluded: Boolean(existingExcludedEarly),
        stillInSh: personStillInSh(event.personId, event.fullName),
        stillInEjoos: personStillInEjoos(event.personId, event.fullName),
      });
      if (
        skipExternalIfAlreadyProcessed({
          stillInEjoos: personStillInEjoos(event.personId, event.fullName),
          unrecordedTransit: unrecordedSameMonthTransit,
        })
      ) {
        return;
      }
      // Скасоване переведення не є чинним: новий рядок у Виключені не пишемо,
      // навіть якщо історичний запис від того наказу вже є.
      if (cancelledExternalTransferRows.has(event.excelRow)) return;
      {
        const cancel = transferCancelOf(event);
        if (
          cancel &&
          movementEventTime(cancel) >= movementEventTime(event)
        ) {
          return;
        }
      }
      // Вибуття підтверджуємо не лише рядком РУХ: людини вже не повинно бути
      // в актуальному sh. Інакше старий ПЕРЕВ/РОЗПОРЯДЖ створює хибне виключення.
      if (personStillInSh(event.personId, event.fullName)) return;
      const existingExcluded = existingExcludedEarly;
      const alreadyExcluded = Boolean(existingExcluded);
      const destinationRaw = event.destination || "";
      const destinationUpper = destinationRaw.toUpperCase();
      const rawDestination =
        !destinationRaw ||
        destinationUpper.includes("РОЗПОР") ||
        destinationUpper.includes("ПЕРЕВ") ||
        destinationUpper === event.type ||
        isOwnFirstPbDestination(destinationRaw)
          ? ""
          : destinationRaw;
      // Якщо «Куди» = 0 або лишилось 1ПБ, військова частина часто в S / примітці (А7400).
      const unitFromS = unitCodeFromMovement(event);
      const exclusionPlace =
        rawDestination ||
        (/[АA]\s*\d{4}/iu.test(unitFromS) ? unitFromS : "") ||
        event.note ||
        event.changeText;
      // Табель: «вибув до/у» + підрозділ з «Яка зміна» без назви посади.
      // Код в/ч А#### лишається для Виключених (AE / підстава), не для Табеля.
      const timesheetDestination = (() => {
        const positionSource =
          positionChangeDestination(event) || event.changeText;
        const fromPosition =
          extractTimesheetDestinationFromPosition(positionSource);
        if (fromPosition) return fromPosition;
        const unitPhrase =
          formatTransferDestinationForTimesheet(
            [rawDestination, event.note, unitFromS, event.changeText]
              .filter(Boolean)
              .join(" "),
          ) || rawDestination;
        if (/[АA]\s*\d{4}/iu.test(unitPhrase) || /в\s*\/\s*ч/iu.test(unitPhrase)) {
          return event.note && /[АA]\s*\d{4}|в\s*\/\s*ч/iu.test(event.note)
            ? event.note
            : unitPhrase;
        }
        return unitPhrase;
      })();
      const exclusionReason =
        event.type === "ПЕРЕВ" || event.type === "ПОСАДА"
          ? "ПЕРЕВЕДЕННЯ"
          : "Розпорядження";
      const shpoAtOldIndex = (event.previousIndex || inboundPlacement?.nextIndex)
        ? shpoByIndex.get(event.previousIndex || inboundPlacement?.nextIndex || "") ??
          null
        : null;
      // Індекс після вибуття міг зайняти хтось інший (Хубаєв → Атрахов на
      // 2103764). ШПО/Табель чіпаємо лише якщо рядок цієї самої особи.
      const shpo =
        (event.personId &&
          ejoosShpo.find((row) => row.personId === event.personId)) ||
        ejoosShpo.find((row) => isSamePerson(event, row)) ||
        (shpoAtOldIndex && isSamePerson(event, shpoAtOldIndex)
          ? shpoAtOldIndex
          : null);
      const oos =
        (event.personId && oosPersonById.get(event.personId)) ||
        byPersonName(oosPersonByName, event.personId, event.fullName) ||
        null;
      const staffIndexLeft =
        (isPositionIndex(event.previousIndex) && event.previousIndex) ||
        (isPositionIndex(inboundPlacement?.nextIndex || "") &&
          inboundPlacement!.nextIndex) ||
        "";
      // Для вибуття у «Виключені» базові персональні дані мають іти з ЕЖООС.
      // Якщо в серпні було ЗВАННЯ раніше за цей ПЕРЕВ — беремо нове звання,
      // навіть якщо картка ООС/ШПО ще не оновлена.
      const rankBeforeTransfer = [...latestRankEventByPerson.values()].find(
        (rankEvent) =>
          isSamePerson(event, rankEvent) &&
          movementEventTime(rankEvent) <= movementEventTime(event),
      );
      const pendingRank =
        pendingRankByPerson.get(movementPersonKey(event)) ||
        (event.personId ? pendingRankByPerson.get(`id:${event.personId}`) : null);
      const promotedRank =
        pendingRank?.payload.nextRank ||
        (rankBeforeTransfer
          ? parseRankPromotion(rankBeforeTransfer).nextRank
          : "");
      const fromRank =
        promotedRank || oos?.rank || shpo?.rank || event.rank || "";
      const fromName = oos?.fullName || shpo?.fullName || event.fullName || "";
      const occupiedIndex =
        shpo?.positionIndex ||
        String(oos?.positionIndex || "").match(/\d{5,}/)?.[0] ||
        oos?.positionIndex ||
        "";
      // У «Виключені» пишемо посаду, з якої вибув із 1ПБ (РУХ), навіть якщо
      // внутрішню ПОСАДУ в ЕЖООС ще не провели.
      const fromIndex = staffIndexLeft || occupiedIndex;
      const fromId = personIdFromShpo(ejoosShpo, {
        fullName: fromName,
        positionIndex: occupiedIndex || shpo?.positionIndex || "",
        personId: oos?.personId || event.personId || shpo?.personId,
      });
      const arrival =
        arrivalOf(fromId, fromName) ||
        arrivalOf(event.personId, event.fullName);
      const tsCandidate =
        activeTimesheetRowOf(event.personId, event.fullName, occupiedIndex) ||
        ejoosDays.find((row) => isSamePerson(event, row)) ||
        (occupiedIndex
          ? staffIndexTimesheetForPerson(fromId, fromName, occupiedIndex)
          : null) ||
        null;
      const ts =
        tsCandidate &&
        (isVacantStaffRow(tsCandidate) || isSamePerson(event, tsCandidate))
          ? tsCandidate
          : ejoosDays.find((row) => isSamePerson(event, row)) || null;
      // Колонки «наказ» і перша «дата» — стройовий наказ. Окремі реквізити
      // підстави (наприклад 668-РС від 03.08.2026) сюди не підставляємо.
      const excludeDate = event.orderDate;
      const hasRequiredExcludedFields = Boolean(
        exclusionPlace &&
          exclusionReason &&
          excludeDate &&
          event.orderNumber &&
          event.orderDate,
      );

      ops.push({
        id: opId(["excl", event.movementNumber || String(event.excelRow), event.type]),
        kind: "exclude_transfer",
        class:
          hasRequiredExcludedFields && Boolean(shpo || fromName)
            ? "ready"
            : "needs_input",
        sheet: "Виключені → Табель → ШПО/ООС",
        personId: fromId || event.personId,
        fullName: fromName,
        positionIndex: fromIndex,
        rank: fromRank,
        before: shpo
          ? occupiedIndex && occupiedIndex !== fromIndex
            ? `ШПО R${shpo.excelRow}: ${fromRank} ${fromName} · інд. ${occupiedIndex} (кінцева посада РУХ ${fromIndex})`
            : `ШПО R${shpo.excelRow}: ${fromRank} ${fromName} · інд. ${fromIndex}`
          : unrecordedSameMonthTransit
            ? `немає в ШПО/ООС · транзит ${inboundPlacement?.orderDate || "?"} → ${event.orderDate || "?"}`
            : "в обліку (ШПО не знайдено — перевірте вручну)",
        after: `${alreadyExcluded ? "доробити очищення після виключення" : "виключити"}: ${event.type} → ${exclusionPlace || "(куди?)"} · табель: ${timesheetDestination || "(куди?)"} · ${exclusionReason} · наказ №${event.orderNumber || "?"} від ${event.orderDate || "?"}`,
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: alreadyExcluded
            ? "Рядок у Виключених уже є — не дублюємо його, доробляємо Табель → ШПО/ООС. Тимч. відсутні/прибулі не чіпаємо."
          : unrecordedSameMonthTransit
            ? `Постановка ${inboundPlacement?.orderDate || "?"} на ${fromIndex || "штабний індекс"} і вибуття ${event.orderDate || "?"} у цьому місяці, але в ЕЖООС особи не було. Пишемо Виключені + історичний Табель; зайнятий індекс ШПО/ООС не чіпаємо.`
          : hasRequiredExcludedFields
            ? pendingRank
              ? `Спочатку звання ${pendingRank.payload.nextRank} (№${pendingRank.payload.orderNumber || "?"} від ${pendingRank.payload.orderDate || "?"}), потім ${event.type} → ${exclusionPlace}: Виключені → Табель → очистка ШПО/ООС`
              : inboundPlacement
                ? `Ланцюг: ${inboundPlacement.orderDate || "?"} ${inboundPlacement.type} ${inboundPlacement.previousIndex || "?"} → ${inboundPlacement.nextIndex || fromIndex} (лишився б у ООС); далі ${event.type} №${event.orderNumber || "?"} від ${event.orderDate || "?"} — вибув з 1ПБ, в sh немає. Виключені з інд. ${fromIndex}; внутрішню постановку на штат не проводимо.`
                : `${event.type} з «куди»: алгоритм Виключені → Табель (історія) → очистка ШПО/ООС. Відкритий рядок «Тимч. прибулі» закриваємо.`
            : "У РУХ не вистачає фактичного місця вибуття або реквізитів стройового наказу",
        confidence:
          hasRequiredExcludedFields && (shpo || unrecordedSameMonthTransit)
            ? "high"
            : "manual",
        payload: {
          movementNumber: event.movementNumber,
          type: event.type,
          destination: exclusionPlace,
          timesheetDestination,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
          excludeDate,
          timesheetActiveFrom: inboundPlacement?.orderDate || "",
          timesheetCreateHistory:
            !ts || isVacantStaffRow(ts) || !isSamePerson(event, ts)
              ? "1"
              : "",
          transitSameMonth: unrecordedSameMonthTransit ? "1" : "",
          basisNumber: event.basisNumber,
          basisDate: event.basisDate,
          previousIndex: staffIndexLeft || event.previousIndex,
          nextIndex: event.nextIndex,
          occupiedPositionIndex: occupiedIndex,
          priorPlacementRow: inboundPlacement
            ? String(inboundPlacement.excelRow)
            : "",
          priorPlacementType: inboundPlacement?.type || "",
          priorPlacementDate: inboundPlacement?.orderDate || "",
          priorPlacementOrder: inboundPlacement?.orderNumber || "",
          priorPlacementFromIndex: inboundPlacement?.previousIndex || "",
          priorPlacementToIndex: inboundPlacement?.nextIndex || "",
          arrivedFrom:
            inboundPlacement?.previousIndex ||
            inboundPlacement?.destination ||
            "",
          appointmentOrderNumber: inboundPlacement?.orderNumber || "",
          appointmentOrderDate: inboundPlacement?.orderDate || "",
          arrivalExcelRow: arrival ? String(arrival.excelRow) : "",
          arrivalDepartDate:
            inboundPlacement?.orderDate || event.orderDate || "",
          arrivalDepartOrderNumber:
            inboundPlacement?.orderNumber || event.orderNumber || "",
          arrivalDepartOrderDate:
            inboundPlacement?.orderDate || event.orderDate || "",
          excludedExcelRow: existingExcluded ? String(existingExcluded.excelRow) : "",
          oosExcelRow: oos ? String(oos.excelRow) : "",
          shpoExcelRow: shpo ? String(shpo.excelRow) : "",
          timesheetExcelRow:
            ts && isSamePerson(event, ts) ? String(ts.excelRow) : "",
          fromRank,
          fromName,
          fromPersonId: fromId || event.personId,
          fromPositionIndex: fromIndex,
          lastRankOrderNumber:
            pendingRank?.payload.orderNumber ||
            rankBeforeTransfer?.orderNumber ||
            "",
          lastRankOrderDate:
            pendingRank?.payload.orderDate ||
            rankBeforeTransfer?.orderDate ||
            "",
          // AE «Куди вибув»: в/ч з примітки / «Куди», не стара посада з «Яка зміна».
          documentsDest:
            exclusionPlace ||
            positionChangeDestination(event) ||
            event.changeText,
          changeText: event.changeText || "",
          exclusionReason,
          awaitRankChange: pendingRank ? "1" : "",
        },
        movementKey: createMovementKey(event),
        checkedDefault:
          hasRequiredExcludedFields &&
          Boolean(shpo || unrecordedSameMonthTransit),
      });
      if (
        /відсутн.*архів|в архіві/iu.test(`${event.status} ${event.note}`) &&
        !archiveAll.some((period) => isSamePerson(event, period)) &&
        !ops.some(
          (op) =>
            op.payload.mismatchKind === "ARCHIVE_REFERENCE_MISSING" &&
            isSamePerson(event, op),
        )
      ) {
        ops.push({
          id: opId(["archive_missing", event.personId || event.fullName]),
          kind: "data_mismatch",
          class: "needs_input",
          sheet: "Дані джерел / archive",
          personId: fromId || event.personId,
          fullName: fromName,
          positionIndex: fromIndex,
          rank: fromRank,
          before: event.status || "ВІДСУТНІЙ в АРХІВІ",
          after: "ARCHIVE_REFERENCE_MISSING → перевірити archive",
          sourceRef: `Рух!R${event.excelRow} СТАТУС=«${event.status}»`,
          why: "У РУХ стоїть «ВІДСУТНІЙ в АРХІВІ», але в archive немає запису за ПІБ чи ID. ЛІК / ВІД / СЗЧ не вигадуємо — кадровий маршрут вибуття це не скасовує.",
          confidence: "review",
          payload: {
            type: "ARCHIVE_REFERENCE_MISSING",
            mismatchKind: "ARCHIVE_REFERENCE_MISSING",
            statusRaw: event.status,
          },
          checkedDefault: false,
        });
      }
      return;
    }

    if (isPositionChangeRow(event.excelRow)) {
      if (ownUnitMoveSuperseded(event)) return;
      const chained = chainedPositionRows.has(event.excelRow);
      const personChain = positionChainByPerson.get(movementPersonKey(event));
      // Ланцюг має пріоритет: інакше по одній особі вийшло б дві різні
      // операції зміни посади з різних рядків РУХ.
      if (personChain?.length && !chained) return;
      if (personChain?.length) {
        const last = personChain[personChain.length - 1];
        if (event.excelRow !== last.excelRow) return;
      }
      if (
        !chained &&
        latestPositionByName.get(normKey(event.fullName))?.excelRow !==
          event.excelRow
      ) {
        return;
      }
      const shPerson =
        (event.personId && shPersonById.get(event.personId)) ||
        byPersonName(shPersonByName, event.personId, event.fullName) ||
        null;
      // `sh` is authoritative for the current occupant and usually contains the
      // stable ID which may be absent in РУХ.
      const personId = shPerson?.personId || event.personId;
      const fullName = shPerson?.fullName || event.fullName;
      const rank = shPerson?.rank || event.rank;
      const arrival = arrivalOf(personId, fullName);
      // Відкрита відсутність не скасовує постановку з «Тимчасово прибулі».
      if (
        alreadyVacatedForAbsence(personId, fullName) &&
        !arrival
      ) {
        return;
      }
      // У кроці ланцюга особи вже може не бути в sh на цьому індексі
      // (далі за датою — розпорядження), тому цільовий індекс беремо з РУХ.
      const shIndex = isPositionIndex(shPerson?.positionIndex || "")
        ? shPerson!.positionIndex
        : "";
      const nextIndex = chained
        ? event.nextIndex || shIndex || event.previousIndex
        : shIndex || event.nextIndex || event.previousIndex;
      // Без коректного нового індексу писати нову посаду нікуди.
      if (!isPositionIndex(nextIndex)) return;
      const targetShpo = nextIndex ? shpoByIndex.get(nextIndex) ?? null : null;
      const indexTimesheet = nextIndex
        ? dayByIndex.get(nextIndex) ?? null
        : null;
      const existingOos =
        (personId && oosById.get(personId)) ||
        byPersonName(oosByName, personId, fullName) ||
        null;
      const excludedSource =
        existingOos ? null : findLatestExcludedRow(personId, fullName);
      const transferCancel = transferCancelOf(event);
      const cancelledTransfer = cancelledTransferOf({ personId, fullName });
      const personTimesheetRows = timesheetRowsOf(personId, fullName);
      const personStaffTimesheet = staffIndexTimesheetForPerson(
        personId,
        fullName,
        nextIndex,
      );
      const personActiveTimesheet =
        [...personTimesheetRows]
          .filter((row) => !row.hasDepartureText)
          .sort((left, right) => right.plusDays.length - left.plusDays.length)[0] ??
        null;
      const targetTimesheet = (() => {
        if (personStaffTimesheet) return personStaffTimesheet;
        if (personActiveTimesheet) return personActiveTimesheet;
        if (isVacantStaffRow(indexTimesheet)) return indexTimesheet;
        if (indexTimesheet && isSamePerson({ personId, fullName }, indexTimesheet)) {
          return indexTimesheet;
        }
        return indexTimesheet;
      })();
      const actualExclusionEvidence = Boolean(
        cancelledTransfer && findMovementExcludedRow(cancelledTransfer),
      );
      const historyTimesheet =
        personTimesheetRows.find(
          (row) =>
            row.hasDepartureText &&
            row.excelRow !== targetTimesheet?.excelRow,
        ) ?? null;
      const currentTimesheet =
        historyTimesheet ||
        (personId && dayById.get(personId)) ||
        ejoosDays.find((row) => isSamePerson({ personId, fullName }, row)) ||
        null;
      const samePersonRow = (
        row: { personId: string; fullName: string } | null | undefined,
      ) => Boolean(row && isSamePerson({ personId, fullName }, row));
      const previousShpo = event.previousIndex
        ? shpoByIndex.get(event.previousIndex) ?? null
        : null;
      const previousIndexTimesheet = event.previousIndex
        ? dayByIndex.get(event.previousIndex) ?? null
        : null;
      // Рядок особи в ШПО може бути без індексу (лише ПІБ) — його теж треба
      // звільнити, але цільовий рядок нової посади не чіпаємо.
      const personShpoRow =
        (personId && shpoPersonById.get(personId)) ||
        byPersonName(shpoPersonByName, personId, fullName) ||
        null;
      const oldShpo = samePersonRow(previousShpo)
        ? previousShpo
        : personShpoRow && personShpoRow.excelRow !== targetShpo?.excelRow
          ? personShpoRow
          : null;
      const previousShpoTimesheet = samePersonRow(previousIndexTimesheet)
        ? previousIndexTimesheet
        : null;
      // Стару штатну посаду закриваємо як історію: рядок у «3. Виключені»,
      // копія рядка Табеля і звільнення індексу в ШПО. ООС не чіпаємо —
      // особа лишається в частині.
      const closeOldPosition = Boolean(
        event.previousIndex &&
          event.previousIndex !== nextIndex &&
          (oldShpo || previousShpoTimesheet),
      );
      const returningToStaffIndex = Boolean(
        shPerson &&
          nextIndex === shPerson.positionIndex &&
          historyTimesheet?.hasDepartureText &&
          isOwnUnitStaffMove(event) &&
          !samePersonRow(targetShpo),
      );
      // Якщо особа вже стоїть на цільовому індексі в ЕЖООС і стару посаду
      // закривати не треба — у РУХ лише історія, змін немає. Скасоване
      // переведення — виняток: Табель треба розкласти на історію + новий рядок.
      if (samePersonRow(targetShpo) && !closeOldPosition && !returningToStaffIndex) {
        const arrivalStillOpen = Boolean(arrival);
        const oosMissing = !existingOos;
        if (
          !arrivalStillOpen &&
          !oosMissing &&
          (!transferCancel ||
            !timesheetNeedsTransferCancelSplit(
              personId,
              fullName,
              nextIndex,
              transferCancel.orderDate,
            ))
        ) {
          return;
        }
      }
      // Старий рядок «Тимчасово прибулі» лишається історією. Якщо особа вже
      // на штатній посаді, це не постановка з тимчасового прибуття.
      const isTempArrivalPlacement = Boolean(arrival) && !closeOldPosition;
      const openAbsence = openAbsenceOf(personId, fullName);
      const dispositionShpoExcelRow = findNonStaffOccupantExcelRow(
        shpoSheet,
        personId,
        fullName,
        { index: 0, name: 6, id: 7 },
      );
      const dispositionTimesheetExcelRow = findNonStaffOccupantExcelRow(
        timesheetSheet,
        personId,
        fullName,
        { index: 1, name: 6, id: 7 },
      );
      const returningFromDisposition = Boolean(
        dispositionShpoExcelRow ||
          dispositionTimesheetExcelRow ||
          openAbsence ||
          /розпорядж/iu.test(
            `${personChain?.[0]?.previousIndex || ""} ${event.previousIndex}`,
          ),
      );
      const indexHistory = (() => {
        const fromChain = (personChain?.length ? personChain : [event])
          .filter((item) => isPositionIndex(item.nextIndex))
          .map((item) => ({
            index: item.nextIndex,
            date: item.orderDate || item.basisDate || event.orderDate,
          }));
        if (fromChain.length > 1) return fromChain;
        const extra = ownUnitIndexHistoryOf(personId, fullName);
        return extra.length ? extra : fromChain;
      })();
      const newestFirstHistory = [...indexHistory].reverse();
      const oosHistoryIndexes = newestFirstHistory
        .map((item) => item.index)
        .filter((index, idx, all) => all.indexOf(index) === idx)
        .join("\n");
      const oosHistoryDates = newestFirstHistory
        .filter(
          (item, idx, all) =>
            all.findIndex((other) => other.index === item.index) === idx,
        )
        .map((item) => item.date)
        .join("\n");
      const staffTimesheetFrom =
        staffAppointmentDateFor(personId, fullName, nextIndex) ||
        inboundStaffDateFor(personId, fullName) ||
        event.orderDate;
      const monthSpans = augustAbsenceSpansFor(personId, fullName);
      const carryAbsenceFromMonthStart = monthSpans.some(
        (span) => span.fromDay === 1,
      );
      const timesheetActiveFrom = carryAbsenceFromMonthStart
        ? journalMonthStartLabel
        : returningToStaffIndex
          ? event.orderDate
          : staffTimesheetFrom;
      const timesheetPreserveHistory =
        !carryAbsenceFromMonthStart &&
        journalDayFromDateMs(dateMs(timesheetActiveFrom), leadWindowStart) > 1
          ? "1"
          : "";
      const staleTimesheet =
        timesheetRowsOf(personId, fullName).find(
          (row) =>
            row.excelRow !== targetTimesheet?.excelRow && !row.hasDepartureText,
        ) ?? null;
      const newPositionText = positionChangeDestination(event);
      const timesheetDestination =
        extractTimesheetDestinationFromPosition(newPositionText) ||
        newPositionText;
      const chainIndexes = indexHistory.map((item) => item.index).filter(Boolean);
      const chainNote =
        chainIndexes.length > 1
          ? `Ланцюг ${chainIndexes.join(" → ")}: `
          : "";
      const canApply = Boolean(
        (shPerson || chained || returningFromDisposition) &&
          nextIndex &&
          targetShpo &&
          targetTimesheet,
      );
      const applyNow = canApply;
      const processedBefore = wasMovementProcessed(event);

      ops.push({
        id: opId(["pos", event.movementNumber || String(event.excelRow)]),
        kind: "position_change",
        class: applyNow ? "ready" : "needs_input",
        sheet: returningFromDisposition
          ? "5. Тимч. відсутні → 1. ШПО / 2. ООС / 6. Табель"
          : closeOldPosition
            ? "3. Виключені → 6. Табель → 1. ШПО / 2. ООС"
            : isTempArrivalPlacement
              ? "4. Тимч. прибулі → 1. ШПО / 2. ООС / 6. Табель"
              : "1. ШПО / 2. ООС / 6. Табель",
        personId,
        fullName,
        positionIndex: nextIndex,
        rank,
        before: returningFromDisposition
          ? `у розпорядженні / ${openAbsence?.ground || "СЗЧ"}`
          : closeOldPosition
            ? `штатна посада ${event.previousIndex}`
            : isTempArrivalPlacement
              ? `тимчасово прибулий${arrival?.positionIndex ? ` · інд. ${arrival.positionIndex}` : ""}`
              : event.previousIndex || "?",
        after: `штатна посада ${nextIndex || event.changeText || "?"}${
          oosHistoryIndexes && oosHistoryIndexes.includes("\n")
            ? ` · ООС ${oosHistoryIndexes.replaceAll("\n", " → ")}`
            : ""
        }`,
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why:
          chainNote +
          (processedBefore
            ? "Подія є в історії застосувань, але фактичний стан ШПО/ООС/Табеля потребує повторного проведення"
            : !canApply
              ? "Не знайдено однозначний рядок нового індексу в ШПО/Табелі — потрібна ручна перевірка"
              : returningFromDisposition
                ? `Повернення із СЗЧ / розпорядження: закрити період, прибрати з блоку «у розпорядженні», поставити на ${nextIndex}${
                    chainIndexes.length > 1
                      ? `, історія ООС ${[...chainIndexes].reverse().join(", ")}`
                      : ""
                  }`
                : closeOldPosition
                  ? `зміна посади в межах 1ПБ: закрити ${event.previousIndex} у Виключених і Табелі, поставити на ${nextIndex}, ООС лишити активним`
                  : isTempArrivalPlacement
                    ? `ПОСАДА №${event.orderNumber || "?"} від ${event.orderDate || "?"}: закрити тимчасове прибуття, поставити на штат ${nextIndex}; Табель з ${staffTimesheetFrom || event.orderDate || "дати наказу"}, коди archive в цьому ж кроці`
                    : transferCancel
                      ? `Серпневий ланцюг: ПОСАДА ${nextIndex} з ${staffAppointmentDateFor(personId, fullName, nextIndex) || event.orderDate || "?"}; ПЕРЕВ №${cancelledTransfer?.orderNumber || "?"} скасовано №${transferCancel.orderNumber || "?"} — один рядок Табеля, запис у «Виключені» прибираємо`
                    : returningToStaffIndex
                      ? `Повернення на ${nextIndex} з ${event.orderDate || "?"}: історичний рядок з вибуттям лишаємо, новий активний з «+» від дати наказу`
                      : event.type === "ПЕРЕВ"
                      ? "Внутрішній ПЕРЕВ у межах 1ПБ підтверджений поточним sh: змінити штатну позицію, ООС і Табель"
                      : "ПОСАДА підтверджена поточним sh: оновити штатну позицію, ООС і Табель"),
        confidence: applyNow ? "high" : "manual",
        payload: {
          movementNumber: event.movementNumber,
          previousIndex: event.previousIndex,
          nextIndex,
          changeText: event.changeText,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
          basisNumber: event.basisNumber,
          basisDate: event.basisDate,
          nextName: fullName,
          nextRank: rank,
          nextPersonId: personId,
          positionTitle: shPerson?.positionTitle || "",
          statusRaw: shPerson?.status || event.status,
          isTempArrivalPlacement: isTempArrivalPlacement ? "1" : "",
          arrivalExcelRow: arrival ? String(arrival.excelRow) : "",
          oosExcelRow: existingOos ? String(existingOos.excelRow) : "",
          oosHistoryIndexes,
          oosHistoryDates,
          shpoExcelRow: targetShpo ? String(targetShpo.excelRow) : "",
          timesheetExcelRow: targetTimesheet
            ? String(targetTimesheet.excelRow)
            : "",
          previousTimesheetExcelRow:
            returningToStaffIndex && historyTimesheet
              ? String(historyTimesheet.excelRow)
              : currentTimesheet &&
                  currentTimesheet.excelRow !== targetTimesheet?.excelRow
                ? String(currentTimesheet.excelRow)
                : "",
          clearTimesheetExcelRow: staleTimesheet
            ? String(staleTimesheet.excelRow)
            : "",
          timesheetActiveFrom,
          timesheetSkipHistory: "1",
          timesheetPreserveHistory,
          timesheetBindStaffIndex: isTempArrivalPlacement ? nextIndex : "",
          returningToStaffIndex: returningToStaffIndex ? "1" : "",
          returningFromDisposition: returningFromDisposition ? "1" : "",
          dispositionShpoExcelRow: dispositionShpoExcelRow
            ? String(dispositionShpoExcelRow)
            : "",
          dispositionTimesheetExcelRow: dispositionTimesheetExcelRow
            ? String(dispositionTimesheetExcelRow)
            : "",
          openAbsenceExcelRow: openAbsence ? String(openAbsence.excelRow) : "",
          timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
            carryAbsenceFromMonthStart
              ? monthSpans
              : clipAbsenceSpansToActiveEpisode(
                  monthSpans,
                  journalDayFromDateMs(
                    dateMs(timesheetActiveFrom),
                    leadWindowStart,
                  ) || 1,
                ),
          ),
          historyTimesheetExcelRow: (() => {
            if (carryAbsenceFromMonthStart) return "";
            const fromDay = journalDayFromDateMs(
              dateMs(timesheetActiveFrom),
              leadWindowStart,
            );
            if (fromDay <= 1) return "";
            const prior = priorEpisodeTimesheetOf(
              personId,
              fullName,
              targetTimesheet?.excelRow || 0,
              nextIndex,
            );
            return prior ? String(prior.excelRow) : "";
          })(),
          historyTimesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
            carryAbsenceFromMonthStart
              ? []
              : absenceSpansBeforeEpisode(
                  monthSpans,
                  journalDayFromDateMs(
                    dateMs(timesheetActiveFrom),
                    leadWindowStart,
                  ) || 1,
                ),
          ),
          arrivedFrom: event.arrivedFrom || arrival?.fromUnit || "",
          arrivalDepartDate: arrival ? event.orderDate : "",
          arrivalDepartOrderNumber: arrival ? event.orderNumber : "",
          arrivalDepartOrderDate: arrival ? event.orderDate : "",
          cancelledTransferOrder: cancelledTransfer?.orderNumber || "",
          cancelledTransferDate: cancelledTransfer?.orderDate || "",
          cancelledTransferDest:
            cancelledTransfer?.destination ||
            cancelledTransfer?.changeText ||
            "",
          actualExclusionEvidence: actualExclusionEvidence ? "1" : "",
          chainStep: personChain?.length
            ? String(personChain.length)
            : "",
          chainTotal: personChain?.length ? String(personChain.length) : "",
          closeOldPosition: closeOldPosition ? "1" : "",
          previousShpoExcelRow:
            closeOldPosition && oldShpo ? String(oldShpo.excelRow) : "",
          previousIndexTimesheetExcelRow:
            closeOldPosition && previousShpoTimesheet
              ? String(previousShpoTimesheet.excelRow)
              : "",
          excludeDate: closeOldPosition ? event.orderDate : "",
          documentsDest: closeOldPosition ? newPositionText : "",
          timesheetDestination: closeOldPosition ? timesheetDestination : "",
          exclusionReason: closeOldPosition ? "ПЕРЕВЕДЕННЯ 1 ПБ" : "",
          fromRank: oldShpo?.rank || previousShpoTimesheet?.rank || rank,
          fromName: oldShpo?.fullName || fullName,
          fromPersonId: oldShpo?.personId || personId,
          fromPositionIndex: event.previousIndex,
          excludedSourceExcelRow: excludedSource
            ? String(excludedSource.excelRow)
            : "",
          transferCancelOrder: transferCancel?.orderNumber || "",
          transferCancelDate: transferCancel?.orderDate || "",
        },
        movementKey: createMovementKey(event),
        checkedDefault: applyNow,
      });
      return;
    }

    if (event.type === "ПРИБУВ" || event.type === "ЗВІЛЬН") {
      if (
        (event.type === "ПРИБУВ" &&
          personStillInEjoos(event.personId, event.fullName)) ||
        (event.type === "ЗВІЛЬН" &&
          !personStillInEjoos(event.personId, event.fullName))
      ) {
        return;
      }
      const laterPlacement = latestPositionByName.get(normKey(event.fullName));
      if (
        event.type === "ПРИБУВ" &&
        laterPlacement &&
        laterPlacement.excelRow > event.excelRow
      ) {
        // Arrival is historical once a later ПОСАДА has put the person on штат.
        return;
      }
      const processedBefore = wasMovementProcessed(event);
      ops.push({
        id: opId(["mov", event.type, event.movementNumber || String(event.excelRow)]),
        kind: event.type === "ПРИБУВ" ? "arrival" : "other_manual",
        class: "needs_input",
        sheet: event.type === "ПРИБУВ" ? "2. ООС" : "3. Виключені",
        personId: event.personId,
        fullName: event.fullName,
        positionIndex: event.nextIndex || event.previousIndex,
        rank: event.rank,
        before: "—",
        after: `${event.type}: ${event.destination || event.changeText || "потрібні реквізити"}`,
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: processedBefore
          ? "Подію вже застосовували, але поточний стан ЕЖООС їй не відповідає"
          : `${event.type} не застосовується автоматично — заповніть накази/звідки/куди`,
        confidence: "manual",
        payload: {
          movementNumber: event.movementNumber,
          type: event.type,
          destination: event.destination,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
        },
        movementKey: createMovementKey(event),
        checkedDefault: false,
      });
    }
  };

  effectiveMovements.forEach(considerMovement);

  const finalExternalTransferByPerson = new Map<string, PbMovement>();
  for (const event of movementsAll) {
    if (event.type !== "ПЕРЕВ" || isOwnUnitStaffMove(event)) continue;
    if (cancelledExternalTransferRows.has(event.excelRow)) continue;
    if (isCancelledMovementRecord(event) || isTransferCancellation(event)) {
      continue;
    }
    const key = movementPersonKey(event);
    if (!key) continue;
    const previous = finalExternalTransferByPerson.get(key);
    if (
      !previous ||
      movementEventTime(event) > movementEventTime(previous) ||
      (movementEventTime(event) === movementEventTime(previous) &&
        event.excelRow > previous.excelRow)
    ) {
      finalExternalTransferByPerson.set(key, event);
    }
  }
  const finalExternalTransferFor = (personId: string, fullName: string) => {
    if (personId) {
      const byId = finalExternalTransferByPerson.get(`id:${personId}`);
      if (byId) return byId;
    }
    for (const event of finalExternalTransferByPerson.values()) {
      if (isSamePerson({ personId, fullName }, event)) return event;
    }
    return fullName
      ? finalExternalTransferByPerson.get(`name:${normKey(fullName)}`) ?? null
      : null;
  };
  const excludeAlreadyPlanned = (personId: string, fullName: string) =>
    ops.some(
      (op) =>
        op.kind === "exclude_transfer" &&
        isSamePerson({ personId, fullName }, op),
    );
  // Людина ще в ШПО/ООС, в sh уже немає, останній ПЕРЕВ — зовнішній.
  // Статус РУХ «В СТРОЮ» ігноруємо: вирішальні ТИП + в/ч А#### + нема в sh.
  {
    const seen = new Set<number>();
    for (const row of ejoosOccupied) {
      if (personStillInSh(row.personId, row.fullName)) continue;
      if (excludeAlreadyPlanned(row.personId, row.fullName)) continue;
      const transfer = finalExternalTransferFor(row.personId, row.fullName);
      if (!transfer) continue;
      if (!eventInLeadWindow(transfer)) continue;
      if (cancelledExternalTransferRows.has(transfer.excelRow)) continue;
      if (seen.has(transfer.excelRow)) continue;
      seen.add(transfer.excelRow);
      considerMovement(transfer);
    }
  }
  const staffTimesheetRows = new Set(
    [...dayByIndex.values()].map((row) => row.excelRow),
  );
  for (const row of timesheetPeople) {
    if (!staffTimesheetRows.has(row.excelRow)) continue;
    if (personStillInSh(row.personId, row.fullName)) continue;
    if (transferCancelForPerson(row.personId, row.fullName)) continue;
    const transfer = finalExternalTransferFor(row.personId, row.fullName);
    if (!transfer) continue;
    if (excludeAlreadyPlanned(row.personId, row.fullName)) continue;
    const history = timesheetRowsOf(row.personId, row.fullName).find(
      (other) => other.excelRow !== row.excelRow && other.hasDepartureText,
    );
    ops.push({
      id: opId(["stale-tab", row.personId || row.fullName, String(row.excelRow)]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: row.personId,
      fullName: row.fullName,
      positionIndex: row.positionIndex,
      rank: row.rank,
      before: `штатний рядок R${row.excelRow}: ${row.fullName || "ПІБ"} ще стоїть на ${row.positionIndex}`,
      after: "прибрати ПІБ/ID зі штатної позиції; історичний рядок з вибуттям лишити",
      sourceRef: history
        ? `Табель!R${row.excelRow} · історія R${history.excelRow}`
        : `Табель!R${row.excelRow}`,
      why:
        "Особа відсутня в актуальному sh і має чинне зовнішнє переведення, але персональні дані лишилися на старій штатній позиції Табеля. Історичний рядок — не поточний стан; штатний рядок треба очистити.",
      confidence: "high",
      payload: {
        type: "STALE_TAB_PERSON_ROW",
        clearStalePerson: "1",
        excelRow: String(row.excelRow),
        keepHistoryRow: history ? String(history.excelRow) : "",
      },
      movementKey: createMovementKey(transfer),
      checkedDefault: true,
    });
  }

  for (const row of timesheetPeople) {
    if (staffTimesheetRows.has(row.excelRow)) continue;
    const cancel = transferCancelForPerson(row.personId, row.fullName);
    if (!cancel || !eventInLeadWindow(cancel)) continue;
    if (!personStillInSh(row.personId, row.fullName)) continue;
    const staff = row.positionIndex
      ? dayByIndex.get(row.positionIndex)
      : undefined;
    const staffScan = staff
      ? timesheetScanByRow.get(staff.excelRow)
      : undefined;
    // Штатний рядок ще з «вибув» — це не дубль, а стан до відновлення.
    if (staffScan?.hasDepartureText) continue;
    const shPerson =
      (row.personId && shPersonById.get(row.personId)) ||
      byPersonName(shPersonByName, row.personId, row.fullName) ||
      null;
    const personId = row.personId || shPerson?.personId || "";
    const fullName = row.fullName || shPerson?.fullName || "";
    const host =
      ops.find(
        (op) =>
          op.kind === "position_change" &&
          isSamePerson({ personId, fullName }, op),
      ) ||
      ops.find(
        (op) =>
          op.kind === "absent_upsert" &&
          isSamePerson({ personId, fullName }, op),
      ) ||
      ops.find(
        (op) =>
          op.kind !== "data_mismatch" &&
          op.kind !== "timesheet_day" &&
          isSamePerson({ personId, fullName }, op),
      );
    if (host) {
      host.payload.duplicateTimesheetExcelRow = String(row.excelRow);
      host.payload.clearStalePerson = "1";
      host.payload.clearTimesheetIndex = "1";
      continue;
    }
    ops.push({
      id: opId([
        "dup-tab-cancel",
        personId || fullName,
        String(row.excelRow),
      ]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId,
      fullName,
      positionIndex: row.positionIndex || shPerson?.positionIndex || "",
      rank: row.rank || shPerson?.rank || "",
      before: `дубль R${row.excelRow}: ${fullName || "ПІБ"} · «вибув» після скасованого переведення`,
      after: "прибрати другий рядок — за серпень лишається один активний запис",
      sourceRef: `Табель!R${row.excelRow}`,
      why: "Переведення скасовано в тому ж місяці. Другий рядок Табеля з «вибув» не є окремою фактичною наявністю і не потрібен.",
      confidence: "high",
      payload: {
        type: "DUPLICATE_TAB_AFTER_CANCEL",
        clearStalePerson: "1",
        clearTimesheetIndex: "1",
        excelRow: String(row.excelRow),
      },
      movementKey: createMovementKey(cancel),
      checkedDefault: true,
    });
  }

  const timesheetOccupantByIndex = new Map(
    shPeople
      .filter((person) => person.positionIndex)
      .map((person) => [
        person.positionIndex,
        { personId: person.personId, fullName: person.fullName },
      ]),
  );
  const clearedTimesheetRows = new Set(
    ops
      .filter(
        (op) =>
          op.kind === "timesheet_day" &&
          op.payload.clearStalePerson === "1" &&
          Number(op.payload.excelRow || 0) > 0,
      )
      .map((op) => Number(op.payload.excelRow)),
  );
  for (const item of findDuplicateTimesheetExtras(
    timesheetPeople,
    timesheetOccupantByIndex,
  )) {
    if (clearedTimesheetRows.has(item.extra.excelRow)) continue;
    clearedTimesheetRows.add(item.extra.excelRow);
    const keepName = item.keep.fullName || item.keep.personId || "канонічний рядок";
    ops.push({
      id: opId([
        "dup-tab-row",
        item.extra.personId || item.extra.fullName,
        String(item.extra.excelRow),
      ]),
      kind: "timesheet_day",
      class: "ready",
      sheet: "6. Табель",
      personId: item.extra.personId,
      fullName: item.extra.fullName,
      positionIndex: item.extra.positionIndex,
      rank: "",
      before: `дубль R${item.extra.excelRow}: ${item.extra.fullName || "ПІБ"} на ${item.extra.positionIndex || "індексі"}`,
      after: `прибрати рядок — лишається R${item.keep.excelRow} (${keepName})`,
      sourceRef: `Табель!R${item.extra.excelRow} · канон R${item.keep.excelRow}`,
      why:
        item.reason === "same_person"
          ? "У Табелі не може бути двох активних записів однієї особи. Повторне застосування раніше дописувало копію замість штатного рядка."
          : "Один штатний індекс — один активний рядок Табеля. Історія з «вибув» лишається, зайві копії з «+» прибираємо.",
      confidence: "high",
      payload: {
        type: "DUPLICATE_TAB_ROW",
        clearStalePerson: "1",
        clearTimesheetIndex: "1",
        excelRow: String(item.extra.excelRow),
        keepTimesheetExcelRow: String(item.keep.excelRow),
      },
      checkedDefault: true,
    });
  }

  for (const [personId, variants] of nameVariantsById) {
    if (variants.size < 2) continue;
    const displayNames = [...variants.values()].map(
      (variant) => `${variant.display} (${[...variant.sources].join(", ")})`,
    );
    const relevant =
      personStillInSh(personId, "") ||
      personStillInEjoos(personId, "") ||
      activeArrivalIds.has(personId) ||
      openAbsentIds.has(personId);
    if (!relevant) continue;
    const surnames = new Set(
      [...variants.keys()].map((key) => key.split(" ")[0]).filter(Boolean),
    );
    const differentPeople = surnames.size > 1;
    ops.push({
      id: opId(["data-mismatch", personId]),
      kind: "data_mismatch",
      class: "needs_input",
      sheet: "Дані джерел",
      personId,
      fullName:
        shPersonById.get(personId)?.fullName ||
        oosPersonById.get(personId)?.fullName ||
        shpoPersonById.get(personId)?.fullName ||
        "",
      rank: shPersonById.get(personId)?.rank || "",
      positionIndex: shPersonById.get(personId)?.positionIndex || "",
      before: differentPeople
        ? `ID ${personId}: один ID у різних осіб`
        : `ID ${personId}: різне написання ПІБ`,
      after: displayNames.join(" · "),
      sourceRef: `ID ${personId}`,
      why: differentPeople
        ? "Один ID зустрічається у різних прізвищ — перевірте, де ID вказано помилково. Жодних змін у ЕЖООС ця позначка не виконує."
        : "Один ID має різне написання ПІБ у джерелах. Людину зв'язано по ID; написання треба уніфікувати вручну. Жодних змін у ЕЖООС ця позначка не виконує.",
      confidence: "manual",
      payload: {
        type: "DATA_MISMATCH",
        personId,
        mismatchKind: differentPeople ? "ID_COLLISION" : "NAME_SPELLING",
        nameVariants: displayNames.join(" | "),
      },
      checkedDefault: false,
    });
  }

  const rankVariantsById = new Map<
    string,
    Map<string, { display: string; sources: Set<string> }>
  >();
  const addRankVariant = (personId: string, rank: string, source: string) => {
    const id = normId(personId);
    const display = norm(rank);
    const key = normKey(display);
    if (!id || !key) return;
    const variants =
      rankVariantsById.get(id) ??
      new Map<string, { display: string; sources: Set<string> }>();
    const variant = variants.get(key) ?? { display, sources: new Set() };
    variant.sources.add(source);
    variants.set(key, variant);
    rankVariantsById.set(id, variants);
  };
  for (const person of shPeople) {
    addRankVariant(person.personId, person.rank, "sh");
  }
  const latestRankByPerson = new Map<string, string>();
  for (const event of activeMovementsAll) {
    if (!event.personId || !event.rank) continue;
    latestRankByPerson.set(event.personId, event.rank);
  }
  for (const [personId, rank] of latestRankByPerson) {
    addRankVariant(personId, rank, "Рух");
  }
  for (const row of ejoosShpo) addRankVariant(row.personId, row.rank, "ШПО");
  for (const row of ejoosOos) addRankVariant(row.personId, row.rank, "ООС");

  for (const [personId, variants] of rankVariantsById) {
    if (variants.size < 2) continue;
    const relevant =
      personStillInSh(personId, "") ||
      personStillInEjoos(personId, "") ||
      activeArrivalIds.has(personId) ||
      openAbsentIds.has(personId);
    if (!relevant) continue;
    if (
      ops.some(
        (op) =>
          op.kind === "rank_change" &&
          (op.personId === personId ||
            isSamePerson({ personId, fullName: "" }, op)),
      )
    ) {
      continue;
    }
    const displayRanks = [...variants.values()].map(
      (variant) => `${variant.display} (${[...variant.sources].join(", ")})`,
    );
    ops.push({
      id: opId(["rank-mismatch", personId]),
      kind: "data_mismatch",
      class: "needs_input",
      sheet: "Дані джерел",
      personId,
      fullName:
        shPersonById.get(personId)?.fullName ||
        oosPersonById.get(personId)?.fullName ||
        shpoPersonById.get(personId)?.fullName ||
        "",
      rank: shPersonById.get(personId)?.rank || "",
      positionIndex: shPersonById.get(personId)?.positionIndex || "",
      before: `ID ${personId}: різне звання`,
      after: displayRanks.join(" · "),
      sourceRef: `ID ${personId}`,
      why: "Звання в sh/Рух не збігається зі званням в ООС/ШПО. Автоматично з sh не підставляємо — перевірте наказ про присвоєння. Жодних змін у ЕЖООС ця позначка не виконує.",
      confidence: "manual",
      payload: {
        type: "DATA_MISMATCH",
        personId,
        mismatchKind: "RANK",
        rankVariants: displayRanks.join(" | "),
      },
      checkedDefault: false,
    });
  }

  // Фінальна звірка: актуальна sh — джерело правди для зайнятості штатного індексу.
  for (const [positionIndex, shPerson] of shOccupantByIndex) {
    const inTempArrivals = Boolean(
      (shPerson.personId && arrivalById.get(shPerson.personId)) ||
        byPersonName(arrivalByName, shPerson.personId, shPerson.fullName),
    );
    if (
      alreadyVacatedForAbsence(shPerson.personId, shPerson.fullName) &&
      !inTempArrivals
    ) {
      continue;
    }
    const shpo = shpoByIndex.get(positionIndex);
    const shpoMatches =
      Boolean(shpo) &&
      isSamePerson(shPerson, shpo!) &&
      normKey(shpo?.fullName || "") === normKey(shPerson.fullName);
    if (shpoMatches) continue;
    const hasShpoOp = ops.some(
      (op) =>
        (op.kind === "shpo_occupant" || op.kind === "position_change") &&
        isSamePerson(shPerson, op) &&
        op.positionIndex === positionIndex &&
        op.class === "ready",
    );
    if (hasShpoOp) continue;
    const ts =
      staffIndexTimesheetForPerson(
        shPerson.personId,
        shPerson.fullName,
        positionIndex,
      ) || dayByIndex.get(positionIndex);
    const returnEvent = [...activeMovementsAll]
      .filter(
        (event) =>
          isSamePerson(shPerson, event) &&
          eventInLeadWindow(event) &&
          isOwnUnitStaffMove(event) &&
          (event.nextIndex === positionIndex ||
            String(event.changeText || "").includes(positionIndex)),
      )
      .sort(
        (left, right) =>
          movementEventTime(right) - movementEventTime(left) ||
          right.excelRow - left.excelRow,
      )[0];
    ops.push({
      id: opId([
        "shpo-reconcile",
        positionIndex,
        shPerson.personId || shPerson.fullName,
      ]),
      kind: "shpo_occupant",
      class: "ready",
      sheet: "1. ШПО / 6. Табель",
      personId: shPerson.personId,
      fullName: shPerson.fullName,
      positionIndex,
      rank: shPerson.rank,
      before: shpo?.fullName
        ? `${shpo.rank || "?"} ${shpo.fullName} (ID ${shpo.personId || "—"})`
        : "вакантно",
      after: `${shPerson.rank || "?"} ${shPerson.fullName} (ID ${shPerson.personId || "—"})`,
      sourceRef: `sh!R${shPerson.excelRow} · фінальна зайнятість ${positionIndex}`,
      why: "Після всіх рухів на цьому індексі актуальна sh визначає поточного військовослужбовця. Відновлюємо ШПО/Табель за sh, історію попередніх осіб не чіпаємо.",
      confidence: "high",
      payload: {
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
        nextName: shPerson.fullName,
        nextRank: shPerson.rank,
        nextPersonId: shPerson.personId,
        reconcileFromSh: "1",
        timesheetActiveFrom: returnEvent?.orderDate || "",
        timesheetAbsenceSpans: encodeTimesheetAbsenceSpans(
          augustAbsenceSpansFor(shPerson.personId, shPerson.fullName),
        ),
      },
      checkedDefault: true,
    });
  }

  for (const person of shPeople) {
    if (!personStillInSh(person.personId, person.fullName)) continue;
    if (mapStatus(person.status).ruleId !== "absent_archive") continue;
    const foundInArchive = archiveAll.some((period) =>
      isSamePerson(person, period),
    );
    if (foundInArchive) continue;
    ops.push({
      id: opId(["archive_missing", person.personId || person.fullName]),
      kind: "data_mismatch",
      class: "needs_input",
      sheet: "Дані джерел / archive",
      personId: person.personId,
      fullName: person.fullName,
      positionIndex: person.positionIndex,
      rank: person.rank,
      before: person.status || "ВІДСУТНІЙ в АРХІВІ",
      after: "ARCHIVE_REFERENCE_MISSING → перевірити archive",
      sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
      why: "sh каже «ВІДСУТНІЙ в АРХІВІ», але в archive немає запису за ПІБ чи ID. ЛІК / ВІД / СЗЧ / БЕЗВІСТИ з цього статусу не вигадуємо.",
      confidence: "review",
      payload: {
        type: "ARCHIVE_REFERENCE_MISSING",
        mismatchKind: "ARCHIVE_REFERENCE_MISSING",
        statusRaw: person.status,
      },
      checkedDefault: false,
    });
  }

  const summary = {
    ready: ops.filter((op) => op.class === "ready").length,
    needsInput: ops.filter((op) => op.class === "needs_input").length,
    conflict: ops.filter((op) => op.class === "conflict").length,
  };

  return {
    ejoosName: ejoos.fileName,
    pbName: pb.fileName,
    timesheetDay,
    timesheetDayLabel,
    ops,
    summary,
    limitsNote:
      `РУХ: перевірено всі ${movementsAll.length} рядків, активних ` +
      `${activeMovementsAll.length}, останніх кадрових подій ` +
      `${effectiveMovements.length}` +
      (cancelledExternalTransferRows.size
        ? `, скасованих переведень ${cancelledExternalTransferRows.size}`
        : "") +
      `; archive — ${archive.length} періодів цього місяця ` +
      `з ${archiveAll.length}.` +
      (timesheetMonthHeaderMismatch && timesheetMonthHeader
        ? ` Заголовок «6. Табель» зараз «${timesheetHeaderCell?.matched}», має бути «${timesheetMonthHeader}» — оновимо при застосуванні; коди пишемо лише за ${timesheetMonthHeader.replace(" р.", "")}.`
        : ""),
  };
};

export const buildConfirmSummary = (ops: EjoosSyncOp[]) => {
  const byKind: Record<string, number> = {};
  const bySheet: Record<string, number> = {};
  ops.forEach((op) => {
    byKind[op.kind] = (byKind[op.kind] ?? 0) + 1;
    bySheet[op.sheet] = (bySheet[op.sheet] ?? 0) + 1;
  });
  const kindLabels: Record<string, string> = {
    timesheet_day: "Табель (день)",
    shpo_occupant: "ШПО / зайнятість",
    absent_upsert: "Тимч. відсутні (новий/оновлення)",
    absent_close: "Тимч. відсутні (закриття)",
    exclude_transfer: "Виключені",
    move_to_disposition: "Переміщення у розпорядження",
    data_mismatch: "Помилка даних (ПІБ / ID / звання)",
    position_change: "Зміна посади",
    rank_change: "Присвоєння звання",
    arrival: "Прибуття",
    other_manual: "Інше",
  };
  return {
    total: ops.length,
    byKind: Object.entries(byKind).map(([kind, count]) => ({
      kind,
      label: kindLabels[kind] || kind,
      count,
    })),
    bySheet: Object.entries(bySheet).map(([sheet, count]) => ({ sheet, count })),
    names: ops.map((op) => op.fullName || op.personId || "—").filter(Boolean),
  };
};

export const buildProtocolText = (
  plan: EjoosSyncPlan,
  applied: EjoosSyncOp[],
  meta: { version?: number; actor?: string; at?: string },
) => {
  const confirm = buildConfirmSummary(applied);
  const lines = [
    `Протокол змін ЕЖООС`,
    `ЕЖООС: ${plan.ejoosName}`,
    `1ПБ: ${plan.pbName}`,
    `День табеля: ${plan.timesheetDayLabel} (день ${plan.timesheetDay})`,
    `Версія після застосування: ${meta.version ?? "—"}`,
    `Хто: ${meta.actor ?? "—"}`,
    `Коли: ${meta.at ?? new Date().toLocaleString("uk-UA")}`,
    `Застосовано змін: ${applied.length}`,
    ``,
    `Підсумок за типом:`,
    ...confirm.byKind.map((item) => `  - ${item.label}: ${item.count}`),
    ``,
    `Підсумок за аркушем:`,
    ...confirm.bySheet.map((item) => `  - ${item.sheet}: ${item.count}`),
    ``,
    `Деталі:`,
    ...applied.map(
      (op, index) =>
        `${index + 1}. [${op.sheet}] ${op.fullName || "—"} (ID ${op.personId || "—"}, індекс ${op.positionIndex || "—"})\n` +
        `   Було: ${op.before}\n` +
        `   Стане: ${op.after}\n` +
        `   Джерело: ${op.sourceRef}\n` +
        `   Чому: ${op.why}`,
    ),
  ];
  return lines.join("\n");
};

export type { EjoosTimesheetCode };
