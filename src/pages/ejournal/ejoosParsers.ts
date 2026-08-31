import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { canonicalName, isJournalPersonId, normId, normKey, usablePersonId } from "./ejoosIdentity";
import {
  resolveMovementDestination,
  resolveOutboundTransferDestination,
} from "./ejoosMovementRules";
import { isTimesheetDepartureMark } from "./ejoosTimesheetText";

export const norm = (value: CellValue | unknown): string => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const day = String(value.getDate()).padStart(2, "0");
    const month = String(value.getMonth() + 1).padStart(2, "0");
    return `${day}.${month}.${value.getFullYear()}`;
  }
  if (typeof value === "number" && Number.isFinite(value) && value > 20000 && value < 80000) {
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

const idCell = (row: CellValue[] | undefined, index: number) =>
  index >= 0 ? normId(row?.[index]) : "";

const journalIdCell = (row: CellValue[] | undefined, index: number) => {
  const value = idCell(row, index);
  return isJournalPersonId(value) ? value : "";
};

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
  /** «30 діб» / кількість днів зі стовпця «Строк», якщо є. */
  duration: string;
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
  /** 1-based day -> нормалізована позначка (або «вибув»). */
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
  rank: string;
  positionIndex: string;
  orderNumber: string;
  orderDate: string;
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
  const durationCol = findCol(headers, /^строк/, /строк\s*\(?діб/, /кількість\s*діб/);
  const returnCol = findCol(
    headers,
    /^дата прибуття$/,
    /фактичн.*(?:поверн|прибут)/,
    /^дата повернення$/,
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
      duration: cell(row, durationCol),
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
      personId: idCell(row, idCol),
      fullName,
      rank: cell(row, rankCol),
      previousIndex: cell(row, prevIdxCol),
      nextIndex: cell(row, nextIdxCol),
      destination:
        type === "ПЕРЕВ"
          ? resolveOutboundTransferDestination(rawDest, note)
          : resolveMovementDestination(rawDest, note),
      orderNumber: cell(row, orderNumCol),
      orderDate: cell(row, orderDateCol),
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
      rank: norm(row?.[0]),
      positionIndex: norm(row?.[3]) || norm(row?.[4]),
      orderDate: norm(row?.[28]) || norm(row?.[10]),
      orderNumber: norm(row?.[29]) || norm(row?.[11]),
    });
  }
  return rows;
};
