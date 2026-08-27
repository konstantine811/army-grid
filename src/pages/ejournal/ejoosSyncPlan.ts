import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { mapPbStatusToEjoosWithRules, readOperatorSettings } from "./ejoosStatusMap";
import type { EjoosTimesheetCode } from "./ejoosRules";

export type EjoosOpClass = "ready" | "needs_input" | "conflict";

export type EjoosOpKind =
  | "timesheet_day"
  | "shpo_occupant"
  | "absent_upsert"
  | "absent_close"
  | "exclude_transfer"
  | "position_change"
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
  changeText: string;
  status: string;
  note: string;
  /** Колонка «Звідки» / «Звідки прибув», якщо є в Рух. */
  arrivedFrom: string;
};

/** «Куди» у Рух часто = «x» / порожньо; реальне місце — у примітці. */
const resolveMovementDestination = (rawDest: string, note: string) => {
  const clean = (value: string) => {
    const text = value.replace(/\s+/g, " ").trim();
    if (!text) return "";
    const upper = text.toUpperCase();
    if (
      upper === "X" ||
      upper === "Х" ||
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
  return clean(rawDest) || clean(note);
};

export const parsePbShPeople = (workbook: ExcelWorkbookSnapshot): PbShPerson[] => {
  const sheet = findSheet(workbook, /^sh$/i);
  if (!sheet) return [];
  const headers = headerMap(sheet.rawRows[0] ?? []);
  const idCol = findCol(headers, /^id$/);
  const nameCol = findCol(headers, /^піб$/, /прізвище/);
  const rankCol = findCol(headers, /^зван/);
  const indexCol = findCol(headers, /індекс\s*посад/);
  const posCol = findCol(headers, /^посада$/);
  const statusCol = findCol(headers, /^статус$/);
  const fromCol = findCol(headers, /звідки.*прибув|^звідки$/);
  const people: PbShPerson[] = [];

  sheet.rawRows.slice(1).forEach((row, offset) => {
    const fullName = cell(row, nameCol);
    const positionIndex = cell(row, indexCol);
    const status = cell(row, statusCol);
    const personId = cell(row, idCol);
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
  const returnCol = findCol(
    headers,
    /фактичн.*поверн|дата поверн|повернув|до якої дати|дата прибутт/,
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
      personId: cell(row, idCol),
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
  const rankCol = findCol(headers, /^зван/);
  const statusCol = findCol(headers, /^статус$/);
  const prevIdxCol = findCol(headers, /індекс.*попер|попер.*індекс/);
  const nextIdxCol = findCol(headers, /індекс.*як|яка зміна.*індекс|індекси посад \(яка/);
  // Fallback columns when layout differs (older file had ID in H, type in E)
  const changeCol = findCol(headers, /яка зміна/, /попер/);
  const destCol = findCol(headers, /^куди$/);
  const fromCol = findCol(headers, /звідки.*прибув|^звідки$/);
  const noteCol = findCol(headers, /^примітка$/);
  const orderNumCol = findCol(headers, /^наказ$/);
  const orderDateCol = findCol(headers, /^дата$/);
  const movements: PbMovement[] = [];

  const normalizeType = (value: string) => {
    const text = value.toUpperCase();
    if (text.includes("ПОСАД")) return "ПОСАДА";
    if (text.includes("РОЗПОР")) return "РОЗПОРЯДЖ";
    if (text.includes("ПЕРЕВ")) return "ПЕРЕВ";
    if (text.includes("ПРИБ")) return "ПРИБУВ";
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
      // У новому Рух ID у колонці D (індекс 3); старий layout — H (7).
      personId: cell(row, idCol >= 0 ? idCol : 3),
      fullName,
      rank: cell(row, rankCol >= 0 ? rankCol : 5),
      previousIndex: cell(row, prevIdxCol),
      nextIndex: cell(row, nextIdxCol),
      destination: resolveMovementDestination(rawDest, note),
      orderNumber: cell(row, orderNumCol),
      orderDate: cell(row, orderDateCol),
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

export type EjoosShpoRow = {
  excelRow: number;
  positionIndex: string;
  personId: string;
  fullName: string;
  rank: string;
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
      personId: norm(row?.[2]),
      fullName,
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
      personId: norm(row?.[7]),
      fullName,
      rank: norm(row?.[5]),
      positionIndex,
      dayValue: norm(row?.[dayCol]),
    });
  }
  return rows;
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
      personId: norm(row?.[7]),
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

const opId = (parts: string[]) =>
  parts
    .map((part) => part.replace(/\s+/g, "_").slice(0, 40))
    .filter(Boolean)
    .join("__");

export const buildEjoosSyncPlan = (
  ejoos: ExcelWorkbookSnapshot,
  pb: ExcelWorkbookSnapshot,
  options?: { statusRules?: import("./ejoosRules").EjoosStatusRule[] },
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
  // Рух накопичує тисячі рядків — для ПОСАДА/ПРИБУВ беремо хвіст.
  // Але ПЕРЕВ «вибули з sh, ще є в ШПО» підхоплюємо по всьому Рух (інакше
  // люди на кшталт БАСОВСЬКОГО ~300 рядків від кінця ніколи не потрапляють у план).
  const MOVEMENT_TAIL = 200;
  const movementsTail = movementsAll.slice(-MOVEMENT_TAIL);

  const absentSheet = findSheet(ejoos, /тимчасов.*відсут/i);
  const timesheetSheet = findSheet(ejoos, /табель/i);
  const shpoSheet = findSheet(ejoos, /шпо|штатно.?посад/i);
  const ejoosAbsents = parseEjoosAbsents(absentSheet);
  const ejoosDays = parseEjoosTimesheetDay(timesheetSheet, timesheetDay);
  const ejoosShpo = parseEjoosShpo(shpoSheet);

  const shIds = new Set(
    shPeople.map((person) => person.personId).filter(Boolean),
  );
  const shNames = new Set(
    shPeople.map((person) => normKey(person.fullName)).filter(Boolean),
  );
  const ejoosOccupied = ejoosShpo.filter((row) => row.fullName || row.personId);
  const ejoosIds = new Set(
    ejoosOccupied.map((row) => row.personId).filter(Boolean),
  );
  const ejoosNames = new Set(
    ejoosOccupied.map((row) => normKey(row.fullName)).filter(Boolean),
  );

  const personStillInEjoos = (personId: string, fullName: string) =>
    Boolean(
      (personId && ejoosIds.has(personId)) ||
        (fullName && ejoosNames.has(normKey(fullName))),
    );
  const personStillInSh = (personId: string, fullName: string) =>
    Boolean(
      (personId && shIds.has(personId)) ||
        (fullName && shNames.has(normKey(fullName))),
    );

  /** Останній ПЕРЕВ/РОЗПОРЯДЖ для тих, хто ще в ШПО, але вже немає в sh. */
  const outboundTransfers = new Map<string, (typeof movementsAll)[number]>();
  for (const event of movementsAll) {
    if (event.type !== "ПЕРЕВ" && event.type !== "РОЗПОРЯДЖ") continue;
    if (!personStillInEjoos(event.personId, event.fullName)) continue;
    if (personStillInSh(event.personId, event.fullName)) continue;
    const key = event.personId
      ? `id:${event.personId}`
      : `name:${normKey(event.fullName)}`;
    if (!key || key.endsWith(":")) continue;
    outboundTransfers.set(key, event);
  }

  const movementRows = new Set(movementsTail.map((event) => event.excelRow));
  const movements = [...movementsTail];
  for (const event of outboundTransfers.values()) {
    if (!movementRows.has(event.excelRow)) {
      movements.push(event);
      movementRows.add(event.excelRow);
    }
  }

  const dayByIndex = new Map(ejoosDays.map((row) => [row.positionIndex, row]));
  const dayById = new Map(
    ejoosDays.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const shpoByIndex = new Map(ejoosShpo.map((row) => [row.positionIndex, row]));
  const openAbsents = ejoosAbsents.filter((row) => !row.actualReturn);
  const openById = new Map(
    openAbsents.filter((row) => row.personId).map((row) => [row.personId, row]),
  );
  const openByName = new Map(
    openAbsents.map((row) => [normKey(row.fullName), row]),
  );

  // Archive: лише люди з sh, які зараз не «в строю» (інакше тисячі історичних періодів).
  const absentFocusKeys = new Set<string>();
  shPeople.forEach((person) => {
    const mapped = mapStatus(person.status);
    if (mapped.timesheetCode && mapped.timesheetCode !== "+") {
      if (person.personId) absentFocusKeys.add(person.personId);
      if (person.fullName) absentFocusKeys.add(normKey(person.fullName));
    }
  });
  const archive = archiveAll.filter((period) => {
    if (period.personId && absentFocusKeys.has(period.personId)) return true;
    if (period.fullName && absentFocusKeys.has(normKey(period.fullName))) return true;
    return false;
  });

  const ops: EjoosSyncOp[] = [];

  // SHPO / Табель occupant identity from sh (by position index)
  shPeople.forEach((person) => {
    if (!person.positionIndex || !person.fullName) return;
    const shpo = shpoByIndex.get(person.positionIndex);
    const ts = dayByIndex.get(person.positionIndex);
    if (!shpo && !ts) return;

    const beforeName = shpo?.fullName || ts?.fullName || "—";
    const beforeRank = shpo?.rank || ts?.rank || "—";
    const beforeId = shpo?.personId || ts?.personId || "—";
    const afterName = person.fullName;
    const afterRank = person.rank || "—";
    const afterId = person.personId || "—";

    const nameChanged = normKey(beforeName) !== normKey(afterName);
    const rankChanged = normKey(beforeRank) !== normKey(afterRank);
    const idChanged =
      Boolean(person.personId) &&
      beforeId !== "—" &&
      beforeId !== person.personId;

    if (!nameChanged && !rankChanged && !idChanged) return;

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
      why: "Зайнятість посади в ЕЖООС відрізняється від sh — оновити ПІБ/звання/ID",
      confidence: "high",
      payload: {
        shpoExcelRow: shpo ? String(shpo.excelRow) : "",
        timesheetExcelRow: ts ? String(ts.excelRow) : "",
        nextName: afterName,
        nextRank: afterRank,
        nextPersonId: person.personId,
      },
      checkedDefault: true,
    });
  });

  shPeople.forEach((person) => {
    if (!person.status || person.status === "0") return;
    const mapped = mapStatus(person.status);
    const timesheetRow =
      (person.positionIndex && dayByIndex.get(person.positionIndex)) ||
      (person.personId && dayById.get(person.personId)) ||
      null;
    const before = timesheetRow?.dayValue || "—";
    const afterCode = mapped.timesheetCode;

    if (afterCode) {
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
    } else if (mapped.confidence !== "high") {
      ops.push({
        id: opId(["ts_manual", person.personId || person.fullName, person.status]),
        kind: "timesheet_day",
        class: "needs_input",
        sheet: "6. Табель",
        personId: person.personId,
        fullName: person.fullName,
        positionIndex: person.positionIndex,
        rank: person.rank,
        before,
        after: "(оберіть код)",
        sourceRef: `sh!R${person.excelRow} СТАТУС=«${person.status}»`,
        why: mapped.reason,
        confidence: "manual",
        payload: {
          day: String(timesheetDay),
          excelRow: timesheetRow ? String(timesheetRow.excelRow) : "",
          statusRaw: person.status,
          timesheetCode: "",
        },
        checkedDefault: false,
      });
    }

    // Close open absence when back on duty
    if (mapped.timesheetCode === "+") {
      const open =
        (person.personId && openById.get(person.personId)) ||
        openByName.get(normKey(person.fullName));
      if (open) {
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
          after: `фактичне прибуття: ${timesheetDayLabel}`,
          sourceRef: `sh!R${person.excelRow} → В СТРОЮ; ЕЖООС sheet5 R${open.excelRow}`,
          why: "У 1ПБ знову «в строю», а в «Тимч. відсутні» період ще відкритий",
          confidence: "high",
          payload: {
            excelRow: String(open.excelRow),
            returnDate: timesheetDayLabel,
          },
          checkedDefault: true,
        });
      }
    }
  });

  archive.forEach((period) => {
    const open =
      (period.personId && openById.get(period.personId)) ||
      openByName.get(normKey(period.fullName));
    const complete = Boolean(period.absenceType && (period.departDate || period.plannedReturn));
    const before = open
      ? `${open.ground} / ${open.place} / ${open.departDate}`
      : "(немає відкритого періоду)";
    const after = `${period.absenceType || "?"} → ${period.place || "?"} з ${period.departDate || "?"}`;
    if (open && before === after) return;

    ops.push({
      id: opId(["absent_up", period.personId || period.fullName, period.periodNumber || String(period.excelRow)]),
      kind: "absent_upsert",
      class: complete ? "ready" : "needs_input",
      sheet: "5. Тимчасово відсутні",
      personId: period.personId,
      fullName: period.fullName,
      positionIndex: "",
      rank: period.rank,
      before,
      after,
      sourceRef: `archive!R${period.excelRow} №${period.periodNumber || "—"}`,
      why: complete
        ? "Період з archive пропонується внести/оновити у «Тимч. відсутні»"
        : "В archive неповні поля (дата/підстава) — дозаповніть перед застосуванням",
      confidence: complete ? "high" : "manual",
      payload: {
        absenceType: period.absenceType,
        place: period.place,
        departDate: period.departDate,
        orderNumber: period.orderNumber,
        orderDate: period.orderDate,
        plannedReturn: period.plannedReturn || "?",
        periodNumber: period.periodNumber,
        positionTitle: period.positionTitle,
        existingExcelRow: open ? String(open.excelRow) : "",
      },
      checkedDefault: complete && !open,
    });
  });

  movements.forEach((event) => {
    if (event.type === "ПЕРЕВ" || event.type === "РОЗПОРЯДЖ") {
      const destinationRaw = event.destination || "";
      const destinationUpper = destinationRaw.toUpperCase();
      const destination =
        !destinationRaw ||
        destinationUpper.includes("РОЗПОР") ||
        destinationUpper.includes("ПЕРЕВ") ||
        destinationUpper === event.type
          ? ""
          : destinationRaw;
      const hasDest = Boolean(destination);
      const shpo =
        (event.personId &&
          ejoosShpo.find((row) => row.personId === event.personId)) ||
        (event.previousIndex && shpoByIndex.get(event.previousIndex)) ||
        ejoosShpo.find((row) => normKey(row.fullName) === normKey(event.fullName)) ||
        null;
      const ts =
        (event.previousIndex && dayByIndex.get(event.previousIndex)) ||
        (event.personId && dayById.get(event.personId)) ||
        ejoosDays.find((row) => normKey(row.fullName) === normKey(event.fullName)) ||
        null;
      const fromRank = shpo?.rank || event.rank || "";
      const fromName = shpo?.fullName || event.fullName || "";
      const fromId = shpo?.personId || event.personId || "";
      const fromIndex =
        shpo?.positionIndex || event.previousIndex || event.nextIndex || "";
      const excludeDate = event.orderDate || timesheetDayLabel;

      ops.push({
        id: opId(["excl", event.movementNumber || String(event.excelRow), event.type]),
        kind: "exclude_transfer",
        class: hasDest && Boolean(shpo || fromName) ? "ready" : "needs_input",
        sheet: "Виключені → Табель → ШПО/ООС",
        personId: fromId,
        fullName: fromName,
        positionIndex: fromIndex,
        rank: fromRank,
        before: shpo
          ? `ШПО R${shpo.excelRow}: ${fromRank} ${fromName} · інд. ${fromIndex}`
          : "в обліку (ШПО не знайдено — перевірте вручну)",
        after: `виключити: ${event.type} → ${destination || "(куди?)"} · дата ${excludeDate}`,
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: hasDest
          ? "ПЕРЕВ/РОЗПОРЯДЖ з «куди»: алгоритм Виключені → Табель (історія) → очистка ШПО/ООС. Не чіпаємо Тимч. відсутні/прибулі."
          : "Немає «куди» у Рух — вкажіть місце вибуття вручну в картці",
        confidence: hasDest && shpo ? "high" : "manual",
        payload: {
          movementNumber: event.movementNumber,
          type: event.type,
          destination,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
          excludeDate,
          previousIndex: event.previousIndex,
          nextIndex: event.nextIndex,
          shpoExcelRow: shpo ? String(shpo.excelRow) : "",
          timesheetExcelRow: ts ? String(ts.excelRow) : "",
          fromRank,
          fromName,
          fromPersonId: fromId,
          fromPositionIndex: fromIndex,
          documentsDest: destination,
        },
        checkedDefault: hasDest && Boolean(shpo),
      });
      return;
    }

    if (event.type === "ПОСАДА") {
      ops.push({
        id: opId(["pos", event.movementNumber || String(event.excelRow)]),
        kind: "position_change",
        class: "needs_input",
        sheet: "1. ШПО / 2. ООС",
        personId: event.personId,
        fullName: event.fullName,
        positionIndex: event.nextIndex || event.previousIndex,
        rank: event.rank,
        before: event.previousIndex || "?",
        after: event.nextIndex || event.changeText || "?",
        sourceRef: `Рух!R${event.excelRow} №${event.movementNumber}`,
        why: "Зміна посади завжди потребує підтвердження індексів",
        confidence: "manual",
        payload: {
          movementNumber: event.movementNumber,
          previousIndex: event.previousIndex,
          nextIndex: event.nextIndex,
          changeText: event.changeText,
        },
        checkedDefault: false,
      });
      return;
    }

    if (event.type === "ПРИБУВ" || event.type === "ЗВІЛЬН") {
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
        why: `${event.type} не застосовується автоматично — заповніть накази/звідки/куди`,
        confidence: "manual",
        payload: {
          movementNumber: event.movementNumber,
          type: event.type,
          destination: event.destination,
          orderNumber: event.orderNumber,
          orderDate: event.orderDate,
        },
        checkedDefault: false,
      });
    }
  });

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
      movementsAll.length > MOVEMENT_TAIL ||
      archiveAll.length > archive.length ||
      outboundTransfers.size > 0
        ? `Для швидкості: Рух — останні ${Math.min(MOVEMENT_TAIL, movementsAll.length)} з ${movementsAll.length}` +
          (outboundTransfers.size
            ? ` + ${outboundTransfers.size} ПЕРЕВ/РОЗПОРЯДЖ (є в ШПО, немає в sh)`
            : "") +
          `; archive — ${archive.length} актуальних з ${archiveAll.length} (лише відсутні зараз у sh).`
        : undefined,
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
    position_change: "Зміна посади",
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
