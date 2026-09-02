import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  findEjoosSheet,
  parseEjoosAbsents,
  parseEjoosOos,
  parseEjoosShpo,
  parseEjoosTimesheetDay,
  parseEjoosTimesheetPeople,
  parseTimesheetDayFromPbName,
  type EjoosAbsentRow,
  type EjoosShpoRow,
  type EjoosTimesheetPersonScan,
  type EjoosTimesheetRow,
} from "./ejoosSyncPlan";
import {
  findDuplicateOosById,
  findPossibleDuplicateOosByName,
} from "./ejoosOosText";
import {
  findDuplicateTimesheetExtras,
  isClosedTimesheetHistoryRow,
} from "./ejoosTimesheetDuplicates";
import type { BackendEjournalLiveVersion } from "../../api";
import type { EjoosDiffSession } from "./ejoosPersonDiff";
import { formatValueForDisplay } from "../../shared/format";

const norm = (value: CellValue | unknown) => {
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

const normDate = (value: CellValue | unknown) => {
  const formatted = formatValueForDisplay(value);
  return String(formatted ?? "").replace(/\s+/g, " ").trim();
};

const normKey = (value: string) =>
  value
    .toLowerCase()
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

export type EjoosRegisterPerson = {
  key: string;
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  dayCode: string;
  isVacant: boolean;
};

export type EjoosExcludedRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  destination: string;
  orderNumber: string;
  orderDate: string;
  excludeDate: string;
  note: string;
};

export type EjoosArrivalRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  fromUnit: string;
  arriveDate: string;
  note: string;
};

export type EjoosIrrevocableLossRow = {
  excelRow: number;
  personId: string;
  fullName: string;
  rank: string;
  positionIndex: string;
  serviceType: string;
  birthDate: string;
  birthPlace: string;
  recruitedBy: string;
  relatives: string;
  lossType: string;
  lossDate: string;
  circumstances: string;
  lossPlace: string;
};

export type EjoosCheckItem = {
  id: string;
  severity: "ok" | "warn" | "error";
  title: string;
  detail: string;
};

export type EjoosTodayChange = {
  fullName: string;
  before: string;
  after: string;
  sheet: string;
};

export type EjoosLiveView = {
  timesheetDay: number;
  timesheetDayLabel: string;
  roster: EjoosRegisterPerson[];
  shpo: EjoosShpoRow[];
  timesheet: EjoosTimesheetRow[];
  absentsOpen: EjoosAbsentRow[];
  absentsClosed: EjoosAbsentRow[];
  arrivals: EjoosArrivalRow[];
  irrevocableLosses: EjoosIrrevocableLossRow[];
  excluded: EjoosExcludedRow[];
  checks: EjoosCheckItem[];
  todayChanges: EjoosTodayChange[];
  counts: {
    roster: number;
    occupied: number;
    vacant: number;
    absentsOpen: number;
    arrivals: number;
    irrevocableLosses: number;
    excluded: number;
    onDutyToday: number;
  };
};

const personKey = (personId: string, fullName: string, positionIndex: string) => {
  if (personId) return `id:${personId}`;
  if (fullName) return `name:${normKey(fullName)}`;
  return `idx:${positionIndex}`;
};

/** Канонічний шаблон «3. Виключені»: A звання, B ПІБ, C ID, AB–AF вибуття. */
export const parseExcluded = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosExcludedRow[] => {
  if (!sheet) return [];
  const rows: EjoosExcludedRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]);
    const personId = norm(row?.[2]);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]),
      excludeDate: norm(row?.[27]),
      orderDate: norm(row?.[28]),
      orderNumber: norm(row?.[29]),
      destination: norm(row?.[30]),
      note: norm(row?.[31]),
    });
  }
  return rows;
};

const parseArrivals = (sheet: ExcelSheetSnapshot | undefined): EjoosArrivalRow[] => {
  if (!sheet) return [];
  const rows: EjoosArrivalRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]) || norm(row?.[6]);
    const personId = norm(row?.[2]);
    const positionIndex = norm(row?.[4]) || norm(row?.[3]);
    if (!fullName && !personId && !/^\d/.test(positionIndex)) continue;
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]) || norm(row?.[5]),
      positionIndex,
      fromUnit: norm(row?.[5]) || norm(row?.[8]),
      arriveDate: normDate(row?.[7]) || normDate(row?.[10]),
      note: norm(row?.[15]) || norm(row?.[11]) || norm(row?.[12]),
    });
  }
  return rows;
};

const parseIrrevocableLosses = (
  sheet: ExcelSheetSnapshot | undefined,
): EjoosIrrevocableLossRow[] => {
  if (!sheet) return [];
  const rows: EjoosIrrevocableLossRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[1]);
    const personId = norm(row?.[4]);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[0]),
      positionIndex: norm(row?.[2]),
      serviceType: norm(row?.[3]),
      birthDate: normDate(row?.[7]),
      birthPlace: norm(row?.[8]),
      recruitedBy: [norm(row?.[9]), normDate(row?.[10])].filter(Boolean).join(" "),
      relatives: norm(row?.[11]),
      lossType: norm(row?.[12]),
      lossDate: normDate(row?.[13]),
      circumstances: norm(row?.[14]),
      lossPlace: norm(row?.[15]),
    });
  }
  return rows;
};

const resolveDay = (input: {
  asOfDate?: string | null;
  pbFileName?: string | null;
}) => {
  if (input.asOfDate) {
    const match = String(input.asOfDate).match(/(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
    if (match) {
      const day = Number(match[1]);
      const month = Number(match[2]);
      const yearRaw = Number(match[3]);
      const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
      if (day >= 1 && day <= 31) {
        return {
          day,
          label: `${String(day).padStart(2, "0")}.${String(month).padStart(2, "0")}.${year}`,
        };
      }
    }
  }
  if (input.pbFileName) {
    const fromName = parseTimesheetDayFromPbName(input.pbFileName);
    if (!fromName.sourceDateUnknown) return fromName;
  }
  return { day: 0, label: "", sourceDateUnknown: true };
};

const sameRosterPerson = (
  left: { personId?: string; fullName?: string },
  right: { personId?: string; fullName?: string },
) => {
  const leftId = String(left.personId || "").trim();
  const rightId = String(right.personId || "").trim();
  if (leftId && rightId) return leftId === rightId;
  const leftName = normKey(left.fullName || "");
  const rightName = normKey(right.fullName || "");
  return Boolean(leftName && rightName && leftName === rightName);
};

const pickTimesheetDayRow = (
  candidates: EjoosTimesheetRow[],
  occupant: { personId?: string; fullName?: string } | null,
  scanByRow: Map<number, EjoosTimesheetPersonScan>,
) => {
  if (!candidates.length) return null;
  const identity = occupant
    ? candidates.filter((row) => sameRosterPerson(row, occupant))
    : [];
  const pool = identity.length ? identity : occupant ? [] : candidates;
  if (!pool.length) return null;
  const active = pool.filter((row) => {
    const scan = scanByRow.get(row.excelRow);
    return !scan || !isClosedTimesheetHistoryRow(scan);
  });
  return (active[0] ?? pool[0]) ?? null;
};

const applyTimesheetToRoster = (
  entry: EjoosRegisterPerson,
  dayRow: EjoosTimesheetRow | null,
) => {
  if (!dayRow) return;
  entry.dayCode = dayRow.dayValue;
  if (!entry.fullName) entry.fullName = dayRow.fullName;
  if (!entry.rank) entry.rank = dayRow.rank;
  if (!entry.personId) entry.personId = dayRow.personId;
  entry.isVacant = !entry.fullName && !entry.personId;
  entry.key = personKey(entry.personId, entry.fullName, entry.positionIndex);
};

const buildRoster = (
  shpo: EjoosShpoRow[],
  timesheet: EjoosTimesheetRow[],
  timesheetPeople: EjoosTimesheetPersonScan[] = [],
): EjoosRegisterPerson[] => {
  const byIndex = new Map<string, EjoosRegisterPerson>();
  const scanByRow = new Map(
    timesheetPeople.map((row) => [row.excelRow, row]),
  );
  const byPosition = new Map<string, EjoosTimesheetRow[]>();
  for (const row of timesheet) {
    const list = byPosition.get(row.positionIndex) ?? [];
    list.push(row);
    byPosition.set(row.positionIndex, list);
  }
  const used = new Set<number>();

  shpo.forEach((row) => {
    const occupant = row.fullName || row.personId ? row : null;
    const dayRow = pickTimesheetDayRow(
      byPosition.get(row.positionIndex) ?? [],
      occupant,
      scanByRow,
    );
    if (dayRow) used.add(dayRow.excelRow);
    const entry: EjoosRegisterPerson = {
      key: personKey(row.personId, row.fullName, row.positionIndex),
      excelRow: row.excelRow,
      personId: row.personId,
      fullName: row.fullName,
      rank: row.rank,
      positionIndex: row.positionIndex,
      dayCode: "",
      isVacant: !row.fullName && !row.personId,
    };
    applyTimesheetToRoster(entry, dayRow);
    byIndex.set(row.positionIndex, entry);
  });

  const leftoverByIndex = new Map<string, EjoosTimesheetRow[]>();
  for (const row of timesheet) {
    if (used.has(row.excelRow) || byIndex.has(row.positionIndex)) continue;
    const list = leftoverByIndex.get(row.positionIndex) ?? [];
    list.push(row);
    leftoverByIndex.set(row.positionIndex, list);
  }
  for (const [index, rows] of leftoverByIndex) {
    const dayRow = pickTimesheetDayRow(rows, null, scanByRow);
    if (!dayRow) continue;
    byIndex.set(index, {
      key: personKey(dayRow.personId, dayRow.fullName, dayRow.positionIndex),
      excelRow: dayRow.excelRow,
      personId: dayRow.personId,
      fullName: dayRow.fullName,
      rank: dayRow.rank,
      positionIndex: dayRow.positionIndex,
      dayCode: dayRow.dayValue,
      isVacant: !dayRow.fullName && !dayRow.personId,
    });
  }

  return [...byIndex.values()].sort((a, b) =>
    a.positionIndex.localeCompare(b.positionIndex, "uk", { numeric: true }),
  );
};

const buildChecks = (input: {
  hasLive: boolean;
  roster: EjoosRegisterPerson[];
  absentsOpen: EjoosAbsentRow[];
  session: EjoosDiffSession | null;
  duplicateTimesheet?: number;
  duplicateOosById?: Array<Array<{ excelRow: number; personId: string }>>;
  possibleDuplicateOosByName?: Array<
    Array<{ excelRow: number; personId: string; fullName: string }>
  >;
}): EjoosCheckItem[] => {
  const checks: EjoosCheckItem[] = [];

  if (!input.hasLive) {
    checks.push({
      id: "no-live",
      severity: "error",
      title: "Немає канонічного ЕЖООС",
      detail: "Завантажте файл у БД на Головній або через 1ПБ.",
    });
    return checks;
  }

  checks.push({
    id: "live-ok",
    severity: "ok",
    title: "ЕЖООС у БД відкрито",
    detail: "Можна переглядати реєстри та експортувати.",
  });

  const occupiedNoDay = input.roster.filter(
    (row) => !row.isVacant && !row.dayCode,
  );
  if (occupiedNoDay.length) {
    checks.push({
      id: "empty-day",
      severity: "warn",
      title: `Порожній код табеля: ${occupiedNoDay.length}`,
      detail: occupiedNoDay
        .slice(0, 8)
        .map((row) => row.fullName || row.positionIndex)
        .join(", "),
    });
  } else {
    checks.push({
      id: "day-ok",
      severity: "ok",
      title: "Зайняті посади мають код табеля",
      detail: "Для обраного дня немає порожніх кодів у зайнятих рядках.",
    });
  }

  const incompleteAbsents = input.absentsOpen.filter(
    (row) => !row.ground || !row.departDate || !row.place,
  );
  if (incompleteAbsents.length) {
    checks.push({
      id: "absent-incomplete",
      severity: "warn",
      title: `Неповні відкриті відсутності: ${incompleteAbsents.length}`,
      detail: incompleteAbsents
        .slice(0, 8)
        .map((row) => row.fullName || row.personId)
        .join(", "),
    });
  } else {
    checks.push({
      id: "absent-ok",
      severity: "ok",
      title: "Відкриті відсутності заповнені",
      detail: "Підстава, місце й дата вибуття є у відкритих періодах.",
    });
  }

  if (input.duplicateTimesheet) {
    checks.push({
      id: "dup-timesheet",
      severity: "error",
      title: `Дублі Табеля: ${input.duplicateTimesheet}`,
      detail:
        "Одна особа або один штатний індекс не може мати два активні рядки. Перебудуйте операції — зайві копії підуть у чергу на очищення.",
    });
  } else {
    checks.push({
      id: "dup-timesheet-ok",
      severity: "ok",
      title: "Табель без дублів",
      detail: "Немає двох активних рядків на одну особу чи один індекс.",
    });
  }

  const duplicateOosById = input.duplicateOosById ?? [];
  if (duplicateOosById.length) {
    checks.push({
      id: "dup-oos-id",
      severity: "error",
      title: `DUPLICATE_OOS_ID: ${duplicateOosById.length}`,
      detail: duplicateOosById
        .slice(0, 6)
        .map(
          (group) =>
            `ID ${group[0]?.personId}: ${group
              .map((row) => `R${row.excelRow}`)
              .join(", ")}`,
        )
        .join("; "),
    });
  } else {
    checks.push({
      id: "dup-oos-id-ok",
      severity: "ok",
      title: "ООС без дублів ID",
      detail: "Кожен ID має рівно одну активну картку в «2. ООС».",
    });
  }

  const possibleOosNameDups = input.possibleDuplicateOosByName ?? [];
  if (possibleOosNameDups.length) {
    checks.push({
      id: "dup-oos-name",
      severity: "warn",
      title: `POSSIBLE_DUPLICATE_OOS_NAME: ${possibleOosNameDups.length}`,
      detail: possibleOosNameDups
        .slice(0, 6)
        .map(
          (group) =>
            `${group[0]?.fullName}: ${group
              .map((row) => `R${row.excelRow}${row.personId ? ` ID${row.personId}` : " без ID"}`)
              .join(", ")}`,
        )
        .join("; "),
    });
  }

  if (input.session) {
    const queued = input.session.people.filter(
      (person) => person.decision === "accepted",
    ).length;
    const conflicts = input.session.people.filter(
      (person) => person.severity === "conflict",
    ).length;
    if (conflicts) {
      checks.push({
        id: "session-open",
        severity: "error",
        title: `Конфлікти в операціях: ${conflicts}`,
        detail: "Розберіть конфлікти на вкладці «Операції» перед експортом.",
      });
    } else if (queued) {
      checks.push({
        id: "session-queued",
        severity: "warn",
        title: `У черзі до застосування: ${queued}`,
        detail: "Застосуйте чергу на вкладці «Операції → До застосування» або експортуйте поточну версію як є.",
      });
    } else {
      checks.push({
        id: "session-ready",
        severity: "ok",
        title: "Немає змін у черзі застосування",
        detail: "Позначте людей у «Операціях» і застосуйте чергу, або експортуйте поточну версію.",
      });
    }
  } else {
    checks.push({
      id: "no-session",
      severity: "ok",
      title: "Немає активної сесії аналізу 1ПБ",
      detail: "Експорт піде з поточної версії ЕЖООС у БД.",
    });
  }

  return checks;
};

const todayChangesFromVersion = (
  version: BackendEjournalLiveVersion | null | undefined,
): EjoosTodayChange[] => {
  const protocol = version?.changeProtocol as
    | { ops?: Array<Record<string, string>> }
    | null
    | undefined;
  const ops = protocol?.ops;
  if (!Array.isArray(ops)) return [];
  return ops.slice(0, 40).map((op) => ({
    fullName: String(op.fullName || "—"),
    before: String(op.before || "—"),
    after: String(op.after || "—"),
    sheet: String(op.sheet || "—"),
  }));
};

const todayChangesFromSession = (
  session: EjoosDiffSession | null,
): EjoosTodayChange[] => {
  if (!session) return [];
  return session.people
    .filter((person) => person.decision === "accepted")
    .slice(0, 40)
    .map((person) => ({
      fullName: person.fullName,
      before: person.summaryBefore,
      after: person.summaryAfter,
      sheet: person.ejoosWillDo[0] || "—",
    }));
};

export const buildEjoosLiveView = (input: {
  workbook: ExcelWorkbookSnapshot | null;
  asOfDate?: string | null;
  pbFileName?: string | null;
  session?: EjoosDiffSession | null;
  currentVersion?: BackendEjournalLiveVersion | null;
}): EjoosLiveView => {
  const dayInfo = resolveDay({
    asOfDate: input.asOfDate,
    pbFileName: input.pbFileName || input.session?.pbFileName,
  });

  if (!input.workbook) {
    return {
      timesheetDay: dayInfo.day,
      timesheetDayLabel: dayInfo.label,
      roster: [],
      shpo: [],
      timesheet: [],
      absentsOpen: [],
      absentsClosed: [],
      arrivals: [],
      irrevocableLosses: [],
      excluded: [],
      checks: buildChecks({
        hasLive: false,
        roster: [],
        absentsOpen: [],
        session: input.session ?? null,
      }),
      todayChanges: todayChangesFromSession(input.session ?? null).length
        ? todayChangesFromSession(input.session ?? null)
        : todayChangesFromVersion(input.currentVersion),
      counts: {
        roster: 0,
        occupied: 0,
        vacant: 0,
        absentsOpen: 0,
        arrivals: 0,
        irrevocableLosses: 0,
        excluded: 0,
        onDutyToday: 0,
      },
    };
  }

  const absentSheet = findEjoosSheet(input.workbook, /тимчасов.*відсут/i);
  const arrivalSheet = findEjoosSheet(input.workbook, /тимчасов.*прибул/i);
  const irrevocableLossSheet = findEjoosSheet(input.workbook, /безповорот/i);
  const excludedSheet = findEjoosSheet(input.workbook, /виключ/i);
  const timesheetSheet = findEjoosSheet(input.workbook, /табель/i);
  const shpoSheet = findEjoosSheet(input.workbook, /шпо|штатно.?посад/i);
  const oosSheet = findEjoosSheet(input.workbook, /оос/i);

  const absents = parseEjoosAbsents(absentSheet);
  const absentsOpen = absents.filter((row) => !row.actualReturn);
  const absentsClosed = absents.filter((row) => Boolean(row.actualReturn));
  const shpo = parseEjoosShpo(shpoSheet);
  const timesheet = parseEjoosTimesheetDay(timesheetSheet, dayInfo.day);
  const timesheetPeople = parseEjoosTimesheetPeople(timesheetSheet);
  const duplicateTimesheet = findDuplicateTimesheetExtras(timesheetPeople).length;
  const oosPeople = parseEjoosOos(oosSheet);
  const duplicateOosById = findDuplicateOosById(oosPeople);
  const possibleDuplicateOosByName = findPossibleDuplicateOosByName(oosPeople);
  const roster = buildRoster(shpo, timesheet, timesheetPeople);
  const arrivals = parseArrivals(arrivalSheet);
  const irrevocableLosses = parseIrrevocableLosses(irrevocableLossSheet);
  const excluded = parseExcluded(excludedSheet);
  const occupied = roster.filter((row) => !row.isVacant);
  const onDutyToday = occupied.filter((row) => row.dayCode === "+").length;

  const sessionChanges = todayChangesFromSession(input.session ?? null);
  const todayChanges = sessionChanges.length
    ? sessionChanges
    : todayChangesFromVersion(input.currentVersion);

  return {
    timesheetDay: dayInfo.day,
    timesheetDayLabel: dayInfo.label,
    roster,
    shpo,
    timesheet,
    absentsOpen,
    absentsClosed,
    arrivals,
    irrevocableLosses,
    excluded,
    checks: buildChecks({
      hasLive: true,
      roster,
      absentsOpen,
      session: input.session ?? null,
      duplicateTimesheet,
      duplicateOosById,
      possibleDuplicateOosByName,
    }),
    todayChanges,
    counts: {
      roster: roster.length,
      occupied: occupied.length,
      vacant: roster.length - occupied.length,
      absentsOpen: absentsOpen.length,
      arrivals: arrivals.length,
      irrevocableLosses: irrevocableLosses.length,
      excluded: excluded.length,
      onDutyToday,
    },
  };
};

export const filterByQuery = <T extends { fullName?: string; personId?: string; positionIndex?: string }>(
  rows: T[],
  query: string,
) => {
  const q = normKey(query);
  if (!q) return rows;
  return rows.filter((row) => {
    const hay = normKey(
      [row.fullName, row.personId, row.positionIndex].filter(Boolean).join(" "),
    );
    return hay.includes(q);
  });
};

export const findPersonAbsents = (
  absents: EjoosAbsentRow[],
  person: Pick<EjoosRegisterPerson, "personId" | "fullName">,
) => {
  const nameKey = normKey(person.fullName);
  return absents.filter((row) => {
    if (person.personId && row.personId === person.personId) return true;
    if (nameKey && normKey(row.fullName) === nameKey) return true;
    return false;
  });
};
