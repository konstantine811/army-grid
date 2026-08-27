import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  findEjoosSheet,
  parseEjoosAbsents,
  parseEjoosShpo,
  parseEjoosTimesheetDay,
  parseTimesheetDayFromPbName,
  type EjoosAbsentRow,
  type EjoosShpoRow,
  type EjoosTimesheetRow,
} from "./ejoosSyncPlan";
import type { EjoosDiffSession } from "./ejoosPersonDiff";
import type { BackendEjournalLiveVersion } from "../../api";

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
  excluded: EjoosExcludedRow[];
  checks: EjoosCheckItem[];
  todayChanges: EjoosTodayChange[];
  counts: {
    roster: number;
    occupied: number;
    vacant: number;
    absentsOpen: number;
    arrivals: number;
    excluded: number;
    onDutyToday: number;
  };
};

const personKey = (personId: string, fullName: string, positionIndex: string) => {
  if (personId) return `id:${personId}`;
  if (fullName) return `name:${normKey(fullName)}`;
  return `idx:${positionIndex}`;
};

const parseExcluded = (sheet: ExcelSheetSnapshot | undefined): EjoosExcludedRow[] => {
  if (!sheet) return [];
  const rows: EjoosExcludedRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[6]) || norm(row?.[1]);
    const personId = norm(row?.[7]) || norm(row?.[2]);
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[5]) || norm(row?.[3]),
      destination: norm(row?.[10]) || norm(row?.[8]),
      orderNumber: norm(row?.[11]) || norm(row?.[9]),
      orderDate: norm(row?.[12]) || norm(row?.[10]),
      note: norm(row?.[13]) || norm(row?.[4]),
    });
  }
  return rows;
};

const parseArrivals = (sheet: ExcelSheetSnapshot | undefined): EjoosArrivalRow[] => {
  if (!sheet) return [];
  const rows: EjoosArrivalRow[] = [];
  for (let i = 5; i < sheet.rawRows.length; i += 1) {
    const row = sheet.rawRows[i];
    const fullName = norm(row?.[6]) || norm(row?.[1]);
    const personId = norm(row?.[7]) || norm(row?.[2]);
    const positionIndex = norm(row?.[3]) || norm(row?.[1]);
    if (!fullName && !personId && !/^\d/.test(positionIndex)) continue;
    if (!fullName && !personId) continue;
    rows.push({
      excelRow: i + 1,
      personId,
      fullName,
      rank: norm(row?.[5]) || norm(row?.[4]),
      positionIndex,
      fromUnit: norm(row?.[8]) || norm(row?.[9]),
      arriveDate: norm(row?.[10]) || norm(row?.[6]),
      note: norm(row?.[11]) || norm(row?.[12]),
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
    return parseTimesheetDayFromPbName(input.pbFileName);
  }
  return parseTimesheetDayFromPbName("");
};

const buildRoster = (
  shpo: EjoosShpoRow[],
  timesheet: EjoosTimesheetRow[],
): EjoosRegisterPerson[] => {
  const byIndex = new Map<string, EjoosRegisterPerson>();

  shpo.forEach((row) => {
    byIndex.set(row.positionIndex, {
      key: personKey(row.personId, row.fullName, row.positionIndex),
      excelRow: row.excelRow,
      personId: row.personId,
      fullName: row.fullName,
      rank: row.rank,
      positionIndex: row.positionIndex,
      dayCode: "",
      isVacant: !row.fullName && !row.personId,
    });
  });

  timesheet.forEach((row) => {
    const existing = byIndex.get(row.positionIndex);
    if (existing) {
      existing.dayCode = row.dayValue;
      if (!existing.fullName) existing.fullName = row.fullName;
      if (!existing.rank) existing.rank = row.rank;
      if (!existing.personId) existing.personId = row.personId;
      existing.isVacant = !existing.fullName && !existing.personId;
      existing.key = personKey(
        existing.personId,
        existing.fullName,
        existing.positionIndex,
      );
      return;
    }
    byIndex.set(row.positionIndex, {
      key: personKey(row.personId, row.fullName, row.positionIndex),
      excelRow: row.excelRow,
      personId: row.personId,
      fullName: row.fullName,
      rank: row.rank,
      positionIndex: row.positionIndex,
      dayCode: row.dayValue,
      isVacant: !row.fullName && !row.personId,
    });
  });

  return [...byIndex.values()].sort((a, b) =>
    a.positionIndex.localeCompare(b.positionIndex, "uk", { numeric: true }),
  );
};

const buildChecks = (input: {
  hasLive: boolean;
  roster: EjoosRegisterPerson[];
  absentsOpen: EjoosAbsentRow[];
  session: EjoosDiffSession | null;
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

  if (input.session) {
    const pending = input.session.people.filter(
      (person) => person.decision === "pending",
    ).length;
    const conflicts = input.session.people.filter(
      (person) => person.severity === "conflict",
    ).length;
    if (pending || conflicts) {
      checks.push({
        id: "session-open",
        severity: conflicts ? "error" : "warn",
        title: `Є непідтверджені зміни з 1ПБ: ${pending}`,
        detail: conflicts
          ? `Конфліктів: ${conflicts}. Перейдіть у «Зміни» перед експортом.`
          : "Підтвердіть або відхиліть зміни перед застосуванням.",
      });
    } else {
      checks.push({
        id: "session-ready",
        severity: "ok",
        title: "Сесія 1ПБ без відкритих конфліктів",
        detail: "Можна застосовувати підтверджені або експортувати поточну версію.",
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
        excluded: 0,
        onDutyToday: 0,
      },
    };
  }

  const absentSheet = findEjoosSheet(input.workbook, /тимчасов.*відсут/i);
  const arrivalSheet = findEjoosSheet(input.workbook, /тимчасов.*прибул/i);
  const excludedSheet = findEjoosSheet(input.workbook, /виключ/i);
  const timesheetSheet = findEjoosSheet(input.workbook, /табель/i);
  const shpoSheet = findEjoosSheet(input.workbook, /шпо|штатно.?посад/i);

  const absents = parseEjoosAbsents(absentSheet);
  const absentsOpen = absents.filter((row) => !row.actualReturn);
  const absentsClosed = absents.filter((row) => Boolean(row.actualReturn));
  const shpo = parseEjoosShpo(shpoSheet);
  const timesheet = parseEjoosTimesheetDay(timesheetSheet, dayInfo.day);
  const roster = buildRoster(shpo, timesheet);
  const arrivals = parseArrivals(arrivalSheet);
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
    excluded,
    checks: buildChecks({
      hasLive: true,
      roster,
      absentsOpen,
      session: input.session ?? null,
    }),
    todayChanges,
    counts: {
      roster: roster.length,
      occupied: occupied.length,
      vacant: roster.length - occupied.length,
      absentsOpen: absentsOpen.length,
      arrivals: arrivals.length,
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
