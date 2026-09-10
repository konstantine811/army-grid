import type {
  BackendPersonnelOverview,
  BackendPersonnelOverviewRow,
} from "../../api";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";

/** Parses ISO and Ukrainian dd.MM.yyyy[, HH:mm] used in Overview rows. */
export const parseOverviewCalendarDate = (
  value: string | null | undefined,
): Date | null => {
  const text = String(value ?? "").trim();
  if (!text) return null;

  const ukMatch = text.match(
    /^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:[,\s]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/,
  );
  if (ukMatch) {
    const [, day, month, year, hour = "0", minute = "0", second = "0"] = ukMatch;
    const parsed = new Date(
      Number(year),
      Number(month) - 1,
      Number(day),
      Number(hour),
      Number(minute),
      Number(second),
    );
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const isSameCalendarDay = (value: string | null | undefined, now = new Date()) => {
  const parsed = parseOverviewCalendarDate(value);
  if (!parsed) return false;
  return (
    parsed.getFullYear() === now.getFullYear() &&
    parsed.getMonth() === now.getMonth() &&
    parsed.getDate() === now.getDate()
  );
};

const overviewRowTouchDates = (row: BackendPersonnelOverviewRow) =>
  [
    row.updatedAt,
    row.validFrom,
    row.fighterEntryDate,
    row.fighterExitDate,
    row.fighterReturnDate,
  ].filter(Boolean);

export const overviewRowChangedToday = (
  row: BackendPersonnelOverviewRow,
  now = new Date(),
) => overviewRowTouchDates(row).some((value) => isSameCalendarDay(value, now));

const isReturnOverdue = (plannedReturn: string | null | undefined, now = new Date()) => {
  const text = String(plannedReturn ?? "").trim();
  if (!text) return false;
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return false;
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return parsed < today;
};

const rowLookupKeys = (rows: BackendPersonnelOverviewRow[]) => {
  const keys = new Set<string>();
  for (const row of rows) {
    if (row.id) keys.add(row.id);
    if (row.externalId) keys.add(row.externalId);
    const name = normalizeRosterMatchText(row.name);
    if (name) keys.add(name);
  }
  return keys;
};

export const filterOverviewCriticalForRows = (
  critical: BackendPersonnelOverview["critical"],
  rows: BackendPersonnelOverviewRow[],
) => {
  const keys = rowLookupKeys(rows);
  return critical.filter((item) => {
    if (keys.has(item.id)) return true;
    const name = normalizeRosterMatchText(item.text.split(":")[0] ?? "");
    return Boolean(name && keys.has(name));
  });
};

const isCriticalAbsence = (row: BackendPersonnelOverviewRow) => {
  if (
    row.status === "AWOL" ||
    row.status === "MISSING" ||
    row.status === "CAPTIVITY"
  ) {
    return true;
  }
  const label = normalizeRosterMatchText(row.statusLabel);
  if (
    label.includes("сзч") ||
    label.includes("смерт") ||
    label.includes("загиб") ||
    label.includes("безв")
  ) {
    return true;
  }
  if (isReturnOverdue(row.plannedReturn)) return true;
  return row.days != null && row.days >= 30;
};

const formatCriticalText = (row: BackendPersonnelOverviewRow) => {
  const parts = [row.statusLabel || row.status];
  if (row.days != null) parts.push(`${row.days} дн.`);
  if (isReturnOverdue(row.plannedReturn)) parts.push("строк повернення минув");
  return `${row.name}: ${parts.join(" · ")}`;
};

export const buildOverviewCriticalFromRows = (
  rows: BackendPersonnelOverviewRow[],
): BackendPersonnelOverview["critical"] =>
  rows
    .filter(isCriticalAbsence)
    .sort((left, right) => (right.days ?? 0) - (left.days ?? 0))
    .slice(0, 50)
    .map((row) => ({
      id: row.id,
      severity:
        row.status === "AWOL" ||
        row.status === "MISSING" ||
        row.status === "CAPTIVITY"
          ? "danger"
          : "warning",
      text: formatCriticalText(row),
      status: String(row.status),
      days: row.days,
      daysToReturn: null,
    }));

export const buildOverviewTodayStats = (
  rows: BackendPersonnelOverviewRow[],
  now = new Date(),
) => {
  const todayRows = rows.filter((row) => overviewRowChangedToday(row, now));
  const count = (status: string) =>
    todayRows.filter((row) => row.status === status).length;

  return {
    todayChanges: {
      total: todayRows.length,
      onDuty: count("ON_DUTY"),
      businessTrip: count("BUSINESS_TRIP"),
      leave: count("LEAVE"),
      medical: count("MEDICAL"),
      awol: count("AWOL"),
      other: todayRows.filter(
        (row) =>
          !["ON_DUTY", "BUSINESS_TRIP", "LEAVE", "MEDICAL", "AWOL"].includes(
            String(row.status),
          ),
      ).length,
    },
    todayUpdates: todayRows.length,
  };
};

export const buildOverviewSideStats = (
  rows: BackendPersonnelOverviewRow[],
  overview: BackendPersonnelOverview | null,
) => {
  const fromBackend = overview
    ? filterOverviewCriticalForRows(overview.critical, rows)
    : [];
  const covered = new Set(
    fromBackend.flatMap((item) =>
      [item.id, normalizeRosterMatchText(item.text.split(":")[0] ?? "")].filter(
        Boolean,
      ),
    ),
  );
  const extra = buildOverviewCriticalFromRows(rows).filter(
    (item) =>
      !covered.has(item.id) &&
      !covered.has(normalizeRosterMatchText(item.text.split(":")[0] ?? "")),
  );
  return {
    critical: [...fromBackend, ...extra].sort(
      (left, right) => (right.days ?? 0) - (left.days ?? 0),
    ),
    ...buildOverviewTodayStats(rows),
  };
};
