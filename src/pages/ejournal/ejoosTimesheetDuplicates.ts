import { isInternalStaffTimesheetDeparture } from "./ejoosTimesheetText";

export type TimesheetDuplicateScan = {
  excelRow: number;
  personId: string;
  fullName: string;
  positionIndex: string;
  hasDepartureText: boolean;
  plusDays: number[];
  firstDepartureDay?: number;
  departureText?: string;
};

const nameKey = (value: string) =>
  value
    .toLocaleLowerCase("uk-UA")
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

const personKeyOf = (row: TimesheetDuplicateScan) => {
  const name = nameKey(row.fullName);
  if (name) return `n:${name}`;
  const id = String(row.personId || "").trim();
  return id ? `id:${id}` : "";
};

const isStaffPositionIndex = (value: string) =>
  /^\d{5,}$/.test(String(value || "").trim());

/**
 * Закритий історичний рядок: є «вибув», і немає «+» після дня вибуття.
 * Плюси до вибуття (01–20 + , 21 вибув) — нормальна історія, не дубль.
 */
export const isClosedTimesheetHistoryRow = (row: TimesheetDuplicateScan) => {
  if (!row.hasDepartureText) return false;
  const departDay = Number(row.firstDepartureDay || 0);
  if (departDay > 0) {
    return !row.plusDays.some((day) => day > departDay);
  }
  return row.plusDays.length === 0;
};

export const isNamedTimesheetRow = (row: TimesheetDuplicateScan) =>
  Boolean(row.personId || row.fullName);

/**
 * Канонічний рядок індексу / особи: збіг з sh, більше «+», інакше менший номер
 * рядка (шаблонний штатний, не дописаний знизу).
 */
export const pickCanonicalTimesheetRow = (
  rows: TimesheetDuplicateScan[],
  currentOccupant?: { personId?: string; fullName?: string } | null,
  occupantIndex?: string,
) => {
  if (!rows.length) return null;
  const occupantId = String(currentOccupant?.personId || "").trim();
  const occupantName = nameKey(currentOccupant?.fullName || "");
  const staffIndex = String(occupantIndex || "").trim();
  const score = (row: TimesheetDuplicateScan) => {
    const lastPlus = row.plusDays.length ? Math.max(...row.plusDays) : 0;
    let value = row.plusDays.length * 10 + lastPlus * 8;
    if (isClosedTimesheetHistoryRow(row)) value -= 5000;
    if (staffIndex && row.positionIndex === staffIndex) value += 2500;
    if (occupantId && row.personId === occupantId) value += 1000;
    else if (occupantName && nameKey(row.fullName) === occupantName) value += 800;
    value -= row.excelRow;
    return value;
  };
  return [...rows].sort((left, right) => score(right) - score(left))[0];
};

export type TimesheetDuplicateExtra = {
  keep: TimesheetDuplicateScan;
  extra: TimesheetDuplicateScan;
  reason: "same_person" | "same_index" | "leftover_history" | "internal_hop";
};

/** Після вибуття лишаємо один іменний рядок — зайву копію з ПІБ прибираємо. */
export const pickLeftoverTimesheetKeep = (rows: TimesheetDuplicateScan[]) => {
  if (!rows.length) return null;
  const score = (row: TimesheetDuplicateScan) => {
    let value = 0;
    if (isClosedTimesheetHistoryRow(row)) value += 200;
    if (row.hasDepartureText) value += 80;
    value += row.plusDays.length;
    value -= row.excelRow / 1000;
    return value;
  };
  return [...rows].sort((left, right) => score(right) - score(left))[0];
};

const isFreshActiveEpisode = (
  row: TimesheetDuplicateScan,
  history?: TimesheetDuplicateScan,
) => {
  if (row.hasDepartureText || isClosedTimesheetHistoryRow(row)) return false;
  if (!row.plusDays.length) return false;
  const departDay = Number(history?.firstDepartureDay || 0);
  if (departDay > 0) return row.plusDays.some((day) => day > departDay);
  return true;
};

/** Другий активний запис тієї ж особи або той самий штатний індекс з іншим ПІБ. */
export const findDuplicateTimesheetExtras = (
  rows: TimesheetDuplicateScan[],
  occupantByIndex?: Map<string, { personId?: string; fullName?: string }>,
): TimesheetDuplicateExtra[] => {
  const extras: TimesheetDuplicateExtra[] = [];
  const seen = new Set<number>();

  const takeExtras = (
    group: TimesheetDuplicateScan[],
    reason: TimesheetDuplicateExtra["reason"],
    occupant?: { personId?: string; fullName?: string } | null,
    occupantIndex?: string,
  ) => {
    const keep = pickCanonicalTimesheetRow(group, occupant, occupantIndex);
    if (!keep) return;
    for (const row of group) {
      if (row.excelRow === keep.excelRow || seen.has(row.excelRow)) continue;
      seen.add(row.excelRow);
      extras.push({ keep, extra: row, reason });
    }
  };

  const byPerson = new Map<string, TimesheetDuplicateScan[]>();
  for (const row of rows) {
    if (!isNamedTimesheetRow(row) || isClosedTimesheetHistoryRow(row)) continue;
    const key = personKeyOf(row);
    if (!key) continue;
    const list = byPerson.get(key) ?? [];
    list.push(row);
    byPerson.set(key, list);
  }
  const occupants = [...(occupantByIndex?.values() ?? [])];
  for (const group of byPerson.values()) {
    if (group.length < 2) continue;
    const sample = group[0];
    const occupant =
      occupants.find(
        (item) =>
          (sample.personId && item.personId === sample.personId) ||
          (sample.fullName &&
            nameKey(item.fullName || "") === nameKey(sample.fullName)),
      ) ?? occupantByIndex?.get(sample.positionIndex) ?? null;
    const occupantIndex = [...(occupantByIndex?.entries() ?? [])].find(
      ([, item]) =>
        (sample.personId && item.personId === sample.personId) ||
        (sample.fullName &&
          nameKey(item.fullName || "") === nameKey(sample.fullName)),
    )?.[0];
    takeExtras(group, "same_person", occupant, occupantIndex);
  }

  const matchesPerson = (
    row: TimesheetDuplicateScan,
    person?: { personId?: string; fullName?: string } | null,
  ) => {
    if (!person) return false;
    const id = String(person.personId || "").trim();
    if (id && row.personId) return id === row.personId;
    return Boolean(
      person.fullName && nameKey(row.fullName) === nameKey(person.fullName),
    );
  };

  const byIndex = new Map<string, TimesheetDuplicateScan[]>();
  for (const row of rows) {
    if (!isStaffPositionIndex(row.positionIndex) || !isNamedTimesheetRow(row)) continue;
    if (isClosedTimesheetHistoryRow(row)) continue;
    const list = byIndex.get(row.positionIndex) ?? [];
    list.push(row);
    byIndex.set(row.positionIndex, list);
  }
  for (const [index, group] of byIndex) {
    if (group.length < 2) continue;
    const occupant = occupantByIndex?.get(index) ?? null;
    const keep = pickCanonicalTimesheetRow(group, occupant, index);
    if (!keep) continue;
    for (const row of group) {
      if (row.excelRow === keep.excelRow || seen.has(row.excelRow)) continue;
      // СЗЧ / розпорядження лишає власний рядок Табеля на старому індексі —
      // це не дубль штатного occupant.
      if (
        personKeyOf(row) !== personKeyOf(keep) &&
        (!occupant || !matchesPerson(row, occupant))
      ) {
        continue;
      }
      seen.add(row.excelRow);
      extras.push({ keep, extra: row, reason: "same_index" });
    }
  }

  const leftoverByPerson = new Map<string, TimesheetDuplicateScan[]>();
  for (const row of rows) {
    if (!isNamedTimesheetRow(row)) continue;
    const key = personKeyOf(row);
    if (!key) continue;
    const list = leftoverByPerson.get(key) ?? [];
    list.push(row);
    leftoverByPerson.set(key, list);
  }
  for (const group of leftoverByPerson.values()) {
    if (group.length < 2) continue;
    const historyRows = group.filter(isClosedTimesheetHistoryRow);
    const activeRows = group.filter(
      (row) => !isClosedTimesheetHistoryRow(row),
    );
    if (!historyRows.length || !activeRows.length) continue;
    const occupant =
      occupants.find((item) => {
        if (
          group[0].fullName &&
          nameKey(item.fullName || "") === nameKey(group[0].fullName)
        ) {
          return true;
        }
        return Boolean(
          group[0].personId && item.personId === group[0].personId,
        );
      }) ?? null;
    if (!occupant) continue;
    const keep = pickCanonicalTimesheetRow(activeRows, occupant);
    if (!keep) continue;
    for (const row of historyRows) {
      if (seen.has(row.excelRow)) continue;
      if (!isInternalStaffTimesheetDeparture(row.departureText || "")) continue;
      seen.add(row.excelRow);
      extras.push({ keep, extra: row, reason: "internal_hop" });
    }
  }
  for (const group of leftoverByPerson.values()) {
    if (group.length < 2) continue;
    const history = group.find((row) => isClosedTimesheetHistoryRow(row));
    const occupant =
      occupants.find((item) => {
        if (
          group[0].fullName &&
          nameKey(item.fullName || "") === nameKey(group[0].fullName)
        ) {
          return true;
        }
        return Boolean(
          !group[0].fullName &&
            group[0].personId &&
            item.personId === group[0].personId,
        );
      }) ?? null;
    if (occupant) continue;
    if (group.some((row) => isFreshActiveEpisode(row, history))) continue;
    const keep = pickLeftoverTimesheetKeep(group);
    if (!keep) continue;
    for (const row of group) {
      if (row.excelRow === keep.excelRow || seen.has(row.excelRow)) continue;
      seen.add(row.excelRow);
      extras.push({ keep, extra: row, reason: "leftover_history" });
    }
  }

  return extras;
};
