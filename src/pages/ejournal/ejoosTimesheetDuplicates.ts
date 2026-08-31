export type TimesheetDuplicateScan = {
  excelRow: number;
  personId: string;
  fullName: string;
  positionIndex: string;
  hasDepartureText: boolean;
  plusDays: number[];
};

const nameKey = (value: string) =>
  value
    .toLocaleLowerCase("uk-UA")
    .replace(/[''`´]/g, "")
    .replace(/\s+/g, " ")
    .trim();

const personKeyOf = (row: TimesheetDuplicateScan) =>
  String(row.personId || "").trim() || nameKey(row.fullName);

/** Закритий історичний рядок: є «вибув», немає чинних «+». */
export const isClosedTimesheetHistoryRow = (row: TimesheetDuplicateScan) =>
  Boolean(row.hasDepartureText) && row.plusDays.length === 0;

export const isNamedTimesheetRow = (row: TimesheetDuplicateScan) =>
  Boolean(row.personId || row.fullName);

/**
 * Канонічний рядок індексу / особи: збіг з sh, більше «+», інакше менший номер
 * рядка (шаблонний штатний, не дописаний знизу).
 */
export const pickCanonicalTimesheetRow = (
  rows: TimesheetDuplicateScan[],
  currentOccupant?: { personId?: string; fullName?: string } | null,
) => {
  if (!rows.length) return null;
  const occupantId = String(currentOccupant?.personId || "").trim();
  const occupantName = nameKey(currentOccupant?.fullName || "");
  const score = (row: TimesheetDuplicateScan) => {
    let value = row.plusDays.length * 10;
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
  reason: "same_person" | "same_index";
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
  ) => {
    const keep = pickCanonicalTimesheetRow(group, occupant);
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
    takeExtras(group, "same_person", occupant);
  }

  const byIndex = new Map<string, TimesheetDuplicateScan[]>();
  for (const row of rows) {
    if (!row.positionIndex || !isNamedTimesheetRow(row)) continue;
    if (isClosedTimesheetHistoryRow(row)) continue;
    const list = byIndex.get(row.positionIndex) ?? [];
    list.push(row);
    byIndex.set(row.positionIndex, list);
  }
  for (const [index, group] of byIndex) {
    if (group.length < 2) continue;
    takeExtras(group, "same_index", occupantByIndex?.get(index) ?? null);
  }

  return extras;
};
