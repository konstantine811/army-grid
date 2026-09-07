import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeRosterMatchText } from "./fighterStatusImport";
import {
  buildPersonListSummary,
  collectPersonCallSignFieldValues,
  getPersonFieldValue,
  isLikelyPersonnelRow,
  normalizePersonnelSearchText,
  resolvePersonBirthDate,
  resolvePersonDisplayNameFromRoster,
  type PersonnelRecord,
} from "./personnelUtils";
import {
  getRosterPersonName,
  isPersonnelFromArchive,
  isPersonnelInStaffRoster,
} from "./personnelRosterMerge";

const normalizeRosterText = normalizeRosterMatchText;

export type PersonnelListRecord = PersonnelRecord & {
  inStaff: boolean;
  inArchive: boolean;
  /** Pre-normalized search corpus without per-user phone overrides. */
  searchBase: string;
};

export type PersonnelListIndex = {
  records: PersonnelListRecord[];
  staffCounts: { all: number; in: number; archive: number };
};

const buildSearchBase = (record: PersonnelRecord) =>
  normalizePersonnelSearchText(
    [
      record.summary.name,
      resolvePersonDisplayNameFromRoster(record.row),
      getRosterPersonName(record.row),
      record.summary.callSign,
      getPersonFieldValue(record.row, ["позивний"]),
      getPersonFieldValue(record.row, ["позив"]),
      ...collectPersonCallSignFieldValues(record.row),
      record.summary.rank,
      record.summary.externalId,
      getPersonFieldValue(record.row, ["індекс", "посади"]),
      getPersonFieldValue(record.row, ["місце_дислокації"]),
      getPersonFieldValue(record.row, ["рнокпп_за_наявності"]),
      getPersonFieldValue(record.row, ["додаткова_інформація"]),
      getPersonFieldValue(record.row, ["військового", "квитка"]),
    ]
      .filter(Boolean)
      .join(" "),
  );

/** Build list rows once per dataset load — flags and search text included. */
export const buildPersonnelListIndex = (
  rows: EjournalPreviewRow[],
): PersonnelListIndex => {
  const baseRecords = rows
    .filter(isLikelyPersonnelRow)
    .map((row) => ({ row, summary: buildPersonListSummary(row) }));

  const nameCounts = new Map<string, number>();
  for (const record of baseRecords) {
    const key = normalizeRosterText(record.summary.name);
    if (!key) continue;
    nameCounts.set(key, (nameCounts.get(key) ?? 0) + 1);
  }

  let inStaff = 0;
  let archive = 0;
  const records = baseRecords.map((record) => {
    const key = normalizeRosterText(record.summary.name);
    const withBirth =
      key && (nameCounts.get(key) ?? 0) > 1
        ? {
            ...record,
            summary: {
              ...record.summary,
              birthDate: resolvePersonBirthDate(record.row),
            },
          }
        : record;
    const inStaffRow = isPersonnelInStaffRoster(withBirth.row);
    const inArchiveRow = isPersonnelFromArchive(withBirth.row);
    if (inStaffRow) inStaff += 1;
    if (inArchiveRow) archive += 1;
    return {
      ...withBirth,
      inStaff: inStaffRow,
      inArchive: inArchiveRow,
      searchBase: buildSearchBase(withBirth),
    };
  });

  return {
    records,
    staffCounts: {
      all: records.length,
      in: inStaff,
      archive,
    },
  };
};
