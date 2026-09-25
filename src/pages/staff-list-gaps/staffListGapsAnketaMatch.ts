import {
  anketaNameKeyVariants,
  listAnketaNameLookupKeys,
} from "../anketa-data/anketaPersonMatch";
import {
  createEmptyAnketaRow,
  type AnketaRow,
} from "../anketa-data/anketaSheet";
import { normalizePersonBirthKey } from "../personnel/personnelUtils";
import { looksLikePersonName } from "../soc-passport/socPassportFields";
import type { StaffListRow } from "./staffListGapsParse";

export type AnketaRowLookup = {
  byName: Map<string, AnketaRow[]>;
};

export const buildAnketaRowLookup = (rows: AnketaRow[]): AnketaRowLookup => {
  const byName = new Map<string, AnketaRow[]>();
  for (const row of rows) {
    for (const key of listAnketaNameLookupKeys(row.fullName, row.birthDate)) {
      for (const variant of anketaNameKeyVariants(key)) {
        const list = byName.get(variant) ?? [];
        list.push(row);
        byName.set(variant, list);
      }
    }
  }
  return { byName };
};

const dedupeAnketaRows = (list: AnketaRow[]) =>
  [...new Map(list.map((row) => [row.__rowId, row])).values()];

const pickUniqueAnketaMatch = (list: AnketaRow[]) => {
  const unique = dedupeAnketaRows(list);
  if (!unique.length) return null;
  if (unique.length === 1) return unique[0]!;
  const withData = unique.filter((row) =>
    [row.rnokpp, row.birthPlace, row.relatives, row.education].some((value) =>
      String(value ?? "").trim(),
    ),
  );
  if (withData.length === 1) return withData[0]!;
  return null;
};

const disambiguateAnketaByBirth = (list: AnketaRow[], birthDate: string) => {
  const unique = dedupeAnketaRows(list);
  if (unique.length <= 1) return unique[0] ?? null;
  const birthKey = normalizePersonBirthKey(birthDate);
  if (!birthKey) return pickUniqueAnketaMatch(unique);
  const byBirth = unique.filter(
    (row) => normalizePersonBirthKey(row.birthDate) === birthKey,
  );
  if (byBirth.length === 1) return byBirth[0]!;
  return pickUniqueAnketaMatch(byBirth.length ? byBirth : unique);
};

export const matchAnketaRowForStaffList = (
  lookup: AnketaRowLookup,
  pib: string,
  birthDate = "",
): AnketaRow | null => {
  if (!looksLikePersonName(pib)) return null;
  const keys = listAnketaNameLookupKeys(pib, birthDate);
  let bestCandidates: AnketaRow[] = [];
  for (const key of keys) {
    for (const variant of anketaNameKeyVariants(key)) {
      const list = lookup.byName.get(variant) ?? [];
      if (list.length) bestCandidates.push(...list);
      const uniqueMatch = disambiguateAnketaByBirth(list, birthDate);
      if (uniqueMatch) return uniqueMatch;
    }
  }
  if (bestCandidates.length) {
    return disambiguateAnketaByBirth(bestCandidates, birthDate);
  }
  return null;
};

/** Мінімальний рядок анкети з Excel-списку — для картки через особовий склад. */
export const buildStaffListSyntheticAnketaRow = (
  staffRow: StaffListRow,
): AnketaRow => {
  const row = createEmptyAnketaRow(staffRow.__rowNumber);
  row.__rowId = `staff-list:${staffRow.__rowId}`;
  row.fullName = staffRow.pib;
  row.birthDate = staffRow.birthDate;
  return row;
};

export const resolveStaffListAnketaRow = (
  lookup: AnketaRowLookup,
  staffRow: StaffListRow,
): AnketaRow =>
  matchAnketaRowForStaffList(lookup, staffRow.pib, staffRow.birthDate) ??
  buildStaffListSyntheticAnketaRow(staffRow);

export const isStaffListSyntheticAnketaRow = (row: AnketaRow) =>
  row.__rowId.startsWith("staff-list:");
