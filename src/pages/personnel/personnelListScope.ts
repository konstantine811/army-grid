import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import type { PersonnelListIndex } from "./personnelListIndex";
import {
  isPersonnelFromArchive,
  isPersonnelInStaffRoster,
} from "./personnelRosterMerge";

export type PersonnelListScope = {
  all: boolean;
  archive: boolean;
};

export const filterPersonnelScopeRows = (
  rows: EjournalPreviewRow[],
  scope: PersonnelListScope,
) => {
  if (scope.all) return rows;
  return rows.filter((row) => {
    if (isPersonnelFromArchive(row)) return scope.archive;
    return isPersonnelInStaffRoster(row);
  });
};

export const sourceHasPeopleOutsideStaff = (rows: EjournalPreviewRow[]) =>
  rows.some(
    (row) => !isPersonnelInStaffRoster(row) && !isPersonnelFromArchive(row),
  );

export const withClosedScopeCounts = (
  index: PersonnelListIndex,
  scope: PersonnelListScope,
) => {
  if (!scope.all) index.staffCounts.all = -1;
  if (!scope.archive) index.staffCounts.archive = -1;
  return index;
};

/** «Усі» довантажує повний набір, «Архів» лише відкриває вже завантажені рядки. */
export const staffFilterScopeChange = (
  value: "all" | "in" | "archive",
  sourceRows: EjournalPreviewRow[],
) => {
  if (value === "all") {
    const hasFullList = sourceHasPeopleOutsideStaff(sourceRows);
    return {
      all: true,
      archive: true,
      publish: hasFullList,
      reload: !hasFullList,
    };
  }
  if (value === "archive") {
    return { all: false, archive: true, publish: true, reload: false };
  }
  return { all: false, archive: false, publish: false, reload: false };
};
