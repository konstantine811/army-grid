import { api, type BackendPersonQuestionnaireMeta } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  readRosterColumnValue,
} from "../excel-fill/rosterSourceSnapshot";
import {
  applyAnketaEditsToRows,
  loadAnketaEdits,
} from "./anketaEdits";
import { loadPersonnelRowsForAnketa, normalizeAnketaNameKey } from "./anketaPersonMatch";
import { loadAnketaSheetPreferCache, type AnketaRow } from "./anketaSheet";
import { getPersonExternalId } from "../personnel/personnelUtils";
import { loadStaffSheetImport } from "./staffSheetImport";

const sortRosterRows = (rows: EjournalPreviewRow[]) =>
  [...rows].sort(
    (left, right) =>
      (Number(left.__rowNumber) || 0) - (Number(right.__rowNumber) || 0),
  );

export type StaffSheetEnrichmentContext = {
  /** Усі рядки імпортованої «Штатки» в порядку Excel. */
  rosterRows: EjournalPreviewRow[];
  /** ООС + roster для полів, яких немає в штатці. */
  mergedPersonnelRows: EjournalPreviewRow[];
  anketaRows: AnketaRow[];
  questionnaires: BackendPersonQuestionnaireMeta[];
  importFileName: string;
};

export const loadStaffSheetEnrichmentContext =
  async (): Promise<StaffSheetEnrichmentContext> => {
    const [importedStaffSheet, mergedPersonnelRows, anketaSnapshot, anketaEdits, questionnaires] =
      await Promise.all([
        loadStaffSheetImport(),
        loadPersonnelRowsForAnketa(),
        loadAnketaSheetPreferCache(),
        loadAnketaEdits(),
        api.listPersonQuestionnaires().catch(() => []),
      ]);

    const rosterRows = sortRosterRows(importedStaffSheet?.rows ?? []);
    if (!rosterRows.length) {
      throw new Error(
        "Спочатку імпортуйте «Штатку» (.xlsx) кнопкою «Імпорт Штатки» на цій сторінці.",
      );
    }

    const anketaRows = applyAnketaEditsToRows(
      anketaSnapshot?.rows ?? [],
      anketaEdits,
    );

    return {
      rosterRows,
      mergedPersonnelRows,
      anketaRows,
      questionnaires,
      importFileName: importedStaffSheet?.fileName ?? "",
    };
  };

export const findMergedPersonnelRow = (
  rosterRow: EjournalPreviewRow,
  mergedPersonnelRows: EjournalPreviewRow[],
) => {
  const externalId = getPersonExternalId(rosterRow).trim();
  if (externalId) {
    const byId = mergedPersonnelRows.find(
      (row) => getPersonExternalId(row).trim() === externalId,
    );
    if (byId) return byId;
  }

  const nameKey = normalizeAnketaNameKey(readRosterColumnValue(rosterRow, 14));
  if (!nameKey) return rosterRow;

  const byName = mergedPersonnelRows.find(
    (row) =>
      normalizeAnketaNameKey(readRosterColumnValue(row, 14)) === nameKey,
  );
  return byName ?? rosterRow;
};
