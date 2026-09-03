import { api, type BackendPersonQuestionnaireMeta } from "../../api";
import { loadSharedRosterLatest } from "../../data/sharedAppData";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { runStaffSheetRosterImportHeavy } from "./runStaffSheetHeavyJobs";
import {
  pullStaffSheetGvizRosterTable,
  STAFF_SHEET_ID,
} from "../excel-fill/staffSheet";
import {
  applyAnketaEditsToRows,
  loadAnketaEdits,
} from "./anketaEdits";
import {
  loadPersonnelIndexForAnketa,
  normalizeAnketaNameKey,
  resolvePersonnelRowForStaffRoster,
  type AnketaPersonnelIndex,
} from "./anketaPersonMatch";
import { loadAnketaSheetPreferCache, type AnketaRow } from "./anketaSheet";
import { getPersonExternalId } from "../personnel/personnelUtils";
import {
  loadStaffSheetImport,
  rosterLatestToStaffSheetImportSnapshot,
  rosterRowsFromStaffSheetPayload,
} from "./staffSheetImport";

const sortRosterRows = (rows: EjournalPreviewRow[]) =>
  [...rows].sort(
    (left, right) =>
      (Number(left.__rowNumber) || 0) - (Number(right.__rowNumber) || 0),
  );

export type StaffSheetEnrichmentContext = {
  /** Усі рядки «Штатки» в порядку Excel. */
  rosterRows: EjournalPreviewRow[];
  personnelIndex: AnketaPersonnelIndex;
  anketaRows: AnketaRow[];
  questionnaires: BackendPersonQuestionnaireMeta[];
  importFileName: string;
};

const loadAnketaContext = async () => {
  const [anketaSnapshot, anketaEdits, questionnaires] = await Promise.all([
    loadAnketaSheetPreferCache(),
    loadAnketaEdits(),
    api.listPersonQuestionnaires().catch(() => []),
  ]);
  return {
    anketaRows: applyAnketaEditsToRows(
      anketaSnapshot?.rows ?? [],
      anketaEdits,
    ),
    questionnaires,
  };
};

/** Повна «Штатка» з Google gviz — усі рядки (підрозділи, вакансії) для експорту. */
export const loadStaffSheetExportRosterRows = async (): Promise<{
  rosterRows: EjournalPreviewRow[];
  importFileName: string;
}> => {
  const table = await pullStaffSheetGvizRosterTable();
  const payload = await runStaffSheetRosterImportHeavy(table, {
    source: "gviz",
    sourceLabel: `Штатка (gviz) · ${STAFF_SHEET_ID}`,
    includeAllRows: true,
  });
  return {
    rosterRows: sortRosterRows(rosterRowsFromStaffSheetPayload(payload)),
    importFileName: "Штатка з Google",
  };
};

export const loadStaffSheetEnrichmentContext =
  async (): Promise<StaffSheetEnrichmentContext> => {
    const [
      latestRoster,
      importedStaffSheet,
      personnelIndex,
      anketaContext,
    ] = await Promise.all([
      loadSharedRosterLatest(),
      loadStaffSheetImport(),
      loadPersonnelIndexForAnketa(),
      loadAnketaContext(),
    ]);

    const fromDb = rosterLatestToStaffSheetImportSnapshot(latestRoster);
    const snapshot = fromDb ?? importedStaffSheet;
    let rosterRows = sortRosterRows(snapshot?.rows ?? []);
    let importFileName = snapshot?.fileName ?? "";
    if (!rosterRows.length) {
      const fresh = await loadStaffSheetExportRosterRows();
      rosterRows = fresh.rosterRows;
      importFileName = fresh.importFileName;
    }
    if (!rosterRows.length) {
      throw new Error(
        "Немає «Штатки». Оновіть з Google Sheets в Особовому складі або перевірте доступ до таблиці.",
      );
    }

    return {
      rosterRows,
      personnelIndex,
      anketaRows: anketaContext.anketaRows,
      questionnaires: anketaContext.questionnaires,
      importFileName,
    };
  };

/** Контекст для експорту: свіжа gviz-«Штатка» з усіма рядками + індекс ООС. */
export const loadStaffSheetExportContext =
  async (): Promise<StaffSheetEnrichmentContext> => {
    const [exportRoster, personnelIndex, anketaContext] = await Promise.all([
      loadStaffSheetExportRosterRows(),
      loadPersonnelIndexForAnketa(),
      loadAnketaContext(),
    ]);
    if (!exportRoster.rosterRows.length) {
      throw new Error("Google «Штатка» порожня або недоступна.");
    }
    return {
      rosterRows: exportRoster.rosterRows,
      personnelIndex,
      anketaRows: anketaContext.anketaRows,
      questionnaires: anketaContext.questionnaires,
      importFileName: exportRoster.importFileName,
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

export const resolveStaffSheetPersonnelRow = (
  rosterRow: EjournalPreviewRow,
  personnelIndex: AnketaPersonnelIndex,
) => resolvePersonnelRowForStaffRoster(rosterRow, personnelIndex);
