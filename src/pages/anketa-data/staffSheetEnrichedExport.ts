import { downloadBlob, sanitizeFileName } from "../../shared/browserExport";
import {
  buildStaffSheetEnrichmentReport,
  formatStaffSheetEnrichmentReport,
  type StaffSheetEnrichmentReport,
} from "./staffSheetEnrichment";
import { loadStaffSheetExportContext } from "./staffSheetEnrichmentContext";
import { loadStaffSheetVkIndex } from "./staffSheetMilitaryIdMerge";
import {
  runStaffSheetEnrichmentHeavy,
  runStaffSheetExportWorkbookHeavy,
} from "./runStaffSheetHeavyJobs";

export const downloadEnrichedStaffSheetExcel = async (options?: {
  onProgress?: (phase: string) => void;
}): Promise<StaffSheetEnrichmentReport> => {
  options?.onProgress?.("Завантаження «Штатки» з Google…");
  const [context, vkIndex] = await Promise.all([
    loadStaffSheetExportContext(),
    loadStaffSheetVkIndex(),
  ]);

  options?.onProgress?.("Зіставлення з анкетами…");
  const entries = await runStaffSheetEnrichmentHeavy({
    rosterRows: context.rosterRows,
    personnelIndex: context.personnelIndex,
    anketaRows: context.anketaRows,
    questionnaires: context.questionnaires,
    vkIndex,
  });

  if (!context.rosterRows.length) {
    throw new Error("Немає рядків «Штатки» для експорту.");
  }

  options?.onProgress?.("Формування Excel…");
  const buffer = await runStaffSheetExportWorkbookHeavy(
    context.rosterRows,
    entries,
  );
  const stamp = new Date().toISOString().slice(0, 10);
  const fileName = sanitizeFileName(`Штатка для Google ${stamp}.xlsx`);
  downloadBlob(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName,
  );

  return buildStaffSheetEnrichmentReport(entries, context.rosterRows);
};

export { formatStaffSheetEnrichmentReport };
