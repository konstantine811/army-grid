import { downloadBlob, sanitizeFileName } from "../../shared/browserExport";
import { downloadStaffSheetXlsxBytes } from "../excel-fill/staffSheet";
import {
  buildStaffSheetEnrichmentReport,
  formatStaffSheetAnketaVkReport,
  type StaffSheetEnrichmentReport,
} from "./staffSheetEnrichment";
import { loadStaffSheetEnrichmentContext } from "./staffSheetEnrichmentContext";
import { loadStaffSheetVkIndex } from "./staffSheetMilitaryIdMerge";
import { loadStaffSheetImport } from "./staffSheetImport";
import {
  runStaffSheetEnrichmentHeavy,
  runStaffSheetAnketaVkOverlayHeavy,
} from "./runStaffSheetHeavyJobs";

const resolveStaffSheetBaseWorkbook = async () => {
  const imported = await loadStaffSheetImport();
  if (imported?.fileData?.byteLength) {
    return {
      fileData: imported.fileData,
      fileName: imported.fileName || "Штатка.xlsx",
      source: "import" as const,
    };
  }

  const downloaded = await downloadStaffSheetXlsxBytes();
  return {
    fileData: downloaded.fileData,
    fileName: downloaded.fileName,
    source: "google" as const,
  };
};

/** Доповнює імпортовану «Штатку»: лише колонки «Анкета» та «Військовий квиток». */
export const downloadEnrichedStaffSheetExcel = async (options?: {
  onProgress?: (phase: string) => void;
}): Promise<StaffSheetEnrichmentReport> => {
  options?.onProgress?.("Завантаження «Штатки»…");
  const [context, vkIndex, baseWorkbook] = await Promise.all([
    loadStaffSheetEnrichmentContext(),
    loadStaffSheetVkIndex(),
    resolveStaffSheetBaseWorkbook(),
  ]);

  options?.onProgress?.("Зіставлення з анкетами…");
  const entries = await runStaffSheetEnrichmentHeavy({
    rosterRows: context.rosterRows,
    personnelIndex: context.personnelIndex,
    anketaRows: context.anketaRows,
    questionnaires: context.questionnaires,
    vkIndex,
  });

  if (!entries.length) {
    throw new Error("Немає рядків з ПІБ для доповнення «Штатки».");
  }

  options?.onProgress?.("Запис «Анкета» та «Військовий квиток»…");
  const stamp = new Date().toISOString().slice(0, 10);
  const buffer = await runStaffSheetAnketaVkOverlayHeavy(
    entries,
    baseWorkbook.fileData,
  );
  const fileName = sanitizeFileName(
    baseWorkbook.fileName.replace(/\.(xlsx|xlsm)$/i, "") +
      ` Анкета+ВК ${stamp}.xlsx`,
  );
  downloadBlob(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName,
  );

  return buildStaffSheetEnrichmentReport(entries, context.rosterRows);
};

export { formatStaffSheetAnketaVkReport as formatStaffSheetEnrichmentReport };
