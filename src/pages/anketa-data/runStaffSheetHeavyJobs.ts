import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { runHeavyJob } from "../../workers/runHeavyJob";
import {
  toStaffSheetEnrichmentJobPayload,
  vkEntriesToIndex,
  type StaffSheetEnrichmentWorkerInput,
} from "../../workers/staffSheetWorkerSerialize";
import type { StaffSheetEnrichmentEntry } from "./staffSheetEnrichment";
import type { StaffSheetRosterImportPayload } from "../excel-fill/staffSheet";
import type { VkTpvDovidkyNameEntry, VkTpvDovidkyRecord } from "../personnel/vkTpvDovidkyImport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { loadStaffSheetExportTemplate } from "./staffSheetExportWorkbook";

export const runStaffSheetEnrichmentHeavy = async (
  input: StaffSheetEnrichmentWorkerInput,
): Promise<StaffSheetEnrichmentEntry[]> =>
  runHeavyJob({
    type: "staffSheetEnrichment",
    ...toStaffSheetEnrichmentJobPayload(input),
  });

export const runStaffSheetVkIndexHeavy = async (
  snapshot: ExcelWorkbookSnapshot,
): Promise<Map<string, VkTpvDovidkyNameEntry>> => {
  const entries = await runHeavyJob({
    type: "staffSheetVkIndex",
    snapshot,
  });
  return vkEntriesToIndex(entries);
};

export const runParseVkTpvDovidkyHeavy = async (
  snapshot: ExcelWorkbookSnapshot,
): Promise<VkTpvDovidkyRecord[]> =>
  runHeavyJob({
    type: "parseVkTpvDovidky",
    snapshot,
  });

export const runStaffSheetRosterImportHeavy = async (
  table: string[][],
  meta: {
    source: "apps-script" | "gviz";
    sourceLabel: string;
    fighterStatusTable?: string[][] | null;
    includeAllRows?: boolean;
  },
): Promise<StaffSheetRosterImportPayload> =>
  runHeavyJob({
    type: "staffSheetRosterImport",
    table,
    meta,
  });

export const runStaffSheetExportWorkbookHeavy = async (
  rosterRows: EjournalPreviewRow[],
  entries: StaffSheetEnrichmentEntry[],
  templateData?: ArrayBuffer,
): Promise<ArrayBuffer> => {
  const template = templateData ?? (await loadStaffSheetExportTemplate());
  return runHeavyJob({
    type: "staffSheetExportWorkbook",
    rosterRows,
    entries,
    templateData: template,
  });
};
