import { readWorkbookSnapshot, type ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  CacheKeys,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { normalizeAnketaNameKey } from "./anketaPersonMatch";
import {
  extractMilitaryIdFromText,
  type VkTpvDovidkyNameEntry,
} from "../personnel/vkTpvDovidkyImport";
import { sanitizeStaffSheetCellValue } from "../excel-fill/staffSheet";
import { runStaffSheetVkIndexHeavy } from "./runStaffSheetHeavyJobs";

export type StaffSheetVkCache = {
  fileName: string;
  importedAt: string;
  entries: VkTpvDovidkyNameEntry[];
};

export type StaffSheetMilitaryIdMergeAction =
  | "kept_staff"
  | "from_vk"
  | "unchanged";

export type StaffSheetMilitaryIdMergeReport = {
  rosterRows: number;
  vkNamesInFile: number;
  fromVk: number;
  keptStaff: number;
  unchanged: number;
  skippedNoName: number;
};

export const resolveStaffSheetMilitaryIdValue = (
  currentValue: string,
  vkEntry: VkTpvDovidkyNameEntry | undefined,
): { value: string; action: StaffSheetMilitaryIdMergeAction } => {
  const current = sanitizeStaffSheetCellValue(String(currentValue ?? ""), 11);
  const currentId = extractMilitaryIdFromText(current);
  const vkId = extractMilitaryIdFromText(String(vkEntry?.militaryId ?? ""));

  if (currentId) {
    return { value: currentId, action: "kept_staff" };
  }

  if (vkId) {
    return { value: vkId, action: "from_vk" };
  }

  return { value: current, action: "unchanged" };
};

export const planStaffSheetMilitaryIdMerges = (
  rosterRows: EjournalPreviewRow[],
  vkIndex: Map<string, VkTpvDovidkyNameEntry>,
) => {
  const report: StaffSheetMilitaryIdMergeReport = {
    rosterRows: rosterRows.length,
    vkNamesInFile: vkIndex.size,
    fromVk: 0,
    keptStaff: 0,
    unchanged: 0,
    skippedNoName: 0,
  };

  for (const row of rosterRows) {
    const fullName = readRosterColumnValue(row, 14).trim();
    if (fullName.length < 3) continue;

    const nameKey = normalizeAnketaNameKey(fullName);
    if (!nameKey) {
      report.skippedNoName += 1;
      continue;
    }

    const resolved = resolveStaffSheetMilitaryIdValue(
      readRosterColumnValue(row, 11),
      vkIndex.get(nameKey),
    );

    switch (resolved.action) {
      case "from_vk":
        report.fromVk += 1;
        break;
      case "kept_staff":
        report.keptStaff += 1;
        break;
      case "unchanged":
        report.unchanged += 1;
        break;
    }
  }

  return report;
};

export const formatStaffSheetMilitaryIdMergeReport = (
  report: StaffSheetMilitaryIdMergeReport,
) => {
  const parts = [
    `ПІБ у файлі ВК: ${report.vkNamesInFile}`,
    report.fromVk ? `ВК додано в Штатку: ${report.fromVk}` : "",
    report.keptStaff ? `ВК уже в Штатці: ${report.keptStaff}` : "",
    report.unchanged ? `без ВК: ${report.unchanged}` : "",
    report.skippedNoName ? `без ПІБ: ${report.skippedNoName}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

export const loadStaffSheetVkIndex = async (): Promise<
  Map<string, VkTpvDovidkyNameEntry>
> => {
  const cached = await readDataCache<StaffSheetVkCache>(
    CacheKeys.staffSheetVkIndex,
  );
  if (!cached?.entries?.length) return new Map();
  return new Map(cached.entries.map((entry) => [entry.nameKey, entry]));
};

export const loadStaffSheetVkCache = async () =>
  readDataCache<StaffSheetVkCache>(CacheKeys.staffSheetVkIndex);

export const saveStaffSheetVkIndex = async (
  fileName: string,
  index: Map<string, VkTpvDovidkyNameEntry>,
) => {
  await writeDataCache<StaffSheetVkCache>(CacheKeys.staffSheetVkIndex, {
    fileName,
    importedAt: new Date().toISOString(),
    entries: [...index.values()],
  });
};

export const importStaffSheetVkWorkbook = async (
  snapshot: ExcelWorkbookSnapshot,
  fileName: string,
  rosterRows: EjournalPreviewRow[],
) => {
  const index = await runStaffSheetVkIndexHeavy(snapshot);
  await saveStaffSheetVkIndex(fileName, index);
  const report = planStaffSheetMilitaryIdMerges(rosterRows, index);
  return { index, report };
};

export const importStaffSheetVkFile = async (
  file: File,
  rosterRows: EjournalPreviewRow[],
) => {
  const snapshot = await readWorkbookSnapshot(file);
  return importStaffSheetVkWorkbook(snapshot, file.name, rosterRows);
};
