import { readWorkbookSnapshot } from "../../excelRoundTrip";
import type { BchsMorningDailySnapshot } from "./overviewRotaBchsMorningSnapshot";
import { parseBchsMorningBaselineFile } from "./overviewRotaBchsMorningBaselineImport";
import { parseStaffSheetMorningBaselineFile } from "./overviewRotaBchsMorningStaffBaselineImport";

export type MorningComparisonBaselineKind = "bchs" | "staff";

const findRosterSheet = (
  sheets: Awaited<ReturnType<typeof readWorkbookSnapshot>>["sheets"],
) =>
  sheets.find((sheet) => /загальний\s*список/i.test(sheet.sheetName)) ??
  sheets.find((sheet) => /^sh$/i.test(sheet.sheetName.trim()));

export const detectMorningComparisonBaselineKind = async (
  file: File,
): Promise<MorningComparisonBaselineKind> => {
  const snapshot = await readWorkbookSnapshot(file);
  if (findRosterSheet(snapshot.sheets)) return "staff";
  const hasBchsSheet = snapshot.sheets.some(
    (sheet) =>
      /^\d/.test(sheet.sheetName.trim()) || /ПР$/iu.test(sheet.sheetName.trim()),
  );
  if (hasBchsSheet) return "bchs";
  throw new Error(
    "Невідомий формат. Завантажте файл «БЧС (ранковий ПБ)» або «Штатку» (.xlsx).",
  );
};

export const parseMorningComparisonBaselineFile = async (
  file: File,
  unitLabel?: string,
): Promise<BchsMorningDailySnapshot> => {
  const kind = await detectMorningComparisonBaselineKind(file);
  if (kind === "staff") {
    return parseStaffSheetMorningBaselineFile(file, unitLabel?.trim() ?? "");
  }
  const snapshot = await parseBchsMorningBaselineFile(file, unitLabel);
  return {
    ...snapshot,
    baselineKind: "bchs",
  };
};
