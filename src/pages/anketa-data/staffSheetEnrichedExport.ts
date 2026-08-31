import { downloadBlob, sanitizeFileName } from "../../shared/browserExport";
import {
  STAFF_SHEET_ENRICHMENT_COLUMNS,
  STAFF_SHEET_EXPORT_COLUMN_NUMBERS,
  staffSheetBirthDateColumn,
} from "../excel-fill/staffSheet";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  MORNING_GENERAL_LIST_COLUMN_LABELS,
  isLikelyCallSignToken,
  looksLikePersonBirthDate,
} from "../personnel/personnelUtils";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildStaffSheetEnrichmentEntries,
  formatStaffSheetEnrichmentReport,
  type StaffSheetEnrichmentEntry,
  type StaffSheetEnrichmentReport,
} from "./staffSheetEnrichment";
import { loadStaffSheetEnrichmentContext } from "./staffSheetEnrichmentContext";
import {
  loadStaffSheetImport,
  type StaffSheetImportSnapshot,
} from "./staffSheetImport";

const loadXlsxPopulate = async () => {
  const module = await import(
    "xlsx-populate/browser/xlsx-populate-no-encryption"
  );
  return module.default;
};

const isValidBirthDate = (value: string) => {
  const text = value.trim();
  if (!text || isLikelyCallSignToken(text)) return false;
  return looksLikePersonBirthDate(text);
};

const applyEnrichmentToWorkbook = (
  sheet: {
    cell: (row: number, column: number) => {
      value: (value?: string | number) => unknown;
      formula: (value?: string) => unknown;
      style?: (key: string, value: string) => unknown;
    };
  },
  entries: StaffSheetEnrichmentEntry[],
) => {
  for (const entry of entries) {
    const excelRow = entry.excelRowNumber;
    STAFF_SHEET_ENRICHMENT_COLUMNS.forEach((columnNumber, columnIndex) => {
      const value = entry.values[columnIndex] ?? "";
      const cell = sheet.cell(excelRow, columnNumber);

      if (columnNumber === staffSheetBirthDateColumn + 1) {
        if (!value) {
          cell.value("");
          return;
        }
        const year = Number(value);
        cell.value(Number.isFinite(year) ? year : value);
        return;
      }

      if (columnNumber === staffSheetBirthDateColumn + 2) {
        if (!value) {
          cell.value("");
          return;
        }
        const age = Number(value);
        cell.value(Number.isFinite(age) ? age : value);
        return;
      }

      if (columnNumber === 10) {
        cell.value(value === "так" ? "так" : "");
        return;
      }

      if (columnNumber === 19) {
        const inn = value.replace(/\D/g, "");
        cell.value(inn.length === 10 ? inn : "");
        return;
      }

      if (columnNumber === staffSheetBirthDateColumn) {
        const text = value.trim();
        if (
          text &&
          !text.startsWith("=") &&
          looksLikePersonBirthDate(text) &&
          !isLikelyCallSignToken(text)
        ) {
          cell.value(text);
        } else if (!text) {
          cell.value("");
        }
        return;
      }

      cell.value(value);
    });
  }
};

const downloadImportedStaffSheet = async (
  importSnapshot: StaffSheetImportSnapshot,
  entries: StaffSheetEnrichmentEntry[],
) => {
  if (!importSnapshot.fileData) {
    throw new Error("Немає бінарного файлу «Штатки».");
  }
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(importSnapshot.fileData);
  const sheet =
    workbook
      .sheets()
      .find((item: { name: () => string }) =>
        /загальний\s*список/i.test(item.name()),
      ) ?? workbook.sheet(0);

  applyEnrichmentToWorkbook(sheet, entries);

  const stamp = new Date().toISOString().slice(0, 10);
  const baseName = importSnapshot.fileName.replace(/\.(xlsx|xlsm)$/i, "");
  const fileName = sanitizeFileName(`${baseName} доповнено ${stamp}.xlsx`);
  const buffer = await workbook.outputAsync();
  downloadBlob(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName,
  );
};

/** Збірка .xlsx з рядків Google (коли немає вихідного файлу). */
const downloadBuiltStaffSheet = async (
  rosterRows: EjournalPreviewRow[],
  entries: StaffSheetEnrichmentEntry[],
  label: string,
) => {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromBlankAsync();
  const sheet = workbook.sheet(0).name("1.ОС Загальний список");
  const enrichmentByRow = new Map(
    entries.map((entry) => [entry.excelRowNumber, entry.values]),
  );

  // Пишемо в рідні Excel-номери колонок (10, 14, 16…), щоб формули віку працювали.
  STAFF_SHEET_EXPORT_COLUMN_NUMBERS.forEach((columnNumber) => {
    const header =
      MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber] ??
      `Колонка ${columnNumber}`;
    sheet.cell(1, columnNumber).value(header);
  });

  for (const rosterRow of rosterRows) {
    const excelRow = Number(rosterRow.__rowNumber);
    if (!excelRow) continue;
    const enrichRow = enrichmentByRow.get(excelRow) ?? [];

    STAFF_SHEET_EXPORT_COLUMN_NUMBERS.forEach((columnNumber) => {
      const cell = sheet.cell(excelRow, columnNumber);
      const enrichedBirth = (enrichRow[1] ?? "").trim();
      const baseBirth = readRosterColumnValue(
        rosterRow,
        staffSheetBirthDateColumn,
      ).trim();
      const birthDate = [enrichedBirth, baseBirth].find(isValidBirthDate);

      if (columnNumber === staffSheetBirthDateColumn + 1) {
        const yearText = (enrichRow[2] ?? "").trim();
        if (yearText && /^\d{4}$/.test(yearText)) {
          cell.value(Number(yearText));
          return;
        }
        return;
      }
      if (columnNumber === staffSheetBirthDateColumn + 2) {
        const ageText = (enrichRow[3] ?? "").trim();
        if (ageText && /^\d{1,3}$/.test(ageText)) {
          cell.value(Number(ageText));
          return;
        }
        return;
      }
      if (columnNumber === staffSheetBirthDateColumn) {
        if (birthDate) cell.value(birthDate);
        return;
      }
      if (columnNumber === 10) {
        if (enrichRow[0] === "так") cell.value("так");
        return;
      }
      if (columnNumber === 19) {
        const inn = (enrichRow[4] ?? "").replace(/\D/g, "");
        if (inn.length === 10) cell.value(inn);
        return;
      }

      const base = readRosterColumnValue(rosterRow, columnNumber).trim();
      if (!base) return;
      cell.value(base);
    });
  }

  const stamp = new Date().toISOString().slice(0, 10);
  const fileName = sanitizeFileName(
    `${label.replace(/\.(xlsx|xlsm)$/i, "")} доповнено ${stamp}.xlsx`,
  );
  const buffer = await workbook.outputAsync();
  downloadBlob(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName,
  );
};

export const downloadEnrichedStaffSheetExcel = async (options?: {
  onProgress?: (phase: string) => void;
}): Promise<StaffSheetEnrichmentReport> => {
  options?.onProgress?.("Завантаження даних…");
  const context = await loadStaffSheetEnrichmentContext();
  const importSnapshot = await loadStaffSheetImport();

  const entries = buildStaffSheetEnrichmentEntries({
    rosterRows: context.rosterRows,
    mergedPersonnelRows: context.mergedPersonnelRows,
    anketaRows: context.anketaRows,
    questionnaires: context.questionnaires,
  });

  if (!entries.length) {
    throw new Error("Немає рядків з ПІБ для експорту «Штатки».");
  }

  if (!importSnapshot) {
    throw new Error(
      "Немає «Штатки». Імпортуйте .xlsx кнопкою «Імпорт Штатки» на Анкетних даних.",
    );
  }

  options?.onProgress?.("Доповнення Excel…");
  if (importSnapshot.fileData) {
    await downloadImportedStaffSheet(importSnapshot, entries);
  } else {
    await downloadBuiltStaffSheet(
      context.rosterRows,
      entries,
      importSnapshot.fileName || "Штатка",
    );
  }

  return {
    rows: entries.length,
    anketaYes: entries.filter((entry) => entry.values[0] === "так").length,
    withBirthDate: entries.filter((entry) =>
      looksLikePersonBirthDate(String(entry.values[1] ?? "")),
    ).length,
    withInn: entries.filter(
      (entry) => entry.values[4]?.replace(/\D/g, "").length === 10,
    ).length,
  };
};

export { formatStaffSheetEnrichmentReport };
