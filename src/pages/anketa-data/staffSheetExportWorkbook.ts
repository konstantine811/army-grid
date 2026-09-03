import { downloadBlob, sanitizeFileName } from "../../shared/browserExport";
import {
  STAFF_SHEET_ENRICHMENT_VALUE_INDEX,
  STAFF_SHEET_EXPORT_COLUMN_NUMBERS,
  staffSheetBirthDateColumn,
} from "../excel-fill/staffSheet";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  MORNING_GENERAL_LIST_COLUMN_LABELS,
  isLikelyCallSignToken,
  looksLikePersonBirthDate,
} from "../personnel/personnelUtils";
import { normalizeMilitaryIdCellValue } from "../personnel/vkTpvDovidkyImport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import type { StaffSheetEnrichmentEntry } from "./staffSheetEnrichment";

export const STAFF_SHEET_EXPORT_TEMPLATE_URL = `${
  import.meta.env.BASE_URL
}templates/staffSheetExportTemplate.xlsx`;

/** Рядок-еталон стилів даних (жовті/білі комірки, рамки) з «Штатки». */
export const STAFF_SHEET_EXPORT_STYLE_ROW = 2;

export const STAFF_SHEET_EXPORT_MAX_COLUMN = Math.max(
  ...STAFF_SHEET_EXPORT_COLUMN_NUMBERS,
  46,
);

const STYLE_PROP_NAMES = [
  "bold",
  "italic",
  "underline",
  "strikethrough",
  "fontSize",
  "fontFamily",
  "fontColor",
  "horizontalAlignment",
  "verticalAlignment",
  "wrapText",
  "shrinkToFit",
  "fill",
  "border",
  "leftBorder",
  "rightBorder",
  "topBorder",
  "bottomBorder",
  "numberFormat",
] as const;

let templateCache: ArrayBuffer | null = null;

const loadXlsxPopulate = async () => {
  const module = await import(
    "xlsx-populate/browser/xlsx-populate-no-encryption"
  );
  return module.default;
};

export const loadStaffSheetExportTemplate = async (
  templateUrl = STAFF_SHEET_EXPORT_TEMPLATE_URL,
) => {
  if (templateCache) return templateCache;
  const response = await fetch(templateUrl);
  if (!response.ok) {
    throw new Error(
      `Не вдалося завантажити шаблон «Штатки» (${response.status}).`,
    );
  }
  templateCache = await response.arrayBuffer();
  return templateCache;
};

export const resetStaffSheetExportTemplateCache = () => {
  templateCache = null;
};

type StyleCell = {
  style: (
    names?: readonly string[] | Record<string, string | boolean>,
  ) => Record<string, string | boolean>;
};

const copyCellStyle = (sourceCell: StyleCell, targetCell: StyleCell) => {
  targetCell.style(sourceCell.style([...STYLE_PROP_NAMES]));
};

const findGeneralListSheet = (workbook: {
  sheets: () => Array<{ name: () => string }>;
  sheet: (index: number) => StaffSheetExportSheet;
}) =>
  workbook
    .sheets()
    .find((item) => /загальний\s*список/i.test(item.name())) ??
  workbook.sheet(0);

type StaffSheetExportSheet = {
  cell: (
    row: number,
    column: number,
  ) => StyleCell & { value: (value?: unknown) => unknown };
};

const ensureDataRowStyles = (
  sheet: StaffSheetExportSheet,
  excelRow: number,
  styleRow = STAFF_SHEET_EXPORT_STYLE_ROW,
) => {
  if (excelRow === styleRow) return;
  for (let column = 1; column <= STAFF_SHEET_EXPORT_MAX_COLUMN; column += 1) {
    copyCellStyle(
      sheet.cell(styleRow, column),
      sheet.cell(excelRow, column),
    );
  }
};

const isValidBirthDate = (value: string) => {
  const text = value.trim();
  if (!text || isLikelyCallSignToken(text)) return false;
  return looksLikePersonBirthDate(text);
};

export const resolveStaffSheetExportCellValue = (
  rosterRow: EjournalPreviewRow | undefined,
  columnNumber: number,
  enrichRow: string[],
) => {
  const enrichedBirth = (
    enrichRow[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.birthDate] ?? ""
  ).trim();
  const baseBirth = rosterRow
    ? readRosterColumnValue(rosterRow, staffSheetBirthDateColumn).trim()
    : "";
  const birthDate = [enrichedBirth, baseBirth].find(isValidBirthDate);

  if (columnNumber === staffSheetBirthDateColumn + 1) {
    const yearText = (
      enrichRow[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.birthYear] ?? ""
    ).trim();
    if (yearText && /^\d{4}$/.test(yearText)) return Number(yearText);
    const base = rosterRow
      ? readRosterColumnValue(rosterRow, columnNumber).trim()
      : "";
    return base && /^\d{4}$/.test(base) ? Number(base) : base;
  }

  if (columnNumber === staffSheetBirthDateColumn + 2) {
    const ageText = (
      enrichRow[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.fullYears] ?? ""
    ).trim();
    if (ageText && /^\d{1,3}$/.test(ageText)) return Number(ageText);
    const base = rosterRow
      ? readRosterColumnValue(rosterRow, columnNumber).trim()
      : "";
    return base && /^\d{1,3}$/.test(base) ? Number(base) : base;
  }

  if (columnNumber === staffSheetBirthDateColumn) {
    return birthDate ?? "";
  }

  if (columnNumber === 10) {
    if (enrichRow[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.anketa] === "так") {
      return "так";
    }
    return rosterRow ? readRosterColumnValue(rosterRow, 10).trim() : "";
  }

  if (columnNumber === 11) {
    const merged = (
      enrichRow[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.militaryId] ?? ""
    ).trim();
    if (merged) return normalizeMilitaryIdCellValue(merged);
    const base = rosterRow ? readRosterColumnValue(rosterRow, 11).trim() : "";
    return normalizeMilitaryIdCellValue(base);
  }

  if (columnNumber === 19) {
    const inn = (
      enrichRow[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.inn] ?? ""
    ).replace(/\D/g, "");
    if (inn.length === 10) return inn;
    const base = rosterRow ? readRosterColumnValue(rosterRow, 19).trim() : "";
    const baseDigits = base.replace(/\D/g, "");
    return baseDigits.length === 10 ? baseDigits : "";
  }

  return rosterRow ? readRosterColumnValue(rosterRow, columnNumber).trim() : "";
};

export const writeStaffSheetExportWorkbook = async (
  rosterRows: EjournalPreviewRow[],
  entries: StaffSheetEnrichmentEntry[],
  options?: {
    templateData?: ArrayBuffer;
    download?: boolean;
  },
) => {
  const XlsxPopulate = await loadXlsxPopulate();
  const templateData =
    options?.templateData ?? (await loadStaffSheetExportTemplate());
  const workbook = await XlsxPopulate.fromDataAsync(templateData);
  const sheet = findGeneralListSheet(workbook);

  const enrichmentByRow = new Map(
    entries.map((entry) => [entry.excelRowNumber, entry.values]),
  );
  const rowByNumber = new Map(
    rosterRows
      .map((row) => [Number(row.__rowNumber) || 0, row] as const)
      .filter(([rowNumber]) => rowNumber > 0),
  );
  const maxRow = Math.max(
    STAFF_SHEET_EXPORT_STYLE_ROW,
    ...[...rowByNumber.keys()],
    ...[...enrichmentByRow.keys()],
  );

  STAFF_SHEET_EXPORT_COLUMN_NUMBERS.forEach((columnNumber) => {
    const header =
      MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber] ??
      `Колонка ${columnNumber}`;
    sheet.cell(1, columnNumber).value(header);
  });

  for (let excelRow = STAFF_SHEET_EXPORT_STYLE_ROW; excelRow <= maxRow; excelRow += 1) {
    ensureDataRowStyles(sheet, excelRow);
    const rosterRow = rowByNumber.get(excelRow);
    const enrichRow = enrichmentByRow.get(excelRow) ?? [];
    STAFF_SHEET_EXPORT_COLUMN_NUMBERS.forEach((columnNumber) => {
      const value = resolveStaffSheetExportCellValue(
        rosterRow,
        columnNumber,
        enrichRow,
      );
      sheet.cell(excelRow, columnNumber).value(value ?? "");
    });
  }

  const buffer = await workbook.outputAsync();
  if (options?.download === false) {
    return buffer;
  }

  const stamp = new Date().toISOString().slice(0, 10);
  const fileName = sanitizeFileName(`Штатка для Google ${stamp}.xlsx`);
  downloadBlob(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName,
  );
  return buffer;
};
