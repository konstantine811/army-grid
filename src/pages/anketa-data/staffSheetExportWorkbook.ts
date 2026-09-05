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
import { looksLikePersonName } from "../soc-passport/socPassportFields";
import { normalizeMilitaryIdCellValue } from "../personnel/vkTpvDovidkyImport";
import { sanitizeStaffSheetCellValue } from "../excel-fill/staffSheet";
import { normalizeAnketaNameKey } from "./anketaPersonMatch";
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

export const STAFF_SHEET_EXPORT_PIB_COLUMN = 14;

/** Лише рядки з ПІБ — без вакантних позицій і підрозділів. */
export const filterStaffSheetExportRowsWithPib = (
  rosterRows: EjournalPreviewRow[],
): EjournalPreviewRow[] =>
  [...rosterRows]
    .filter((row) => readRosterColumnValue(row, STAFF_SHEET_EXPORT_PIB_COLUMN).trim())
    .sort(
      (left, right) =>
        (Number(left.__rowNumber) || 0) - (Number(right.__rowNumber) || 0),
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

type XlsxPopulateModule = typeof import(
  "xlsx-populate/browser/xlsx-populate-no-encryption"
);
type StaffSheetExportWorkbook = Awaited<
  ReturnType<XlsxPopulateModule["default"]["fromDataAsync"]>
>;
type StaffSheetExportSheet = ReturnType<StaffSheetExportWorkbook["sheet"]>;
type StaffSheetExportCell = ReturnType<StaffSheetExportSheet["cell"]>;

const copyCellStyle = (
  sourceCell: StaffSheetExportCell,
  targetCell: StaffSheetExportCell,
) => {
  const styles = Object.fromEntries(
    STYLE_PROP_NAMES.map((name) => [name, sourceCell.style(name)]),
  );
  targetCell.style(styles);
};

const findGeneralListSheet = (
  workbook: StaffSheetExportWorkbook,
): StaffSheetExportSheet =>
  workbook
    .sheets()
    .find((item) => /загальний\s*список/i.test(item.name())) ??
  workbook.sheet(0);

const HEADER_LIKE_FILL_RGB = new Set(["FFC000", "FFFF00"]);

const normalizeDataRowAppearance = (
  sheet: StaffSheetExportSheet,
  excelRow: number,
  rosterRow: EjournalPreviewRow | undefined,
) => {
  const hasPersonName = Boolean(
    readRosterColumnValue(rosterRow ?? {}, 14).trim(),
  );
  for (let column = 1; column <= STAFF_SHEET_EXPORT_MAX_COLUMN; column += 1) {
    const cell = sheet.cell(excelRow, column);
    const fill = cell.style("fill") as { color?: { rgb?: string } } | string;
    const rgb =
      typeof fill === "string"
        ? fill.replace(/^#?FF?/i, "").toUpperCase()
        : fill?.color?.rgb?.replace(/^FF/i, "").toUpperCase() ?? "";
    if (!HEADER_LIKE_FILL_RGB.has(rgb)) continue;
    if (column === 14 && hasPersonName) {
      cell.style("fill", { type: "solid", color: "FFC6EFCE" });
      continue;
    }
    cell.style("fill", { type: "solid", color: "FFFFFFFF" });
  }
};

const ensureDataRowStyles = (
  sheet: StaffSheetExportSheet,
  excelRow: number,
  styleRow = STAFF_SHEET_EXPORT_STYLE_ROW,
) => {
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

const readSheetCellText = (
  sheet: StaffSheetExportSheet,
  row: number,
  column: number,
) => String(sheet.cell(row, column).value() ?? "").trim();

/** Індекс ПІБ → номер рядка у завантаженому .xlsx (не покладатися лише на __rowNumber з gviz). */
const buildStaffSheetPibRowIndex = (sheet: StaffSheetExportSheet) => {
  const index = new Map<string, number>();
  let emptyStreak = 0;
  for (let row = STAFF_SHEET_EXPORT_STYLE_ROW; row <= 8000; row += 1) {
    const pib = readSheetCellText(sheet, row, STAFF_SHEET_EXPORT_PIB_COLUMN);
    if (looksLikePersonName(pib)) {
      emptyStreak = 0;
      const key = normalizeAnketaNameKey(pib);
      if (key && !index.has(key)) index.set(key, row);
      continue;
    }
    const hasRowContent = [5, 8, 10, 11, 13, 14].some(
      (column) => readSheetCellText(sheet, row, column).length > 0,
    );
    if (!hasRowContent) {
      emptyStreak += 1;
      if (emptyStreak >= 80) break;
    } else {
      emptyStreak = 0;
    }
  }
  return index;
};

const resolveStaffSheetOverlayRow = (
  pibRows: Map<string, number>,
  entry: StaffSheetEnrichmentEntry,
  usedRows: Set<number>,
): number | null => {
  const key = normalizeAnketaNameKey(entry.pib);
  if (!key) return null;
  const row = pibRows.get(key);
  if (!row || usedRows.has(row)) return null;
  return row;
};

/** Доповнює лише «Анкета» (10) та «Військовий квиток» (11) у вже імпортованому .xlsx. */
export const writeStaffSheetAnketaVkOverlay = async (
  baseWorkbookData: ArrayBuffer,
  entries: StaffSheetEnrichmentEntry[],
  options?: {
    download?: boolean;
    fileName?: string;
  },
) => {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(baseWorkbookData);
  const sheet = findGeneralListSheet(workbook);
  const pibRows = buildStaffSheetPibRowIndex(sheet);
  const usedRows = new Set<number>();

  for (const entry of entries) {
    const excelRow = resolveStaffSheetOverlayRow(pibRows, entry, usedRows);
    if (!excelRow) continue;
    usedRows.add(excelRow);

    const anketa = String(
      entry.values[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.anketa] ?? "",
    ).trim();
    const militaryId = normalizeMilitaryIdCellValue(
      sanitizeStaffSheetCellValue(
        String(entry.values[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.militaryId] ?? ""),
        11,
      ),
    );

    if (anketa === "так") {
      sheet.cell(excelRow, 10).value("так");
    } else if (
      !sanitizeStaffSheetCellValue(
        String(sheet.cell(excelRow, 10).value() ?? ""),
        10,
      )
    ) {
      sheet.cell(excelRow, 10).value("");
    }

    if (militaryId) {
      sheet.cell(excelRow, 11).value(militaryId);
    } else if (
      !sanitizeStaffSheetCellValue(
        String(sheet.cell(excelRow, 11).value() ?? ""),
        11,
      )
    ) {
      sheet.cell(excelRow, 11).value("");
    }
  }

  const buffer = await workbook.outputAsync();
  if (options?.download === false) {
    return buffer;
  }

  const stamp = new Date().toISOString().slice(0, 10);
  const fileName = sanitizeFileName(
    options?.fileName ?? `Штатка Анкета+ВК ${stamp}.xlsx`,
  );
  downloadBlob(
    new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    fileName,
  );
  return buffer;
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
  const exportRows = filterStaffSheetExportRowsWithPib(rosterRows);

  STAFF_SHEET_EXPORT_COLUMN_NUMBERS.forEach((columnNumber) => {
    const header =
      MORNING_GENERAL_LIST_COLUMN_LABELS[columnNumber] ??
      `Колонка ${columnNumber}`;
    sheet.cell(1, columnNumber).value(header);
  });

  exportRows.forEach((rosterRow, index) => {
    const excelRow = STAFF_SHEET_EXPORT_STYLE_ROW + index;
    const sourceRowNumber = Number(rosterRow.__rowNumber) || 0;
    const enrichRow = enrichmentByRow.get(sourceRowNumber) ?? [];
    ensureDataRowStyles(sheet, excelRow);
    normalizeDataRowAppearance(sheet, excelRow, rosterRow);
    STAFF_SHEET_EXPORT_COLUMN_NUMBERS.forEach((columnNumber) => {
      const value = sanitizeStaffSheetCellValue(
        String(
          resolveStaffSheetExportCellValue(
            rosterRow,
            columnNumber,
            enrichRow,
          ) ?? "",
        ),
        columnNumber,
      );
      sheet.cell(excelRow, columnNumber).value(value);
    });
  });

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
