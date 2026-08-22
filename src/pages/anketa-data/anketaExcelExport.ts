import writeXlsxFile, {
  type CellObject,
  type SheetData,
} from "write-excel-file/browser";
import { ANKETA_COLUMNS } from "./anketaSheet";
import { ANKETA_MISSING_VALUE_PRESETS } from "./anketaGaps";

/** Visual style close to Google Sheets anketa workbook. */
const HEADER_BG = "#E8EAED";
const HEADER_TEXT = "#202124";
const BORDER = "#BDBDBD";
const WHITE = "#FFFFFF";
const ZEBRA = "#F8F9FA";
const TEXT = "#202124";
const MUTED = "#5F6368";
const EMPTY_BG = "#FCE8E6";
const EMPTY_TEXT = "#A50E0E";
const PRESET_BG = "#FEF7E0";
const PRESET_TEXT = "#8A5A00";
const ROW_NUM_BG = "#F1F3F4";

const MISSING_PRESET_SET = new Set<string>(ANKETA_MISSING_VALUE_PRESETS);

const border = {
  borderColor: BORDER,
  borderStyle: "thin" as const,
};

const cell = (
  value: string | number | null | undefined,
  extra: Omit<CellObject, "value"> = {},
): CellObject => ({
  value: value == null || value === "" ? undefined : value,
  fontFamily: "Arial",
  fontSize: 10,
  alignVertical: "center",
  textColor: TEXT,
  ...border,
  ...extra,
});

const widthForColumn = (columnId: string, label: string) => {
  if (columnId === "__rowNumber") return 6;
  const def = ANKETA_COLUMNS.find((column) => column.key === columnId);
  if (def) return Math.max(10, Math.min(42, Math.round(def.size / 8)));
  return Math.max(12, Math.min(36, Math.round(label.length * 0.9) + 4));
};

const bodyStyleForValue = (
  columnId: string,
  raw: string,
  zebra: { backgroundColor: string },
) => {
  const value = raw.trim();
  if (columnId === "__rowNumber") {
    return {
      ...zebra,
      backgroundColor: ROW_NUM_BG,
      align: "center" as const,
      textColor: MUTED,
      fontWeight: "bold" as const,
    };
  }
  if (!value) {
    return {
      backgroundColor: EMPTY_BG,
      textColor: EMPTY_TEXT,
      wrap: true,
      align: "left" as const,
    };
  }
  if (MISSING_PRESET_SET.has(value)) {
    return {
      backgroundColor: PRESET_BG,
      textColor: PRESET_TEXT,
      wrap: true,
      align: "left" as const,
    };
  }
  if (
    columnId === "additionalInfo" ||
    columnId === "relatives" ||
    columnId === "education" ||
    columnId === "birthPlace"
  ) {
    return { ...zebra, wrap: true, align: "left" as const };
  }
  if (columnId === "fullName") {
    return {
      ...zebra,
      wrap: true,
      align: "left" as const,
      fontWeight: "bold" as const,
    };
  }
  return { ...zebra, wrap: true, align: "left" as const };
};

export type AnketaExcelExportColumn = {
  id: string;
  label: string;
  value: (row: unknown) => string;
};

export const exportAnketaSheetExcel = async <TRow>({
  columns,
  rows,
  fileName,
}: {
  columns: AnketaExcelExportColumn[];
  rows: TRow[];
  fileName?: string;
}) => {
  const stamp = new Date().toISOString().slice(0, 10);
  const headerHeight = 36;

  const headerRow = columns.map((column) =>
    cell(column.label, {
      fontWeight: "bold",
      fontSize: 10,
      textColor: HEADER_TEXT,
      backgroundColor: HEADER_BG,
      align: "center",
      wrap: true,
      height: headerHeight,
    }),
  );

  const bodyRows: SheetData = rows.map((row, index) => {
    const zebra =
      index % 2 === 1
        ? { backgroundColor: ZEBRA }
        : { backgroundColor: WHITE };
    return columns.map((column) => {
      const raw = String(column.value(row) ?? "");
      const style = bodyStyleForValue(column.id, raw, zebra);
      const display =
        !raw.trim() && column.id !== "__rowNumber" ? "—" : raw;
      return cell(display, {
        height:
          column.id === "additionalInfo" || column.id === "relatives"
            ? 28
            : 18,
        ...style,
      });
    });
  });

  const sheetData: SheetData = [headerRow, ...bodyRows];
  const hasRowNumber = columns.some((column) => column.id === "__rowNumber");
  const hasFullName = columns.some((column) => column.id === "fullName");
  const stickyColumnsCount = (hasRowNumber ? 1 : 0) + (hasFullName ? 1 : 0);

  await writeXlsxFile(sheetData, {
    sheet: "Анкети",
    orientation: "landscape",
    stickyRowsCount: 1,
    stickyColumnsCount: Math.max(1, stickyColumnsCount),
    showGridLines: true,
    dateFormat: "DD.MM.YYYY",
    columns: columns.map((column) => ({
      width: widthForColumn(column.id, column.label),
    })),
  }).toFile(fileName ?? `anketni-dani-${stamp}.xlsx`);
};
