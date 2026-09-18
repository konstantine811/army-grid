import type { CellValue, ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  formatUkDate,
  tryParseExcelSerialDate,
} from "../../shared/format";

const MAX_DISPLAY_COLUMNS = 40;

const splitPackedStaffIndexes = (text: string) => {
  const packed = text.replace(/\s+/g, "");
  if (/^(?:\d{7}){2,}$/.test(packed)) {
    return packed.match(/\d{7}/g)?.join("\n") ?? text;
  }
  return text;
};

export const displayEjoosCell = (
  value: CellValue | undefined,
): string => {
  if (value == null) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toLocaleDateString("uk-UA");
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const parsedSerial = tryParseExcelSerialDate(value);
    if (parsedSerial) return formatUkDate(parsedSerial);
    return splitPackedStaffIndexes(String(value));
  }
  const parsedSerial = tryParseExcelSerialDate(value);
  if (parsedSerial) return formatUkDate(parsedSerial);
  const text = String(value)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n");
  if (text.trim() === "[object Object]") return "#N/A";
  return splitPackedStaffIndexes(text)
    .split("\n")
    .map((line) => line.replace(/[^\S\n]+/g, " ").trim())
    .join("\n")
    .trim();
};

export const findEjoosHeaderRowIndex = (
  sheet: ExcelSheetSnapshot,
): number => {
  const limit = Math.min(12, sheet.rawRows.length);
  for (let index = 0; index < limit; index += 1) {
    const texts = (sheet.rawRows[index] ?? [])
      .map((cell) => displayEjoosCell(cell).toLocaleLowerCase("uk-UA"))
      .filter(Boolean);
    if (texts.length < 2) continue;
    if (
      /прізвище|піб|звання|посада|індекс|статус|дата|звідки|табель/.test(
        texts.join(" "),
      )
    ) {
      return index;
    }
  }
  return 0;
};

const looksLikeFieldNumberRow = (values: string[]) => {
  const filled = values.filter(Boolean);
  return (
    filled.length >= 2 &&
    filled.every((value) => /^\d+(?:\.\d+)?$/.test(value))
  );
};

export type EjoosDataTableRow = {
  __rowId: string;
  __excelRowNumber: number;
  [columnId: string]: string | number;
};

export type EjoosDataTableColumn = {
  id: string;
  label: string;
  sourceIndex: number;
  compact: boolean;
};

export type EjoosDataTableModel = {
  sheetName: string;
  headerIndex: number;
  dataStartIndex: number;
  nameColumnId: string | null;
  columns: EjoosDataTableColumn[];
  rows: EjoosDataTableRow[];
};

export const buildEjoosDataTableModel = (
  sheet: ExcelSheetSnapshot,
): EjoosDataTableModel => {
  const headerIndex = findEjoosHeaderRowIndex(sheet);
  const headerRow = sheet.rawRows[headerIndex] ?? [];
  let lastColumn = 0;
  for (let rowIndex = headerIndex; rowIndex < sheet.rawRows.length; rowIndex += 1) {
    (sheet.rawRows[rowIndex] ?? []).forEach((cell, columnIndex) => {
      if (displayEjoosCell(cell)) {
        lastColumn = Math.max(lastColumn, columnIndex + 1);
      }
    });
  }
  lastColumn = Math.min(Math.max(lastColumn, 4), MAX_DISPLAY_COLUMNS);

  const rawLabels = Array.from({ length: lastColumn }, (_, index) =>
    displayEjoosCell(headerRow[index]),
  );
  const labels = rawLabels.map(
    (label, index) => label || `Колонка ${index + 1}`,
  );
  let dataStartIndex = headerIndex + 1;
  const possibleNumberRow = Array.from({ length: lastColumn }, (_, index) =>
    displayEjoosCell(sheet.rawRows[dataStartIndex]?.[index]),
  );
  if (looksLikeFieldNumberRow(possibleNumberRow)) dataStartIndex += 1;

  const rawData = sheet.rawRows.slice(dataStartIndex).map((rawRow, offset) => ({
    excelRowNumber: dataStartIndex + offset + 1,
    values: Array.from({ length: lastColumn }, (_, index) =>
      displayEjoosCell(rawRow?.[index]),
    ),
  }));
  const nonEmptyData = rawData.filter(({ values }) =>
    values.some((value) => value),
  );
  const columns = labels.map((label, sourceIndex) => {
    const values = nonEmptyData
      .map((row) => row.values[sourceIndex]?.trim() ?? "")
      .filter(Boolean);
    return {
      id: `column_${sourceIndex}`,
      label,
      sourceIndex,
      compact:
        (!rawLabels[sourceIndex] ||
          /^[-–—]+$/.test(rawLabels[sourceIndex]!.trim())) &&
        values.length > 0 &&
        values.every((value) => /^[-–—]+$/.test(value)),
    };
  });
  const nameColumn =
    columns.find((column) => /прізвище|піб/i.test(column.label)) ??
    columns[1] ??
    null;
  const rows = nonEmptyData.map(({ excelRowNumber, values }) => {
    const record: EjoosDataTableRow = {
      __rowId: `${sheet.sheetName}:${excelRowNumber}`,
      __excelRowNumber: excelRowNumber,
    };
    columns.forEach((column) => {
      record[column.id] = values[column.sourceIndex] ?? "";
    });
    return record;
  });

  return {
    sheetName: sheet.sheetName,
    headerIndex,
    dataStartIndex,
    nameColumnId: nameColumn?.id ?? null,
    columns,
    rows,
  };
};
