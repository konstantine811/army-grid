import type { ExcelRow, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  getColumnHeader,
  getColumnLabel,
  valueToDisplay,
} from "../../excelRoundTrip";
import { normalizeDatasetKey } from "../../shared/format";
import type { EjournalColumn, EjournalPreviewRow } from "./ejournalTypes";
import { resolveMorningGeneralListColumnLabel } from "../personnel/personnelUtils";

export const buildImportColumns = (
  sheet: ExcelWorkbookSnapshot | ExcelWorkbookSnapshot["sheets"][number],
) =>
  Array.from({ length: sheet.columnCount }, (_, index) => {
    const originalIndex = sheet.columnIndexes[index] ?? index;
    const rawLabel = getColumnHeader(sheet, index).trim();
    const fallbackKey = `column_${originalIndex + 1}`;
    const label =
      rawLabel ||
      resolveMorningGeneralListColumnLabel(fallbackKey) ||
      rawLabel;
    const baseKey =
      normalizeDatasetKey(label || fallbackKey) || fallbackKey;

    return {
      key: baseKey,
      label,
      order: index,
      originalIndex,
      letter: getColumnLabel(originalIndex),
    };
  }).reduce<
    Array<{
      key: string;
      label: string;
      order: number;
      originalIndex: number;
      letter: string;
    }>
  >((columns, column) => {
    const duplicateCount = columns.filter(
      (existingColumn) =>
        existingColumn.key === column.key ||
        existingColumn.key.startsWith(`${column.key}_`),
    ).length;

    columns.push({
      ...column,
      key:
        duplicateCount === 0
          ? column.key
          : `${column.key}_${duplicateCount + 1}`,
    });

    return columns;
  }, []);

export const isEjournalColumn = (value: unknown): value is EjournalColumn =>
  Boolean(
    value && typeof value === "object" && "key" in value && "label" in value,
  );

export const parseDbColumns = (columns: unknown): EjournalColumn[] => {
  if (!Array.isArray(columns)) return [];

  return columns
    .filter(isEjournalColumn)
    .map((column, index) => ({
      key: String(column.key || `column_${index + 1}`),
      label: String(column.label ?? ""),
      order: Number.isFinite(column.order) ? Number(column.order) : index,
      originalIndex:
        typeof column.originalIndex === "number"
          ? column.originalIndex
          : undefined,
      letter: column.letter ? String(column.letter) : undefined,
    }))
    .sort((left, right) => left.order - right.order);
};

export const localRowsToPreviewRows = (
  rows: ExcelRow[],
  columns: EjournalColumn[],
): EjournalPreviewRow[] =>
  rows.map((row) => ({
    __rowNumber: row.excelRowNumber,
    ...Object.fromEntries(
      columns.map((column, index) => [column.key, row.values[index]]),
    ),
  }));

export const previewValueToDisplay = (value: unknown) =>
  value instanceof Date ||
  typeof value === "string" ||
  typeof value === "number" ||
  typeof value === "boolean" ||
  value === null ||
  value === undefined
    ? valueToDisplay(value)
    : JSON.stringify(value);

export const getRowValueByKeyPart = (
  row: EjournalPreviewRow | null,
  keyParts: string[],
) => {
  if (!row) return "";

  const foundKey = Object.keys(row).find(
    (key) =>
      !key.startsWith("__") &&
      keyParts.every((part) => key.toLowerCase().includes(part.toLowerCase())),
  );
  return foundKey ? previewValueToDisplay(row[foundKey]) : "";
};

export const getRowKeyByKeyPart = (
  row: EjournalPreviewRow | null,
  keyParts: string[],
) => {
  if (!row) return "";

  return (
    Object.keys(row).find(
      (key) =>
        !key.startsWith("__") &&
        keyParts.every((part) =>
          key.toLowerCase().includes(part.toLowerCase()),
        ),
    ) ?? ""
  );
};
