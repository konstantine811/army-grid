import {
  readWorkbookSnapshot,
  type CellValue,
  type ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { formatUkDate, tryParseExcelSerialDate } from "../../shared/format";

export type StaffListColumn = {
  id: string;
  label: string;
  index: number;
};

export type StaffListRow = {
  __rowId: string;
  __rowNumber: number;
  __excelRowNumber: number;
  pib: string;
  birthDate: string;
  values: Record<string, string>;
};

export type StaffListSnapshot = {
  fileName: string;
  sheetName: string;
  sheetIndex: number;
  columnIndexes: number[];
  columns: StaffListColumn[];
  rows: StaffListRow[];
  pibColumnId: string;
  birthDateColumnId: string;
};

export type ParsedStaffListFile = {
  snapshot: StaffListSnapshot;
  fileData: ArrayBuffer;
};

export const normalizeStaffListHeader = (value: unknown) =>
  String(value ?? "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const normalizeHeader = normalizeStaffListHeader;

export const staffListCellToDisplayString = (value: CellValue): string => {
  if (value == null || value === "") return "";
  if (value instanceof Date) return formatUkDate(value);
  if (typeof value === "boolean") return value ? "так" : "ні";
  if (typeof value === "number") {
    const asDate = tryParseExcelSerialDate(value);
    if (asDate) return formatUkDate(asDate);
    return String(value);
  }
  return String(value).trim();
};

const buildColumnId = (label: string, index: number) => {
  const slug = normalizeHeader(label)
    .replace(/[^a-zа-яіїєґ0-9]+/gi, "_")
    .replace(/^_+|_+$/g, "");
  return slug ? `${slug}_${index}` : `col_${index}`;
};

const findColumnByHeader = (
  columns: StaffListColumn[],
  pattern: RegExp,
  fallbackIndex = -1,
) => {
  const match = columns.find((column) => pattern.test(normalizeHeader(column.label)));
  return match?.id ?? columns[fallbackIndex]?.id ?? "";
};

export const buildStaffListSnapshot = (
  workbook: ExcelWorkbookSnapshot,
): StaffListSnapshot => {
  const activeSheet =
    workbook.sheets.find((sheet) => sheet.sheetName === workbook.sheetName) ??
    workbook.sheets[0];
  const headerRow = workbook.headerRows.at(-1) ?? [];
  const columns: StaffListColumn[] = headerRow.map((label, index) => ({
    id: buildColumnId(staffListCellToDisplayString(label), index),
    label: staffListCellToDisplayString(label) || `Колонка ${index + 1}`,
    index,
  }));

  const pibColumnId = findColumnByHeader(columns, /п\.?\s*і\.?\s*б|піб|прізвищ/);
  const birthDateColumnId = findColumnByHeader(
    columns,
    /дата.*народ|народж/,
  );

  const rows: StaffListRow[] = workbook.rows
    .map((row, index) => {
      const values = Object.fromEntries(
        columns.map((column) => [
          column.id,
          staffListCellToDisplayString(row.values[column.index] ?? ""),
        ]),
      );
      const pib = pibColumnId ? values[pibColumnId] ?? "" : "";
      if (!pib.trim()) return null;
      return {
        __rowId: row.id,
        __rowNumber: index + 1,
        __excelRowNumber: row.excelRowNumber,
        pib,
        birthDate: birthDateColumnId ? values[birthDateColumnId] ?? "" : "",
        values,
      };
    })
    .filter((row): row is StaffListRow => Boolean(row));

  return {
    fileName: workbook.fileName,
    sheetName: workbook.sheetName,
    sheetIndex: activeSheet?.sheetIndex ?? 0,
    columnIndexes: activeSheet?.columnIndexes ?? workbook.columnIndexes,
    columns,
    rows,
    pibColumnId,
    birthDateColumnId,
  };
};

export const parseStaffListExcelFile = async (
  file: File,
): Promise<ParsedStaffListFile> => {
  const [workbook, fileData] = await Promise.all([
    readWorkbookSnapshot(file, {
      maxColumns: 40,
      skipStyleFills: true,
    }),
    file.arrayBuffer(),
  ]);
  if (!workbook.rows.length) {
    throw new Error("У файлі не знайдено рядків з даними.");
  }
  return {
    snapshot: buildStaffListSnapshot(workbook),
    fileData,
  };
};

export const isStaffListStructuralColumn = (
  snapshot: StaffListSnapshot,
  column: StaffListColumn,
) => {
  if (column.id === snapshot.pibColumnId) return true;
  const header = normalizeHeader(column.label);
  if (/^№$|^номер$/i.test(header)) return true;
  if (/^посад/i.test(header)) return true;
  if (/^зван/i.test(header)) return true;
  if (/підрозділ|підр\.|subunit/i.test(header)) return true;
  return false;
};

const STAFF_LIST_DEFAULT_GAP_HEADER_PATTERNS = [
  /телефон|phone|моб/i,
  /місце.*народ|place.*birth|народж.*місце/i,
  /дата.*народ|birth.*date|народж.*дата/i,
];

export const defaultStaffListGapColumnIds = (snapshot: StaffListSnapshot) => {
  const byPattern = snapshot.columns
    .filter(
      (column) =>
        !isStaffListStructuralColumn(snapshot, column) &&
        STAFF_LIST_DEFAULT_GAP_HEADER_PATTERNS.some((pattern) =>
          pattern.test(normalizeHeader(column.label)),
        ),
    )
    .sort((left, right) => left.index - right.index)
    .map((column) => column.id);

  if (byPattern.length) return byPattern;

  return snapshot.columns
    .filter((column) => !isStaffListStructuralColumn(snapshot, column))
    .sort((left, right) => left.index - right.index)
    .map((column) => column.id);
};
