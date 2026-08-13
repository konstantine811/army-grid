export type CellValue = string | number | boolean | Date | null | undefined;

type RawCellValue =
  | CellValue
  | {
      text?: () => string;
      value?: () => unknown;
      children?: unknown[];
      name?: string;
      constructor?: { name?: string };
    }
  | unknown[];

export type ExcelRow = {
  id: string;
  excelRowNumber: number;
  values: CellValue[];
  source: "template" | "merged";
};

export type ExcelSheetSnapshot = {
  sheetIndex: number;
  sheetName: string;
  rawRows: CellValue[][];
  headerRows: CellValue[][];
  rows: ExcelRow[];
  columnCount: number;
  columnIndexes: number[];
  dataStartRow: number;
};

export type ExcelWorkbookSnapshot = {
  file: File;
  fileName: string;
  sheetName: string;
  headerRows: CellValue[][];
  rows: ExcelRow[];
  columnCount: number;
  columnIndexes: number[];
  dataStartRow: number;
  sheets: ExcelSheetSnapshot[];
};

type DebugCellValue = string | number | boolean | null;

const DEFAULT_DATA_START_ROW = 5;
const DEFAULT_HEADER_ROW_INDEX = 3;
const KEY_COLUMN_INDEX = 1;
const FALLBACK_COLUMN_COUNT = 30;
const HEADER_SCAN_LIMIT = 12;
const HEADER_KEYWORDS = [
  "аркуш",
  "колонка",
  "№ колонки",
  "назва стовпця",
  "обовязковість",
  "обовʼязковість",
  "обов’язковість",
  "правило",
  "коментар",
  "піб",
  "прізвище",
  "посада",
  "статус",
  "підрозділ",
  "звання",
  "дата",
];

const loadXlsxPopulate = async () => {
  const module =
    await import("xlsx-populate/browser/xlsx-populate-no-encryption");
  return module.default;
};

const isEmptyValue = (value: CellValue) =>
  value === null || value === undefined || String(value).trim() === "";

const normalizeKey = (value: CellValue) =>
  String(value ?? "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();

export const hasRowData = (row: CellValue[]) =>
  row.some((value) => !isEmptyValue(value));

const pickColumns = (row: CellValue[], columnIndexes: number[]) =>
  columnIndexes.map((columnIndex) => row[columnIndex] ?? null);

const normalizeHeaderText = (value: CellValue) =>
  valueToDisplay(value)
    .replace(/[ʼ’']/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

const scoreHeaderRow = (row: CellValue[]) => {
  const values = row.filter((value) => !isEmptyValue(value));
  const joined = values.map(normalizeHeaderText).join(" ");
  const keywordScore = HEADER_KEYWORDS.reduce((score, keyword) => {
    const normalizedKeyword = keyword.replace(/[ʼ’']/g, "").toLowerCase();

    return joined.includes(normalizedKeyword) ? score + 8 : score;
  }, 0);
  const textScore = values.filter(
    (value) => typeof value === "string" && String(value).trim().length <= 90,
  ).length;
  const numericPenalty =
    values.filter((value) => typeof value === "number").length * 2;

  return values.length + textScore + keywordScore - numericPenalty;
};

const detectHeaderRowIndex = (rows: CellValue[][]) => {
  const candidates = rows.slice(0, HEADER_SCAN_LIMIT).map((row, index) => ({
    index,
    score: scoreHeaderRow(row),
  }));
  const bestCandidate = candidates.reduce(
    (best, candidate) => (candidate.score > best.score ? candidate : best),
    { index: DEFAULT_HEADER_ROW_INDEX, score: Number.NEGATIVE_INFINITY },
  );
  const defaultCandidate = candidates.find(
    (candidate) => candidate.index === DEFAULT_HEADER_ROW_INDEX,
  );

  if (
    bestCandidate.score >= 18 ||
    bestCandidate.score >= (defaultCandidate?.score ?? 0) + 8
  ) {
    return bestCandidate.index;
  }

  return rows[DEFAULT_HEADER_ROW_INDEX]
    ? DEFAULT_HEADER_ROW_INDEX
    : Math.max(0, bestCandidate.index);
};

const extractXmlText = (value: unknown): string => {
  if (value === null || value === undefined) return "";
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  )
    return String(value);
  if (value instanceof Date) return value.toISOString();
  if (Array.isArray(value)) return value.map(extractXmlText).join("");

  if (typeof value === "object") {
    const node = value as { children?: unknown[]; value?: () => unknown };
    if (typeof node.value === "function") return extractXmlText(node.value());
    if (Array.isArray(node.children))
      return node.children.map(extractXmlText).join("");
  }

  return String(value);
};

const sanitizeCellValue = (value: RawCellValue): CellValue => {
  if (value === undefined || value === null) return null;
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  )
    return value;
  if (value instanceof Date) return value;

  if (Array.isArray(value)) {
    const text = value.map(extractXmlText).join("").trim();
    
    return text || null;
  }

  if (typeof value === "object") {
    const maybeRichText = value as {
      text?: () => string;
      value?: () => unknown;
      constructor?: { name?: string };
    };

    if (typeof maybeRichText.text === "function") return maybeRichText.text();
    if (typeof maybeRichText.value === "function")
      return extractXmlText(maybeRichText.value());

    return extractXmlText(value);
  }

  return String(value);
};

const normalizeDateDisplay = (value: CellValue) => {
  if (!(value instanceof Date)) return value ?? "";

  return new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(value);
};

export const valueToDisplay = (value: CellValue) =>
  String(normalizeDateDisplay(value));

const valueToDebug = (value: CellValue): DebugCellValue => {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
};

export function createWorkbookDebugPayload(snapshot: ExcelWorkbookSnapshot) {
  const createSheetPayload = (sheet: ExcelSheetSnapshot) => {
    const columns = Array.from({ length: sheet.columnCount }, (_, index) => ({
      index,
      originalIndex: sheet.columnIndexes[index],
      letter: getColumnLabel(sheet.columnIndexes[index] ?? index),
      header: getColumnHeader(sheet, index),
    }));

    const parsedRows = sheet.rows
      .filter((row) => hasRowData(row.values))
      .map((row) => ({
        excelRowNumber: row.excelRowNumber,
        values: row.values.map(valueToDebug),
        byColumn: Object.fromEntries(
          columns.map((column, index) => [
            column.header || `column_${index + 1}`,
            valueToDebug(row.values[index]),
          ]),
        ),
      }));

    return {
      sheetIndex: sheet.sheetIndex,
      sheetName: sheet.sheetName,
      dataStartRow: sheet.dataStartRow,
      columnCount: sheet.columnCount,
      rawRowCount: sheet.rawRows.length,
      parsedRowCount: parsedRows.length,
      rawRows: sheet.rawRows.map((row, index) => ({
        excelRowNumber: index + 1,
        values: row.map(valueToDebug),
      })),
      headerRows: sheet.headerRows.map((row) => row.map(valueToDebug)),
      columns,
      rows: parsedRows,
    };
  };

  return {
    fileName: snapshot.fileName,
    sheetCount: snapshot.sheets.length,
    activeSheetName: snapshot.sheetName,
    sheets: snapshot.sheets.map(createSheetPayload),
  };
}

export async function readWorkbookSnapshot(
  file: File,
): Promise<ExcelWorkbookSnapshot> {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(file);
  const sheets = workbook.sheets().map((sheet, sheetIndex) => {
    const values = (sheet.usedRange()?.value() ?? []) as RawCellValue[][];
    const rawColumnCount = Math.max(
      FALLBACK_COLUMN_COUNT,
      ...values.map((row) => row.length),
    );
    const fullRows = values.map((row) =>
      Array.from({ length: rawColumnCount }, (_, index) =>
        sanitizeCellValue(row[index] ?? null),
      ),
    );
    const headerRowIndex = detectHeaderRowIndex(fullRows);
    const dataStartRow = headerRowIndex + 2;
    const fullDataRows = fullRows.slice(dataStartRow - 1);
    const usedColumnIndexes = Array.from(
      { length: rawColumnCount },
      (_, index) => index,
    ).filter((columnIndex) => {
      const hasHeader = !isEmptyValue(fullRows[headerRowIndex]?.[columnIndex]);
      const hasData = fullDataRows.some(
        (row) => !isEmptyValue(row[columnIndex]),
      );

      return hasHeader || hasData;
    });
    const firstUsedColumnIndex = usedColumnIndexes[0] ?? 0;
    const lastUsedColumnIndex =
      usedColumnIndexes.at(-1) ?? Math.max(0, rawColumnCount - 1);
    const columnIndexes =
      usedColumnIndexes.length > 0
        ? Array.from(
            { length: lastUsedColumnIndex - firstUsedColumnIndex + 1 },
            (_, index) => firstUsedColumnIndex + index,
          )
        : [];
    const normalizedRows = fullRows.map((row) =>
      pickColumns(row, columnIndexes),
    );
    const columnCount = columnIndexes.length;

    return {
      sheetIndex,
      sheetName: sheet.name(),
      rawRows: normalizedRows,
      headerRows: normalizedRows.slice(0, headerRowIndex + 1),
      rows: normalizedRows.slice(dataStartRow - 1).map((values, index) => ({
        id: `${file.name}-${sheetIndex}-${dataStartRow + index}`,
        excelRowNumber: dataStartRow + index,
        values,
        source: "template" as const,
      })),
      columnCount,
      columnIndexes,
      dataStartRow,
    };
  });

  const activeSheet = sheets[0] ?? {
    sheetIndex: 0,
    sheetName: "Аркуш1",
    rawRows: [],
    headerRows: [],
    rows: [],
    columnCount: 0,
    columnIndexes: [],
    dataStartRow: DEFAULT_DATA_START_ROW,
  };

  return {
    file,
    fileName: file.name,
    sheetName: activeSheet.sheetName,
    headerRows: activeSheet.headerRows,
    rows: activeSheet.rows,
    columnCount: activeSheet.columnCount,
    columnIndexes: activeSheet.columnIndexes,
    dataStartRow: activeSheet.dataStartRow,
    sheets,
  };
}

export function getColumnLabel(index: number) {
  let label = "";
  let value = index + 1;

  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - 1) / 26);
  }

  return label;
}

export function getColumnHeader(
  snapshot: Pick<ExcelWorkbookSnapshot, "headerRows">,
  index: number,
) {
  const rawHeader = snapshot.headerRows.at(-1)?.[index];
  const header = valueToDisplay(rawHeader).replace(/\s+/g, " ").trim();

  return header;
}

export function updateCell(
  rows: ExcelRow[],
  rowId: string,
  columnIndex: number,
  value: string,
) {
  return rows.map((row) =>
    row.id === rowId
      ? {
          ...row,
          values: row.values.map((cellValue, index) =>
            index === columnIndex ? value : cellValue,
          ),
        }
      : row,
  );
}

export function mergeRows(baseRows: ExcelRow[], incomingRows: ExcelRow[]) {
  const rows = baseRows.map((row) => ({ ...row, values: [...row.values] }));
  const indexByKey = new Map<string, number>();

  rows.forEach((row, index) => {
    const key = normalizeKey(row.values[KEY_COLUMN_INDEX]);
    if (key) indexByKey.set(key, index);
  });

  const stats = {
    updated: 0,
    created: 0,
    skipped: 0,
  };

  incomingRows
    .filter((row) => hasRowData(row.values))
    .forEach((incomingRow) => {
      const key = normalizeKey(incomingRow.values[KEY_COLUMN_INDEX]);

      if (!key) {
        stats.skipped += 1;
        return;
      }

      const existingIndex = indexByKey.get(key);
      if (existingIndex !== undefined) {
        rows[existingIndex] = {
          ...rows[existingIndex],
          values: rows[existingIndex].values.map((value, index) =>
            isEmptyValue(incomingRow.values[index])
              ? value
              : incomingRow.values[index],
          ),
        };
        stats.updated += 1;
        return;
      }

      const emptyIndex = rows.findIndex((row) => !hasRowData(row.values));
      const nextRowNumber =
        emptyIndex >= 0
          ? rows[emptyIndex].excelRowNumber
          : Math.max(...rows.map((row) => row.excelRowNumber)) + 1;
      const nextRow = {
        id: `merged-${key}-${nextRowNumber}`,
        excelRowNumber: nextRowNumber,
        values: incomingRow.values,
        source: "merged" as const,
      };

      if (emptyIndex >= 0) {
        rows[emptyIndex] = nextRow;
      } else {
        rows.push(nextRow);
      }

      indexByKey.set(
        key,
        rows.findIndex((row) => row.id === nextRow.id),
      );
      stats.created += 1;
    });

  return { rows, stats };
}

export async function exportStyledWorkbook(
  snapshot: ExcelWorkbookSnapshot,
  rows: ExcelRow[],
  sheetIndex = 0,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(snapshot.file);
  const sheet = workbook.sheet(sheetIndex);
  const targetSnapshot =
    snapshot.sheets.find((item) => item.sheetIndex === sheetIndex) ?? snapshot;

  rows.forEach((row) => {
    row.values.forEach((value, columnIndex) => {
      const originalColumnIndex =
        targetSnapshot.columnIndexes[columnIndex] ?? columnIndex;
      sheet
        .cell(row.excelRowNumber, originalColumnIndex + 1)
        .value(isEmptyValue(value) ? null : value);
    });
  });

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const baseName = snapshot.fileName.replace(/\.xlsx$/i, "");
  link.href = url;
  link.download = `${baseName}_edited.xlsx`;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportWorkbookWithMutations(
  snapshot: ExcelWorkbookSnapshot,
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(snapshot.file);

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportWorkbookFileWithMutations(
  file: File,
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromDataAsync(file);

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportBlankWorkbookWithMutations(
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook = await XlsxPopulate.fromBlankAsync();

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportBlankWorkbookSheetWithMutations(
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  sheetName: string,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const workbook: any = await XlsxPopulate.fromBlankAsync();

  await mutateWorkbook(workbook);

  const targetSheet = workbook.sheet(sheetName);
  if (!targetSheet) {
    throw new Error(`Не знайдено аркуш "${sheetName}" для експорту.`);
  }

  [...workbook.sheets()].reverse().forEach((sheet: any) => {
    if (sheet.name() !== sheetName) workbook.deleteSheet(sheet);
  });

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

export async function exportTemplateWorkbookWithMutations(
  templateUrl: string,
  mutateWorkbook: (workbook: any) => void | Promise<void>,
  fileName: string,
) {
  const XlsxPopulate = await loadXlsxPopulate();
  const response = await fetch(templateUrl);
  if (!response.ok) {
    throw new Error(`Не вдалося завантажити Excel-шаблон: ${templateUrl}`);
  }
  const workbook: any = await XlsxPopulate.fromDataAsync(
    await response.arrayBuffer(),
  );

  await mutateWorkbook(workbook);

  const blob = await workbook.outputAsync("blob");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");

  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}
