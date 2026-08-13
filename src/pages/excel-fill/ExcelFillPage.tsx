import { useMemo, useState } from "react";
import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import {
  FileDownloadOutlinedIcon,
  FileUploadOutlinedIcon,
  SyncAltOutlinedIcon,
} from "@/components/sci/icons";
import {
  type CellValue,
  type ExcelSheetSnapshot,
  type ExcelWorkbookSnapshot,
  exportBlankWorkbookWithMutations,
  exportWorkbookFileWithMutations,
  getColumnHeader,
  readWorkbookSnapshot,
  valueToDisplay,
} from "../../excelRoundTrip";

type ColumnMeta = {
  index: number;
  originalIndex: number;
  header: string;
  key: string;
};

type FillChange = {
  rowNumber: number;
  columnIndex: number;
  originalColumnIndex: number;
  header: string;
  from: CellValue;
  to: CellValue;
  person: string;
};

type SourceRecord = {
  row: CellValue[];
  columns: ColumnMeta[];
  weaponKind: "rifle" | "pistol" | "other";
};

type FillResult = {
  changes: FillChange[];
  matchedRows: number;
  unmatchedRows: number;
  skippedFilled: number;
  sourceColumns: ColumnMeta[];
  targetColumns: ColumnMeta[];
};

type MorningReportRow = {
  sourceRowNumber: number;
  sequence: number;
  staffUnit: string;
  name: string;
  callsign: string;
  actualUnit: string;
  location: string;
  activity: string;
  reportStatus: string;
  sourceStatus: string;
  note: string;
  submittedBy: string;
};

type MorningReportResult = {
  rows: MorningReportRow[];
  counts: Record<string, number>;
  skippedRows: number;
  missingIdRows: number;
};

const isEmpty = (value: CellValue) =>
  value === null || value === undefined || valueToDisplay(value).trim() === "";

const isWeaponPlaceholder = (value: CellValue) =>
  /^(втрачений|втрачено|втрачена|втрата)$/i.test(valueToDisplay(value).trim());

const toUppercaseCallsign = (value: string) =>
  value.replace(/\s+/g, " ").trimStart().toLocaleUpperCase("uk-UA");

const normalizeText = (value: CellValue | string) =>
  valueToDisplay(value as CellValue)
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

const normalizeRank = (value: CellValue) => normalizeText(value);

const STATUS_OPTIONS = [
  "2ШБ",
  "3ШБ (2ПБ)",
  "4ШБ",
  "В зброярні",
  "Видана в.с.",
  "Втрата",
  "ДБР",
  "Задіяна в охороні ППД",
  "Знищена",
  "Наградний",
  "не бг",
  "не бг, в ремонті в полку",
  "не бг, пошкодження",
  "Передана",
  "Прикомандирований, не на обліку",
  "Прикомандирований, не на обліку, 184 НЦ",
  "Сзч",
  "Шквал",
] as const;

const MORNING_REPORT_SOURCE_STATUSES = new Set([
  "В строю",
  "Відком. за межі ПБ",
  "Відрядження",
  "Новоприбулий",
]);

const MORNING_REPORT_STATUSES = [
  "На виконанні",
  "БГ",
  "Екіпажі техніки",
  "Розрахунки колективного озброєння",
  "Екіпажі БПЛА",
  "Управління",
  "Забезпечення",
  "Без БЗВП",
  "БРЕЗ",
  "Навчання",
  "Відмовники",
  "Не БГ",
  "Вибув з локації",
] as const;

const MORNING_REPORT_STATUS_DESCRIPTIONS: Array<[string, string]> = [
  ["На виконанні", "ті, хто фактично виконує бойову задачу на позиції"],
  ["БГ", "здорові в/с, які пройшли БЗВП і можуть виконувати бойові задачі"],
  ["Екіпажі техніки", "мехводи, навідники та екіпажі техніки"],
  ["Розрахунки колективного озброєння", "розрахунки колективного озброєння"],
  ["Екіпажі БПЛА", "в/с, які працюють з БПЛА"],
  ["Управління", "командування, штаб, управління локації або підрозділу"],
  ["Забезпечення", "кухарі, комірники, охорона, водії, медики та інше забезпечення"],
  ["Без БЗВП", "в/с, які не пройшли БЗВП"],
  ["БРЕЗ", "окрема категорія БРЕЗ"],
  ["Відмовники", ""],
  ["Не БГ", "тимчасово небоєготові в/с"],
  ["Навчання", "в/с, які проходять навчання"],
  ["Вибув з локації", "в/с, які вибули за межі локації/ПБ"],
];

const normalizeStatus = (value: CellValue) => {
  const text = normalizeText(value);
  if (!text) return null;

  const exact = STATUS_OPTIONS.find((status) => normalizeText(status) === text);
  if (exact) return exact;
  if (text === "сзч") return "Сзч";
  if (text === "не бг") return "не бг";
  if (text.includes("ремонт") && text.includes("полк")) return "не бг, в ремонті в полку";
  if (text.includes("пошкод")) return "не бг, пошкодження";
  if (text.includes("прикомандирований") && text.includes("184")) {
    return "Прикомандирований, не на обліку, 184 НЦ";
  }
  if (text.includes("прикомандирований")) return "Прикомандирований, не на обліку";

  return null;
};

const normalizeReportText = (value: CellValue | string) => normalizeText(value);

const isPlaceholderUnit = (value: string) => {
  const text = normalizeReportText(value);
  return !text || text === "ж" || text === "нова";
};

const findMorningSheet = (workbook: ExcelWorkbookSnapshot) =>
  workbook.sheets.find((sheet) => normalizeReportText(sheet.sheetName).includes("ос загальний")) ??
  workbook.sheets[0];

const getMorningValue = (row: CellValue[], columnNumber: number) =>
  valueToDisplay(row[columnNumber - 1]).trim();

const makeStaffUnit = (row: CellValue[]) => {
  const possibleId = getMorningValue(row, 1);
  if (possibleId && !isPlaceholderUnit(possibleId)) return possibleId;

  return [2, 3, 4]
    .map((column) => getMorningValue(row, column))
    .filter((value) => !isPlaceholderUnit(value))
    .join(" / ");
};

const makeMorningActivity = (row: CellValue[]) =>
  getMorningValue(row, 22) || getMorningValue(row, 5) || getMorningValue(row, 7);

const makeMorningLocation = (row: CellValue[]) =>
  getMorningValue(row, 31) || getMorningValue(row, 33) || getMorningValue(row, 29);

const classifyMorningReportStatus = (row: CellValue[]) => {
  const activity = normalizeReportText(getMorningValue(row, 22));
  const readiness = normalizeReportText(getMorningValue(row, 23));
  const bzvp = normalizeReportText(getMorningValue(row, 24));
  const unit = normalizeReportText(getMorningValue(row, 29));
  const location = normalizeReportText(makeMorningLocation(row));
  const position = normalizeReportText(`${getMorningValue(row, 5)} ${getMorningValue(row, 7)}`);

  if (bzvp.includes("брез")) return "БРЕЗ";
  if (bzvp.includes("без бзвп")) return "Без БЗВП";
  if (readiness.includes("тимчасово не бг") || readiness === "не бг") return "Не БГ";
  if (/навчан|полігон/.test(location)) return "Навчання";
  if (/на виконан|позиці|одрадне|бажан/.test(location)) return "На виконанні";
  if (/бпла|пілот|fpv|бомбер|розвідка/.test(activity) || /бпла/.test(position)) return "Екіпажі БПЛА";
  if (/екіпаж|мехвод|навідник|брон|танк|бтр|бмп|хамер|hmmwv|технік/.test(activity)) return "Екіпажі техніки";
  if (/гранатомет|кулемет|міномет|гаубиц|розрахунок/.test(activity)) {
    return "Розрахунки колективного озброєння";
  }
  if (/упр|штаб/.test(activity) || /управл|штаб/.test(unit)) return "Управління";
  if (/забезпеч|охорон|водій|медик|кухар|майстер|зв'яз|зв яз|вмтз|ппо|реб/.test(activity) || /вмтз|ппо|реб|зв/.test(unit)) {
    return "Забезпечення";
  }
  if (readiness === "бг") return "БГ";

  return "Забезпечення";
};

function analyzeMorningReport(
  source?: ExcelWorkbookSnapshot | null,
  submittedBy = "",
  idLookup = new Map<string, string>(),
): MorningReportResult | null {
  const sourceSheet = source ? findMorningSheet(source) : undefined;
  if (!sourceSheet) return null;

  const rows: MorningReportRow[] = [];
  let skippedRows = 0;
  let missingIdRows = 0;

  sourceSheet.rows.forEach((excelRow) => {
    const row = excelRow.values;
    const name = getMorningValue(row, 14);
    const sourceStatus = getMorningValue(row, 21);
    if (!name || !MORNING_REPORT_SOURCE_STATUSES.has(sourceStatus)) {
      if (name) skippedRows += 1;
      return;
    }
    const nameKey = buildFullNameKey(name);
    const staffUnit = idLookup.get(nameKey) || makeStaffUnit(row);
    if (!idLookup.has(nameKey)) missingIdRows += 1;

    rows.push({
      sourceRowNumber: excelRow.excelRowNumber,
      sequence: rows.length + 1,
      staffUnit,
      name,
      callsign: getMorningValue(row, 15),
      actualUnit: "_5 1ПБ",
      location: makeMorningLocation(row),
      activity: makeMorningActivity(row),
      reportStatus: classifyMorningReportStatus(row),
      sourceStatus,
      note: getMorningValue(row, 32) || getMorningValue(row, 34),
      submittedBy,
    });
  });

  const counts = Object.fromEntries(MORNING_REPORT_STATUSES.map((status) => [status, 0]));
  rows.forEach((row) => {
    counts[row.reportStatus] = (counts[row.reportStatus] ?? 0) + 1;
  });

  return { rows, counts, skippedRows, missingIdRows };
}

const buildNameKey = (value: CellValue) => {
  const text = normalizeText(value);
  if (!text) return "";
  const parts = text.split(" ").filter(Boolean);
  const surname = parts[0] ?? "";
  const initials = parts
    .slice(1)
    .map((part) => part[0])
    .filter(Boolean)
    .join("");

  return surname ? `${surname}|${initials}` : "";
};

const buildFullNameKey = (value: CellValue | string) =>
  normalizeText(value)
    .replace(/\b\d{1,2}\s*\d{1,2}\s*\d{2,4}\b/g, " ")
    .replace(/\bр\s*н\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();

const headerKey = (header: string, previousHeader = "") => {
  const rawHeader = header.replace(/\s+/g, " ").trim().toLowerCase();
  const text = normalizeText(header);
  const previous = normalizeText(previousHeader);
  const isNumberHeader = rawHeader === "№" || rawHeader === "номер";

  if (!text && !isNumberHeader) return "";
  if (/(піб|піп|фио|прізвище|закріплена зброя)/.test(text)) return "personName";
  if (text.includes("звання")) return "rank";
  if (text.includes("позив")) return "callsign";
  if (text.includes("статус")) return "status";
  if (text.includes("підрозділ")) return "unit";
  if (text.includes("напрям")) return "direction";
  if (text.includes("приміт")) return "note";
  if (text.includes("стан збро")) return "weaponState";
  if (text.includes("номер збро") || text === "номер зброї") return "weaponNumber";
  if (isNumberHeader && /(автомат|збро|найменування)/.test(previous)) return "weaponNumber";
  if (isNumberHeader && previous.includes("пістолет")) return "pistolNumber";
  if (isNumberHeader) return "sequence";
  if (text.includes("назва пістолет")) return "pistolName";
  if (text.includes("назва автомат") || text.includes("назва збро")) return "weaponName";
  if (text.includes("найменування")) return "weaponName";

  return `same:${text}`;
};

const getColumns = (sheet: ExcelSheetSnapshot): ColumnMeta[] =>
  Array.from({ length: sheet.columnCount }, (_, index) => {
    const header = getColumnHeader(sheet, index);
    const previousHeader = index > 0 ? getColumnHeader(sheet, index - 1) : "";

    return {
      index,
      originalIndex: sheet.columnIndexes[index] ?? index,
      header,
      key: headerKey(header, previousHeader),
    };
  });

const findColumn = (columns: ColumnMeta[], keys: string[]) =>
  columns.find((column) => keys.includes(column.key));

const buildMorningIdLookup = (workbook?: ExcelWorkbookSnapshot | null) => {
  const lookup = new Map<string, string>();
  if (!workbook) return lookup;

  workbook.sheets.forEach((sheet) => {
    const columns = getColumns(sheet);
    const idColumn =
      columns.find((column) => normalizeText(column.header) === "id") ??
      columns.find((column) => column.header.trim().toLocaleUpperCase("uk-UA") === "ID");
    const nameColumn = findColumn(columns, ["personName"]);
    if (!idColumn || !nameColumn) return;

    sheet.rows.forEach((row) => {
      const id = valueToDisplay(row.values[idColumn.index]).trim();
      const nameKey = buildFullNameKey(row.values[nameColumn.index]);
      if (!nameKey || !id || id === "0" || id === "[object Object]") return;
      if (!lookup.has(nameKey)) lookup.set(nameKey, id);
    });
  });

  return lookup;
};

const rowValue = (row: CellValue[], column?: ColumnMeta) =>
  column ? row[column.index] : null;

const getWeaponKind = (weaponName: CellValue): SourceRecord["weaponKind"] => {
  const text = normalizeText(weaponName);
  if (/пістолет|glock|sig sauer|browning/.test(text)) return "pistol";
  if (/автомат|гвинтівк|карабін|ак-|акс|scar|colt|haenel|cz bren|м4|m4/.test(text)) return "rifle";
  return "other";
};

const sourceRowScore = (record: SourceRecord) => {
  const row = record.row;
  const columns = record.columns;
  const status = normalizeStatus(rowValue(row, findColumn(columns, ["status"]))) ?? "";
  const weaponName = rowValue(row, findColumn(columns, ["weaponName"]));
  const weaponNumber = rowValue(row, findColumn(columns, ["weaponNumber"]));
  let score = 0;

  if (!isEmpty(weaponName)) score += 2;
  if (!isEmpty(weaponNumber)) score += 2;
  if (record.weaponKind === "rifle") score += 1;
  if (/видана|на руках|закріп/.test(status)) score += 4;
  if (/втрата|втрач|передана|сзч/.test(status)) score -= 3;

  return score;
};

function analyzeFill(source?: ExcelWorkbookSnapshot | null, target?: ExcelWorkbookSnapshot | null): FillResult | null {
  const sourceSheet = source?.sheets[0];
  const targetSheet = target?.sheets[0];
  if (!sourceSheet || !targetSheet) return null;

  const sourceColumns = getColumns(sourceSheet);
  const targetColumns = getColumns(targetSheet);
  const sourceNameColumn = findColumn(sourceColumns, ["personName"]);
  const sourceRankColumn = findColumn(sourceColumns, ["rank"]);
  const targetNameColumn = findColumn(targetColumns, ["personName"]);
  const targetRankColumn = findColumn(targetColumns, ["rank"]);

  if (!sourceNameColumn || !targetNameColumn) {
    return {
      changes: [],
      matchedRows: 0,
      unmatchedRows: targetSheet.rows.filter((row) => !isEmpty(rowValue(row.values, targetNameColumn))).length,
      skippedFilled: 0,
      sourceColumns,
      targetColumns,
    };
  }

  const sourceByStrictKey = new Map<string, SourceRecord[]>();
  const sourceByNameKey = new Map<string, SourceRecord[]>();

  sourceSheet.rows.forEach((row) => {
    const nameKey = buildNameKey(rowValue(row.values, sourceNameColumn));
    if (!nameKey) return;
    const rank = normalizeRank(rowValue(row.values, sourceRankColumn));
    const strictKey = `${nameKey}|${rank}`;
    const record: SourceRecord = {
      row: row.values,
      columns: sourceColumns,
      weaponKind: getWeaponKind(rowValue(row.values, findColumn(sourceColumns, ["weaponName"]))),
    };

    sourceByStrictKey.set(strictKey, [...(sourceByStrictKey.get(strictKey) ?? []), record]);
    sourceByNameKey.set(nameKey, [...(sourceByNameKey.get(nameKey) ?? []), record]);
  });

  const changes: FillChange[] = [];
  let matchedRows = 0;
  let unmatchedRows = 0;
  let skippedFilled = 0;

  targetSheet.rows.forEach((targetRow) => {
    const targetName = rowValue(targetRow.values, targetNameColumn);
    const nameKey = buildNameKey(targetName);
    if (!nameKey) return;

    const rank = normalizeRank(rowValue(targetRow.values, targetRankColumn));
    const sourceRecords = sourceByStrictKey.get(`${nameKey}|${rank}`) ?? sourceByNameKey.get(nameKey);
    if (!sourceRecords?.length) {
      unmatchedRows += 1;
      return;
    }

    let rowChanged = false;
    targetColumns.forEach((targetColumn) => {
      if (["personName", "rank", "sequence"].includes(targetColumn.key)) return;
      const currentValue = targetRow.values[targetColumn.index];
      const canReplacePlaceholder =
        ["weaponName", "weaponNumber", "pistolName", "pistolNumber"].includes(targetColumn.key) &&
        isWeaponPlaceholder(currentValue);

      if (!isEmpty(currentValue) && !canReplacePlaceholder) {
        skippedFilled += 1;
        return;
      }

      const sourceRecord = pickSourceRecordForTargetColumn(sourceRecords, targetColumn);
      if (!sourceRecord) return;

      const sourceColumn = sourceColumns.find((column) => column.key === sourceKeyForTargetColumn(targetColumn));
      if (!sourceColumn) return;

      const sourceValue = sourceRecord.row[sourceColumn.index];
      const nextValue =
        targetColumn.key === "status"
          ? normalizeStatus(sourceValue)
          : sourceValue;
      if (isEmpty(nextValue)) return;

      changes.push({
        rowNumber: targetRow.excelRowNumber,
        columnIndex: targetColumn.index,
        originalColumnIndex: targetColumn.originalIndex,
        header: targetColumn.header || `Колонка ${targetColumn.originalIndex + 1}`,
        from: targetRow.values[targetColumn.index],
        to: nextValue,
        person: valueToDisplay(targetName),
      });
      rowChanged = true;
    });

    if (rowChanged) matchedRows += 1;
  });

  return { changes, matchedRows, unmatchedRows, skippedFilled, sourceColumns, targetColumns };
}

function sourceKeyForTargetColumn(targetColumn: ColumnMeta) {
  if (targetColumn.key === "pistolName") return "weaponName";
  if (targetColumn.key === "pistolNumber") return "weaponNumber";
  return targetColumn.key;
}

function pickSourceRecordForTargetColumn(records: SourceRecord[], targetColumn: ColumnMeta) {
  const preferredKind =
    targetColumn.key === "pistolName" || targetColumn.key === "pistolNumber"
      ? "pistol"
      : targetColumn.key === "weaponName" || targetColumn.key === "weaponNumber"
        ? "rifle"
        : undefined;

  const candidates = preferredKind
    ? records.filter((record) => record.weaponKind === preferredKind)
    : records;

  return candidates.sort((a, b) => sourceRowScore(b) - sourceRowScore(a))[0];
}

const setCell = (sheet: any, row: number, column: number, value: CellValue) => {
  sheet.cell(row, column).value(value ?? "");
};

const styleRange = (sheet: any, range: string, style: Record<string, unknown>) => {
  sheet.range(range).style(style);
};

const morningThinBorder = {
  style: "thin",
  color: "B7B7B7",
};

const morningMediumBorder = {
  style: "medium",
  color: "000000",
};

const MORNING_REPORT_COLUMN_WIDTHS = [8, 18, 34, 16, 20, 24, 28, 16, 24, 14, 10, 10, 10, 14];
const MORNING_REPORT_MIN_COLUMN_WIDTHS = [8, 28, 34, 16, 20, 24, 28, 16, 24, 14, 10, 10, 10, 14];
const MORNING_REPORT_MAX_COLUMN_WIDTHS = [8, 42, 42, 18, 24, 30, 42, 22, 34, 16, 10, 10, 10, 16];

const estimateTextWidth = (value: CellValue) =>
  Math.max(
    0,
    ...valueToDisplay(value)
      .split(/\r?\n/)
      .map((part) => part.trim().length),
  );

const calculateMorningColumnWidths = (rows: CellValue[][]) =>
  MORNING_REPORT_MIN_COLUMN_WIDTHS.map((minWidth, index) => {
    const maxTextWidth = Math.max(...rows.map((row) => estimateTextWidth(row[index])));
    const paddedWidth = Math.ceil(maxTextWidth * 1.08) + 2;
    return Math.min(MORNING_REPORT_MAX_COLUMN_WIDTHS[index] ?? 32, Math.max(minWidth, paddedWidth));
  });

const estimateMorningRowHeight = (values: CellValue[], columnWidths: number[]) => {
  const maxLines = values.reduce<number>((max, value, index) => {
    const text = valueToDisplay(value).trim();
    if (!text) return max;
    const columnWidth = columnWidths[index] ?? MORNING_REPORT_COLUMN_WIDTHS[index] ?? 16;
    const charsPerLine = Math.max(8, Math.floor(columnWidth * 1.15));
    const lines = text
      .split(/\r?\n/)
      .reduce((sum, part) => sum + Math.max(1, Math.ceil(part.length / charsPerLine)), 0);
    return Math.max(max, lines);
  }, 1);

  return Math.min(96, Math.max(22, maxLines * 17));
};

function getMorningLocationCounts(rows: MorningReportRow[]) {
  const counts = new Map<string, number>();
  rows.forEach((row) => {
    const location = row.location.trim() || "(пусто)";
    counts.set(location, (counts.get(location) ?? 0) + 1);
  });

  return [...counts.entries()].sort((left, right) => {
    if (left[0] === "(пусто)") return 1;
    if (right[0] === "(пусто)") return -1;
    return right[1] - left[1] || left[0].localeCompare(right[0], "uk");
  });
}

function writeMorningCompositionSheet(workbook: any, result: MorningReportResult) {
  const sheet = workbook.addSheet("Чисельний склад");
  const locationCounts = getMorningLocationCounts(result.rows);

  setCell(sheet, 1, 1, "Кількість");
  setCell(sheet, 2, 2, "(пусто)");
  setCell(sheet, 2, 3, "Загальний підсумок");
  setCell(sheet, 3, 1, "_5 1ПБ");
  setCell(sheet, 3, 2, result.rows.length);
  setCell(sheet, 3, 3, result.rows.length);

  locationCounts.forEach(([location, count], index) => {
    const row = index + 4;
    setCell(sheet, row, 1, location);
    setCell(sheet, row, 2, count);
    setCell(sheet, row, 3, count);
  });

  const totalRow = locationCounts.length + 4;
  setCell(sheet, totalRow, 1, "Загальний підсумок");
  setCell(sheet, totalRow, 2, result.rows.length);
  setCell(sheet, totalRow, 3, result.rows.length);

  styleRange(sheet, "A1:C1", {
    bold: true,
    horizontalAlignment: "center",
    border: morningThinBorder,
  });
  styleRange(sheet, `A2:C${totalRow}`, {
    border: morningThinBorder,
    horizontalAlignment: "center",
    verticalAlignment: "top",
    wrapText: true,
  });
  styleRange(sheet, "A2:C3", {
    bold: true,
    horizontalAlignment: "center",
  });
  styleRange(sheet, `A${totalRow}:C${totalRow}`, {
    bold: true,
    horizontalAlignment: "center",
  });
  sheet.row(1).height(18);
  sheet.row(2).height(18);
  sheet.row(3).height(18);
  sheet.column(1).width(32);
  sheet.column(2).width(14);
  sheet.column(3).width(18);
}

function writeMorningReportWorkbook(workbook: any, result: MorningReportResult) {
  const sheet = workbook.sheet(0);
  sheet.name("Поіменний");

  const summaryHeaders = ["Всього", ...MORNING_REPORT_STATUSES];
  summaryHeaders.forEach((header, index) => {
    setCell(sheet, 1, index + 1, header);
    setCell(
      sheet,
      2,
      index + 1,
      index === 0 ? result.rows.length : result.counts[header] ?? 0,
    );
  });

  const tableHeaders = [
    "№ зп",
    "Підрозділ за штатом",
    "ПІБ",
    "Позивний",
    "Підрозділ фактичний",
    "Локація",
    "Чим займається",
    "Статус",
    "Примітка",
    "Хто подав",
  ];
  tableHeaders.forEach((header, index) => setCell(sheet, 4, index + 1, header));

  const dataRows = result.rows.map((row) => [
    row.sequence,
    row.staffUnit,
    row.name,
    row.callsign,
    row.actualUnit,
    row.location,
    row.activity,
    row.reportStatus,
    row.note,
    row.submittedBy,
  ]);
  const columnWidths = calculateMorningColumnWidths([summaryHeaders, tableHeaders, ...dataRows]);

  result.rows.forEach((_, index) => {
    const excelRow = index + 5;
    const rowValues = dataRows[index];
    rowValues.forEach((value, columnIndex) => setCell(sheet, excelRow, columnIndex + 1, value));
    sheet.row(excelRow).height(estimateMorningRowHeight(rowValues, columnWidths));
  });

  const lastRow = Math.max(result.rows.length + 4, 5);
  const lastSummaryColumn = summaryHeaders.length;
  styleRange(sheet, `A1:${sheet.cell(1, lastSummaryColumn).address()}`, {
    bold: true,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    border: morningThinBorder,
    wrapText: true,
  });
  styleRange(sheet, `A2:${sheet.cell(2, lastSummaryColumn).address()}`, {
    horizontalAlignment: "center",
    verticalAlignment: "center",
    border: morningThinBorder,
  });
  styleRange(sheet, "A4:J4", {
    bold: true,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    border: morningThinBorder,
    wrapText: true,
  });
  styleRange(sheet, `A5:J${lastRow}`, {
    border: morningThinBorder,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  });
  styleRange(sheet, `A2:${sheet.cell(2, lastSummaryColumn).address()}`, {
    bottomBorder: morningMediumBorder,
  });
  sheet.range("A4:J4").autoFilter();
  sheet.freezePanes(5, 1);
  sheet.row(1).height(20);
  sheet.row(2).height(18);
  sheet.row(3).height(10);
  sheet.row(4).height(20);
  columnWidths.forEach((width, index) => {
    sheet.column(index + 1).width(width);
  });

  writeMorningCompositionSheet(workbook, result);

  const statusSheet = workbook.addSheet("Статуси");
  MORNING_REPORT_STATUS_DESCRIPTIONS.forEach(([status, description], index) => {
    setCell(statusSheet, index + 1, 1, status);
    setCell(statusSheet, index + 1, 2, description);
  });
  styleRange(statusSheet, `A1:B${MORNING_REPORT_STATUS_DESCRIPTIONS.length}`, {
    border: true,
    wrapText: true,
    verticalAlignment: "top",
  });
  statusSheet.column(1).width(34);
  statusSheet.column(2).width(90);
}

export function ExcelFillPage() {
  const [source, setSource] = useState<ExcelWorkbookSnapshot | null>(null);
  const [target, setTarget] = useState<ExcelWorkbookSnapshot | null>(null);
  const [morningIdSource, setMorningIdSource] = useState<ExcelWorkbookSnapshot | null>(null);
  const [morningSubmittedBy, setMorningSubmittedBy] = useState("");
  const [isBusy, setIsBusy] = useState(false);
  const [message, setMessage] = useState("Завантажте файл-джерело і файл, який треба дозаповнити.");
  const result = useMemo(() => analyzeFill(source, target), [source, target]);
  const morningSubmittedByValue = morningSubmittedBy.trim();
  const morningIdLookup = useMemo(
    () => buildMorningIdLookup(morningIdSource),
    [morningIdSource],
  );
  const morningReport = useMemo(
    () => analyzeMorningReport(source, morningSubmittedByValue, morningIdLookup),
    [morningIdLookup, morningSubmittedByValue, source],
  );
  const morningLocationCounts = useMemo(
    () => (morningReport ? getMorningLocationCounts(morningReport.rows) : []),
    [morningReport],
  );

  const loadFile = async (file: File | undefined, role: "source" | "target" | "idSource") => {
    if (!file) return;
    setIsBusy(true);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      if (role === "source") setSource(snapshot);
      else if (role === "target") setTarget(snapshot);
      else setMorningIdSource(snapshot);
      setMessage(`Завантажено ${file.name}.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Не вдалося прочитати Excel.");
    } finally {
      setIsBusy(false);
    }
  };

  const exportFilled = async () => {
    if (!target || !result?.changes.length) return;
    setIsBusy(true);
    try {
      await exportWorkbookFileWithMutations(
        target.file,
        (workbook) => {
          const sheet = workbook.sheet(0);
          result.changes.forEach((change) => {
            sheet.cell(change.rowNumber, change.originalColumnIndex + 1).value(change.to);
          });
        },
        `${target.fileName.replace(/\.xlsx$/i, "")}_filled.xlsx`,
      );
      setMessage(`Експортовано: заповнено клітинок ${result.changes.length}.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Не вдалося експортувати Excel.");
    } finally {
      setIsBusy(false);
    }
  };

  const exportMorningReport = async () => {
    if (!morningReport?.rows.length) return;
    if (!morningSubmittedByValue) {
      setMessage("Введіть позивний у полі «Хто подав» для ранкового звіту.");
      return;
    }
    setIsBusy(true);
    try {
      await exportBlankWorkbookWithMutations(
        (workbook) => writeMorningReportWorkbook(workbook, morningReport),
        "1ПБ с id.xlsx",
      );
      setMessage(`Експортовано Ранковий звіт 1ПБ: ${morningReport.rows.length} осіб.`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : "Не вдалося експортувати ранковий звіт.");
    } finally {
      setIsBusy(false);
    }
  };

  return (
    <main className="main-panel excel-fill-page">
      <header className="topbar">
        <Box>
          <Typography component="h1" variant="h5">Заповнення Excel</Typography>
          <Typography variant="body2" color="text.secondary">
            Джерело → цільова таблиця: пошук по ПІБ і званню, дозаповнення порожніх клітинок
          </Typography>
        </Box>
        <Button
          disabled={!result?.changes.length || isBusy}
          startIcon={<FileDownloadOutlinedIcon />}
          variant="contained"
          onClick={() => void exportFilled()}
        >
          Експортувати
        </Button>
        <Button
          disabled={!morningReport?.rows.length || !morningSubmittedByValue || isBusy}
          startIcon={<FileDownloadOutlinedIcon />}
          variant="outlined"
          onClick={() => void exportMorningReport()}
        >
          Ранковий звіт 1ПБ
        </Button>
      </header>

      <div className="dashboard-grid">
        <section className="panel">
          <div className="panel-heading">1. Файл-джерело</div>
          <Typography variant="body2" color="text.secondary">
            Таблиця, звідки беремо дані. Наприклад, 1ПБ по ЗББР.
          </Typography>
          <Stack direction="row" spacing={1} sx={{ mt: 2, alignItems: "center" }}>
            <Button component="label" startIcon={<FileUploadOutlinedIcon />}>
              Завантажити
              <input hidden type="file" accept=".xlsx" onChange={(event) => void loadFile(event.target.files?.[0], "source")} />
            </Button>
            <Typography variant="body2">{source?.fileName ?? "Файл не вибрано"}</Typography>
          </Stack>
        </section>

        <section className="panel">
          <div className="panel-heading">2. Таблиця для заповнення</div>
          <Typography variant="body2" color="text.secondary">
            Саме цей файл буде експортований з дозаповненими пропусками.
          </Typography>
          <Stack direction="row" spacing={1} sx={{ mt: 2, alignItems: "center" }}>
            <Button component="label" startIcon={<FileUploadOutlinedIcon />}>
              Завантажити
              <input hidden type="file" accept=".xlsx" onChange={(event) => void loadFile(event.target.files?.[0], "target")} />
            </Button>
            <Typography variant="body2">{target?.fileName ?? "Файл не вибрано"}</Typography>
          </Stack>
        </section>

        <section className="panel">
          <div className="panel-heading">Ранковий звіт 1ПБ</div>
          <Stack spacing={1}>
            <Typography variant="body2">
              Для експорту беруться статуси: В строю, Відком. за межі ПБ,
              Відрядження, Новоприбулий.
            </Typography>
            <label className="excel-fill-inline-field">
              <span>Хто подав / позивний</span>
              <input
                value={morningSubmittedBy}
                onChange={(event) => setMorningSubmittedBy(toUppercaseCallsign(event.target.value))}
                placeholder="Наприклад: БУКЛЯ"
              />
            </label>
            <Stack direction="row" spacing={1} sx={{ alignItems: "center" }}>
              <Button component="label" startIcon={<FileUploadOutlinedIcon />}>
                Файл з ID
                <input hidden type="file" accept=".xlsx,.xlsm" onChange={(event) => void loadFile(event.target.files?.[0], "idSource")} />
              </Button>
              <Typography variant="body2">{morningIdSource?.fileName ?? "ID-файл не вибрано"}</Typography>
            </Stack>
            <Typography variant="body2">Осіб у звіті: {morningReport?.rows.length ?? 0}</Typography>
            <Typography variant="body2">ID у довіднику: {morningIdLookup.size}</Typography>
            <Typography variant="body2">Без знайденого ID: {morningReport?.missingIdRows ?? 0}</Typography>
            <Typography variant="body2">Пропущено інших статусів: {morningReport?.skippedRows ?? 0}</Typography>
            <Typography variant="body2">
              Локацій у чисельному складі: {morningLocationCounts.length}
            </Typography>
            <Button
              disabled={!morningReport?.rows.length || !morningSubmittedByValue || isBusy}
              startIcon={<FileDownloadOutlinedIcon />}
              variant="outlined"
              onClick={() => void exportMorningReport()}
            >
              Експортувати Ранковий звіт 1ПБ
            </Button>
          </Stack>
        </section>

        <section className="panel">
          <div className="panel-heading">Результат</div>
          <Stack spacing={1}>
            <Typography variant="body2">{message}</Typography>
            <Typography variant="body2">Збігів з оновленнями: {result?.matchedRows ?? 0}</Typography>
            <Typography variant="body2">Клітинок до заповнення: {result?.changes.length ?? 0}</Typography>
            <Typography variant="body2">Не знайдено людей: {result?.unmatchedRows ?? 0}</Typography>
          </Stack>
        </section>
      </div>

      <section className="panel excel-fill-preview-panel">
        <div className="panel-heading">
          <SyncAltOutlinedIcon fontSize="small" /> Preview змін
        </div>
        <div className="excel-fill-preview-wrap">
          <table className="excel-preview-table">
            <thead>
              <tr>
                <th>Рядок</th>
                <th>ПІБ</th>
                <th>Колонка</th>
                <th>Буде записано</th>
              </tr>
            </thead>
            <tbody>
              {(result?.changes.slice(0, 300) ?? []).map((change, index) => (
                <tr key={`${change.rowNumber}-${change.columnIndex}-${index}`}>
                  <td>{change.rowNumber}</td>
                  <td>{change.person}</td>
                  <td>{change.header}</td>
                  <td>{valueToDisplay(change.to)}</td>
                </tr>
              ))}
              {!result?.changes.length ? (
                <tr>
                  <td colSpan={4}>Поки немає змін для експорту.</td>
                </tr>
              ) : null}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel excel-fill-preview-panel">
        <div className="panel-heading">
          <FileDownloadOutlinedIcon fontSize="small" /> Preview Ранковий звіт 1ПБ
        </div>
        <div className="excel-fill-preview-wrap">
          <table className="excel-preview-table">
            <thead>
              <tr>
                <th>№</th>
                <th>ПІБ</th>
                <th>Локація</th>
                <th>Чим займається</th>
                <th>Статус звіту</th>
                <th>Статус джерела</th>
              </tr>
            </thead>
            <tbody>
              {(morningReport?.rows.slice(0, 300) ?? []).map((row) => (
                <tr key={`${row.sourceRowNumber}-${row.sequence}`}>
                  <td>{row.sequence}</td>
                  <td>{row.name}</td>
                  <td>{row.location}</td>
                  <td>{row.activity}</td>
                  <td>{row.reportStatus}</td>
                  <td>{row.sourceStatus}</td>
                </tr>
              ))}
              {!morningReport?.rows.length ? (
                <tr>
                  <td colSpan={6}>Завантажте першу таблицю, щоб сформувати ранковий звіт.</td>
                </tr>
              ) : null}
            </tbody>
          </table>
        </div>
      </section>
    </main>
  );
}
