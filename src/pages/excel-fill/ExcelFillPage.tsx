import { useEffect, useMemo, useState } from "react";
import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import {
  FileDownloadOutlinedIcon,
  FileUploadOutlinedIcon,
  SyncAltOutlinedIcon,
} from "@/components/sci/icons";
import { api } from "../../api";
import {
  type CellValue,
  type ExcelSheetSnapshot,
  type ExcelWorkbookSnapshot,
  exportTemplateWorkbookWithMutations,
  exportWorkbookFileWithMutations,
  getColumnHeader,
  readWorkbookSnapshot,
  valueToDisplay,
} from "../../excelRoundTrip";
import {
  analyzePositionsVozFill,
  DEFAULT_POSITIONS_VOZ_RULES,
  loadPositionsVozRules,
  parseRosterLatestToPeople,
  savePositionsVozRules,
  type PositionsVozFillRule,
  type PositionsVozPerson,
} from "./positionsVozFill";

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

const isMorningTransiterNote = (note: CellValue | string) => {
  const normalized = normalizeReportText(note);
  return (
    normalized === "транзитер" ||
    normalized === "тринзитер" ||
    normalized.includes("транзитер")
  );
};

const findMorningSheet = (workbook: ExcelWorkbookSnapshot) =>
  workbook.sheets.find((sheet) => normalizeReportText(sheet.sheetName).includes("ос загальний")) ??
  workbook.sheets[0];

const getMorningValue = (row: CellValue[], columnNumber: number) =>
  valueToDisplay(row[columnNumber - 1]).trim();

const findMorningReportTable = (sheet: ExcelSheetSnapshot) => {
  for (const excelRow of sheet.rows.slice(0, 12)) {
    const headers = excelRow.values.map((value) => normalizeReportText(value));
    const hasName = headers.some((header) => header === "піб" || header === "пип");
    const hasStatus = headers.some((header) => header === "статус");
    const hasLocation = headers.some((header) => header.includes("локац"));
    if (!hasName || !hasStatus || !hasLocation) continue;

    const findHeader = (patterns: RegExp[]) =>
      headers.findIndex((header) => patterns.some((pattern) => pattern.test(header)));

    return {
      headerRowNumber: excelRow.excelRowNumber,
      sequence: findHeader([/^№/, /номер/]),
      name: findHeader([/^піб$/, /^пип$/]),
      callsign: findHeader([/позив/]),
      actualUnit: findHeader([/підрозділ фактич/]),
      location: findHeader([/локац/]),
      activity: findHeader([/чим займа/]),
      reportStatus: findHeader([/^статус$/]),
      note: findHeader([/приміт/]),
    };
  }

  return null;
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
): MorningReportResult | null {
  const sourceSheet = source ? findMorningSheet(source) : undefined;
  if (!sourceSheet) return null;

  const rows: MorningReportRow[] = [];
  let skippedRows = 0;
  const existingReportTable = findMorningReportTable(sourceSheet);

  if (existingReportTable) {
    sourceSheet.rows
      .filter((excelRow) => excelRow.excelRowNumber > existingReportTable.headerRowNumber)
      .forEach((excelRow) => {
        const row = excelRow.values;
        const name = valueToDisplay(row[existingReportTable.name]).trim();
        if (!name) return;
        const reportStatus = valueToDisplay(row[existingReportTable.reportStatus]).trim();

        const note =
          existingReportTable.note >= 0
            ? valueToDisplay(row[existingReportTable.note]).trim()
            : "";
        if (isMorningTransiterNote(note)) {
          skippedRows += 1;
          return;
        }

        rows.push({
          sourceRowNumber: excelRow.excelRowNumber,
          sequence: rows.length + 1,
          staffUnit: "",
          name,
          callsign:
            existingReportTable.callsign >= 0
              ? valueToDisplay(row[existingReportTable.callsign]).trim()
              : "",
          actualUnit:
            existingReportTable.actualUnit >= 0
              ? valueToDisplay(row[existingReportTable.actualUnit]).trim()
              : "_5 1ПБ",
          location:
            existingReportTable.location >= 0
              ? valueToDisplay(row[existingReportTable.location]).trim()
              : "",
          activity:
            existingReportTable.activity >= 0
              ? valueToDisplay(row[existingReportTable.activity]).trim()
              : "",
          reportStatus: reportStatus || "Забезпечення",
          sourceStatus: reportStatus || "—",
          note,
          submittedBy,
        });
      });

    const counts = Object.fromEntries(MORNING_REPORT_STATUSES.map((status) => [status, 0]));
    rows.forEach((row) => {
      counts[row.reportStatus] = (counts[row.reportStatus] ?? 0) + 1;
    });

    return { rows, counts, skippedRows };
  }

  sourceSheet.rows.forEach((excelRow) => {
    const row = excelRow.values;
    const name = getMorningValue(row, 14);
    const sourceStatus = getMorningValue(row, 21);
    if (!name || !MORNING_REPORT_SOURCE_STATUSES.has(sourceStatus)) {
      if (name) skippedRows += 1;
      return;
    }

    const note = getMorningValue(row, 32) || getMorningValue(row, 34);
    if (isMorningTransiterNote(note)) {
      skippedRows += 1;
      return;
    }

    rows.push({
      sourceRowNumber: excelRow.excelRowNumber,
      sequence: rows.length + 1,
      staffUnit: "",
      name,
      callsign: getMorningValue(row, 15),
      actualUnit: "_5 1ПБ",
      location: makeMorningLocation(row),
      activity: makeMorningActivity(row),
      reportStatus: classifyMorningReportStatus(row),
      sourceStatus,
      note,
      submittedBy,
    });
  });

  const counts = Object.fromEntries(MORNING_REPORT_STATUSES.map((status) => [status, 0]));
  rows.forEach((row) => {
    counts[row.reportStatus] = (counts[row.reportStatus] ?? 0) + 1;
  });

  return { rows, counts, skippedRows };
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

const MORNING_REPORT_TEMPLATE_URL = "/templates/Поіменний список 1 ПБ.xlsm";
const MORNING_REPORT_DATA_START_ROW = 5;
const MORNING_REPORT_DATA_COLUMN_COUNT = 10;
const MORNING_REPORT_TEMPLATE_CLEAR_UNTIL_ROW = 600;

const setMorningCell = (sheet: any, row: number, column: number, value: CellValue) => {
  sheet.cell(row, column).value(value ?? "");
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

/** Лише підставляє значення в шаблон .xlsm — формат/листи/зв'язки не чіпаємо. */
function writeMorningReportIntoTemplate(workbook: any, result: MorningReportResult) {
  const sheet =
    workbook.sheet("Поіменний") ??
    workbook.sheets().find((item: any) => /поімен/i.test(String(item.name()))) ??
    workbook.sheet(0);

  // Рядок 2 — формули шаблону (COUNTIF по таблиці «Поіменний_Список»). Не перезаписуємо.

  result.rows.forEach((row, index) => {
    const excelRow = MORNING_REPORT_DATA_START_ROW + index;
    setMorningCell(sheet, excelRow, 1, row.sequence);
    setMorningCell(sheet, excelRow, 2, row.staffUnit);
    setMorningCell(sheet, excelRow, 3, row.name);
    setMorningCell(sheet, excelRow, 4, row.callsign);
    setMorningCell(sheet, excelRow, 5, row.actualUnit);
    setMorningCell(sheet, excelRow, 6, row.location);
    setMorningCell(sheet, excelRow, 7, row.activity);
    setMorningCell(sheet, excelRow, 8, row.reportStatus);
    setMorningCell(sheet, excelRow, 9, row.note);
    setMorningCell(sheet, excelRow, 10, row.submittedBy);
  });

  const firstClearRow = MORNING_REPORT_DATA_START_ROW + result.rows.length;
  for (let row = firstClearRow; row <= MORNING_REPORT_TEMPLATE_CLEAR_UNTIL_ROW; row += 1) {
    for (let column = 1; column <= MORNING_REPORT_DATA_COLUMN_COUNT; column += 1) {
      setMorningCell(sheet, row, column, "");
    }
  }
}

export function ExcelFillPage() {
  const [source, setSource] = useState<ExcelWorkbookSnapshot | null>(null);
  const [target, setTarget] = useState<ExcelWorkbookSnapshot | null>(null);
  const [positionsTemplate, setPositionsTemplate] =
    useState<ExcelWorkbookSnapshot | null>(null);
  const [positionsPeople, setPositionsPeople] = useState<PositionsVozPerson[]>(
    [],
  );
  const [positionsRules, setPositionsRules] = useState<PositionsVozFillRule[]>(
    () => loadPositionsVozRules(),
  );
  const [positionsRulesText, setPositionsRulesText] = useState(() =>
    JSON.stringify(loadPositionsVozRules(), null, 2),
  );
  const [positionsRosterLabel, setPositionsRosterLabel] = useState("");
  const [morningSubmittedBy, setMorningSubmittedBy] = useState("");
  const [isBusy, setIsBusy] = useState(false);
  const [message, setMessage] = useState(
    "Завантажте файл-джерело і файл, який треба дозаповнити.",
  );
  const result = useMemo(() => analyzeFill(source, target), [source, target]);
  const positionsResult = useMemo(
    () =>
      analyzePositionsVozFill(
        positionsTemplate,
        positionsPeople,
        positionsRules,
      ),
    [positionsPeople, positionsRules, positionsTemplate],
  );
  const morningSubmittedByValue = morningSubmittedBy.trim();
  const morningReport = useMemo(
    () => analyzeMorningReport(source, morningSubmittedByValue),
    [morningSubmittedByValue, source],
  );
  const morningLocationCounts = useMemo(
    () => (morningReport ? getMorningLocationCounts(morningReport.rows) : []),
    [morningReport],
  );

  useEffect(() => {
    setPositionsRulesText(JSON.stringify(positionsRules, null, 2));
  }, [positionsRules]);

  const loadFile = async (
    file: File | undefined,
    role: "source" | "target" | "positions",
  ) => {
    if (!file) return;
    setIsBusy(true);
    try {
      const snapshot = await readWorkbookSnapshot(file);
      if (role === "source") setSource(snapshot);
      else if (role === "target") setTarget(snapshot);
      else setPositionsTemplate(snapshot);
      setMessage(`Завантажено ${file.name}.`);
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося прочитати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const loadPositionsFromDb = async () => {
    setIsBusy(true);
    try {
      const latest = await api.getLatestPersonnelRoster();
      if (!latest?.sheet) {
        setPositionsPeople([]);
        setPositionsRosterLabel("");
        setMessage(
          "У БД немає імпортованого «Загального списку». Спочатку імпортуйте ранковий звіт у персонал.",
        );
        return;
      }
      const people = parseRosterLatestToPeople(latest);
      setPositionsPeople(people);
      setPositionsRosterLabel(
        latest.sourceFileName || latest.importName || "Загальний список",
      );
      setMessage(
        `З БД завантажено ${people.length} осіб · ${latest.sourceFileName || latest.importName}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося завантажити БД: ${error.message}`
          : "Не вдалося завантажити БД.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const applyPositionsRulesText = () => {
    try {
      const parsed = JSON.parse(positionsRulesText) as PositionsVozFillRule[];
      if (!Array.isArray(parsed) || !parsed.length) {
        setMessage("Правила мають бути непорожнім JSON-масивом.");
        return;
      }
      const next = parsed.map((rule, index) => ({
        id: String(rule.id || `rule_${index + 1}`),
        enabled: Boolean(rule.enabled),
        targetColumn: Number(rule.targetColumn) || 1,
        label: String(rule.label || `Колонка ${rule.targetColumn || index + 1}`),
        source: rule.source,
        customKeys: Array.isArray(rule.customKeys)
          ? rule.customKeys.map(String)
          : [],
        onlyIfEmpty: Boolean(rule.onlyIfEmpty),
      }));
      setPositionsRules(next);
      savePositionsVozRules(next);
      setMessage(
        `Правила оновлено: ${next.filter((rule) => rule.enabled).length} активних.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Невалідний JSON правил: ${error.message}`
          : "Невалідний JSON правил.",
      );
    }
  };

  const resetPositionsRules = () => {
    const next = structuredClone(DEFAULT_POSITIONS_VOZ_RULES);
    setPositionsRules(next);
    savePositionsVozRules(next);
    setMessage("Правила скинуто до стандартних.");
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
            sheet
              .cell(change.rowNumber, change.originalColumnIndex + 1)
              .value(change.to);
          });
        },
        `${target.fileName.replace(/\.xlsx$/i, "")}_filled.xlsx`,
      );
      setMessage(`Експортовано: заповнено клітинок ${result.changes.length}.`);
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportPositionsFilled = async () => {
    if (!positionsTemplate || !positionsResult?.changes.length) return;
    setIsBusy(true);
    try {
      await exportWorkbookFileWithMutations(
        positionsTemplate.file,
        (workbook) => {
          const sheet = workbook.sheet(0);
          positionsResult.changes.forEach((change) => {
            sheet.cell(change.rowNumber, change.column).value(change.to);
          });
        },
        `${positionsTemplate.fileName.replace(/\.xlsx$/i, "")}_з_бд.xlsx`,
      );
      setMessage(
        `Експортовано «Посади, ВОЗ»: ${positionsResult.changes.length} клітинок.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати «Посади, ВОЗ».",
      );
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
      await exportTemplateWorkbookWithMutations(
        MORNING_REPORT_TEMPLATE_URL,
        (workbook) => writeMorningReportIntoTemplate(workbook, morningReport),
        "Поіменний список 1 ПБ.xlsm",
      );
      setMessage(
        `Експортовано Ранковий звіт 1ПБ: ${morningReport.rows.length} осіб.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати ранковий звіт.",
      );
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
            Джерело → цільова таблиця: пошук по ПІБ і званню, дозаповнення порожніх
            клітинок. Окремо: «Посади, ВОЗ» з БД та змінними правилами.
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
              <input hidden type="file" accept=".xlsx,.xlsm" onChange={(event) => void loadFile(event.target.files?.[0], "source")} />
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
            <Typography variant="body2" color="text.secondary">
              ID-файл більше не потрібен: колонка «Підрозділ за штатом»
              лишається пустою.
            </Typography>
            <label className="excel-fill-inline-field">
              <span>Хто подав / позивний</span>
              <input
                value={morningSubmittedBy}
                onChange={(event) => setMorningSubmittedBy(toUppercaseCallsign(event.target.value))}
                placeholder="Наприклад: БУКЛЯ"
              />
            </label>
            <Typography variant="body2">Осіб у звіті: {morningReport?.rows.length ?? 0}</Typography>
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

      <section className="panel excel-fill-positions-panel">
        <div className="panel-heading">Посади, ВОЗ · з БД</div>
        <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5 }}>
          Завантажте шаблон «Посади, ВОЗ.xlsx», підтягніть людей з останнього
          «Загального списку» в БД і заповніть колонки за правилами нижче.
          Правила можна міняти (JSON) без зміни коду — зберігаються в браузері.
        </Typography>
        <div className="excel-fill-positions-grid">
          <Stack spacing={1.25}>
            <Stack
              direction="row"
              spacing={1}
              sx={{ alignItems: "center", flexWrap: "wrap" }}
            >
              <Button
                component="label"
                startIcon={<FileUploadOutlinedIcon />}
                disabled={isBusy}
              >
                Шаблон Excel
                <input
                  hidden
                  type="file"
                  accept=".xlsx,.xlsm"
                  onChange={(event) =>
                    void loadFile(event.target.files?.[0], "positions")
                  }
                />
              </Button>
              <Typography variant="body2">
                {positionsTemplate?.fileName ?? "Файл не вибрано"}
              </Typography>
            </Stack>
            <Stack
              direction="row"
              spacing={1}
              sx={{ alignItems: "center", flexWrap: "wrap" }}
            >
              <Button
                variant="outlined"
                disabled={isBusy}
                startIcon={<SyncAltOutlinedIcon />}
                onClick={() => void loadPositionsFromDb()}
              >
                Завантажити з БД
              </Button>
              <Typography variant="body2">
                {positionsRosterLabel
                  ? `${positionsPeople.length} осіб · ${positionsRosterLabel}`
                  : "БД ще не завантажена"}
              </Typography>
            </Stack>
            <Typography variant="body2">
              Збігів у шаблоні: {positionsResult?.matchedRows ?? 0} · змін:{" "}
              {positionsResult?.changes.length ?? 0} · без збігу в шаблоні:{" "}
              {positionsResult?.unmatchedTemplateRows ?? 0} · у БД без рядка в
              шаблоні: {positionsResult?.unmatchedDbPeople ?? 0}
            </Typography>
            <Button
              variant="contained"
              disabled={!positionsResult?.changes.length || isBusy}
              startIcon={<FileDownloadOutlinedIcon />}
              onClick={() => void exportPositionsFilled()}
              sx={{ alignSelf: "flex-start", color: "#1a1a14" }}
            >
              Експортувати заповнений файл
            </Button>
          </Stack>

          <label className="excel-fill-rules-editor">
            <span>Правила заповнення (JSON)</span>
            <textarea
              value={positionsRulesText}
              onChange={(event) => setPositionsRulesText(event.target.value)}
              spellCheck={false}
            />
            <Stack direction="row" spacing={1} sx={{ mt: 1 }}>
              <Button variant="outlined" onClick={applyPositionsRulesText}>
                Застосувати правила
              </Button>
              <Button variant="text" onClick={resetPositionsRules}>
                Скинути стандарт
              </Button>
            </Stack>
            <Typography
              variant="body2"
              color="text.secondary"
              sx={{ mt: 1 }}
            >
              source: rank | name | callsign | vos | position | fullPosition |
              positionGroup | roleType | customKeys. onlyIfEmpty — писати лише в
              порожні клітинки. targetColumn — номер колонки Excel (1…). Для
              fullPosition: якщо посади немає або там дата → Підрозділ
              (напр. БРЕЗ); якщо людини немає в БД або посади немає взагалі →
              «відсутній у списку».
            </Typography>
          </label>
        </div>
      </section>

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
        <div className="panel-heading">Preview · Посади / ВОЗ</div>
        <div className="excel-fill-preview-wrap">
          <table className="excel-preview-table">
            <thead>
              <tr>
                <th>Рядок</th>
                <th>ПІБ</th>
                <th>Колонка</th>
                <th>Було</th>
                <th>Стане</th>
              </tr>
            </thead>
            <tbody>
              {(positionsResult?.changes.slice(0, 300) ?? []).map(
                (change, index) => (
                  <tr key={`pos-${change.rowNumber}-${change.column}-${index}`}>
                    <td>{change.rowNumber}</td>
                    <td>{change.person}</td>
                    <td>
                      {change.label} ({change.column})
                    </td>
                    <td>{change.from || "—"}</td>
                    <td>{change.to}</td>
                  </tr>
                ),
              )}
              {!positionsResult?.changes.length ? (
                <tr>
                  <td colSpan={5}>
                    Завантажте шаблон «Посади, ВОЗ» і дані з БД, щоб побачити
                    зміни.
                  </td>
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
