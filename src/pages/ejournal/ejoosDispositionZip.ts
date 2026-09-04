import type {
  CellValue,
  ExcelSheetSnapshot,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  parseEjoosShpo,
  personIdFromShpo,
  type EjoosSyncOp,
  type EjoosSyncPlan,
} from "./ejoosSyncPlan";
import { canonicalName, normId } from "./ejoosIdentity";
import {
  dayFromOrderLabel,
  formatDispositionTimesheetDeparture,
  formatTimesheetDeparture,
  formatTimesheetMonthHeader,
  parseTimesheetAbsenceSpans,
  parseTimesheetMonthHeaderText,
  timesheetMarkForOpenDispositionDay,
  timesheetMarkFromArchive,
  timesheetMonthHeaderTextFromCell,
} from "./ejoosTimesheetText";
import {
  applyInlineStringWritesToWorkbook,
  type ZipCellWrite,
} from "./ejoosZipCellWrites";
import {
  TimesheetLayoutError,
  assertTimesheetLayoutReadyForApply,
  buildTimesheetLayout,
  resolveCanonicalTimesheetSlot,
  resolveHistoryTimesheetRow,
  stampTimesheetHistoryInserts,
  takeHistoryTimesheetRow,
  type PendingHistoryInsert,
  type TimesheetLayout,
} from "./ejoosTimesheetLayout";
import {
  loadTimesheetSheetArtifacts,
  mergeTimesheetGrids,
  placementSheetFromMergedGrid,
} from "./ejoosTimesheetPersonRows";

const valueOf = (value: CellValue | unknown): string | number | null => {
  if (value === undefined || value === null || value === "") return null;
  if (typeof value === "number" || typeof value === "string") return value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Math.floor(
      (Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) -
        Date.UTC(1899, 11, 30)) /
        86400000,
    );
  }
  const rich = value as { text?: () => string; value?: () => unknown };
  if (typeof rich.text === "function") return rich.text() || null;
  if (typeof rich.value === "function") return valueOf(rich.value());
  return null;
};

const findSheet = (book: ExcelWorkbookSnapshot, pattern: RegExp) =>
  book.sheets.find((sheet) => pattern.test(sheet.sheetName));

const textAt = (sheet: ExcelSheetSnapshot, row: number, column: number) =>
  String(valueOf(sheet.rawRows[row - 1]?.[column - 1]) ?? "").trim();

const findStaffTimesheetRow = (
  sheet: ExcelSheetSnapshot,
  op: EjoosSyncOp,
) => {
  const previousIndex = op.payload.previousIndex || op.positionIndex || "";
  const id = normId(op.personId);
  const name = canonicalName(op.fullName);
  for (let row = 7; row <= sheet.rawRows.length; row += 1) {
    const index = textAt(sheet, row, 2);
    const rowId = normId(textAt(sheet, row, 8));
    const rowName = canonicalName(textAt(sheet, row, 7));
    const onStaffIndex =
      Boolean(previousIndex && index === previousIndex) ||
      /^\d{5,}$/.test(index);
    if (!onStaffIndex) continue;
    if (id && rowId && id === rowId) return row;
    if (name && rowName && (rowName === name || rowName.includes(name))) {
      return row;
    }
    if (previousIndex && index === previousIndex) return row;
  }
  return 0;
};

const findTimesheetStyleRow = (sheet: ExcelSheetSnapshot) => {
  for (let row = 7; row <= sheet.rawRows.length; row += 1) {
    if (textAt(sheet, row, 2) || textAt(sheet, row, 7)) return row;
  }
  return 7;
};

/** Штатний рядок тільки за індексом. Новий рядок не створюємо. */
const findOrAllocateStaffIndexRow = (
  sheet: ExcelSheetSnapshot,
  staffIndex: string,
  _reserved: Set<number>,
  _positionTitle = "",
  _sourceRow = 0,
  _grid?: Array<unknown[] | undefined>,
  layout?: TimesheetLayout,
) => {
  if (!staffIndex) return 0;
  if (layout) return layout.byIndex[staffIndex]?.row ?? 0;
  for (let row = 7; row <= sheet.rawRows.length; row += 1) {
    if (textAt(sheet, row, 2) === staffIndex) return row;
  }
  return 0;
};

const writeStaffTimesheetIdentity = (
  writes: ZipCellWrite[],
  sheet: ExcelSheetSnapshot,
  targetRow: number,
  styleRow: number,
  op: EjoosSyncOp,
) => {
  const staffIndex =
    op.payload.timesheetStaffIndex ||
    op.payload.previousIndex ||
    op.positionIndex ||
    "";
  const entries: Array<[number, string | null]> = [
    [2, staffIndex || null],
    [6, op.rank || null],
    [7, op.fullName || null],
    [8, op.personId || null],
  ];
  for (const [column, value] of entries) {
    if (!value) continue;
    if (column === 2 || !textAt(sheet, targetRow, column)) {
      writes.push({
        row: targetRow,
        column,
        value,
        styleSourceRow: styleRow,
      });
    }
  }
};

const DISPOSITION_TIMESHEET_HEADER_RE =
  /ВИБУВ\s+У\s+РОЗПОРЯДЖЕННЯ\s+КОМАНДИРА/iu;
const DISPOSITION_TIMESHEET_SCAN_COLUMNS = 48;

const isCommanderDispositionLabel = (text: string) =>
  /РОЗПОРЯДЖ/iu.test(text) && !/^\d{5,}$/.test(text);

type DispositionTimesheetCols = {
  division: number;
  specialty: number;
  status: number;
  rank: number;
  name: number;
  id: number;
};

/** Реальний Табель (рядок 1145): A, C посада, D код, F звання, G ПІБ, H ID. */
const DEFAULT_DISPOSITION_TIMESHEET_COLS: DispositionTimesheetCols = {
  division: 1,
  specialty: 3,
  status: 4,
  rank: 6,
  name: 7,
  id: 8,
};

const COMPACT_DISPOSITION_TIMESHEET_COLS: DispositionTimesheetCols = {
  division: 1,
  specialty: 0,
  status: 3,
  rank: 4,
  name: 5,
  id: 8,
};

const looksLikePersonName = (text: string) =>
  /[А-ЯІЇЄҐ][а-яіїєґ'`-]+\s+[А-ЯІЇЄҐ]/u.test(text);

const DISPOSITION_MONTH_MARKER_ALIASES: Record<string, number> = {
  май: 5,
};

const UK_MONTH_TITLES = [
  "Січень",
  "Лютий",
  "Березень",
  "Квітень",
  "Травень",
  "Червень",
  "Липень",
  "Серпень",
  "Вересень",
  "Жовтень",
  "Листопад",
  "Грудень",
] as const;

const parseDispositionMonthMarker = (text: string) => {
  const key = text.trim().toLocaleLowerCase("uk-UA");
  if (!key || key.length > 20) return 0;
  if (DISPOSITION_MONTH_MARKER_ALIASES[key]) {
    return DISPOSITION_MONTH_MARKER_ALIASES[key];
  }
  const idx = UK_MONTH_TITLES.findIndex(
    (title) => title.toLocaleLowerCase("uk-UA") === key,
  );
  return idx >= 0 ? idx + 1 : 0;
};

const isDispositionMonthMarkerRow = (
  sheet: ExcelSheetSnapshot,
  row: number,
) => {
  const markerMonth = parseDispositionMonthMarker(textAt(sheet, row, 7));
  if (!markerMonth) return false;
  if (/РОЗПОРЯДЖ/iu.test(textAt(sheet, row, 1))) return false;
  if (looksLikeDispositionStatus(textAt(sheet, row, 4))) return false;
  if (textAt(sheet, row, 6) && looksLikePersonName(textAt(sheet, row, 7))) {
    return false;
  }
  return true;
};

const readSheetTimesheetMonth = (sheet: ExcelSheetSnapshot) => {
  for (let row = 1; row <= Math.min(6, sheet.rawRows.length); row += 1) {
    for (let column = 1; column <= 40; column += 1) {
      const text = timesheetMonthHeaderTextFromCell(
        sheet.rawRows[row - 1]?.[column - 1],
      );
      const parsed = parseTimesheetMonthHeaderText(text);
      if (parsed) return parsed.month;
    }
  }
  return 0;
};

const targetMonthFromPlan = (plan: EjoosSyncPlan) => {
  const header = formatTimesheetMonthHeader(plan.timesheetDayLabel);
  return parseTimesheetMonthHeaderText(header)?.month ?? 0;
};

type DispositionMonthSlice = {
  startRow: number;
  endRow: number;
  month: number;
};

const findDispositionMonthSlices = (
  sheet: ExcelSheetSnapshot,
  headerRow: number,
  primaryMonth: number,
): DispositionMonthSlice[] => {
  const slices: DispositionMonthSlice[] = [];
  let start = headerRow + 1;
  let sliceMonth = primaryMonth;
  for (let row = headerRow + 1; row <= sheet.rawRows.length; row += 1) {
    if (isTimesheetDispositionHeaderRow(sheet, row)) break;
    if (isStaffTimesheetRow(sheet, row)) break;
    if (!isDispositionMonthMarkerRow(sheet, row)) continue;
    const markerMonth = parseDispositionMonthMarker(textAt(sheet, row, 7));
    if (!markerMonth) continue;
    if (row > start) {
      slices.push({ startRow: start, endRow: row - 1, month: sliceMonth });
    }
    start = row + 1;
    sliceMonth = markerMonth;
  }
  if (start <= sheet.rawRows.length) {
    slices.push({
      startRow: start,
      endRow: sheet.rawRows.length,
      month: sliceMonth,
    });
  }
  return slices;
};

const isDispositionSubsectionDataRow = (
  sheet: ExcelSheetSnapshot,
  row: number,
  cols: DispositionTimesheetCols,
) => {
  if (isTimesheetDispositionHeaderRow(sheet, row)) return false;
  if (isStaffTimesheetRow(sheet, row)) return false;
  if (isDispositionMonthMarkerRow(sheet, row)) return false;
  if (isDispositionDataRow(sheet, row, cols)) return true;
  if (looksLikePersonName(textAt(sheet, row, cols.name))) return true;
  if (looksLikePersonName(textAt(sheet, row, 7))) return true;
  return rowHasTimesheetDayMarks(sheet, row);
};


const looksLikeDispositionStatus = (text: string) =>
  /^(?:ЗБ|СЗЧ|лік|від|\+|-|РОЗПОРЯДЖ)/iu.test(text);

const rowHasDispositionDivision = (sheet: ExcelSheetSnapshot, row: number) =>
  /РОЗПОРЯДЖ/iu.test(textAt(sheet, row, 1)) ||
  /РОЗПОРЯДЖ/iu.test(textAt(sheet, row, 2));

const rowTextThrough = (
  sheet: ExcelSheetSnapshot,
  row: number,
  lastColumn = DISPOSITION_TIMESHEET_SCAN_COLUMNS,
) => {
  const chunks: string[] = [];
  for (let column = 1; column <= lastColumn; column += 1) {
    const text = textAt(sheet, row, column);
    if (text) chunks.push(text);
  }
  return chunks.join(" ");
};

const isTimesheetDispositionHeaderRow = (
  sheet: ExcelSheetSnapshot,
  row: number,
) => DISPOSITION_TIMESHEET_HEADER_RE.test(rowTextThrough(sheet, row));

const findTimesheetDispositionHeaderRow = (sheet: ExcelSheetSnapshot) => {
  let headerRow = 0;
  for (let row = 1; row <= sheet.rawRows.length; row += 1) {
    if (isTimesheetDispositionHeaderRow(sheet, row)) headerRow = row;
  }
  return headerRow;
};

const isStaffTimesheetRow = (sheet: ExcelSheetSnapshot, row: number) =>
  /^\d{5,}$/.test(textAt(sheet, row, 2)) &&
  Boolean(textAt(sheet, row, 6) || textAt(sheet, row, 7));

const isDispositionDataRow = (
  sheet: ExcelSheetSnapshot,
  row: number,
  cols: DispositionTimesheetCols,
) => {
  if (isTimesheetDispositionHeaderRow(sheet, row)) return false;
  if (isStaffTimesheetRow(sheet, row)) return false;
  if (/^\d{5,}$/.test(textAt(sheet, row, 2))) return false;
  if (isCommanderDispositionLabel(textAt(sheet, row, 2))) return true;
  if (
    /РОЗПОРЯДЖ/iu.test(textAt(sheet, row, 1)) &&
    !/^\d{5,}$/.test(textAt(sheet, row, 2))
  ) {
    return true;
  }
  if (
    isCommanderDispositionLabel(textAt(sheet, row, cols.division)) &&
    cols.division !== 1
  ) {
    return true;
  }
  return false;
};

const findLastCommanderDispositionDataRow = (sheet: ExcelSheetSnapshot) => {
  for (let row = sheet.rawRows.length; row >= 1; row -= 1) {
    if (isStaffTimesheetRow(sheet, row)) continue;
    if (isDispositionDataRow(sheet, row, DEFAULT_DISPOSITION_TIMESHEET_COLS)) {
      return row;
    }
  }
  return 0;
};

const rowHasTimesheetDayMarks = (sheet: ExcelSheetSnapshot, row: number) => {
  for (let day = 1; day <= 31; day += 1) {
    if (textAt(sheet, row, 8 + day)) return true;
  }
  return false;
};

const inferDispositionTimesheetColumns = (
  sheet: ExcelSheetSnapshot,
  headerRow: number,
): DispositionTimesheetCols => {
  const start = headerRow > 0 ? headerRow + 1 : 7;
  let compactCandidate = false;
  for (let row = start; row <= sheet.rawRows.length; row += 1) {
    if (isTimesheetDispositionHeaderRow(sheet, row)) continue;
    if (isStaffTimesheetRow(sheet, row)) continue;
    if (
      !/РОЗПОРЯДЖ/iu.test(textAt(sheet, row, 1)) ||
      /^\d{5,}$/.test(textAt(sheet, row, 2))
    ) {
      continue;
    }
    if (looksLikePersonName(textAt(sheet, row, 7))) {
      return DEFAULT_DISPOSITION_TIMESHEET_COLS;
    }
    if (
      looksLikePersonName(textAt(sheet, row, 5)) &&
      !looksLikePersonName(textAt(sheet, row, 7))
    ) {
      compactCandidate = true;
    }
  }
  for (let row = sheet.rawRows.length; row >= start; row -= 1) {
    if (!isDispositionDataRow(sheet, row, DEFAULT_DISPOSITION_TIMESHEET_COLS)) {
      continue;
    }
    if (looksLikePersonName(textAt(sheet, row, 7))) {
      return DEFAULT_DISPOSITION_TIMESHEET_COLS;
    }
  }
  if (compactCandidate) return COMPACT_DISPOSITION_TIMESHEET_COLS;
  return DEFAULT_DISPOSITION_TIMESHEET_COLS;
};

const occupantMatchesOp = (
  sheet: ExcelSheetSnapshot,
  row: number,
  nameCol: number,
  idCol: number,
  op: EjoosSyncOp,
) => {
  const currentId = textAt(sheet, row, idCol);
  const currentName = textAt(sheet, row, nameCol);
  if (!currentId && !currentName) return false;
  if (op.personId && currentId && currentId === op.personId) return true;
  return Boolean(
    op.fullName &&
      currentName &&
      canonicalName(currentName) === canonicalName(op.fullName),
  );
};

const dispositionRowHasOccupant = (
  sheet: ExcelSheetSnapshot,
  row: number,
  cols: DispositionTimesheetCols,
) =>
  Boolean(
    textAt(sheet, row, cols.name) ||
      looksLikePersonName(textAt(sheet, row, 6)) ||
      looksLikePersonName(textAt(sheet, row, 7)) ||
      textAt(sheet, row, cols.id) ||
      textAt(sheet, row, cols.rank),
  );

const rowHasDispositionSectionContent = (
  sheet: ExcelSheetSnapshot,
  row: number,
  cols: DispositionTimesheetCols,
) => {
  if (isTimesheetDispositionHeaderRow(sheet, row)) return false;
  if (isStaffTimesheetRow(sheet, row)) return false;
  if (rowHasDispositionDivision(sheet, row)) return true;
  if (dispositionRowHasOccupant(sheet, row, cols)) return true;
  if (looksLikeDispositionStatus(textAt(sheet, row, cols.status))) return true;
  if (looksLikeDispositionStatus(textAt(sheet, row, 4))) return true;
  return rowHasTimesheetDayMarks(sheet, row);
};

const findTimesheetDispositionSection = (
  sheet: ExcelSheetSnapshot,
  plan: EjoosSyncPlan,
) => {
  const headerRow = findTimesheetDispositionHeaderRow(sheet);
  const cols = inferDispositionTimesheetColumns(
    sheet,
    headerRow || findLastCommanderDispositionDataRow(sheet),
  );
  const targetMonth =
    targetMonthFromPlan(plan) || readSheetTimesheetMonth(sheet);
  const primaryMonth = readSheetTimesheetMonth(sheet) || targetMonth;

  if (headerRow && targetMonth) {
    const slices = findDispositionMonthSlices(sheet, headerRow, primaryMonth);
    const slice =
      slices.find((item) => item.month === targetMonth) ||
      slices.find((item) => item.month === primaryMonth) ||
      slices[0];
    if (slice) {
      let lastRow = slice.startRow - 1;
      let styleRow = slice.startRow;
      for (let row = slice.startRow; row <= slice.endRow; row += 1) {
        if (!isDispositionSubsectionDataRow(sheet, row, cols)) continue;
        lastRow = row;
        styleRow = row;
      }
      return {
        headerRow,
        lastRow,
        styleRow: styleRow || Math.max(slice.startRow, 7),
        cols,
        slice,
      };
    }
  }

  if (headerRow) {
    let lastRow = headerRow;
    let styleRow = 0;
    for (let row = headerRow + 1; row <= sheet.rawRows.length; row += 1) {
      if (isTimesheetDispositionHeaderRow(sheet, row)) break;
      if (isStaffTimesheetRow(sheet, row)) break;
      if (isDispositionMonthMarkerRow(sheet, row)) break;
      if (!isDispositionSubsectionDataRow(sheet, row, cols)) continue;
      lastRow = row;
      styleRow = row;
    }
    return {
      headerRow,
      lastRow,
      styleRow: styleRow || Math.max(headerRow + 1, 7),
      cols,
    };
  }

  const lastDataRow = findLastCommanderDispositionDataRow(sheet);
  if (!lastDataRow) return null;

  let firstDataRow = lastDataRow;
  for (let row = lastDataRow - 1; row >= 1; row -= 1) {
    if (isStaffTimesheetRow(sheet, row)) break;
    if (isTimesheetDispositionHeaderRow(sheet, row)) break;
    if (!isDispositionDataRow(sheet, row, cols)) break;
    firstDataRow = row;
  }

  return {
    headerRow: Math.max(1, firstDataRow - 1),
    lastRow: lastDataRow,
    styleRow: lastDataRow,
    cols,
  };
};

const rowMatchesDispositionOp = (
  sheet: ExcelSheetSnapshot,
  row: number,
  op: EjoosSyncOp,
  cols: DispositionTimesheetCols,
) => occupantMatchesOp(sheet, row, cols.name, cols.id, op);

const DISPOSITION_TIMESHEET_STYLE_LAST_COLUMN = 40;

const seedDispositionTimesheetRowStyles = (
  writes: ZipCellWrite[],
  targetRow: number,
  styleRow: number,
) => {
  for (let column = 1; column <= DISPOSITION_TIMESHEET_STYLE_LAST_COLUMN; column += 1) {
    writes.push({
      row: targetRow,
      column,
      value: null,
      styleOnly: true,
      styleSourceRow: styleRow,
      styleSourceColumn: column,
      copyNeighborStyle: false,
      keepNeighborStyle: true,
      heightSourceRow: styleRow,
    });
  }
};

/** Рядок у блоці «ВИБУВ У РОЗПОРЯДЖЕННЯ…» — існуючий запис або новий рядок у кінці секції. */
const timesheetDispositionTarget = (
  sheet: ExcelSheetSnapshot,
  reserved: Set<number>,
  plan: EjoosSyncPlan,
  op?: EjoosSyncOp,
) => {
  const section = findTimesheetDispositionSection(sheet, plan);
  const cols = section?.cols ?? DEFAULT_DISPOSITION_TIMESHEET_COLS;
  const styleRow = section?.styleRow ?? 7;
  const sliceEnd = section?.slice?.endRow ?? section?.lastRow ?? 0;

  if (section && op) {
    const searchEnd = sliceEnd || section.lastRow;
    for (let row = section.headerRow + 1; row <= searchEnd; row += 1) {
      if (isDispositionMonthMarkerRow(sheet, row)) continue;
      if (!isDispositionSubsectionDataRow(sheet, row, cols)) continue;
      if (rowMatchesDispositionOp(sheet, row, op, cols)) {
        reserved.add(row);
        return { row, styleRow, cols, isNewRow: false };
      }
    }
  }

  let row = section ? section.lastRow + 1 : sheet.rawRows.length + 1;
  while (reserved.has(row)) row += 1;
  if (row < 7) row = Math.max(7, sheet.rawRows.length + 1);
  reserved.add(row);
  return { row, styleRow, cols, isNewRow: true };
};

const writeDispositionTimesheetIdentity = (
  writes: ZipCellWrite[],
  sheet: ExcelSheetSnapshot,
  targetRow: number,
  styleRow: number,
  op: EjoosSyncOp,
  absenceCode: string,
  cols: DispositionTimesheetCols,
  positionTitle?: string,
) => {
  const specialty =
    positionTitle ||
    op.payload.positionTitle ||
    textAt(sheet, targetRow, cols.specialty) ||
    null;
  const entries: Array<[number, string | null]> = [
    [cols.division, "РОЗПОРЯДЖЕННЯ"],
    [cols.status, absenceCode || shpoDispositionMark(op) || null],
    [cols.rank, op.rank || null],
    [cols.name, op.fullName || null],
    [cols.id, op.personId || null],
  ];
  if (cols.specialty > 0 && specialty) {
    entries.splice(1, 0, [cols.specialty, specialty]);
  }
  for (const [column, value] of entries) {
    if (!value) continue;
    writes.push({
      row: targetRow,
      column,
      value,
      styleSourceRow: styleRow,
      styleSourceColumn: column,
      keepNeighborStyle: true,
      copyNeighborStyle: false,
    });
  }
};

const clearTimesheetStaffOccupant = (
  writes: ZipCellWrite[],
  row: number,
  lastDay: number,
) => {
  for (let day = 1; day <= lastDay; day += 1) {
    writes.push({ row, column: 8 + day, value: null });
  }
  for (const column of [6, 7, 8]) {
    writes.push({ row, column, value: null });
  }
};

const dispositionTarget = (
  sheet: ExcelSheetSnapshot,
  reserved: Set<number>,
) => {
  let styleRow = 0;
  let lastRow = 0;
  for (let row = 1; row <= sheet.rawRows.length; row += 1) {
    const location = textAt(sheet, row, 2);
    if (!/РОЗПОРЯДЖ/iu.test(location)) continue;
    lastRow = row;
    if (textAt(sheet, row, 7) || textAt(sheet, row, 8)) styleRow = row;
    if (
      !reserved.has(row) &&
      !textAt(sheet, row, 7) &&
      !textAt(sheet, row, 8)
    ) {
      reserved.add(row);
      return { row, styleRow: styleRow || row };
    }
  }
  let row = lastRow + 1;
  while (reserved.has(row)) row += 1;
  reserved.add(row);
  return { row, styleRow: styleRow || Math.max(1, lastRow) };
};

const nextAbsentRow = (sheet: ExcelSheetSnapshot, reserved: Set<number>) => {
  for (let row = 6; row <= sheet.rawRows.length + 5; row += 1) {
    if (reserved.has(row) || textAt(sheet, row, 2)) continue;
    reserved.add(row);
    return row;
  }
  const row = sheet.rawRows.length + reserved.size + 1;
  reserved.add(row);
  return row;
};

const nextTimesheetRow = (
  sheet: ExcelSheetSnapshot,
  sourceRow: number,
  reserved: Set<number>,
  _positionTitle = "",
  grid?: Array<unknown[] | undefined>,
  layout?: TimesheetLayout,
  sourceIndex = "",
  insertCountBySection?: Map<string, number>,
) => {
  if (!layout) {
    throw new TimesheetLayoutError(
      "SOURCE_SLOT_MISSING",
      "Немає TimesheetLayout для історії розпорядження.",
    );
  }
  const sourceSlot = sourceIndex
    ? resolveCanonicalTimesheetSlot({ index: sourceIndex, layout })
    : null;
  return resolveHistoryTimesheetRow({
    sourceSlot,
    sourceRow,
    layout,
    sheet,
    grid,
    reserved,
    insertCountBySection,
  });
};

/**
 * У РУХ призначення приходить у різних формах: «у розпорядження командира…»,
 * «, який знаходиться у розпорядженні командира…». Зводимо до знахідного
 * відмінка без прийменника, щоб далі підставити потрібну форму.
 */
const dispositionPlace = (destination: string) =>
  String(destination || "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^[,;]\s*/, "")
    .replace(/^який\s+знаходиться\s+/iu, "")
    .replace(/^(?:у|в)\s+розпорядженн\S*/iu, "розпорядження")
    .replace(/^розпорядженн\S*/iu, "розпорядження")
    .replace(/\s+/g, " ")
    .trim();

const dispositionLocation = (place: string) => {
  if (!place) return ", який знаходиться у розпорядженні";
  const tail = place.replace(/^розпорядження\s*/iu, "").trim();
  return tail === place
    ? `, який знаходиться ${place}`
    : `, який знаходиться у розпорядженні ${tail}`.trimEnd();
};

/** У ШПО блок розпорядження тримає короткі коди: ЗБ, СЗЧ, лік — не «БЕЗВІСТИ». */
const shpoDispositionMark = (op: EjoosSyncOp) => {
  const code = String(op.payload.absenceCode || "").trim();
  if (code && code.length <= 8) return code;
  const raw = String(op.payload.absenceType || "").trim();
  if (/безвіст/iu.test(raw)) return "ЗБ";
  if (/сзч|самовіл/iu.test(raw)) return "СЗЧ";
  if (/полон/iu.test(raw)) return "пол";
  return raw || "РОЗПОРЯДЖЕННЯ";
};

type DispositionContext = {
  plan: EjoosSyncPlan;
  shpo: ExcelSheetSnapshot;
  absent: ExcelSheetSnapshot;
  timesheet: ExcelSheetSnapshot;
  timesheetGrid: Array<unknown[] | undefined>;
  shpoWrites: ZipCellWrite[];
  absentWrites: ZipCellWrite[];
  timesheetWrites: ZipCellWrite[];
  reservedShpoRows: Set<number>;
  reservedAbsentRows: Set<number>;
  reservedTimesheetRows: Set<number>;
  pendingHistoryInserts: PendingHistoryInsert[];
  insertCountBySection: Map<string, number>;
  timesheetLayout: TimesheetLayout;
};

const collectWrites = (op: EjoosSyncOp, ctx: DispositionContext) => {
  const { plan, shpo, absent, timesheet } = ctx;
  const shpoWrites = ctx.shpoWrites;
  const absentWrites = ctx.absentWrites;
  const timesheetWrites = ctx.timesheetWrites;
  const oldShpoRow = Number(op.payload.shpoExcelRow || 0);
  if (oldShpoRow > 0 && occupantMatchesOp(shpo, oldShpoRow, 7, 8, op)) {
    for (const column of [6, 7, 8, 18]) {
      shpoWrites.push({ row: oldShpoRow, column, value: null });
    }
  }

  const absenceLabel = op.payload.absenceType || "РОЗПОРЯДЖЕННЯ";
  const statusMark = shpoDispositionMark(op);
  const place = dispositionPlace(op.payload.destination);
  // Якщо особа вже стоїть у блоці «у розпорядженні», другий раз не додаємо.
  if (op.payload.skipShpoDisposition !== "1") {
    const target = dispositionTarget(shpo, ctx.reservedShpoRows);
    for (const [column, value] of [
      [2, dispositionLocation(place)],
      [3, statusMark],
      [6, op.rank],
      [7, op.fullName],
      [8, op.personId],
    ] as Array<[number, string]>) {
      shpoWrites.push({
        row: target.row,
        column,
        value: value || null,
        styleSourceRow: target.styleRow,
      });
    }
  }

  // Новий запис відсутності додаємо лише коли чинного немає, і лише з даними
  // архіву — порожній рядок тільки засмічує аркуш.
  if (op.payload.needsAbsenceRecord === "1" && op.payload.absenceDate) {
    const absentRow = nextAbsentRow(absent, ctx.reservedAbsentRows);
    const absentValues: Array<[number, string | null]> = [
      [1, op.rank || null],
      [2, op.fullName || null],
      [
        3,
        personIdFromShpo(parseEjoosShpo(shpo), {
          fullName: op.fullName,
          positionIndex: op.positionIndex,
          personId: op.personId,
        }) || null,
      ],
      [4, op.positionIndex || null],
      [5, absenceLabel || null],
      [6, op.payload.absencePlace || null],
      [7, op.payload.absenceDate || null],
      [
        9,
        [op.payload.absenceOrderDate, op.payload.absenceOrderNumber]
          .filter(Boolean)
          .join(" ") || null,
      ],
      [12, op.payload.plannedReturn || null],
    ];
    for (const [column, value] of absentValues) {
      absentWrites.push({
        row: absentRow,
        column,
        value,
        styleSourceRow: 6,
      });
    }
  }

  const absenceExcelRow = Number(op.payload.absenceExcelRow || 0);
  if (absenceExcelRow > 0 && op.personId) {
    const currentId = textAt(absent, absenceExcelRow, 3);
    if (!currentId || currentId !== op.personId) {
      absentWrites.push({
        row: absenceExcelRow,
        column: 3,
        value: op.personId,
        styleSourceRow: 6,
      });
    }
  }

  const staffTimesheetRow =
    Number(op.payload.timesheetExcelRow || 0) ||
    findStaffTimesheetRow(timesheet, op);
  const lastDay = Math.min(31, plan.timesheetDay);
  const absenceCode = op.payload.absenceCode || statusMark;
  const absenceSpans = parseTimesheetAbsenceSpans(
    op.payload.timesheetAbsenceSpans || "",
  );
  const journalId =
    op.personId ||
    personIdFromShpo(parseEjoosShpo(shpo), {
      fullName: op.fullName,
      positionIndex: op.positionIndex || op.payload.previousIndex,
      personId: op.personId,
    });
  const dispositionOrderDay = dayFromOrderLabel(op.payload.orderDate);
  const dispositionDeparture = formatDispositionTimesheetDeparture(
    op.payload.destination || op.payload.changeText || "",
    op.payload.orderNumber || "",
    op.payload.orderDate || "",
  );
  const paintOpenSzch = (
    row: number,
    styleSourceRow: number,
    idColumn: number,
  ) => {
    for (let day = 1; day <= lastDay; day += 1) {
      const value = timesheetMarkForOpenDispositionDay(day, {
        dispositionOrderDay,
        dispositionDeparture,
        absenceCode,
        lastDay,
        absenceSpans: absenceSpans.length ? absenceSpans : undefined,
      });
    timesheetWrites.push({
        row,
        column: 8 + day,
        value,
        styleSourceRow,
        styleSourceColumn: 8 + day,
        keepNeighborStyle: day !== dispositionOrderDay,
        wrapText: day === dispositionOrderDay,
        copyNeighborStyle: false,
      });
    }
    if (journalId) {
      timesheetWrites.push({
        row,
        column: idColumn,
        value: journalId,
        styleSourceRow,
        styleSourceColumn: idColumn,
        keepNeighborStyle: true,
        copyNeighborStyle: false,
      });
    }
  };

  const keepOpenSzch = op.payload.keepOpenSzchTimesheet === "1";
  const writeDispositionTimesheet =
    keepOpenSzch ||
    op.payload.timesheetFound === "true" ||
    op.payload.timesheetCreateRow === "1";
  if (writeDispositionTimesheet) {
    const target = timesheetDispositionTarget(
      timesheet,
      ctx.reservedTimesheetRows,
      plan,
      op,
    );
    if (target.isNewRow) {
      seedDispositionTimesheetRowStyles(
        timesheetWrites,
        target.row,
        target.styleRow,
      );
    }
    writeDispositionTimesheetIdentity(
      timesheetWrites,
      timesheet,
      target.row,
      target.styleRow,
      op,
      absenceCode,
      target.cols,
      staffTimesheetRow > 0
        ? textAt(timesheet, staffTimesheetRow, 3) ||
            op.payload.positionTitle ||
            undefined
        : op.payload.positionTitle,
    );
    paintOpenSzch(target.row, target.styleRow, target.cols.id);
    if (
      staffTimesheetRow > 0 &&
      occupantMatchesOp(timesheet, staffTimesheetRow, 7, 8, op)
    ) {
      clearTimesheetStaffOccupant(
        timesheetWrites,
        staffTimesheetRow,
        lastDay,
      );
    }
    return;
  }

  const sourceTimesheetRow =
    staffTimesheetRow ||
    (op.payload.timesheetCreateRow === "1"
      ? findOrAllocateStaffIndexRow(
          timesheet,
          op.payload.timesheetStaffIndex ||
            op.payload.previousIndex ||
            op.positionIndex ||
            "",
          ctx.reservedTimesheetRows,
          op.payload.positionTitle || "",
          staffTimesheetRow,
          ctx.timesheetGrid,
          ctx.timesheetLayout,
        )
      : 0);
  if (sourceTimesheetRow > 0) {
    ctx.reservedTimesheetRows.add(sourceTimesheetRow);
    const styleRow =
      findStaffTimesheetRow(timesheet, op) === sourceTimesheetRow
        ? sourceTimesheetRow
        : findTimesheetStyleRow(timesheet);
    if (
      op.payload.timesheetCreateRow === "1" ||
      op.payload.restorePerson === "1" ||
      !textAt(timesheet, sourceTimesheetRow, 7)
    ) {
      writeStaffTimesheetIdentity(
        timesheetWrites,
        timesheet,
        sourceTimesheetRow,
        styleRow,
        op,
      );
    }

    const historyRow = takeHistoryTimesheetRow(
      nextTimesheetRow(
        timesheet,
        sourceTimesheetRow,
        ctx.reservedTimesheetRows,
        op.payload.positionTitle || "",
        ctx.timesheetGrid,
        ctx.timesheetLayout,
        op.payload.timesheetStaffIndex ||
          op.payload.previousIndex ||
          op.positionIndex ||
          "",
        ctx.insertCountBySection,
      ),
      ctx.pendingHistoryInserts,
    );
    for (let column = 1; column <= 40; column += 1) {
      timesheetWrites.push({
        row: historyRow,
        column,
        value: valueOf(
          timesheet.rawRows[sourceTimesheetRow - 1]?.[column - 1],
        ),
        styleSourceRow: sourceTimesheetRow,
      });
    }
    const orderDay = dayFromOrderLabel(op.payload.orderDate);
    const departMark = [
      formatTimesheetDeparture(place),
      op.payload.orderNumber ? `наказ №${op.payload.orderNumber}` : "",
      op.payload.orderDate ? `від ${op.payload.orderDate}` : "",
    ]
      .filter(Boolean)
      .join(" ");
    for (let day = 1; day <= lastDay; day += 1) {
      let value: string | number | null;
      if (orderDay > 0 && day === orderDay) {
        value = departMark;
      } else if (absenceSpans.length) {
        value =
          timesheetMarkFromArchive(day, {
            activeFromDay: 1,
            lastDay,
            spans: absenceSpans,
            fillBeforeActive: true,
          }) ?? absenceCode;
      } else if (orderDay > 0 && day > orderDay) {
        value = absenceCode;
      } else {
        value = valueOf(
          timesheet.rawRows[sourceTimesheetRow - 1]?.[8 + day - 1],
        );
      }
      timesheetWrites.push({
        row: historyRow,
        column: 8 + day,
        value,
        styleSourceRow: sourceTimesheetRow,
        wrapText: typeof value === "string" && /вибув/iu.test(value),
      });
      timesheetWrites.push({
        row: sourceTimesheetRow,
        column: 8 + day,
        value: null,
      });
    }
    if (journalId && !textAt(timesheet, sourceTimesheetRow, 8)) {
      timesheetWrites.push({
        row: historyRow,
        column: 8,
        value: journalId,
        styleSourceRow: sourceTimesheetRow,
      });
    }
    for (const column of [6, 7, 8]) {
      timesheetWrites.push({
        row: sourceTimesheetRow,
        column,
        value: null,
      });
    }
  }
};

/**
 * Виведення в розпорядження: звільняємо штатну позицію в ШПО, лише якщо там
 * досі ця особа (фінальний sh wins), і додаємо її до блоку «у розпорядженні».
 * У Табелі особа потрапляє в кінець блоку «ВИБУВ У РОЗПОРЯДЖЕННЯ…»; штатний
 * рядок звільняється. ООС і Виключені не змінюються.
 */
export async function applyDispositionWithZip(input: {
  ejoos: ExcelWorkbookSnapshot;
  plan: EjoosSyncPlan;
  ops: EjoosSyncOp[];
}) {
  const { ejoos, plan, ops } = input;
  const shpo = findSheet(ejoos, /шпо|штатно.?посад/i);
  const absent = findSheet(ejoos, /тимчасов.*відсут/i);
  const timesheet = findSheet(ejoos, /табель/i);
  if (!shpo || !absent || !timesheet) {
    throw new Error("Не знайдено ШПО, Тимчасово відсутні або Табель");
  }
  const timesheetArtifacts = await loadTimesheetSheetArtifacts(
    ejoos.file,
    timesheet.sheetName,
  );
  const timesheetGrid = mergeTimesheetGrids(
    timesheet.rawRows,
    timesheetArtifacts.grid,
  );
  const timesheetView = placementSheetFromMergedGrid(timesheet, timesheetGrid);

  const ctx: DispositionContext = {
    plan,
    shpo,
    absent,
    timesheet: timesheetView,
    timesheetGrid,
    shpoWrites: [],
    absentWrites: [],
    timesheetWrites: [],
    reservedShpoRows: new Set<number>(),
    reservedAbsentRows: new Set<number>(),
    reservedTimesheetRows: new Set<number>(),
    pendingHistoryInserts: [],
    insertCountBySection: new Map<string, number>(),
    timesheetLayout: (() => {
      const layout = buildTimesheetLayout(timesheetView, {
        grid: timesheetGrid,
        shpoIndexes: parseEjoosShpo(shpo).map((row) => row.positionIndex),
        formulas: timesheetArtifacts.formulas,
      });
      assertTimesheetLayoutReadyForApply(layout);
      return layout;
    })(),
  };
  for (const op of ops) collectWrites(op, ctx);
  const { shpoWrites, absentWrites, timesheetWrites } = ctx;
  stampTimesheetHistoryInserts(timesheetWrites, ctx.pendingHistoryInserts);

  let blob: Blob | File = ejoos.file;
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    shpo.sheetName,
    shpoWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    absent.sheetName,
    absentWrites,
  );
  blob = await applyInlineStringWritesToWorkbook(
    blob,
    timesheet.sheetName,
    timesheetWrites,
  );
  return blob;
}
