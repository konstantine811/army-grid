import {
  type CellValue,
  type ExcelSheetSnapshot,
  type ExcelWorkbookSnapshot,
  exportWorkbookWithMutations,
} from "../../excelRoundTrip";

const MOVEMENT_REGISTRY_KEY = "army-grid:ejournal-processed-movements";
const INITIAL_CUTOFF_DATE_SERIAL = 46231; // 28.07.2026
const PRESEEDED_PENDING_ROWS = [
  [7082, 7083],
  [7090, 7094],
  [7096, 7103],
  [7105, 7121],
  [7123, 7170],
] as const;
const INITIAL_KNOWN_DONE_AFTER_CUTOFF_END_ROW = 7217;

const col = (letter: string) => {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index - 1;
};

export type EjournalMovementStatus =
  | "pending"
  | "processed"
  | "initial-done"
  | "conflict";

export type EjournalMovementEvent = {
  rowNumber: number;
  movementNumber: string;
  type: string;
  personId: string;
  rank: string;
  name: string;
  orderNumber: string;
  orderDate: string;
  previousPositionIndex: string;
  nextPositionIndex: string;
  changeText: string;
  destination: string;
  note: string;
  status: EjournalMovementStatus;
  conflict?: string;
};

export type EjournalArchivePeriod = {
  periodNumber: string;
  personId: string;
  name: string;
  absenceType: string;
  operationStatus: string;
  departDate: string;
  returnDate: string;
};

export type EjournalWorkflowAnalysis = {
  sourceName: string;
  templateName?: string;
  shRows: number;
  movementRows: number;
  archiveRows: number;
  pending: EjournalMovementEvent[];
  processed: EjournalMovementEvent[];
  conflicts: EjournalMovementEvent[];
  archivePeriods: EjournalArchivePeriod[];
  statsByType: Record<string, number>;
  processedRegistry: string[];
};

export type ProcessedMovementRegistry = Record<
  string,
  {
    processedAt: string;
    exportId: string;
    result: string;
  }
>;

export function readProcessedMovementRegistry(): ProcessedMovementRegistry {
  try {
    const raw = window.localStorage.getItem(MOVEMENT_REGISTRY_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

export function resetProcessedMovementRegistry() {
  window.localStorage.removeItem(MOVEMENT_REGISTRY_KEY);
}

export function analyzeEjournalWorkflow(
  source: ExcelWorkbookSnapshot,
  template?: ExcelWorkbookSnapshot | null,
): EjournalWorkflowAnalysis {
  const sh = getSheet(source, "sh");
  const movement = getSheet(source, "рух");
  const archive = getSheet(source, "archive");
  const registry = readProcessedMovementRegistry();
  const shIndexes = new Set(
    (sh?.rawRows ?? [])
      .slice(1)
      .map((row) => normalizeValue(row[col("R")]))
      .filter(Boolean),
  );
  const shIds = new Set(
    (sh?.rawRows ?? [])
      .slice(1)
      .map((row) => normalizeValue(row[col("J")]))
      .filter((value) => value && value !== "0"),
  );
  const archivePeriods = buildArchivePeriods(archive);
  const archiveByPersonId = new Map<string, EjournalArchivePeriod[]>();
  archivePeriods.forEach((period) => {
    if (!period.personId) return;
    archiveByPersonId.set(period.personId, [
      ...(archiveByPersonId.get(period.personId) ?? []),
      period,
    ]);
  });

  const processed: EjournalMovementEvent[] = [];
  const pending: EjournalMovementEvent[] = [];
  const conflicts: EjournalMovementEvent[] = [];
  const statsByType: Record<string, number> = {};

  (movement?.rawRows ?? []).slice(1).forEach((row, rowIndex) => {
    const rowNumber = rowIndex + 2;
    const movementNumber = normalizeValue(row[col("B")]);
    if (!movementNumber) return;
    const type = normalizeMovementType(row[col("E")]);
    if (!["ПЕРЕВ", "РОЗПОРЯДЖ"].includes(type)) return;

    const event: EjournalMovementEvent = {
      rowNumber,
      movementNumber,
      type,
      rank: normalizeValue(row[col("F")]),
      name: normalizeValue(row[col("G")]),
      personId: normalizeValue(row[col("H")]),
      orderNumber: normalizeValue(row[col("K")]),
      orderDate: formatExcelDate(row[col("L")]),
      previousPositionIndex: normalizeValue(row[col("W")]),
      nextPositionIndex: normalizeValue(row[col("X")]),
      changeText: normalizeValue(row[col("J")]),
      destination: "",
      note: normalizeValue(row[col("V")]),
      status: "pending",
    };
    event.destination = extractMovementDestination(event);

    event.conflict = getMovementConflict(event, shIndexes, shIds, archiveByPersonId);
    if (event.conflict) {
      event.status = "conflict";
      conflicts.push(event);
      return;
    }

    if (registry[movementNumber]) {
      event.status = "processed";
      processed.push(event);
      return;
    }

    if (!isInitiallyPending(rowNumber, row[col("L")])) {
      event.status = "initial-done";
      processed.push(event);
      return;
    }

    statsByType[event.type] = (statsByType[event.type] ?? 0) + 1;
    pending.push(event);
  });

  pending.sort(compareMovementEvents);

  return {
    sourceName: source.fileName,
    templateName: template?.fileName,
    shRows: Math.max(0, (sh?.rawRows.length ?? 1) - 1),
    movementRows: Math.max(0, (movement?.rawRows.length ?? 1) - 1),
    archiveRows: Math.max(0, (archive?.rawRows.length ?? 2) - 2),
    pending,
    processed,
    conflicts,
    archivePeriods,
    statsByType,
    processedRegistry: Object.keys(registry),
  };
}

export async function exportEjournalTemplateWithProtocol(
  template: ExcelWorkbookSnapshot,
  source: ExcelWorkbookSnapshot,
  analysis: EjournalWorkflowAnalysis,
) {
  const exportId = `ejournal-${Date.now()}`;
  const exportedAt = new Date().toISOString();

  await exportWorkbookWithMutations(
    template,
    async (workbook) => {
      applyCurrentShToTemplate(workbook, source);
      applyArchiveToTemporaryAbsences(workbook, source);
      applyMovementEventsToTemplate(workbook, analysis);

      const protocolName = "Протокол конфліктів";
      const existing = workbook.sheet(protocolName);
      if (existing) workbook.deleteSheet(existing);
      const sheet = workbook.addSheet(protocolName);

      writeProtocolSheet(sheet, analysis, exportedAt);
    },
    buildExportName(),
  );

  markPendingMovementsProcessed(analysis.pending, exportId, exportedAt);
}

function applyCurrentShToTemplate(workbook: any, source: ExcelWorkbookSnapshot) {
  const sourceSh = getSheet(source, "sh");
  if (!sourceSh) return;

  const shpoSheet = workbook.sheet("1. ШПО");
  const timesheet = workbook.sheet("6. Табель");
  const shRows = sourceSh.rawRows.slice(1).filter((row) => normalizeValue(row[col("R")]));

  if (shpoSheet) {
    const indexToRow = buildSheetIndex(shpoSheet, "A", 7);
    shRows.forEach((row) => {
      const positionIndex = normalizeValue(row[col("R")]);
      const targetRow = indexToRow.get(positionIndex) ?? nextAppendRow(shpoSheet);
      writeShpoRow(shpoSheet, targetRow, row);
      indexToRow.set(positionIndex, targetRow);
    });
  }

  if (timesheet) {
    const indexToRow = buildSheetIndex(timesheet, "B", 7);
    shRows.forEach((row) => {
      const positionIndex = normalizeValue(row[col("R")]);
      const targetRow = indexToRow.get(positionIndex) ?? nextAppendRow(timesheet);
      writeTimesheetRow(timesheet, targetRow, row);
      indexToRow.set(positionIndex, targetRow);
    });
  }
}

function applyArchiveToTemporaryAbsences(workbook: any, source: ExcelWorkbookSnapshot) {
  const sourceArchive = getSheet(source, "archive");
  const sheet = workbook.sheet("5. Тимчасово відсутні");
  if (!sourceArchive || !sheet) return;

  const existing = buildTemporaryAbsenceIndex(sheet);
  sourceArchive.rawRows.slice(2).forEach((row) => {
    const periodNumber = normalizeValue(row[col("A")]);
    const personId = normalizeValue(row[col("E")]);
    const name = normalizeValue(row[col("G")]);
    const reason = normalizeValue(row[col("I")]);
    const departDate = row[col("J")];
    if (!periodNumber && !personId && !name) return;

    const key = temporaryAbsenceKey(personId, reason, departDate, periodNumber);
    const targetRow = existing.get(key) ?? nextTemporaryAbsenceRow(sheet, existing);
    writeTemporaryAbsenceRow(sheet, targetRow, row, periodNumber);
    existing.set(key, targetRow);
  });
}

function applyMovementEventsToTemplate(workbook: any, analysis: EjournalWorkflowAnalysis) {
  const oosSheet = workbook.sheet("2. ООС");
  const excludedSheet = workbook.sheet("3. Виключені");
  const shpoSheet = workbook.sheet("1. ШПО");
  const timesheet = workbook.sheet("6. Табель");
  const oosByPersonId = oosSheet ? buildSheetIndex(oosSheet, "C", 5) : new Map<string, number>();
  const excludedByMovement = excludedSheet
    ? buildMovementMarkerIndex(excludedSheet, "AG", 6)
    : new Map<string, number>();
  analysis.pending.sort(compareMovementEvents).forEach((event) => {
    if (!["ПЕРЕВ", "РОЗПОРЯДЖ"].includes(event.type)) return;

    if (timesheet) {
      applyTransferToTimesheet(timesheet, event);
    }

    if (excludedSheet) {
      const targetRow = excludedByMovement.get(event.movementNumber) ?? nextAppendRow(excludedSheet);
      writeExcludedMovement(excludedSheet, targetRow, event, oosSheet, oosByPersonId);
      excludedByMovement.set(event.movementNumber, targetRow);
    }

    if (shpoSheet) {
      clearShpoOccupant(shpoSheet, event);
    }

    const oosRow = oosByPersonId.get(event.personId);
    if (oosSheet && oosRow) clearOosPersonnelRow(oosSheet, oosRow);
  });
}

function writeShpoRow(sheet: any, targetRow: number, sourceRow: CellValue[]) {
  const values: Array<[string, CellValue]> = [
    ["A", sourceRow[col("R")]],
    ["B", sourceRow[col("G")]],
    ["C", sourceRow[col("H")]],
    ["D", sourceRow[col("S")]],
    ["E", sourceRow[col("T")]],
    ["F", sourceRow[col("K")]],
    ["G", sourceRow[col("L")]],
    ["H", sourceRow[col("J")]],
    ["Q", sourceRow[col("R")]],
    ["R", sourceRow[col("K")]],
  ];
  values.forEach(([column, value]) => sheet.cell(`${column}${targetRow}`).value(emptyToNull(value)));
}

function writeExcludedMovement(
  sheet: any,
  targetRow: number,
  event: EjournalMovementEvent,
  oosSheet: any,
  oosByPersonId: Map<string, number>,
) {
  const sourceRow = oosByPersonId.get(event.personId);
  const copyValue = (sourceColumn: string, targetColumn: string) => {
    if (!oosSheet || !sourceRow) return;
    const value = oosSheet.cell(`${sourceColumn}${sourceRow}`).value() as CellValue;
    if (normalizeValue(value)) sheet.cell(`${targetColumn}${targetRow}`).value(value);
  };

  copyValue("A", "A");
  copyValue("B", "B");
  copyValue("C", "C");
  copyValue("D", "D");
  copyValue("E", "E");
  copyValue("G", "F");
  copyValue("H", "G");
  copyValue("S", "N");
  copyValue("T", "O");
  copyValue("U", "P");
  copyValue("V", "Q");
  copyValue("Z", "T");
  copyValue("AA", "U");
  copyValue("AB", "V");
  copyValue("AE", "Y");
  copyValue("AF", "Z");
  copyValue("AG", "AA");

  if (event.rank) sheet.cell(`A${targetRow}`).value(event.rank);
  if (event.name) sheet.cell(`B${targetRow}`).value(event.name);
  if (event.personId) sheet.cell(`C${targetRow}`).value(event.personId);
  if (event.previousPositionIndex || event.nextPositionIndex) {
    sheet.cell(`D${targetRow}`).value(event.previousPositionIndex || event.nextPositionIndex);
  }

  sheet.cell(`H${targetRow}`).value(joinDateOrderText(oosSheet, sourceRow, "I", "J"));
  sheet.cell(`J${targetRow}`).value(joinDateOrderText(oosSheet, sourceRow, "L", "K"));
  sheet.cell(`L${targetRow}`).value(joinDateOrderText(oosSheet, sourceRow, "O", "P"));
  sheet.cell(`R${targetRow}`).value(joinDocumentText(oosSheet, sourceRow, "X", "Y"));
  sheet.cell(`W${targetRow}`).value(joinCallupText(oosSheet, sourceRow));
  sheet.cell(`AB${targetRow}`).value(event.orderDate);
  sheet.cell(`AC${targetRow}`).value(event.orderDate);
  sheet.cell(`AD${targetRow}`).value(event.orderNumber);
  sheet.cell(`AE${targetRow}`).value(
    (event.destination || event.note || "")
      .replace(/\s+/g, " ")
      .trim()
      .toLocaleLowerCase("uk-UA") || null,
  );
  sheet.cell(`AF${targetRow}`).value(event.type === "РОЗПОРЯДЖ" ? "Розпорядження" : "Переведення");
  sheet.cell(`AG${targetRow}`).value(`Рух №${event.movementNumber}`);
}

function applyTransferToTimesheet(sheet: any, event: EjournalMovementEvent) {
  const sourceRow = findPersonRow(sheet, event, 7, "G", "H");
  if (!sourceRow) return;

  const targetRow = findTimesheetSectionAppendRow(sheet, sourceRow);
  if (targetRow === sourceRow) return;

  copyRow(sheet, sourceRow, targetRow, col("AN") + 1);
  writeTimesheetDeparture(sheet, targetRow, event);
}

function writeTimesheetDeparture(sheet: any, rowNumber: number, event: EjournalMovementEvent) {
  const departDay = getDayOfMonth(event.orderDate);
  const today = new Date();
  const todayDay = today.getDate();
  const lastDay = Math.min(31, todayDay);
  const departureText = event.destination ? `вибув до ${event.destination}` : "вибув";
  let presentDays = 0;

  for (let day = 1; day <= 31; day += 1) {
    const cell = sheet.cell(rowNumber, col("I") + day);
    if (!departDay || day < departDay) {
      cell.value("+");
      presentDays += 1;
    } else if (day === departDay) {
      cell.value(departureText);
    } else if (day <= lastDay) {
      cell.value("-");
    } else {
      cell.value(null);
    }
  }

  sheet.cell(`AN${rowNumber}`).value(presentDays);
}

function findTimesheetSectionAppendRow(sheet: any, sourceRow: number) {
  const division = normalizeValue(sheet.cell(`A${sourceRow}`).value() as CellValue);
  const endRow = sheet.usedRange()?.endCell().rowNumber() ?? sourceRow;

  for (let rowNumber = sourceRow + 1; rowNumber <= endRow + 50; rowNumber += 1) {
    const rowDivision = normalizeValue(sheet.cell(`A${rowNumber}`).value() as CellValue);
    const positionIndex = normalizeValue(sheet.cell(`B${rowNumber}`).value() as CellValue);
    const name = normalizeValue(sheet.cell(`G${rowNumber}`).value() as CellValue);
    const id = normalizeValue(sheet.cell(`H${rowNumber}`).value() as CellValue);
    const isSameSection = !division || !rowDivision || rowDivision === division;

    if (!isSameSection && (positionIndex || name || id)) break;
    if (isSameSection && !positionIndex && !name && !id && rowNumber > sourceRow) return rowNumber;
  }

  return nextAppendRow(sheet);
}

function clearShpoOccupant(sheet: any, event: EjournalMovementEvent) {
  const rowNumber = findPersonRow(sheet, event, 7, "G", "H");
  if (!rowNumber) return;
  clearRangeValues(sheet, rowNumber, col("F") + 1, col("H") + 1);
  sheet.cell(`R${rowNumber}`).value(null);
}

function findPersonRow(
  sheet: any,
  event: EjournalMovementEvent,
  startRow: number,
  nameColumn: string,
  idColumn: string,
) {
  const endRow = sheet.usedRange()?.endCell().rowNumber() ?? startRow;
  const eventName = normalizePersonName(event.name);
  const eventId = normalizeValue(event.personId);

  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const rowName = normalizePersonName(sheet.cell(`${nameColumn}${rowNumber}`).value() as CellValue);
    const rowId = normalizeValue(sheet.cell(`${idColumn}${rowNumber}`).value() as CellValue);
    if (eventId && rowId === eventId && (!eventName || rowName === eventName)) return rowNumber;
  }

  return 0;
}

function clearOosPersonnelRow(sheet: any, rowNumber: number) {
  clearRangeValues(sheet, rowNumber, col("A") + 1, col("AI") + 1);
}

function copyRow(sheet: any, sourceRow: number, targetRow: number, endColumn: number) {
  sheet.row(targetRow).height(sheet.row(sourceRow).height());
  for (let column = 1; column <= endColumn; column += 1) {
    const sourceCell = sheet.cell(sourceRow, column);
    const targetCell = sheet.cell(targetRow, column);
    targetCell.value(sourceCell.value() as CellValue);
    copyCellStyle(sourceCell, targetCell);
  }
}

function copyCellStyle(sourceCell: any, targetCell: any) {
  const styleNames = [
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
  ];
  targetCell.style(sourceCell.style(styleNames));
}

function clearRangeValues(sheet: any, rowNumber: number, startColumn: number, endColumn: number) {
  for (let column = startColumn; column <= endColumn; column += 1) {
    sheet.cell(rowNumber, column).value(null);
  }
}

function writeTimesheetRow(sheet: any, targetRow: number, sourceRow: CellValue[]) {
  const values: Array<[string, CellValue]> = [
    ["A", sourceRow[col("C")]],
    ["B", sourceRow[col("R")]],
    ["C", sourceRow[col("G")]],
    ["D", sourceRow[col("S")]],
    ["E", sourceRow[col("T")]],
    ["F", sourceRow[col("K")]],
    ["G", sourceRow[col("L")]],
    ["H", sourceRow[col("J")]],
    ["AN", normalizeValue(sourceRow[col("J")]) && normalizeValue(sourceRow[col("J")]) !== "0" ? 11 : 0],
  ];
  values.forEach(([column, value]) => sheet.cell(`${column}${targetRow}`).value(emptyToNull(value)));
}

function writeTemporaryAbsenceRow(
  sheet: any,
  targetRow: number,
  sourceRow: CellValue[],
  periodNumber: string,
) {
  const departureOrder = joinDateOrder(sourceRow[col("L")], sourceRow[col("K")]);
  const returnOrder = joinDateOrder(sourceRow[col("AC")], sourceRow[col("AD")]);
  const isReturned = normalizeValue(sourceRow[col("AA")]).toLowerCase() === "енд";

  const values: Array<[string, CellValue | string]> = [
    ["A", sourceRow[col("F")]],
    ["B", sourceRow[col("G")]],
    ["C", sourceRow[col("E")]],
    ["D", findPositionIndexInText(sourceRow[col("H")])],
    ["E", sourceRow[col("I")]],
    ["F", sourceRow[col("M")]],
    ["G", sourceRow[col("J")]],
    ["H", sourceRow[col("Q")]],
    ["I", departureOrder.date],
    ["J", departureOrder.order],
    ["K", sourceRow[col("N")] || "?"],
    ["L", sourceRow[col("O")] || "?"],
    ["M", isReturned ? sourceRow[col("AB")] : null],
    ["N", isReturned ? sourceRow[col("AB")] : null],
    ["O", isReturned ? returnOrder.date : null],
    ["P", isReturned ? returnOrder.order : null],
    ["Q", isReturned ? "" : sourceRow[col("AA")]],
    ["R", `archive №${periodNumber}`],
  ];
  values.forEach(([column, value]) => sheet.cell(`${column}${targetRow}`).value(emptyToNull(value)));
}

function writeProtocolSheet(sheet: any, analysis: EjournalWorkflowAnalysis, exportedAt: string) {
  const rows = [
    ["ЕЖООС export protocol", exportedAt],
    ["Джерело", analysis.sourceName],
    ["Шаблон", analysis.templateName ?? "—"],
    ["Нових подій", analysis.pending.length],
    ["Конфліктів", analysis.conflicts.length],
    [],
    ["Статус", "№ руху", "Рядок", "Тип", "ID", "ПІБ", "Дата", "Наказ", "W попер.", "X новий", "Коментар"],
    ...analysis.conflicts.map((event) => protocolRow("КОНФЛІКТ", event, event.conflict ?? "")),
    ...analysis.pending.map((event) => protocolRow("ДО ЕКСПОРТУ", event, "")),
  ];

  rows.forEach((row, rowIndex) => {
    row.forEach((value, columnIndex) => sheet.cell(rowIndex + 1, columnIndex + 1).value(value));
  });
  sheet.usedRange()?.style({
    fontFamily: "Times New Roman",
    fontSize: 11,
    border: true,
    wrapText: true,
  });
  sheet.range("A1:K1").style({ bold: true, fill: "1f4e78", fontColor: "ffffff" });
  sheet.range("A7:K7").style({ bold: true, fill: "d9eaf7" });
  [1, 2, 4, 6, 7, 8, 9, 10, 11].forEach((columnIndex) => sheet.column(columnIndex).width(18));
  sheet.column(5).width(10);
  sheet.column(6).width(36);
}

function markPendingMovementsProcessed(
  movements: EjournalMovementEvent[],
  exportId: string,
  processedAt: string,
) {
  const registry = readProcessedMovementRegistry();
  movements.forEach((event) => {
    registry[event.movementNumber] = {
      processedAt,
      exportId,
      result: "exported-to-template",
    };
  });
  window.localStorage.setItem(MOVEMENT_REGISTRY_KEY, JSON.stringify(registry));
}

function protocolRow(status: string, event: EjournalMovementEvent, comment: string) {
  return [
    status,
    event.movementNumber,
    event.rowNumber,
    event.type,
    event.personId,
    event.name,
    event.orderDate,
    event.orderNumber,
    event.previousPositionIndex,
    event.nextPositionIndex,
    comment,
  ];
}

function getMovementConflict(
  event: EjournalMovementEvent,
  shIndexes: Set<string>,
  shIds: Set<string>,
  archiveByPersonId: Map<string, EjournalArchivePeriod[]>,
) {
  if (!event.personId || event.personId === "0") return "Відсутній ID";
  if (!event.orderDate) return "Відсутня дата наказу Рух.L";
  if (!event.orderNumber) return "Відсутній номер наказу Рух.K";
  if (["ПОСАДА", "ПРИБУВ"].includes(event.type) && !event.nextPositionIndex) {
    return "Відсутній новий індекс посади Рух.X";
  }
  if (["ПОСАДА", "ПРИБУВ"].includes(event.type) && event.nextPositionIndex && !shIndexes.has(event.nextPositionIndex)) {
    return "Новий індекс посади не знайдений у sh";
  }
  if (event.type === "ПРИБУВ" && !shIds.has(event.personId)) {
    return "Людина прибула, але її немає в поточному sh";
  }
  if (event.type === "РОЗПОРЯДЖ" && !event.destination && !archiveByPersonId.has(event.personId)) {
    return "РОЗПОРЯДЖ без відповідного періоду archive";
  }
  if (!event.destination) return "Не вдалося визначити куди вибув";
  return "";
}

function buildArchivePeriods(archive?: ExcelSheetSnapshot): EjournalArchivePeriod[] {
  return (archive?.rawRows ?? [])
    .slice(2)
    .map((row, index) => ({
      periodNumber: normalizeValue(row[col("A")]) || String(index + 1),
      personId: normalizeValue(row[col("E")]),
      name: normalizeValue(row[col("G")]),
      absenceType: normalizeValue(row[col("I")]),
      operationStatus: normalizeValue(row[col("AA")]),
      departDate: formatExcelDate(row[col("J")]),
      returnDate: formatExcelDate(row[col("AB")]),
    }))
    .filter((period) => period.periodNumber || period.personId || period.name);
}

function buildSheetIndex(sheet: any, keyColumn: string, startRow: number) {
  const map = new Map<string, number>();
  const endRow = sheet.usedRange()?.endCell().rowNumber() ?? startRow;
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const key = normalizeValue(sheet.cell(`${keyColumn}${rowNumber}`).value() as CellValue);
    if (key) map.set(key, rowNumber);
  }
  return map;
}

function buildTemporaryAbsenceIndex(sheet: any) {
  const map = new Map<string, number>();
  const endRow = sheet.usedRange()?.endCell().rowNumber() ?? 6;
  for (let rowNumber = 6; rowNumber <= endRow; rowNumber += 1) {
    const extra = normalizeValue(sheet.cell(`R${rowNumber}`).value() as CellValue);
    const archiveNumber = extra.match(/archive\s*№\s*([^\s]+)/i)?.[1] ?? "";
    const personId = normalizeValue(sheet.cell(`C${rowNumber}`).value() as CellValue);
    const reason = normalizeValue(sheet.cell(`E${rowNumber}`).value() as CellValue);
    const name = normalizeValue(sheet.cell(`B${rowNumber}`).value() as CellValue);
    const departDate = sheet.cell(`G${rowNumber}`).value() as CellValue;
    const key = temporaryAbsenceKey(personId, reason, departDate, archiveNumber);
    if ((personId && name) || archiveNumber) map.set(key, rowNumber);
  }
  return map;
}

function buildMovementMarkerIndex(sheet: any, keyColumn: string, startRow: number) {
  const map = new Map<string, number>();
  const endRow = sheet.usedRange()?.endCell().rowNumber() ?? startRow;
  for (let rowNumber = startRow; rowNumber <= endRow; rowNumber += 1) {
    const marker = normalizeValue(sheet.cell(`${keyColumn}${rowNumber}`).value() as CellValue);
    const movementNumber = marker.match(/Рух\s*№\s*([^\s]+)/i)?.[1] ?? "";
    if (movementNumber) map.set(movementNumber, rowNumber);
  }
  return map;
}

function nextAppendRow(sheet: any) {
  const used = sheet.usedRange();
  return Math.max(6, (used?.endCell().rowNumber() ?? 5) + 1);
}

function nextTemporaryAbsenceRow(sheet: any, occupiedRows: Map<string, number>) {
  const used = sheet.usedRange();
  const endRow = used?.endCell().rowNumber() ?? 6;
  const reservedRows = new Set(occupiedRows.values());

  for (let rowNumber = 6; rowNumber <= endRow + 200; rowNumber += 1) {
    if (reservedRows.has(rowNumber)) continue;

    const hasVisibleData = ["A", "B", "E", "G", "J", "R"].some((column) =>
      Boolean(normalizeValue(sheet.cell(`${column}${rowNumber}`).value() as CellValue)),
    );
    if (!hasVisibleData) return rowNumber;
  }

  return Math.max(6, endRow + 1);
}

function temporaryAbsenceKey(
  personId: string,
  reason: string,
  departDate: CellValue,
  periodNumber: string,
) {
  if (periodNumber) return `archive:${periodNumber}`;
  return [
    personId,
    normalizeValue(reason).toLowerCase(),
    normalizeValue(formatCellForKey(departDate)),
  ].join("|");
}

function joinDateOrder(date: CellValue, order: CellValue) {
  return {
    date: emptyToNull(date),
    order: normalizeValue(order),
  };
}

function joinDateOrderText(sheet: any, rowNumber: number | undefined, dateColumn: string, orderColumn: string) {
  if (!sheet || !rowNumber) return null;
  const date = formatCellForDisplay(sheet.cell(`${dateColumn}${rowNumber}`).value() as CellValue);
  const order = normalizeValue(sheet.cell(`${orderColumn}${rowNumber}`).value() as CellValue);
  return [date, order && `№${order.replace(/^№\s*/i, "")}`].filter(Boolean).join(" ") || null;
}

function joinDocumentText(sheet: any, rowNumber: number | undefined, numberColumn: string, typeColumn: string) {
  if (!sheet || !rowNumber) return null;
  const number = normalizeValue(sheet.cell(`${numberColumn}${rowNumber}`).value() as CellValue);
  const type = normalizeValue(sheet.cell(`${typeColumn}${rowNumber}`).value() as CellValue);
  return [number, type].filter(Boolean).join(" ") || null;
}

function joinCallupText(sheet: any, rowNumber: number | undefined) {
  if (!sheet || !rowNumber) return null;
  const date = formatCellForDisplay(sheet.cell(`AC${rowNumber}`).value() as CellValue);
  const authority = normalizeValue(sheet.cell(`AD${rowNumber}`).value() as CellValue);
  return [authority, date].filter(Boolean).join(", ") || null;
}

function extractMovementDestination(event: EjournalMovementEvent) {
  const text = normalizeValue(event.changeText || event.note);
  if (!text) return "";

  if (event.type === "РОЗПОРЯДЖ") {
    return text
      .replace(/^,\s*/g, "")
      .replace(/^яки[йхм]\s+знаходиться\s+у\s+/i, "у ")
      .trim();
  }

  return extractUnitDestination(text);
}

function extractUnitDestination(text: string) {
  const normalized = normalizeValue(text)
    .replace(/\s*-\s*/g, " - ")
    .replace(/\s+/g, " ")
    .trim();
  const upper = normalized.toUpperCase();
  const unitMarkers = [
    "ВІЙСЬКОВОЇ ЧАСТИНИ",
    "ОКРЕМОГО БАТАЛЬЙОНУ",
    "ОКРЕМОЇ БРИГАДИ",
    "БАТАЛЬЙОНУ",
    "БАТАРЕЇ",
    "ДИВІЗІОНУ",
    "РОТИ",
    "ВЗВОДУ",
    "ВІДДІЛЕННЯ",
    "ЕКІПАЖУ",
    "ОБСЛУГИ",
    "ГРУПИ",
    "СЛУЖБИ",
    "ШТАБУ",
    "УПРАВЛІННЯ",
    "КОМАНДИ",
  ];

  const markerIndex = unitMarkers
    .map((marker) => ({ marker, index: upper.indexOf(marker) }))
    .filter(({ index }) => index >= 0)
    .sort((a, b) => a.index - b.index)[0];

  if (!markerIndex) return normalized;

  const wordsBefore = normalized.slice(0, markerIndex.index).trim().split(/\s+/).filter(Boolean);
  const previousWord = wordsBefore[wordsBefore.length - 1] ?? "";
  const includePrevious = wordsBefore.length > 0 && isLikelyUnitAdjective(previousWord);
  const start = includePrevious
    ? normalized.lastIndexOf(previousWord, markerIndex.index)
    : markerIndex.index;

  return normalized
    .slice(Math.max(0, start))
    .replace(/^-\s*/, "")
    .trim();
}

function isLikelyUnitAdjective(word: string) {
  return /(ОГО|ОЇ|ИХ|ЬКОГО|ЬКОЇ|НОГО|НОЇ)$/i.test(word) && !/(КОМАНДИРА|НАЧАЛЬНИКА)$/i.test(word);
}

function findPositionIndexInText(value: CellValue) {
  const text = normalizeValue(value);
  const match = text.match(/\b\d{7}\b/);
  return match?.[0] ?? "";
}

function emptyToNull(value: CellValue | string) {
  const normalized = normalizeValue(value as CellValue);
  return normalized && normalized !== "х" && normalized !== "x" && normalized !== "0"
    ? value
    : null;
}

function formatCellForKey(value: CellValue) {
  if (typeof value === "number" || value instanceof Date) return formatExcelDate(value);
  return value;
}

function formatCellForDisplay(value: CellValue) {
  if (typeof value === "number" || value instanceof Date) return formatExcelDate(value);
  return normalizeValue(value);
}

function getDayOfMonth(value: string) {
  const [day, month, year] = value.split(".").map(Number);
  return year && month && day ? day : 0;
}

function normalizePersonName(value: CellValue | unknown) {
  return normalizeValue(value).toUpperCase().replace(/[’']/g, "'").replace(/\s+/g, " ");
}

function isInitiallyPending(rowNumber: number, rawDate: CellValue) {
  const serial = toExcelSerial(rawDate);
  if (PRESEEDED_PENDING_ROWS.some(([start, end]) => rowNumber >= start && rowNumber <= end)) {
    return true;
  }
  return serial > INITIAL_CUTOFF_DATE_SERIAL && rowNumber > INITIAL_KNOWN_DONE_AFTER_CUTOFF_END_ROW;
}

function compareMovementEvents(a: EjournalMovementEvent, b: EjournalMovementEvent) {
  const dateA = toComparableDate(a.orderDate);
  const dateB = toComparableDate(b.orderDate);
  if (dateA !== dateB) return dateA - dateB;
  const numA = Number(a.movementNumber);
  const numB = Number(b.movementNumber);
  if (Number.isFinite(numA) && Number.isFinite(numB) && numA !== numB) return numA - numB;
  return a.rowNumber - b.rowNumber;
}

function normalizeMovementType(value: CellValue) {
  const text = normalizeValue(value).toUpperCase();
  if (text.includes("ПОСАД")) return "ПОСАДА";
  if (text.includes("РОЗПОР")) return "РОЗПОРЯДЖ";
  if (text.includes("ПЕРЕВ")) return "ПЕРЕВ";
  if (text.includes("ПРИБ")) return "ПРИБУВ";
  if (text.includes("ЗВІЛ")) return "ЗВІЛЬН";
  if (text.includes("ЗВАН")) return "ЗВАННЯ";
  return text || "—";
}

function getSheet(workbook: ExcelWorkbookSnapshot, name: string) {
  const normalizedName = name.toLowerCase();
  return workbook.sheets.find((sheet) => sheet.sheetName.toLowerCase() === normalizedName);
}

function normalizeValue(value: CellValue | unknown) {
  if (value && typeof value === "object" && !(value instanceof Date)) {
    const maybeTextValue = value as { text?: () => string; value?: () => unknown };
    if (typeof maybeTextValue.text === "function") return normalizeValue(maybeTextValue.text());
    if (typeof maybeTextValue.value === "function") return normalizeValue(maybeTextValue.value());
    return "";
  }

  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

function toExcelSerial(value: CellValue) {
  if (typeof value === "number") return value;
  if (value instanceof Date) {
    return Math.round((value.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
  }
  const parsed = Date.parse(String(value ?? ""));
  return Number.isFinite(parsed) ? Math.round((parsed - Date.UTC(1899, 11, 30)) / 86400000) : 0;
}

function formatExcelDate(value: CellValue) {
  const serial = toExcelSerial(value);
  if (!serial) return normalizeValue(value);
  const date = new Date(Date.UTC(1899, 11, 30 + serial));
  return date.toLocaleDateString("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    timeZone: "UTC",
  });
}

function toComparableDate(value: string) {
  const [day, month, year] = value.split(".").map(Number);
  return year && month && day ? Date.UTC(year, month - 1, day) : 0;
}

function buildExportName() {
  const now = new Date();
  const date = now
    .toLocaleDateString("uk-UA", { day: "2-digit", month: "2-digit", year: "numeric" })
    .replaceAll(".", "-");
  return `ЄЖООС_станом_на_${date}.xlsx`;
}
