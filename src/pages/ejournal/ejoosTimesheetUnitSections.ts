import type { ExcelSheetSnapshot } from "../../excelRoundTrip";

export type TimesheetGridView = {
  sheet: ExcelSheetSnapshot;
  grid?: Array<unknown[] | undefined>;
};

/** Excel зберігає заголовки секцій Табеля з апострофом: '1 ПІХОТНА РОТА. */
export const stripTimesheetDivisionLabel = (value: string) =>
  value
    .replace(/^['`´ʼ’]+/u, "")
    .replace(/\.+\s*$/u, "")
    .replace(/\s+/g, " ")
    .trim();

const cellAt = (
  view: TimesheetGridView,
  row: number,
  column: number,
) => {
  const fromRaw = String(view.sheet.rawRows[row - 1]?.[column - 1] ?? "").trim();
  const fromGrid = String(view.grid?.[row - 1]?.[column - 1] ?? "").trim();
  const text = fromRaw || fromGrid;
  if ((column === 1 || column === 2) && text) {
    return stripTimesheetDivisionLabel(text);
  }
  return text;
};

const sheetLastRow = (view: TimesheetGridView) =>
  Math.max(view.sheet.rawRows.length, view.grid?.length ?? 0);

export type TimesheetUnitKind =
  | "battalion"
  | "company"
  | "platoon"
  | "squad"
  | "group"
  | "management"
  | "other";

export type TimesheetUnitDescriptor = {
  kind: TimesheetUnitKind;
  number: string;
  text: string;
};

const normUnit = (value: string) =>
  stripTimesheetDivisionLabel(value).toLocaleLowerCase("uk-UA");

const isTimesheetSummaryRow = (text: string) =>
  /на\s*продоволь/iu.test(text) ||
  /перебува(?:є|e)\s*\d/iu.test(text) ||
  /^\s*всього\b/iu.test(text);

const looksLikePersonRow = (view: TimesheetGridView, row: number) =>
  Boolean(
    /^\d{5,}$/.test(cellAt(view, row, 2)) ||
      cellAt(view, row, 7) ||
      cellAt(view, row, 8),
  );

export const parseTimesheetUnitDescriptor = (
  text: string,
): TimesheetUnitDescriptor | null => {
  const normalized = normUnit(text);
  if (!normalized || normalized.length < 3) return null;
  if (isTimesheetSummaryRow(normalized)) return null;
  const number = normalized.match(/(\d+)/)?.[1] || "";
  let kind: TimesheetUnitKind = "other";
  if (/управлінн/.test(normalized)) kind = "management";
  else if (/груп/.test(normalized)) kind = "group";
  else if (/відділен/.test(normalized)) kind = "squad";
  else if (/взвод/.test(normalized)) kind = "platoon";
  else if (/рот/.test(normalized)) kind = "company";
  else if (/батальйон/.test(normalized)) kind = "battalion";
  else return null;
  return { kind, number, text: normalized };
};

const unitDescriptorStem = (text: string) =>
  normUnit(text)
    .replace(/^\d+\s+/u, "")
    .replace(
      /^(управління|група|групу|взводом|взводу|взвод|рота|роти|ротою|батальйоном|батальйону|батальйон|відділенням|відділенні|відділення)\s+/iu,
      "",
    )
    .trim();

const unitStemTokens = (stem: string) =>
  stem.split(/\s+/).filter((token) => token.length > 2);

const unitStemsMatch = (left: string, right: string) => {
  const leftStem = unitDescriptorStem(left);
  const rightStem = unitDescriptorStem(right);
  if (!leftStem || !rightStem) return true;
  if (leftStem === rightStem) return true;
  if (leftStem.includes(rightStem) || rightStem.includes(leftStem)) return true;
  const leftTokens = new Set(unitStemTokens(leftStem));
  const rightTokens = unitStemTokens(rightStem);
  if (!leftTokens.size || !rightTokens.length) return true;
  const overlap = rightTokens.filter((token) => leftTokens.has(token)).length;
  const required = Math.min(
    leftTokens.size,
    rightTokens.length,
    Math.max(2, Math.ceil(rightTokens.length * 0.5)),
  );
  return overlap >= required;
};

export const unitDescriptorsMatch = (left: string, right: string) => {
  const a = parseTimesheetUnitDescriptor(left);
  const b = parseTimesheetUnitDescriptor(right);
  if (!a || !b) return false;
  if (a.kind !== b.kind) return false;
  if (a.number && b.number && a.number !== b.number) return false;
  if (!a.number || !b.number) return unitStemsMatch(left, right);
  return true;
};

const unitKindLevel = (kind: TimesheetUnitKind) => {
  switch (kind) {
    case "battalion":
      return 1;
    case "management":
    case "group":
    case "company":
      return 2;
    case "platoon":
      return 3;
    case "squad":
      return 4;
    default:
      return 5;
  }
};

const extractUnnumberedUnitPhrases = (text: string) => {
  const phrases: string[] = [];
  const platoon = text.match(
    /(?:^|\s)(взводу(?:\s+[а-яіїєґ'`-]+){1,14})(?=\s+\d+\s+(?:[а-яіїєґ'`-]+\s+){0,2}батальйон\w*)/iu,
  )?.[1]?.trim();
  if (platoon) phrases.push(platoon);

  const squad = text.match(
    /(?:^|\s)(відділення(?:\s+[а-яіїєґ'`-]+){1,14})(?=\s+взводу)/iu,
  )?.[1]?.trim();
  if (squad) phrases.push(squad);

  return phrases;
};

/** Фрази підрозділу від найвужчого до найширшого. */
export const extractUnitPhrasesFromPosition = (positionTitle: string) => {
  const text = normUnit(positionTitle);
  if (!text) return [] as string[];
  const numbered = [
    ...text.matchAll(
      /\d+\s+(?:[а-яіїєґ'`-]+\s+){0,8}(?:відділення|відділенні|взводу|взвод|роти|рота|батальйону|батальйон)/giu,
    ),
  ].map((match) => match[0].trim());
  const unique = [...new Set([...numbered, ...extractUnnumberedUnitPhrases(text)])];
  unique.sort(
    (left, right) => unitKindLevel(parseTimesheetUnitDescriptor(right)?.kind || "other") -
      unitKindLevel(parseTimesheetUnitDescriptor(left)?.kind || "other"),
  );
  return unique;
};

/** Для вибуття: рота; якщо роти немає — взвод/відділення з посади; інакше управління/батальйон. */
export const companyUnitPhrasesFromPosition = (positionTitle: string) => {
  const phrases = extractUnitPhrasesFromPosition(positionTitle);
  const byKind = (kind: TimesheetUnitKind) =>
    phrases.filter((phrase) => parseTimesheetUnitDescriptor(phrase)?.kind === kind);

  const companies = byKind("company");
  if (companies.length) return companies;

  const platoons = byKind("platoon");
  if (platoons.length) return platoons;

  const squads = byKind("squad");
  if (squads.length) return squads;

  return phrases.filter((phrase) => {
    const kind = parseTimesheetUnitDescriptor(phrase)?.kind;
    return kind === "management" || kind === "group" || kind === "battalion";
  });
};

export type TimesheetUnitPlacementScope = "any" | "company";

const unitPhrasesForPlacement = (
  positionTitle: string,
  scope: TimesheetUnitPlacementScope,
) =>
  scope === "company"
    ? companyUnitPhrasesFromPosition(positionTitle)
    : extractUnitPhrasesFromPosition(positionTitle);

/** Чи стоїть рядок у секції підрозділу за назвою посади (кол. A). */
export const timesheetRowInExpectedUnitSection = (
  sheet: ExcelSheetSnapshot,
  row: number,
  positionTitle: string,
  grid?: Array<unknown[] | undefined>,
  placementScope: TimesheetUnitPlacementScope = "company",
) => {
  if (row < 7) return true;
  const phrases = unitPhrasesForPlacement(positionTitle, placementScope);
  if (!phrases.length) return true;
  const view = sheetView(sheet, grid);
  let matchedBounds = false;
  for (const phrase of phrases) {
    const bounds = findTimesheetUnitSectionBounds(view, phrase);
    if (!bounds) continue;
    matchedBounds = true;
    if (row >= bounds.headerRow && row <= bounds.endRow) return true;
  }
  return !matchedBounds;
};

export const effectiveTimesheetDivision = (
  view: TimesheetGridView,
  row: number,
) => {
  for (let scan = row; scan >= 7; scan -= 1) {
    const division = cellAt(view, scan, 1);
    if (!division || isTimesheetSummaryRow(division)) continue;
    if (parseTimesheetUnitDescriptor(division)) return division;
  }
  return "";
};

/** Найближча рота/управління/батальйон над рядком — для історії вибуття. */
export const effectiveTimesheetCompanyDivision = (
  view: TimesheetGridView,
  row: number,
) => {
  for (let scan = row; scan >= 7; scan -= 1) {
    const division = cellAt(view, scan, 1);
    if (!division || isTimesheetSummaryRow(division)) continue;
    const kind = parseTimesheetUnitDescriptor(division)?.kind;
    if (
      kind === "company" ||
      kind === "management" ||
      kind === "group" ||
      kind === "battalion"
    ) {
      return division;
    }
  }
  return effectiveTimesheetDivision(view, row);
};

const isStaffIndexText = (text: string) => /^\d{5,}$/.test(text.trim());

export const isTimesheetSummaryRowAt = (view: TimesheetGridView, row: number) => {
  for (let column = 1; column <= 7; column += 1) {
    const text = cellAt(view, row, column);
    if (text && isTimesheetSummaryRow(text)) return true;
  }
  return false;
};

/** Підпис підрозділу в рядку: зазвичай кол. B (взвод/відділення/рота), іноді лише кол. A. */
const timesheetUnitLabelInRow = (view: TimesheetGridView, row: number) => {
  if (isTimesheetSummaryRowAt(view, row)) return "";
  const colB = cellAt(view, row, 2);
  if (colB && !isStaffIndexText(colB)) {
    if (parseTimesheetUnitDescriptor(colB)) return colB;
  }
  const colA = cellAt(view, row, 1);
  if (!colA || !parseTimesheetUnitDescriptor(colA)) return "";
  if (colB && isStaffIndexText(colB)) return "";
  if (looksLikePersonRow(view, row)) return "";
  return colA;
};

/** Заголовок лише в кол. B: взвод, відділення, «1 ПІХОТНА РОТА». */
export const isTimesheetUnitHeaderRow = (
  view: TimesheetGridView,
  row: number,
) => {
  if (looksLikePersonRow(view, row)) return false;
  if (isTimesheetSummaryRowAt(view, row)) return false;
  const colB = cellAt(view, row, 2);
  return Boolean(
    colB && !isStaffIndexText(colB) && parseTimesheetUnitDescriptor(colB),
  );
};

/** Старий формат: назва підрозділу лише в A, B–H порожні. */
const isLegacyColumnAUnitHeaderRow = (
  view: TimesheetGridView,
  row: number,
) => {
  if (looksLikePersonRow(view, row)) return false;
  if (isTimesheetSummaryRowAt(view, row)) return false;
  const colA = cellAt(view, row, 1);
  if (!colA || !parseTimesheetUnitDescriptor(colA)) return false;
  for (let column = 2; column <= 8; column += 1) {
    if (cellAt(view, row, column)) return false;
  }
  const prevA = row > 7 ? cellAt(view, row - 1, 1) : "";
  const prevDesc = prevA ? parseTimesheetUnitDescriptor(prevA) : null;
  const colDesc = parseTimesheetUnitDescriptor(colA);
  // У новому Табелі A повторює роту на кожному рядку — порожній append-слот не заголовок.
  if (
    prevDesc?.kind === "company" &&
    colDesc?.kind === "company" &&
    prevDesc.number &&
    colDesc.number &&
    prevDesc.number === colDesc.number
  ) {
    return false;
  }
  return true;
};

const companyNumberInColumnAUpTo = (
  view: TimesheetGridView,
  row: number,
  stopRow: number,
) => {
  for (let scan = row; scan >= stopRow; scan -= 1) {
    if (isTimesheetSummaryRowAt(view, scan)) break;
    const colA = cellAt(view, scan, 1);
    if (!colA || isTimesheetSummaryRow(colA)) continue;
    const companyA = parseTimesheetUnitDescriptor(colA);
    if (companyA?.kind === "company" && companyA.number) {
      return companyA.number;
    }
  }
  return "";
};

/** Секція роти, коли в A на кожному рядку «1 ПІХОТНА РОТА», а заголовок роти в B може бути відсутній. */
const findCompanySectionBoundsByColumnA = (
  view: TimesheetGridView,
  targetUnit: string,
  startRow = 7,
) => {
  const target = parseTimesheetUnitDescriptor(targetUnit);
  if (!target || target.kind !== "company" || !target.number) return null;

  const rowBelongsToTargetCompany = (row: number, sectionStart = startRow) => {
    const colA = cellAt(view, row, 1);
    const companyA = colA ? parseTimesheetUnitDescriptor(colA) : null;
    if (companyA?.kind === "company") {
      return companyA.number === target.number;
    }
    if (colA) return false;
    return companyNumberInColumnAUpTo(view, row, sectionStart) === target.number;
  };

  let repeatedColAMatches = 0;
  let firstMatchRow = 0;
  for (let row = startRow; row <= sheetLastRow(view); row += 1) {
    const colA = cellAt(view, row, 1);
    const companyA = colA ? parseTimesheetUnitDescriptor(colA) : null;
    if (companyA?.kind === "company" && companyA.number === target.number) {
      repeatedColAMatches += 1;
      if (!firstMatchRow) firstMatchRow = row;
      if (repeatedColAMatches >= 2) break;
    }
  }
  const repeatFormatByColumnA =
    repeatedColAMatches >= 2 ||
    (repeatedColAMatches >= 1 &&
      firstMatchRow > 0 &&
      (looksLikePersonRow(view, firstMatchRow) ||
        isStaffIndexText(cellAt(view, firstMatchRow, 2)) ||
        (isTimesheetUnitHeaderRow(view, firstMatchRow) &&
          unitDescriptorsMatch(
            timesheetUnitLabelInRow(view, firstMatchRow) ||
              cellAt(view, firstMatchRow, 1),
            targetUnit,
          ))));
  if (!repeatFormatByColumnA) return null;

  let headerRow = 0;
  let endRow = 0;
  for (let row = startRow; row <= sheetLastRow(view); row += 1) {
    if (isTimesheetSummaryRowAt(view, row)) {
      if (headerRow) {
        endRow = row - 1;
        break;
      }
      continue;
    }
    const colA = cellAt(view, row, 1);
    const companyA = colA ? parseTimesheetUnitDescriptor(colA) : null;
    const inTargetCompany = rowBelongsToTargetCompany(row);

    if (!headerRow) {
      if (!inTargetCompany) continue;
      if (
        !looksLikePersonRow(view, row) &&
        !isTimesheetUnitHeaderRow(view, row)
      ) {
        const prevA = row > 7 ? cellAt(view, row - 1, 1) : "";
        const prevCompany = prevA ? parseTimesheetUnitDescriptor(prevA) : null;
        if (
          prevCompany?.kind === "company" &&
          prevCompany.number === target.number
        ) {
          continue;
        }
      }
      headerRow = row;
      endRow = row;
      continue;
    }

    if (inTargetCompany) {
      endRow = row;
      continue;
    }

    if (companyA?.kind === "company" && companyA.number !== target.number) {
      break;
    }

    const colB = cellAt(view, row, 2);
    if (colB && !isStaffIndexText(colB)) {
      const bDesc = parseTimesheetUnitDescriptor(colB);
      if (bDesc?.kind === "company" && bDesc.number !== target.number) {
        break;
      }
    }
  }

  if (!headerRow) return null;
  return {
    headerRow,
    endRow: endRow || headerRow,
    division: cellAt(view, headerRow, 1) || targetUnit,
  };
};

export const findTimesheetUnitSectionBounds = (
  view: TimesheetGridView,
  targetUnit: string,
  startRow = 7,
) => {
  const target = parseTimesheetUnitDescriptor(targetUnit);
  if (!target) return null;
  let headerRow = 0;

  if (target.kind === "company") {
    for (let row = startRow; row <= sheetLastRow(view); row += 1) {
      if (isTimesheetSummaryRowAt(view, row)) continue;
      const colB = cellAt(view, row, 2);
      if (
        colB &&
        !isStaffIndexText(colB) &&
        !isTimesheetSummaryRow(colB) &&
        unitDescriptorsMatch(colB, targetUnit)
      ) {
        headerRow = row;
        break;
      }
    }
  }

  if (!headerRow && target.kind === "company") {
    const byColumnA = findCompanySectionBoundsByColumnA(view, targetUnit, startRow);
    if (byColumnA) return byColumnA;
  }

  if (!headerRow) {
    for (let row = startRow; row <= sheetLastRow(view); row += 1) {
      if (isTimesheetSummaryRowAt(view, row)) continue;
      if (!isTimesheetUnitHeaderRow(view, row)) continue;
      const label = timesheetUnitLabelInRow(view, row);
      if (label && unitDescriptorsMatch(label, targetUnit)) {
        headerRow = row;
        break;
      }
    }
  }

  if (!headerRow) {
    for (let row = startRow; row <= sheetLastRow(view); row += 1) {
      if (isTimesheetSummaryRowAt(view, row)) continue;
      const colA = cellAt(view, row, 1);
      if (
        isLegacyColumnAUnitHeaderRow(view, row) &&
        colA &&
        unitDescriptorsMatch(colA, targetUnit)
      ) {
        headerRow = row;
        break;
      }
    }
  }
  if (!headerRow) return null;

  const headerLabel =
    timesheetUnitLabelInRow(view, headerRow) || cellAt(view, headerRow, 1);
  const headerLevel = unitKindLevel(target.kind);
  const headerNumber = parseTimesheetUnitDescriptor(headerLabel)?.number || "";
  let endRow = sheetLastRow(view);
  for (let row = headerRow + 1; row <= sheetLastRow(view); row += 1) {
    if (isTimesheetSummaryRowAt(view, row)) {
      endRow = row - 1;
      break;
    }
    if (target.kind === "company" && headerNumber) {
      const colA = cellAt(view, row, 1);
      const rowCompany = colA ? parseTimesheetUnitDescriptor(colA) : null;
      if (
        rowCompany?.kind === "company" &&
        rowCompany.number &&
        rowCompany.number !== headerNumber
      ) {
        endRow = row - 1;
        break;
      }
    }
    if (!isTimesheetUnitHeaderRow(view, row)) continue;
    const label = timesheetUnitLabelInRow(view, row);
    if (!label) continue;
    const next = parseTimesheetUnitDescriptor(label);
    if (!next) continue;
    if (unitKindLevel(next.kind) <= headerLevel) {
      if (unitDescriptorsMatch(label, headerLabel)) continue;
      endRow = row - 1;
      break;
    }
  }
  return { headerRow, endRow, division: headerLabel };
};

const sectionHeaderLevel = (division: string) =>
  unitKindLevel(parseTimesheetUnitDescriptor(division)?.kind ?? "other");

const shouldStopSectionAppendScan = (
  view: TimesheetGridView,
  row: number,
  sectionDivision: string,
  headerLevel: number,
) => {
  if (isTimesheetSummaryRowAt(view, row)) return true;
  if (!isTimesheetUnitHeaderRow(view, row)) return false;
  const label = timesheetUnitLabelInRow(view, row);
  if (!label) return false;
  if (unitDescriptorsMatch(label, sectionDivision)) return false;
  const next = parseTimesheetUnitDescriptor(label);
  if (!next) return false;
  return unitKindLevel(next.kind) <= headerLevel;
};

const lastOccupiedRowInBounds = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number },
) => {
  let lastOccupied = bounds.headerRow;
  for (let row = bounds.headerRow + 1; row <= bounds.endRow; row += 1) {
    if (looksLikePersonRow(view, row) || rowHasDayMarks(view, row)) {
      lastOccupied = row;
    }
  }
  return lastOccupied;
};

const findFreeAppendRowInSectionBounds = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number; division: string },
  reserved: Set<number>,
  extraRows = 12,
) => {
  const headerLevel = sectionHeaderLevel(bounds.division);
  const lastOccupied = lastOccupiedRowInBounds(view, bounds);
  for (let row = lastOccupied + 1; row <= bounds.endRow + extraRows; row += 1) {
    if (shouldStopSectionAppendScan(view, row, bounds.division, headerLevel)) {
      break;
    }
    if (rowLooksFree(view, row, reserved)) {
      reserved.add(row);
      return row;
    }
  }
  return 0;
};

/** Якщо вільного рядка немає — перший слот після останнього запису секції (перед summary/наступною ротою). */
const appendRowAfterLastOccupiedInSection = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number; division: string },
  reserved: Set<number>,
  extraRows = 30,
) => {
  const headerLevel = sectionHeaderLevel(bounds.division);
  const lastOccupied = lastOccupiedRowInBounds(view, bounds);
  for (let row = lastOccupied + 1; row <= bounds.endRow + extraRows; row += 1) {
    if (shouldStopSectionAppendScan(view, row, bounds.division, headerLevel)) {
      break;
    }
    if (reserved.has(row)) continue;
    if (
      isTimesheetUnitHeaderRow(view, row) ||
      looksLikePersonRow(view, row)
    ) {
      continue;
    }
    reserved.add(row);
    return row;
  }
  return 0;
};

const findSummaryRowInSection = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number },
) => {
  for (
    let row = bounds.headerRow + 1;
    row <= Math.min(bounds.endRow + 5, sheetLastRow(view));
    row += 1
  ) {
    if (isTimesheetSummaryRowAt(view, row)) return row;
  }
  return 0;
};

/** Останній вільний рядок перед «На продовольчому…» (навіть якщо в A повторюється рота). */
const appendRowBeforeSummary = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number; division: string },
  reserved: Set<number>,
) => {
  const summaryRow = findSummaryRowInSection(view, bounds);
  const lastOccupied = lastOccupiedRowInBounds(view, bounds);
  const scanEnd = summaryRow > 0 ? summaryRow - 1 : bounds.endRow;
  for (let row = scanEnd; row > lastOccupied; row -= 1) {
    if (reserved.has(row)) continue;
    if (rowLooksFree(view, row, reserved)) {
      reserved.add(row);
      return row;
    }
    if (
      !looksLikePersonRow(view, row) &&
      !isTimesheetUnitHeaderRow(view, row) &&
      !isTimesheetSummaryRowAt(view, row)
    ) {
      reserved.add(row);
      return row;
    }
  }
  for (let row = lastOccupied + 1; row <= scanEnd; row += 1) {
    if (reserved.has(row)) continue;
    if (rowLooksFree(view, row, reserved)) {
      reserved.add(row);
      return row;
    }
    if (
      !looksLikePersonRow(view, row) &&
      !isTimesheetUnitHeaderRow(view, row) &&
      !isTimesheetSummaryRowAt(view, row)
    ) {
      reserved.add(row);
      return row;
    }
  }
  return 0;
};

/** Кілька вибуттів: вставити новий рядок на позиції summary (не перезаписувати наступну секцію). */
const appendRowInsertBeforeSummary = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number; division: string },
  reserved: Set<number>,
) => {
  const summaryRow = findSummaryRowInSection(view, bounds);
  if (summaryRow <= bounds.headerRow) return 0;
  reserved.add(summaryRow);
  return summaryRow;
};

/** Секція без вільного рядка: вставити перед summary або наступною ротою/взводом. */
const appendRowInsertBeforeSectionBoundary = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number; division: string },
  reserved: Set<number>,
) => {
  const headerLevel = sectionHeaderLevel(bounds.division);
  const lastOccupied = lastOccupiedRowInBounds(view, bounds);
  for (let row = lastOccupied + 1; row <= bounds.endRow + 8; row += 1) {
    if (isTimesheetSummaryRowAt(view, row)) {
      reserved.add(row);
      return row;
    }
    if (shouldStopSectionAppendScan(view, row, bounds.division, headerLevel)) {
      reserved.add(row);
      return row;
    }
  }
  return 0;
};

const appendRowInSectionBounds = (
  view: TimesheetGridView,
  bounds: { headerRow: number; endRow: number; division: string },
  reserved: Set<number>,
) =>
  findFreeAppendRowInSectionBounds(view, bounds, reserved) ||
  appendRowAfterLastOccupiedInSection(view, bounds, reserved) ||
  appendRowBeforeSummary(view, bounds, reserved) ||
  appendRowInsertBeforeSummary(view, bounds, reserved) ||
  appendRowInsertBeforeSectionBoundary(view, bounds, reserved);

/** Чи потрібно зсунути рядки в Excel (вставка перед summary / наступною секцією). */
export const timesheetAppendNeedsRowInsert = (
  view: TimesheetGridView,
  targetRow: number,
  bounds: { headerRow: number; endRow: number },
) => {
  if (targetRow < 7) return false;
  const summaryRow = findSummaryRowInSection(view, bounds);
  if (summaryRow > 0 && targetRow >= summaryRow) return true;
  if (isTimesheetSummaryRowAt(view, targetRow)) return true;
  if (isTimesheetUnitHeaderRow(view, targetRow)) return true;
  return false;
};

export const findDepartureAppendBounds = (
  sheet: ExcelSheetSnapshot,
  positionTitle: string,
  grid?: Array<unknown[] | undefined>,
) => {
  const view = sheetView(sheet, grid);
  for (const candidate of companyUnitPhrasesFromPosition(positionTitle)) {
    const bounds = findTimesheetUnitSectionBounds(view, candidate);
    if (bounds) return bounds;
  }
  return null;
};

const sheetView = (
  sheet: ExcelSheetSnapshot,
  grid?: Array<unknown[] | undefined>,
): TimesheetGridView => ({ sheet, grid });

const rowLooksFree = (
  view: TimesheetGridView,
  row: number,
  reserved: Set<number>,
) => {
  const colB = cellAt(view, row, 2);
  const colBBlocked =
    Boolean(colB) &&
    (isStaffIndexText(colB) || Boolean(parseTimesheetUnitDescriptor(colB)));
  return (
    !reserved.has(row) &&
    !isTimesheetUnitHeaderRow(view, row) &&
    !isTimesheetSummaryRowAt(view, row) &&
    !colBBlocked &&
    !cellAt(view, row, 7) &&
    !cellAt(view, row, 8)
  );
};

const rowHasDayMarks = (view: TimesheetGridView, row: number) => {
  for (let day = 1; day <= 31; day += 1) {
    if (cellAt(view, row, 8 + day)) return true;
  }
  return false;
};

/** Кінець секції підрозділу в колонці A — порожній рядок після останнього запису. */
export const findTimesheetSectionAppendRow = (
  sheet: ExcelSheetSnapshot,
  sourceRow: number,
  reserved = new Set<number>(),
  grid?: Array<unknown[] | undefined>,
  placementScope: TimesheetUnitPlacementScope = "any",
) => {
  const view = sheetView(sheet, grid);
  const division =
    placementScope === "company"
      ? effectiveTimesheetCompanyDivision(view, sourceRow)
      : effectiveTimesheetDivision(view, sourceRow);
  if (division) {
    const bounds = findTimesheetUnitSectionBounds(view, division, 7);
    if (bounds) {
      const row = appendRowInSectionBounds(view, bounds, reserved);
      if (row) return row;
    }
  }

  if (placementScope === "company") return 0;

  for (let row = sourceRow + 1; row <= sheet.rawRows.length + 30; row += 1) {
    const rowDivision = cellAt(view, row, 1);
    const isSameSection =
      !division ||
      !rowDivision ||
      unitDescriptorsMatch(rowDivision, division);
    if (!isSameSection && looksLikePersonRow(view, row)) break;
    if (isSameSection && rowLooksFree(view, row, reserved) && row > sourceRow) {
      reserved.add(row);
      return row;
    }
  }

  const row = sheet.rawRows.length + reserved.size + 1;
  reserved.add(row);
  return row;
};

export const pickPositionTitleForUnitPlacement = (input: {
  positionTitle?: string;
  timesheetDestination?: string;
  documentsDest?: string;
  changeText?: string;
}) =>
  [
    input.positionTitle,
    input.timesheetDestination,
    input.documentsDest,
    input.changeText,
  ]
    .map((value) => String(value || "").trim())
    .find(Boolean) || "";

/** Шукати секцію за посадою/«куди вибув», а не лише за числовим індексом. */
export const findTimesheetAppendRowForUnit = (
  sheet: ExcelSheetSnapshot,
  options: {
    sourceRow?: number;
    positionTitle?: string;
    staffIndex?: string;
    reserved?: Set<number>;
    grid?: Array<unknown[] | undefined>;
    /** При вибутті — кінець усієї роти, не взводу/відділення. */
    placementScope?: TimesheetUnitPlacementScope;
  },
) => {
  const reserved = options.reserved ?? new Set<number>();
  const view = sheetView(sheet, options.grid);
  const scope = options.placementScope ?? "any";
  const title = pickPositionTitleForUnitPlacement({
    positionTitle: options.positionTitle,
  });
  const unitCandidates = unitPhrasesForPlacement(title, scope);
  if (options.sourceRow) {
    const inherited =
      scope === "company"
        ? effectiveTimesheetCompanyDivision(view, options.sourceRow)
        : effectiveTimesheetDivision(view, options.sourceRow);
    if (inherited) {
      const inheritedKind = parseTimesheetUnitDescriptor(inherited)?.kind;
      const titleHasOwnUnit = unitCandidates.some((candidate) => {
        const kind = parseTimesheetUnitDescriptor(candidate)?.kind;
        return kind === "company" || kind === "platoon" || kind === "squad";
      });
      if (!(inheritedKind === "management" && titleHasOwnUnit)) {
        unitCandidates.push(inherited);
      }
    }
  }

  for (const candidate of unitCandidates) {
    const bounds = findTimesheetUnitSectionBounds(view, candidate);
    if (!bounds) continue;
    const row = appendRowInSectionBounds(view, bounds, reserved);
    if (row) return row;
  }

  if (options.sourceRow && options.sourceRow >= 7 && !unitCandidates.length) {
    for (let row = options.sourceRow + 1; row <= options.sourceRow + 12; row += 1) {
      if (rowLooksFree(view, row, reserved)) {
        reserved.add(row);
        return row;
      }
    }
  }

  if (options.sourceRow && options.sourceRow >= 7) {
    if (scope === "company" && companyUnitPhrasesFromPosition(title).length) {
      return 0;
    }
    return findTimesheetSectionAppendRow(
      sheet,
      options.sourceRow,
      reserved,
      options.grid,
      scope,
    );
  }

  if (options.staffIndex && scope !== "company") {
    let anchor = 7;
    for (let row = 7; row <= sheet.rawRows.length; row += 1) {
      const index = cellAt(view, row, 2);
      if (!/^\d{5,}$/.test(index)) continue;
      if (Number(index) <= Number(options.staffIndex)) anchor = row;
    }
    for (let row = anchor + 1; row <= sheet.rawRows.length + 30; row += 1) {
      if (rowLooksFree(view, row, reserved)) {
        reserved.add(row);
        return row;
      }
    }
  }

  if (scope === "company") return 0;

  const row = sheet.rawRows.length + reserved.size + 1;
  reserved.add(row);
  return row;
};
