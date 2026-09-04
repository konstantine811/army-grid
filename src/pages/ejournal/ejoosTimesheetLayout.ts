import type { ExcelSheetSnapshot } from "../../excelRoundTrip";
import {
  isTimesheetSummaryRowAt,
  isTimesheetUnitHeaderRow,
  parseTimesheetUnitDescriptor,
  stripTimesheetDivisionLabel,
  type TimesheetGridView,
  type TimesheetUnitKind,
} from "./ejoosTimesheetUnitSections";

const DATA_START_ROW = 6;
const isPositionIndex = (value: string) => /^\d{5,}$/.test(value.trim());

const cellAt = (view: TimesheetGridView, row: number, column: number) => {
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

const TOP_LEVEL_KINDS = new Set<TimesheetUnitKind>([
  "company",
  "management",
  "group",
  "battalion",
]);

export type TimesheetLayoutErrorCode =
  | "TIMESHEET_SLOT_AMBIGUOUS"
  | "TIMESHEET_SECTION_RANGE_UNKNOWN"
  | "CANONICAL_SLOT_MISSING"
  | "SOURCE_SLOT_MISSING"
  | "SLOT_OCCUPIED_CONFLICT"
  | "SECTION_HISTORY_FULL"
  | "HISTORY_INSIDE_CANONICAL";

export class TimesheetLayoutError extends Error {
  code: TimesheetLayoutErrorCode;
  index: string;
  section: string;
  rows: number[];

  constructor(
    code: TimesheetLayoutErrorCode,
    message: string,
    extra?: { index?: string; section?: string; rows?: number[] },
  ) {
    super(message);
    this.name = "TimesheetLayoutError";
    this.code = code;
    this.index = extra?.index ?? "";
    this.section = extra?.section ?? "";
    this.rows = extra?.rows ?? [];
  }
}

export type TimesheetCanonicalSlot = {
  row: number;
  index: string;
  section: string;
  platoon: string;
  squad: string;
  position: string;
  vos: string;
  tariff: string;
  occupantId: string;
  occupantName: string;
  occupied: boolean;
};

export type TimesheetSection = {
  key: string;
  label: string;
  kind: TimesheetUnitKind;
  headerRow: number;
  canonicalStartRow: number;
  canonicalEndRow: number;
  firstSlotRow: number;
  lastCanonicalRow: number;
  historyStartRow: number;
  historyEndRow: number;
  footerRow: number;
  explicitFooter: boolean;
};

export type TimesheetLayoutIssue = {
  code: "TIMESHEET_SLOT_AMBIGUOUS" | "TIMESHEET_SECTION_RANGE_UNKNOWN";
  index?: string;
  section?: string;
  rows: number[];
};

export type TimesheetLayout = {
  sections: TimesheetSection[];
  byIndex: Record<string, TimesheetCanonicalSlot>;
  sectionByRow: Map<number, TimesheetSection>;
  issues: TimesheetLayoutIssue[];
};

export type TimesheetShpoHint = {
  positionIndex: string;
};

export type TimesheetHistoryPlacement = {
  row: number;
  insertBefore: boolean;
  sectionKey: string;
  footerRow: number;
};

export type PendingHistoryInsert = {
  sectionKey: string;
  footerRow: number;
  targetRow: number;
};

export type TimesheetFormulaMap = Map<string, string> | Record<string, string>;

const sectionKey = (label: string) =>
  stripTimesheetDivisionLabel(label).toLocaleLowerCase("uk-UA");

const formulaEntries = (formulas?: TimesheetFormulaMap) =>
  formulas instanceof Map
    ? [...formulas.entries()]
    : Object.entries(formulas ?? {});

/** COUNTIF(I353:I478,"+") → штатний діапазон 353:478. */
export const parseCountIfStaffRange = (formula: string) => {
  const match = String(formula || "").match(
    /COUNTIF\(\s*\$?[A-Z]{1,3}\$?(\d+)\s*:\s*\$?[A-Z]{1,3}\$?(\d+)/i,
  );
  if (!match) return null;
  const start = Number(match[1]);
  const end = Number(match[2]);
  if (!start || !end || end < start) return null;
  return { start, end };
};

export const extractSheetFormulasByCell = (sheetXml: string) => {
  const formulas = new Map<string, string>();
  const cellRe =
    /<c\b([^<>]*\br="([A-Z]{1,3}\d+)"(?![0-9A-Za-z])[^<>]*)(\/\s*>|>[\s\S]*?<\/c>)/gi;
  for (const cell of sheetXml.matchAll(cellRe)) {
    const ref = cell[2].toUpperCase();
    const inner = cell[3] || "";
    const formula = inner.match(/<f\b[^>]*>([\s\S]*?)<\/f>/i)?.[1]?.trim();
    if (formula) formulas.set(ref, formula);
  }
  return formulas;
};

export const countIfRangeOnRow = (
  formulas: TimesheetFormulaMap | undefined,
  row: number,
) => {
  if (!formulas || !row) return null;
  for (const [ref, formula] of formulaEntries(formulas)) {
    const refRow = Number(String(ref).replace(/^[A-Z]+/i, ""));
    if (refRow !== row) continue;
    const range = parseCountIfStaffRange(formula);
    if (range) return range;
  }
  return null;
};

const headerDescriptor = (view: TimesheetGridView, row: number) => {
  if (isTimesheetSummaryRowAt(view, row)) return null;
  if (isTimesheetUnitHeaderRow(view, row)) {
    return parseTimesheetUnitDescriptor(cellAt(view, row, 2));
  }
  const colA = cellAt(view, row, 1);
  const colB = cellAt(view, row, 2);
  if (colB && isPositionIndex(colB)) return null;
  if (cellAt(view, row, 7) || cellAt(view, row, 8)) return null;
  if (!colA) return null;
  const desc = parseTimesheetUnitDescriptor(colA);
  if (!desc || !TOP_LEVEL_KINDS.has(desc.kind)) return null;
  if (colB && !parseTimesheetUnitDescriptor(colB) && !isPositionIndex(colB)) {
    return null;
  }
  return desc;
};

const rowIsCompletelyEmpty = (view: TimesheetGridView, row: number) => {
  for (let column = 1; column <= 40; column += 1) {
    if (cellAt(view, row, column)) return false;
  }
  return true;
};

const isCanonicalRowOf = (section: TimesheetSection, row: number) =>
  row >= section.canonicalStartRow && row <= section.canonicalEndRow;

export const buildTimesheetLayout = (
  sheet: ExcelSheetSnapshot,
  options?: {
    grid?: Array<unknown[] | undefined>;
    shpoIndexes?: Iterable<string>;
    formulas?: TimesheetFormulaMap;
  },
): TimesheetLayout => {
  const view: TimesheetGridView = { sheet, grid: options?.grid };
  const lastRow = sheetLastRow(view);
  const openSections: Array<{
    label: string;
    kind: TimesheetUnitKind;
    headerRow: number;
    footerRow: number;
    explicitFooter: boolean;
  }> = [];
  let current: (typeof openSections)[number] | null = null;

  const closeSection = (footerRow: number, explicitFooter: boolean) => {
    if (!current) return;
    current.footerRow = footerRow;
    current.explicitFooter = explicitFooter;
    openSections.push(current);
    current = null;
  };

  const openInherited = (row: number) => {
    const inherited = stripTimesheetDivisionLabel(cellAt(view, row, 1));
    const kind = parseTimesheetUnitDescriptor(inherited)?.kind ?? "other";
    current = {
      label: inherited || "—",
      kind,
      headerRow: row,
      footerRow: 0,
      explicitFooter: false,
    };
  };

  for (let row = DATA_START_ROW; row <= lastRow; row += 1) {
    if (isTimesheetSummaryRowAt(view, row)) {
      closeSection(row, true);
      continue;
    }
    const header = headerDescriptor(view, row);
    if (header && TOP_LEVEL_KINDS.has(header.kind)) {
      closeSection(row, true);
      current = {
        label: stripTimesheetDivisionLabel(
          cellAt(view, row, 2) || cellAt(view, row, 1) || header.text,
        ),
        kind: header.kind,
        headerRow: row,
        footerRow: 0,
        explicitFooter: false,
      };
      continue;
    }
    if (!current && isPositionIndex(cellAt(view, row, 2))) {
      openInherited(row);
    }
  }
  closeSection(lastRow + 1, false);

  const issues: TimesheetLayoutIssue[] = [];
  const sections: TimesheetSection[] = openSections.map((section) => {
    const footer = section.footerRow || lastRow + 1;
    const footerIsSummary = isTimesheetSummaryRowAt(view, footer);
    const countIf =
      countIfRangeOnRow(options?.formulas, footer) ||
      (() => {
        for (let row = section.headerRow + 1; row < footer; row += 1) {
          if (!isTimesheetSummaryRowAt(view, row)) continue;
          const range = countIfRangeOnRow(options?.formulas, row);
          if (range) return range;
        }
        return null;
      })();
    let canonicalStart = 0;
    let canonicalEnd = 0;
    if (countIf) {
      canonicalStart = countIf.start;
      canonicalEnd = countIf.end;
    } else if (footerIsSummary) {
      issues.push({
        code: "TIMESHEET_SECTION_RANGE_UNKNOWN",
        section: section.label,
        rows: [footer],
      });
      canonicalStart = section.headerRow + 1;
      canonicalEnd = section.headerRow;
    } else {
      const seen = new Set<string>();
      for (let row = section.headerRow; row < footer; row += 1) {
        if (isTimesheetSummaryRowAt(view, row)) continue;
        if (isTimesheetUnitHeaderRow(view, row)) continue;
        const index = cellAt(view, row, 2);
        if (!isPositionIndex(index) || seen.has(index)) continue;
        seen.add(index);
        if (!canonicalStart) canonicalStart = row;
        canonicalEnd = row;
      }
    }
    if (!canonicalStart) {
      canonicalStart = isPositionIndex(cellAt(view, section.headerRow, 2))
        ? section.headerRow
        : section.headerRow + 1;
    }
    if (!canonicalEnd) canonicalEnd = Math.max(section.headerRow, canonicalStart - 1);
    return {
      key: sectionKey(section.label),
      label: section.label,
      kind: section.kind,
      headerRow: section.headerRow,
      canonicalStartRow: canonicalStart,
      canonicalEndRow: canonicalEnd,
      firstSlotRow: canonicalStart,
      lastCanonicalRow: canonicalEnd,
      footerRow: footer,
      explicitFooter: section.explicitFooter,
      historyStartRow: canonicalEnd + 1,
      historyEndRow: section.explicitFooter ? footer - 1 : canonicalEnd,
    };
  });

  const byIndex: Record<string, TimesheetCanonicalSlot> = {};
  const candidatesByIndex = new Map<string, number[]>();
  const sectionByRow = new Map<number, TimesheetSection>();

  for (const section of sections) {
    const mapEnd = Math.max(section.footerRow, section.headerRow);
    for (let row = section.headerRow; row < mapEnd; row += 1) {
      sectionByRow.set(row, section);
    }
    let platoon = "";
    let squad = "";
    const scanEnd = Math.max(section.canonicalEndRow, section.headerRow);
    for (let row = section.headerRow; row <= scanEnd; row += 1) {
      const header = headerDescriptor(view, row);
      if (header?.kind === "platoon") {
        platoon = stripTimesheetDivisionLabel(cellAt(view, row, 2));
        squad = "";
        continue;
      }
      if (header?.kind === "squad") {
        squad = stripTimesheetDivisionLabel(cellAt(view, row, 2));
        continue;
      }
      if (!isCanonicalRowOf(section, row)) continue;
      const index = cellAt(view, row, 2);
      if (!isPositionIndex(index)) continue;
      const rows = candidatesByIndex.get(index) ?? [];
      rows.push(row);
      candidatesByIndex.set(index, rows);
      if (byIndex[index]) continue;
      byIndex[index] = {
        row,
        index,
        section: section.label,
        platoon,
        squad,
        position: cellAt(view, row, 3),
        vos: cellAt(view, row, 4),
        tariff: cellAt(view, row, 5),
        occupantId: cellAt(view, row, 8),
        occupantName: cellAt(view, row, 7),
        occupied: Boolean(cellAt(view, row, 7) || cellAt(view, row, 8)),
      };
    }
  }

  for (const [index, rows] of candidatesByIndex) {
    const emptyStaff = rows.filter(
      (row) => !cellAt(view, row, 7) && !cellAt(view, row, 8),
    );
    if (emptyStaff.length > 1) {
      issues.push({
        code: "TIMESHEET_SLOT_AMBIGUOUS",
        index,
        rows: emptyStaff,
      });
      delete byIndex[index];
    }
  }

  return { sections, byIndex, sectionByRow, issues };
};

export const sectionForTimesheetRow = (
  layout: TimesheetLayout,
  row: number,
) => layout.sectionByRow.get(row) ?? null;

export const resolveCanonicalTimesheetSlot = (input: {
  index: string;
  layout: TimesheetLayout;
}): TimesheetCanonicalSlot => {
  const index = String(input.index || "").trim();
  if (!index) {
    throw new TimesheetLayoutError(
      "CANONICAL_SLOT_MISSING",
      "Немає індексу посади для штатного рядка Табеля.",
    );
  }
  const ambiguous = input.layout.issues.find(
    (issue) => issue.index === index && issue.code === "TIMESHEET_SLOT_AMBIGUOUS",
  );
  if (ambiguous) {
    throw new TimesheetLayoutError(
      "TIMESHEET_SLOT_AMBIGUOUS",
      `Індекс ${index} має кілька штатних рядків у Табелі (R${ambiguous.rows.join(", R")}). Запис заблоковано.`,
      { index, rows: ambiguous.rows },
    );
  }
  const slot = input.layout.byIndex[index];
  if (!slot) {
    throw new TimesheetLayoutError(
      "CANONICAL_SLOT_MISSING",
      `У Табелі немає штатного рядка з індексом ${index}.`,
      { index },
    );
  }
  return slot;
};

export const assertCanonicalSlotFreeFor = (
  slot: TimesheetCanonicalSlot,
  person: { personId?: string; fullName?: string },
) => {
  if (!slot.occupied) return;
  const wantId = String(person.personId || "").trim();
  const wantName = String(person.fullName || "")
    .toLocaleLowerCase("uk-UA")
    .replace(/\s+/g, " ")
    .trim();
  const sameId = Boolean(wantId && slot.occupantId === wantId);
  const sameName =
    Boolean(wantName) &&
    slot.occupantName.toLocaleLowerCase("uk-UA").replace(/\s+/g, " ").trim() ===
      wantName;
  if (sameId || sameName) return;
  throw new TimesheetLayoutError(
    "SLOT_OCCUPIED_CONFLICT",
    `Штатний рядок ${slot.index} (R${slot.row}) зайнятий: ${slot.occupantName || slot.occupantId}. Спочатку закрийте попереднього, не overwrite.`,
    { index: slot.index, rows: [slot.row] },
  );
};

const resolveHistorySection = (input: {
  sourceSlot?: TimesheetCanonicalSlot | null;
  sourceRow?: number;
  layout: TimesheetLayout;
}): TimesheetSection | null => {
  if (input.sourceSlot) {
    return (
      input.layout.sections.find(
        (item) => item.label === input.sourceSlot!.section,
      ) ?? null
    );
  }
  if (!input.sourceRow) return null;
  const section = sectionForTimesheetRow(input.layout, input.sourceRow);
  if (section && isCanonicalRowOf(section, input.sourceRow)) return section;
  return null;
};

export const assertTimesheetLayoutReadyForApply = (layout: TimesheetLayout) => {
  const unknown = layout.issues.find(
    (issue) => issue.code === "TIMESHEET_SECTION_RANGE_UNKNOWN",
  );
  if (!unknown) return;
  throw new TimesheetLayoutError(
    "TIMESHEET_SECTION_RANGE_UNKNOWN",
    `Не прочитано COUNTIF штатного діапазону для секції «${unknown.section || "—"}» (підсумок R${unknown.rows[0] ?? "?"}). Apply заблоковано.`,
    { section: unknown.section, rows: unknown.rows },
  );
};

export const resolveHistoryTimesheetRow = (input: {
  sourceSlot?: TimesheetCanonicalSlot | null;
  sourceRow?: number;
  layout: TimesheetLayout;
  sheet: ExcelSheetSnapshot;
  grid?: Array<unknown[] | undefined>;
  reserved?: Set<number>;
  insertCountBySection?: Map<string, number>;
}): TimesheetHistoryPlacement => {
  const reserved = input.reserved ?? new Set<number>();
  const insertCountBySection = input.insertCountBySection;
  const section = resolveHistorySection(input);
  if (!section) {
    throw new TimesheetLayoutError(
      "SOURCE_SLOT_MISSING",
      "Немає секції підрозділу для історичного рядка Табеля. Секцію беремо лише з канонічного positionIndex, не з поточного рядка історії.",
      {
        index: input.sourceSlot?.index ?? "",
        rows: input.sourceRow ? [input.sourceRow] : [],
      },
    );
  }
  const rangeUnknown = input.layout.issues.find(
    (issue) =>
      issue.code === "TIMESHEET_SECTION_RANGE_UNKNOWN" &&
      issue.section === section.label,
  );
  if (rangeUnknown) {
    throw new TimesheetLayoutError(
      "TIMESHEET_SECTION_RANGE_UNKNOWN",
      `Не прочитано COUNTIF штатного діапазону для секції «${section.label}». Історію не ставимо.`,
      { section: section.label, rows: rangeUnknown.rows },
    );
  }
  const view: TimesheetGridView = { sheet: input.sheet, grid: input.grid };
  const start = Math.max(section.historyStartRow, section.canonicalEndRow + 1);
  const end = section.explicitFooter
    ? section.historyEndRow
    : Math.max(section.historyEndRow, start - 1);
  if (end >= start) {
    for (let row = start; row <= end; row += 1) {
      if (reserved.has(row)) continue;
      if (isTimesheetSummaryRowAt(view, row)) continue;
      if (isTimesheetUnitHeaderRow(view, row)) continue;
      if (!rowIsCompletelyEmpty(view, row)) continue;
      reserved.add(row);
      return {
        row,
        insertBefore: false,
        sectionKey: section.key,
        footerRow: section.footerRow,
      };
    }
  }
  const offset = insertCountBySection?.get(section.key) ?? 0;
  const row = section.footerRow + offset;
  insertCountBySection?.set(section.key, offset + 1);
  reserved.add(row);
  return {
    row,
    insertBefore: true,
    sectionKey: section.key,
    footerRow: section.footerRow,
  };
};

export const takeHistoryTimesheetRow = (
  placement: TimesheetHistoryPlacement,
  pendingInserts: PendingHistoryInsert[],
) => {
  if (placement.insertBefore) {
    pendingInserts.push({
      sectionKey: placement.sectionKey,
      footerRow: placement.footerRow,
      targetRow: placement.row,
    });
  }
  return placement.row;
};

export const stampTimesheetHistoryInserts = (
  writes: Array<{
    row: number;
    insertRowsBefore?: boolean;
    insertRowCount?: number;
  }>,
  pendingInserts: PendingHistoryInsert[],
) => {
  if (!pendingInserts.length) return;
  const byFooter = new Map<number, PendingHistoryInsert[]>();
  for (const item of pendingInserts) {
    const list = byFooter.get(item.footerRow) ?? [];
    list.push(item);
    byFooter.set(item.footerRow, list);
  }
  for (const [footerRow, group] of byFooter) {
    const first =
      writes.find((write) => write.row === footerRow) ||
      writes.find((write) =>
        group.some((item) => item.targetRow === write.row),
      );
    if (!first) continue;
    first.insertRowsBefore = true;
    first.insertRowCount = group.length;
  }
};

export const assertHistoryRowsOutsideCanonical = (
  layout: TimesheetLayout,
  historyRows: number[],
) => {
  for (const row of historyRows) {
    const section = sectionForTimesheetRow(layout, row);
    if (!section) continue;
    if (isCanonicalRowOf(section, row)) {
      throw new TimesheetLayoutError(
        "HISTORY_INSIDE_CANONICAL",
        `Історичний рядок R${row} потрапив у штатний COUNTIF-діапазон секції «${section.label}» (R${section.canonicalStartRow}:R${section.canonicalEndRow}).`,
        { section: section.label, rows: [row] },
      );
    }
  }
};

export const layoutFromTimesheetSheets = (
  timesheet: ExcelSheetSnapshot,
  options?: {
    grid?: Array<unknown[] | undefined>;
    shpoIndexes?: Iterable<string>;
    formulas?: TimesheetFormulaMap;
  },
) => buildTimesheetLayout(timesheet, options);
