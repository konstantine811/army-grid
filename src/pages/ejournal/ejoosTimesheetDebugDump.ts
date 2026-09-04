import type { CellValue, ExcelSheetSnapshot, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  parseEjoosTimesheetPeople,
  type EjoosTimesheetPersonScan,
} from "./ejoosSyncPlan";
import {
  companyUnitPhrasesFromPosition,
  effectiveTimesheetCompanyDivision,
  effectiveTimesheetDivision,
  findTimesheetAppendRowForUnit,
  findTimesheetUnitSectionBounds,
  isTimesheetUnitHeaderRow,
  parseTimesheetUnitDescriptor,
  stripTimesheetDivisionLabel,
} from "./ejoosTimesheetUnitSections";

/** A=1 … AN=40 у Табелі. */
const TIMESHEET_LAST_COL = 40;

export type TimesheetRowKind =
  | "sheet_header"
  | "section_header"
  | "summary"
  | "person"
  | "staff_index_only"
  | "empty"
  | "other";

export type TimesheetDebugRow = {
  excelRow: number;
  kind: TimesheetRowKind;
  cells: string[];
  colA_raw: string;
  colA_parsed: string;
  unitKind: string;
  colB_index: string;
  colC_vos: string;
  colD: string;
  colE_tariff: string;
  colF_rank: string;
  colG_name: string;
  colH_id: string;
  days: string[];
  colAN_present: string;
  companyAbove: string;
  divisionAbove: string;
  hasDeparture: boolean;
  departureDay: number;
  departureText: string;
  plusDays: number[];
};

export type TimesheetSectionDebug = {
  headerRow: number;
  endRow: number;
  division: string;
  unitKind: string;
  appendRowForDeparture: number;
};

const cell = (row: CellValue[] | undefined, col: number) =>
  String(row?.[col] ?? "").trim();

const rowCells = (row: CellValue[] | undefined) =>
  Array.from({ length: TIMESHEET_LAST_COL }, (_, index) => cell(row, index));

const isSummaryRow = (text: string) =>
  /на\s*продоволь/iu.test(text) ||
  /перебува(?:є|e)\s*\d/iu.test(text) ||
  /^\s*всього\b/iu.test(text);

const classifyRow = (
  sheet: ExcelSheetSnapshot,
  excelRow: number,
  personByRow: Map<number, EjoosTimesheetPersonScan>,
): TimesheetDebugRow => {
  const raw = sheet.rawRows[excelRow - 1] ?? [];
  const cells = rowCells(raw);
  const colA_raw = cells[0] ?? "";
  const colA_parsed = colA_raw ? stripTimesheetDivisionLabel(colA_raw) : "";
  const colB = cells[1] ?? "";
  const view = { sheet };
  const person = personByRow.get(excelRow);
  const days = cells.slice(8, 39);

  let kind: TimesheetRowKind = "empty";
  if (excelRow < 7) kind = "sheet_header";
  else if (colA_parsed && isSummaryRow(colA_parsed)) kind = "summary";
  else if (isTimesheetUnitHeaderRow(view, excelRow)) kind = "section_header";
  else if (person) kind = "person";
  else if (/^\d{5,}$/.test(colB)) kind = "staff_index_only";
  else if (cells.some(Boolean)) kind = "other";

  const descriptor = colA_parsed ? parseTimesheetUnitDescriptor(colA_parsed) : null;

  return {
    excelRow,
    kind,
    cells,
    colA_raw,
    colA_parsed,
    unitKind: descriptor?.kind ?? "",
    colB_index: colB,
    colC_vos: cells[2] ?? "",
    colD: cells[3] ?? "",
    colE_tariff: cells[4] ?? "",
    colF_rank: cells[5] ?? "",
    colG_name: cells[6] ?? "",
    colH_id: cells[7] ?? "",
    days,
    colAN_present: cells[39] ?? "",
    companyAbove:
      excelRow >= 7 ? effectiveTimesheetCompanyDivision(view, excelRow) : "",
    divisionAbove: excelRow >= 7 ? effectiveTimesheetDivision(view, excelRow) : "",
    hasDeparture: Boolean(person?.hasDepartureText),
    departureDay: person?.firstDepartureDay ?? 0,
    departureText: person?.departureText ?? "",
    plusDays: person?.plusDays ?? [],
  };
};

export type TimesheetDebugDumpOptions = {
  positionTitle?: string;
  sourceRow?: number;
  maxRow?: number;
  /** Друкувати rawRows/people/rows/sections масивами (важко для консолі). */
  verbose?: boolean;
  /** Рахувати append≈R для кожної секції (повільно без positionTitle). */
  includeSectionAppend?: boolean;
  /** Скільки рядків у текстовому дампі; решта — «… ще N рядків». */
  maxFormattedRows?: number;
};

export const isTimesheetVerboseDebugEnabled = () =>
  typeof localStorage !== "undefined" &&
  localStorage.getItem("ejoosTimesheetDebug") === "1";

/** Знімок структури Табеля для дебагу правил пошуку секцій. */
export const buildTimesheetDebugDump = (
  sheet: ExcelSheetSnapshot | undefined,
  options?: TimesheetDebugDumpOptions,
) => {
  if (!sheet) {
    return {
      error: "Табель не знайдено",
      rows: [] as TimesheetDebugRow[],
      sections: [] as TimesheetSectionDebug[],
      people: [] as EjoosTimesheetPersonScan[],
      rawRows: [] as CellValue[][],
    };
  }

  const people = parseEjoosTimesheetPeople(sheet);
  const personByRow = new Map(people.map((row) => [row.excelRow, row]));
  const maxRow = Math.min(
    options?.maxRow ?? sheet.rawRows.length,
    sheet.rawRows.length,
  );

  const rows: TimesheetDebugRow[] = [];
  for (let excelRow = 1; excelRow <= maxRow; excelRow += 1) {
    rows.push(classifyRow(sheet, excelRow, personByRow));
  }

  const includeSectionAppend =
    options?.includeSectionAppend ?? Boolean(options?.positionTitle);
  const sectionHeaders = rows.filter((row) => row.kind === "section_header");
  const sections: TimesheetSectionDebug[] = sectionHeaders.map((header) => {
    const bounds = findTimesheetUnitSectionBounds({ sheet }, header.colA_parsed);
    const appendRowForDeparture =
      includeSectionAppend && bounds
        ? findTimesheetAppendRowForUnit(sheet, {
            sourceRow: options?.sourceRow ?? header.excelRow,
            positionTitle: options?.positionTitle ?? header.colA_parsed,
            placementScope: "company",
          })
        : 0;
    return {
      headerRow: bounds?.headerRow ?? header.excelRow,
      endRow: bounds?.endRow ?? header.excelRow,
      division: header.colA_parsed,
      unitKind: header.unitKind,
      appendRowForDeparture,
    };
  });

  const placementHint = options?.positionTitle
    ? {
        positionTitle: options.positionTitle,
        companyPhrases: companyUnitPhrasesFromPosition(options.positionTitle),
        appendRow: findTimesheetAppendRowForUnit(sheet, {
          sourceRow: options.sourceRow,
          positionTitle: options.positionTitle,
          placementScope: "company",
        }),
      }
    : null;

  return {
    sheetName: sheet.sheetName,
    totalRawRows: sheet.rawRows.length,
    peopleCount: people.length,
    sectionHeaderCount: sectionHeaders.length,
    rows,
    sections,
    people,
    rawRows: sheet.rawRows.slice(0, maxRow),
    placementHint,
  };
};

const formatDayMarks = (days: string[]) => {
  const compact = days
    .map((mark, index) => (mark ? `${index + 1}:${mark}` : ""))
    .filter(Boolean);
  return compact.length ? compact.join(" ") : "";
};

export const formatTimesheetDebugDump = (
  dump: ReturnType<typeof buildTimesheetDebugDump>,
  formatOptions?: { maxFormattedRows?: number; includeSections?: boolean },
) => {
  const lines: string[] = [];
  lines.push(`=== Табель: ${dump.sheetName ?? "?"} ===`);
  if ("error" in dump && dump.error) {
    lines.push(String(dump.error));
    return lines.join("\n");
  }
  lines.push(
    `рядків: ${dump.totalRawRows}, осіб (parseEjoosTimesheetPeople): ${dump.peopleCount}, секцій: ${dump.sectionHeaderCount}`,
  );
  lines.push(
    "A=підрозділ B=індекс C=VOS D=тариф? E=тариф F=звання G=ПІБ H=ID I..=дні 1..31 AN=дні+",
  );
  lines.push("");

  const maxFormattedRows = formatOptions?.maxFormattedRows;
  const visibleRows =
    maxFormattedRows && maxFormattedRows > 0
      ? dump.rows.slice(0, maxFormattedRows)
      : dump.rows;
  for (const row of visibleRows) {
    const parts = [
      `R${String(row.excelRow).padStart(4, " ")}`,
      row.kind.padEnd(16, " "),
      row.colA_raw ? `A:${row.colA_raw}` : "",
      row.colB_index ? `B:${row.colB_index}` : "",
      row.colC_vos ? `C:${row.colC_vos}` : "",
      row.colD ? `D:${row.colD}` : "",
      row.colE_tariff ? `E:${row.colE_tariff}` : "",
      row.colF_rank ? `F:${row.colF_rank}` : "",
      row.colG_name ? `G:${row.colG_name}` : "",
      row.colH_id ? `H:${row.colH_id}` : "",
      formatDayMarks(row.days) ? `days:${formatDayMarks(row.days)}` : "",
      row.colAN_present ? `AN:${row.colAN_present}` : "",
      row.companyAbove ? `company↑:${row.companyAbove}` : "",
      row.hasDeparture
        ? `DEP:${row.departureDay} "${row.departureText}"`
        : "",
    ].filter(Boolean);
    lines.push(parts.join(" | "));
  }
  if (maxFormattedRows && dump.rows.length > maxFormattedRows) {
    lines.push(
      `… ще ${dump.rows.length - maxFormattedRows} рядків (localStorage.ejoosTimesheetDebug=1 для повного дампу)`,
    );
  }

  if (dump.placementHint) {
    lines.push("");
    lines.push("=== placementHint ===");
    lines.push(`positionTitle: ${dump.placementHint.positionTitle}`);
    lines.push(`companyPhrases: ${dump.placementHint.companyPhrases.join(" | ")}`);
    lines.push(`appendRow: ${dump.placementHint.appendRow}`);
  }

  const includeSections = formatOptions?.includeSections ?? true;
  if (includeSections && dump.sections.length) {
    lines.push("");
    lines.push("=== Секції (company scope) ===");
    for (const section of dump.sections) {
      lines.push(
        `'${section.division}' (${section.unitKind}) R${section.headerRow}–${section.endRow} → append≈R${section.appendRowForDeparture}`,
      );
    }
  }

  return lines.join("\n");
};

export const dumpTimesheetFromWorkbook = (
  workbook: ExcelWorkbookSnapshot,
  options?: Parameters<typeof buildTimesheetDebugDump>[1],
) => {
  const sheet = workbook.sheets.find((item) => /табель/i.test(item.sheetName));
  return buildTimesheetDebugDump(sheet, options);
};

export const logTimesheetDebugDump = (
  sheet: ExcelSheetSnapshot | undefined,
  options?: TimesheetDebugDumpOptions,
) => {
  const verbose = options?.verbose ?? isTimesheetVerboseDebugEnabled();
  const dump = buildTimesheetDebugDump(sheet, options);
  const maxFormattedRows =
    options?.maxFormattedRows ?? (verbose ? undefined : 100);
  const includeSections =
    options?.includeSectionAppend ?? Boolean(options?.positionTitle);
  console.log(
    formatTimesheetDebugDump(dump, { maxFormattedRows, includeSections }),
  );
  if (verbose) {
    console.log("[ЕЖООС Табель] rawRows (усі комірки)", dump.rawRows);
    console.log("[ЕЖООС Табель] parseEjoosTimesheetPeople", dump.people);
    console.log("[ЕЖООС Табель] rows (класифікація)", dump.rows);
    console.log("[ЕЖООС Табель] sections", dump.sections);
  }
  return dump;
};

const positionTitleFromApplyError = (message: string) =>
  message.match(/секції підрозділу:\s*(.+?)\.\s*(?:Деталі|$)/iu)?.[1]?.trim() ||
  "";

export const formatApplyErrorWithTimesheetDump = (
  workbook: ExcelWorkbookSnapshot | null | undefined,
  err: unknown,
  fallback = "Не вдалося застосувати",
) => {
  const base = err instanceof Error ? err.message : fallback;
  if (/консолі \(F12\)/i.test(base)) {
    return base;
  }
  const sheet = workbook?.sheets.find((item) => /табель/i.test(item.sheetName));
  if (sheet) {
    const positionTitle = positionTitleFromApplyError(base);
    console.group("[ЕЖООС Табель] dump після помилки застосування");
    logTimesheetDebugDump(sheet, {
      positionTitle: positionTitle || undefined,
      maxRow: Math.min(sheet.rawRows.length, 220),
      maxFormattedRows: positionTitle ? undefined : 80,
      includeSectionAppend: Boolean(positionTitle),
    });
    console.groupEnd();
  }
  if (sheet) {
    return `${base} Деталі Табеля — у консолі (F12).`;
  }
  return base;
};
