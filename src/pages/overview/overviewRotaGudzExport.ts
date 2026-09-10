import JSZip from "jszip";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import type { BackendPersonnelOverviewRow } from "../../api";
import { exportTemplateWorkbookWithMutations } from "../../excelRoundTrip";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import {
  cleanPersonDisplayName,
  formatExcelDateDisplay,
  isLikelyCallSignToken,
  looksLikePersonBirthDate,
  looksLikePersonnelName,
  looksLikePersonnelRankTitle,
  normalizePersonBirthKey,
} from "../personnel/personnelUtils";
import { overviewStatusFilterLabel } from "./overviewRosterMerge";

export const GUDZ_ROTA_REPORT_TEMPLATE_URL =
  "/templates/gudz-rota-report-template.xlsx";

const TEMPLATE_STYLE_ROWS = {
  command: 6,
  platoonHeader: 8,
  platoonSpacer: 9,
  soldier: 11,
  newcomerHeader: 82,
  newcomerRow: 83,
  exitRow: 2,
  guardRow: 2,
} as const;

const COMMAND_SLOTS = [
  { label: "Командир роти", pattern: /командир\s+рот/i },
  { label: "Заступник командира роти", pattern: /заступник\s+командир/i },
  { label: "Головний сержант", pattern: /головн.*сержант/i },
  {
    label: "Сержант з мат. забезпечення",
    pattern: /сержант.*мат|мат\.?\s*забезп/i,
  },
  { label: "Діловод", pattern: /діловод/i },
] as const;

export type RotaGudzMergePlan = {
  mainSheetXmlPath: string;
  headerRows: number[];
  spacerRows: number[];
};

export type RotaGudzPersonRow = {
  sourceOrder: number;
  rank: string;
  name: string;
  callsign: string;
  ipn: string;
  bzvp: string;
  note: string;
  status: string;
  platoon: string;
  position: string;
  isNewcomer: boolean;
};

const staffValue = (row: BackendPersonnelOverviewRow, columnNumber: number) =>
  row.staffSheetColumns?.[`staff_${columnNumber}`]?.trim() ?? "";

const STAFF_NAME_COLUMN_ORDER = [14, 13] as const;

const isValidRotaGudzPersonName = (cleaned: string) => {
  if (!cleaned || cleaned === "Без ПІБ") return false;
  if (!looksLikePersonnelName(cleaned)) return false;
  const parts = cleaned.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return false;
  if (parts.length === 2 && parts[1]!.length <= 2) return false;
  if (/^(полігон|ксп|на\s*виконан|вихід|шпитал|охорон)/i.test(cleaned)) {
    return false;
  }
  return true;
};

const stripEmbeddedCallSignFromName = (raw: string) =>
  String(raw ?? "")
    .replace(/\(([^)]+)\)/g, (match, inner: string) =>
      isLikelyCallSignToken(inner.trim()) ? " " : match,
    )
    .replace(/\s+/g, " ")
    .trim();

export const resolveRotaGudzPersonName = (
  row: BackendPersonnelOverviewRow,
): string => {
  const candidates = [
    row.name.trim(),
    ...STAFF_NAME_COLUMN_ORDER.map((columnNumber) =>
      staffValue(row, columnNumber),
    ),
  ];
  for (const raw of candidates) {
    const cleaned = cleanPersonDisplayName(stripEmbeddedCallSignFromName(raw));
    if (isValidRotaGudzPersonName(cleaned)) return cleaned;
  }
  return "";
};

const normalizeNoteText = (value: string) =>
  normalizeRosterMatchText(value).replace(/\s+/g, " ");

export const resolveRotaGudzNote = (
  row: BackendPersonnelOverviewRow,
): string => {
  const location = staffValue(row, 31);
  const direction = staffValue(row, 33) || row.fighterDirection?.trim() || "";
  const notes = staffValue(row, 32);
  const status = staffValue(row, 21) || row.staffStatus?.trim() || "";
  const blob = `${location} ${direction} ${notes} ${status}`.toLowerCase();

  if (/охорон/.test(blob)) return "охорона";
  if (/на\s*виконан/.test(blob) || /^вихід$/i.test(direction.trim())) {
    return "вихід";
  }
  if (/бпла|fpv|бомбер|пілот/.test(blob)) return "БПЛА";
  if (/шпитал|госпітал|лікуван/.test(blob)) return "шпиталь";
  if (/наметов/.test(blob)) return "наметове місто";
  if (/сзч|самовіл/.test(blob)) return "СЗЧ";
  if (/відпуст/.test(blob)) return "Відпустка лікув.";
  if (/повернувся\s+з\s+виходу/.test(blob)) {
    const match = `${location} ${direction} ${notes}`.match(
      /повернувся\s+з\s+виходу[^.]*/i,
    );
    return match?.[0]?.trim() || "повернувся з виходу";
  }
  if (/відкоманд/.test(blob)) return "відкомандирован до іншого ППД";
  if (/мед\.?\s*рот/.test(blob)) return "Мед. рота";

  return location || direction || notes || status;
};

const formatRotaGudzDate = (raw: string) => {
  const key = normalizePersonBirthKey(raw);
  if (!key) return "";
  const [year, month, day] = key.split("-");
  return `${day}.${month}.${year}`;
};

/** «Курс БЗВП» (кол. 26): дата DD.MM.YYYY, інакше — текст як є. */
export const extractRotaGudzBzvpDate = (raw: string): string => {
  const text = String(raw ?? "").trim();
  if (!text || looksLikePersonBirthDate(text)) return "";

  const formatted = formatExcelDateDisplay(text).trim();
  const match = formatted.match(/(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/);
  if (match) return formatRotaGudzDate(match[1]!);

  const asNumber = Number(text.replace(",", "."));
  if (
    Number.isFinite(asNumber) &&
    asNumber > 40000 &&
    asNumber < 80000 &&
    !/[./-]/.test(text)
  ) {
    const dateFromSerial = formatRotaGudzDate(formatExcelDateDisplay(asNumber));
    if (dateFromSerial) return dateFromSerial;
  }

  return formatted || text;
};

export const resolveRotaGudzBzvp = (row: BackendPersonnelOverviewRow) =>
  extractRotaGudzBzvpDate(staffValue(row, 26));

const resolvePositionText = (row: BackendPersonnelOverviewRow) =>
  `${staffValue(row, 5)} ${staffValue(row, 7)} ${row.positionTitle ?? ""}`.trim();

export const isPlatoonCommanderPosition = (position: string) =>
  /командир\s+(?:\d+\s*)?(?:піхотн\S*\s+)?взвод/i.test(
    normalizeNoteText(position),
  );

export const isSectionCommanderPosition = (position: string) =>
  /командир\s+відділення|ком\.?\s*від/i.test(normalizeNoteText(position));

export const resolveRotaGudzRankLabel = (
  position: string,
  rank: string,
): string => {
  if (isPlatoonCommanderPosition(position)) return "Командир взводу";
  if (isSectionCommanderPosition(position)) return "Командир відділення";
  return rank;
};

const isLikelyRankToken = (value: string) => {
  const text = value.trim();
  if (!text || /^\d+$/.test(text)) return false;
  if (looksLikePersonnelRankTitle(text)) return true;
  if (looksLikePersonnelName(text)) return false;
  if (isLikelyCallSignToken(text)) return false;
  return false;
};

/** Кол. 14 = № рядка, кол. 13 = позивний — типовий зсув штатки. */
export const isStaffNameColumnShifted = (row: BackendPersonnelOverviewRow) => {
  const col14 = staffValue(row, 14);
  const col13 = staffValue(row, 13);
  return /^\d+$/.test(col14) && isLikelyCallSignToken(col13);
};

const sanitizeRankCandidate = (value: string) => {
  const text = value.trim();
  if (!text || /^\d+$/.test(text)) return "";
  if (looksLikePersonnelRankTitle(text)) return text;
  if (looksLikePersonnelName(text)) return "";
  if (isLikelyCallSignToken(text)) return "";
  return "";
};

export const resolveRotaGudzCallsign = (row: BackendPersonnelOverviewRow) => {
  const from15 = staffValue(row, 15);
  if (from15) return from15;
  if (isStaffNameColumnShifted(row)) return staffValue(row, 13);
  return "";
};

const resolveRank = (row: BackendPersonnelOverviewRow) => {
  const shifted = isStaffNameColumnShifted(row);
  const staffRank = staffValue(row, 13);
  const rank = shifted
    ? sanitizeRankCandidate(staffValue(row, 12)) ||
      sanitizeRankCandidate(row.rank?.trim() || "")
    : sanitizeRankCandidate(staffRank) ||
      sanitizeRankCandidate(row.rank?.trim() || "") ||
      "";
  return resolveRotaGudzRankLabel(resolvePositionText(row), rank);
};

const platoonPersonSortKey = (person: RotaGudzPersonRow) => {
  if (isPlatoonCommanderPosition(person.position)) return 0;
  if (isSectionCommanderPosition(person.position)) return 1;
  return 2;
};

const sortPlatoonPeople = (people: RotaGudzPersonRow[]) =>
  [...people].sort((left, right) => {
    const roleDelta =
      platoonPersonSortKey(left) - platoonPersonSortKey(right);
    if (roleDelta !== 0) return roleDelta;
    return left.sourceOrder - right.sourceOrder;
  });

const resolvePlatoonLabel = (row: BackendPersonnelOverviewRow) => {
  const raw = staffValue(row, 3) || staffValue(row, 4);
  const match = raw.match(/(\d+)\s*(?:піхотн\S*\s+)?взвод/i);
  if (match) return `${match[1]} взвод`;
  const short = raw.match(/^(\d+)\s*взвод/i);
  if (short) return `${short[1]} взвод`;
  return raw.trim() || "Без взводу";
};

const platoonSortKey = (label: string) => {
  const match = label.match(/(\d+)/);
  return match ? Number(match[1]) : 999;
};

const isNewcomerRow = (row: BackendPersonnelOverviewRow) => {
  const status = normalizeNoteText(
    `${staffValue(row, 21)} ${row.staffStatusLabel ?? ""} ${row.staffStatus ?? ""}`,
  );
  return status.includes("новоприбул");
};

export const buildRotaGudzPersonRows = (
  rows: BackendPersonnelOverviewRow[],
): RotaGudzPersonRow[] =>
  rows
    .map((row, index) => ({
      sourceOrder: index,
      rank: resolveRank(row),
      name: resolveRotaGudzPersonName(row),
      callsign: resolveRotaGudzCallsign(row),
      ipn: staffValue(row, 19),
      bzvp: String(resolveRotaGudzBzvp(row) ?? ""),
      note: resolveRotaGudzNote(row),
      status: overviewStatusFilterLabel(row),
      platoon: resolvePlatoonLabel(row),
      position: resolvePositionText(row),
      isNewcomer: isNewcomerRow(row),
    }))
    .filter((row) => isValidRotaGudzPersonName(row.name));

export const isRotaGudzExitPerson = (person: RotaGudzPersonRow) =>
  /вихід/i.test(person.note) ||
  /на\s*виконан/i.test(normalizeNoteText(person.status));

export const isRotaGudzGuardPerson = (person: RotaGudzPersonRow) =>
  /охорон/i.test(person.note);

const copyCellStyle = (sourceCell: any, targetCell: any) => {
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
};

const copyRowStyle = (
  sheet: any,
  sourceRow: number,
  targetRow: number,
  endColumn = 8,
) => {
  sheet.row(targetRow).height(sheet.row(sourceRow).height());
  for (let column = 1; column <= endColumn; column += 1) {
    copyCellStyle(sheet.cell(sourceRow, column), sheet.cell(targetRow, column));
  }
};

const columnLetter = (columnNumber: number) =>
  columnNumber <= 26
    ? String.fromCharCode(64 + columnNumber)
    : `A${String.fromCharCode(64 + columnNumber - 26)}`;

const unmergeRow = (sheet: any, rowNumber: number, endColumn = 8) => {
  try {
    sheet
      .range(`A${rowNumber}`, `${columnLetter(endColumn)}${rowNumber}`)
      .merged(false);
  } catch {
    /* row was not merged */
  }
};

const clearRowValues = (sheet: any, rowNumber: number, endColumn = 8) => {
  unmergeRow(sheet, rowNumber, endColumn);
  for (let column = 1; column <= endColumn; column += 1) {
    sheet.cell(rowNumber, column).value(null);
  }
};

const cellText = (value: string | null | undefined) => {
  const text = String(value ?? "").trim();
  return text || null;
};

const writePersonCells = (
  sheet: any,
  rowNumber: number,
  person: Pick<
    RotaGudzPersonRow,
    "rank" | "name" | "callsign" | "ipn" | "bzvp" | "note"
  >,
  columns: {
    global?: number;
    local?: number;
  },
) => {
  unmergeRow(sheet, rowNumber, 8);
  sheet.cell(rowNumber, 1).value(columns.global ?? null);
  sheet.cell(rowNumber, 2).value(columns.local ?? null);
  sheet.cell(rowNumber, 3).value(cellText(person.rank));
  sheet.cell(rowNumber, 4).value(cellText(person.name));
  sheet.cell(rowNumber, 5).value(cellText(person.callsign));
  sheet.cell(rowNumber, 6).value(cellText(person.ipn));
  sheet.cell(rowNumber, 7).value(cellText(person.bzvp));
  sheet.cell(rowNumber, 8).value(cellText(person.note));
};

const writePersonRow = (
  sheet: any,
  rowNumber: number,
  person: RotaGudzPersonRow,
  columns: {
    global?: number;
    local?: number;
  },
) => {
  clearRowValues(sheet, rowNumber);
  copyRowStyle(sheet, TEMPLATE_STYLE_ROWS.soldier, rowNumber);
  writePersonCells(sheet, rowNumber, person, columns);
};

const findCommandPerson = (
  people: RotaGudzPersonRow[],
  pattern: RegExp,
  used: Set<string>,
) => {
  const match = people.find(
    (person) =>
      !used.has(person.name) && pattern.test(normalizeNoteText(person.position)),
  );
  if (match) used.add(match.name);
  return match;
};

const writeSectionHeader = (
  sheet: any,
  rowNumber: number,
  styleRow: number,
  label: string,
  endColumn = 8,
) => {
  clearRowValues(sheet, rowNumber, endColumn);
  copyRowStyle(sheet, styleRow, rowNumber, endColumn);
  sheet.cell(rowNumber, 1).value(label);
  for (let column = 2; column <= endColumn; column += 1) {
    copyCellStyle(sheet.cell(rowNumber, 1), sheet.cell(rowNumber, column));
  }
  sheet.cell(rowNumber, 1).style("horizontalAlignment", "center");
  try {
    sheet
      .range(`A${rowNumber}`, `${columnLetter(endColumn)}${rowNumber}`)
      .merged(true);
  } catch {
    /* ignore */
  }
  unmergeRow(sheet, rowNumber + 1, endColumn);
};

/** Як у шаблоні: рядок «КОМАНДИР» (merge A:D) між заголовком взводу і списком. */
const writePlatoonSpacerRow = (sheet: any, rowNumber: number) => {
  clearRowValues(sheet, rowNumber, 8);
  copyRowStyle(sheet, TEMPLATE_STYLE_ROWS.platoonSpacer, rowNumber, 8);
  sheet.cell(rowNumber, 1).value("КОМАНДИР");
  try {
    sheet.range(`A${rowNumber}`, `D${rowNumber}`).merged(true);
  } catch {
    /* ignore */
  }
  unmergeRow(sheet, rowNumber + 1, 8);
};

const columnNumber = (letters: string) => {
  let value = 0;
  for (const letter of letters.toUpperCase()) {
    value = value * 26 + (letter.charCodeAt(0) - 64);
  }
  return value;
};

const parseMergeRef = (ref: string) => {
  const match = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
  if (!match) return null;
  return {
    row: Number(match[2]),
    endRow: Number(match[4]),
    startCol: columnNumber(match[1]!),
    endCol: columnNumber(match[3]!),
  };
};

export const isAllowedRotaGudzMerge = (
  ref: string,
  allowed: { headerRows: Set<number>; spacerRows: Set<number> },
) => {
  const parsed = parseMergeRef(ref);
  if (!parsed || parsed.row !== parsed.endRow) return false;
  if (allowed.headerRows.has(parsed.row) && parsed.startCol === 1 && parsed.endCol === 8) {
    return true;
  }
  if (allowed.spacerRows.has(parsed.row) && parsed.startCol === 1 && parsed.endCol === 4) {
    return true;
  }
  return false;
};

export const sanitizeRotaGudzMainSheetMergeXml = (
  sheetXml: string,
  allowed: { headerRows: number[]; spacerRows: number[] },
): string => {
  const headerRows = new Set(allowed.headerRows);
  const spacerRows = new Set(allowed.spacerRows);
  const mergeBlockRe = /<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/;
  const blockMatch = sheetXml.match(mergeBlockRe);
  if (!blockMatch) return sheetXml;

  const kept: string[] = [];
  for (const match of blockMatch[1]!.matchAll(/<mergeCell ref="([^"]+)"\s*\/>/g)) {
    const ref = match[1]!;
    const parsed = parseMergeRef(ref);
    if (parsed?.row && parsed.row >= 8 && parsed.row <= 220) {
      if (isAllowedRotaGudzMerge(ref, { headerRows, spacerRows })) {
        kept.push(match[0]!);
      }
      continue;
    }
    kept.push(match[0]!);
  }

  if (!kept.length) {
    return sheetXml.replace(mergeBlockRe, "");
  }
  return sheetXml.replace(
    mergeBlockRe,
    `<mergeCells count="${kept.length}">${kept.join("")}</mergeCells>`,
  );
};

export const sanitizeRotaGudzWorkbookBuffer = async (
  buffer: ArrayBuffer,
  mergePlan: RotaGudzMergePlan,
) => {
  const zip = await JSZip.loadAsync(buffer);
  const sheetFile = zip.file(mergePlan.mainSheetXmlPath);
  if (!sheetFile) return buffer;
  const sheetXml = await sheetFile.async("string");
  zip.file(
    mergePlan.mainSheetXmlPath,
    sanitizeRotaGudzMainSheetMergeXml(sheetXml, mergePlan),
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const writeMainSheet = (
  sheet: any,
  unitLabel: string,
  people: RotaGudzPersonRow[],
): RotaGudzMergePlan => {
  const mergePlan: RotaGudzMergePlan = {
    mainSheetXmlPath: "xl/worksheets/sheet1.xml",
    headerRows: [],
    spacerRows: [],
  };
  sheet.name(`ОС ${unitLabel}`);
  sheet.cell(1, 1).value(`${unitLabel.toUpperCase()} 1ПБ`);
  try {
    sheet.range("A8", "M220").merged(false);
  } catch {
    /* ignore */
  }
  for (let rowNumber = 8; rowNumber <= 220; rowNumber += 1) {
    clearRowValues(sheet, rowNumber, 13);
  }

  const usedCommand = new Set<string>();
  COMMAND_SLOTS.forEach((slot, index) => {
    const rowNumber = 3 + index;
    const person = findCommandPerson(people, slot.pattern, usedCommand);
    copyRowStyle(sheet, TEMPLATE_STYLE_ROWS.command, rowNumber);
    clearRowValues(sheet, rowNumber);
    sheet.cell(rowNumber, 1).value(slot.label);
    if (person) {
      sheet.cell(rowNumber, 3).value(person.rank || null);
      sheet.cell(rowNumber, 4).value(person.name);
      sheet.cell(rowNumber, 5).value(person.callsign || null);
      sheet.cell(rowNumber, 6).value(person.ipn || null);
      sheet.cell(rowNumber, 7).value(person.bzvp?.trim() || null);
      sheet.cell(rowNumber, 8).value(person.note || null);
    }
  });

  const regularPeople = people.filter((person) => !usedCommand.has(person.name));
  const newcomers = regularPeople.filter((person) => person.isNewcomer);
  const platoonPeople = regularPeople.filter((person) => !person.isNewcomer);

  const platoons = [...new Set(platoonPeople.map((person) => person.platoon))].sort(
    (left, right) => platoonSortKey(left) - platoonSortKey(right),
  );

  let nextRow = 8;
  let globalIndex = 1;

  platoons.forEach((platoonLabel) => {
    writeSectionHeader(
      sheet,
      nextRow,
      TEMPLATE_STYLE_ROWS.platoonHeader,
      platoonLabel,
    );
    mergePlan.headerRows.push(nextRow);
    nextRow += 1;
    writePlatoonSpacerRow(sheet, nextRow);
    mergePlan.spacerRows.push(nextRow);
    nextRow += 1;
    unmergeRow(sheet, nextRow, 8);
    let platoonIndex = 1;
    sortPlatoonPeople(
      platoonPeople.filter((person) => person.platoon === platoonLabel),
    ).forEach((person) => {
        writePersonRow(sheet, nextRow, person, {
          global: globalIndex,
          local: platoonIndex,
        });
        globalIndex += 1;
        platoonIndex += 1;
        nextRow += 1;
      });
  });

  if (newcomers.length) {
    writeSectionHeader(
      sheet,
      nextRow,
      TEMPLATE_STYLE_ROWS.newcomerHeader,
      "Новоприбулі",
    );
    mergePlan.headerRows.push(nextRow);
    nextRow += 1;
    writePlatoonSpacerRow(sheet, nextRow);
    mergePlan.spacerRows.push(nextRow);
    nextRow += 1;
    unmergeRow(sheet, nextRow, 8);
    let newcomerIndex = 1;
    newcomers.forEach((person) => {
      writePersonRow(sheet, nextRow, person, {
        global: globalIndex,
        local: newcomerIndex,
      });
      globalIndex += 1;
      newcomerIndex += 1;
      nextRow += 1;
    });
  }

  for (let rowNumber = nextRow; rowNumber <= 220; rowNumber += 1) {
    clearRowValues(sheet, rowNumber, 13);
  }

  return mergePlan;
};

const writeFlatSheet = (
  sheet: any,
  people: RotaGudzPersonRow[],
  includeStatus: boolean,
  styleRow: number,
) => {
  for (let rowNumber = 2; rowNumber <= 120; rowNumber += 1) {
    clearRowValues(sheet, rowNumber, includeStatus ? 8 : 7);
  }
  people.forEach((person, index) => {
    const rowNumber = 2 + index;
    const endColumn = includeStatus ? 8 : 7;
    clearRowValues(sheet, rowNumber, endColumn);
    copyRowStyle(sheet, styleRow, rowNumber, endColumn);
    sheet.cell(rowNumber, 1).value(index + 1);
    sheet.cell(rowNumber, 2).value(cellText(person.rank));
    sheet.cell(rowNumber, 3).value(cellText(person.name));
    sheet.cell(rowNumber, 4).value(cellText(person.callsign));
    sheet.cell(rowNumber, 5).value(cellText(person.ipn));
    sheet.cell(rowNumber, 6).value(cellText(person.bzvp));
    sheet.cell(rowNumber, 7).value(cellText(person.note));
    if (includeStatus) {
      sheet.cell(rowNumber, 8).value(cellText(person.status));
    }
  });
};

export const writeRotaGudzReportWorkbook = (
  workbook: any,
  options: {
    unitLabel: string;
    rows: BackendPersonnelOverviewRow[];
  },
): RotaGudzMergePlan => {
  const people = buildRotaGudzPersonRows(options.rows);
  const mainSheet =
    workbook.sheet("ОС 3 роти") ??
    workbook.sheets().find((sheet: any) => /ос\s+\d/i.test(String(sheet.name()))) ??
    workbook.sheet(0);
  const exitSheet = workbook.sheet("Вихід") ?? workbook.sheet(1);
  const guardSheet = workbook.sheet("Охорона") ?? workbook.sheet(2);

  const mergePlan = writeMainSheet(mainSheet, options.unitLabel, people);
  writeFlatSheet(
    exitSheet,
    people.filter(isRotaGudzExitPerson),
    true,
    TEMPLATE_STYLE_ROWS.exitRow,
  );
  writeFlatSheet(
    guardSheet,
    people.filter(isRotaGudzGuardPerson),
    false,
    TEMPLATE_STYLE_ROWS.guardRow,
  );
  return mergePlan;
};

export const resolveSelectedOverviewUnit = (
  filters: SciDataTableExportContext<BackendPersonnelOverviewRow>["filters"],
) => {
  const values =
    filters?.find((filter) => filter.id === "unit")?.values.filter(Boolean) ??
    [];
  if (values.length !== 1) return null;
  return values[0]!.trim();
};

export const buildRotaGudzExportFileName = (
  unitLabel: string,
  reportDate = new Date(),
) => {
  const dateLabel = new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
  }).format(reportDate);
  const cleaned = unitLabel.replace(/\s+/g, " ").trim();
  return `${cleaned} станом на ${dateLabel}..xlsx`;
};

export const exportOverviewRotaGudzReport = async (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
) => {
  const unitLabel = resolveSelectedOverviewUnit(context.filters);
  if (!unitLabel) {
    throw new Error(
      "Оберіть одну роту у фільтрі колонки «Підрозділ» перед експортом звіту.",
    );
  }

  const selectedKey = normalizeRosterMatchText(unitLabel);
  const rows = (context.allRows ?? context.rows).filter(
    (row) => normalizeRosterMatchText(row.unit) === selectedKey,
  );
  if (!rows.length) {
    throw new Error(`Немає осіб для роти «${unitLabel}».`);
  }

  let mergePlan: RotaGudzMergePlan = {
    mainSheetXmlPath: "xl/worksheets/sheet1.xml",
    headerRows: [],
    spacerRows: [],
  };

  await exportTemplateWorkbookWithMutations(
    GUDZ_ROTA_REPORT_TEMPLATE_URL,
    (workbook) => {
      mergePlan = writeRotaGudzReportWorkbook(workbook, {
        unitLabel,
        rows,
      });
    },
    buildRotaGudzExportFileName(unitLabel),
    (buffer) => sanitizeRotaGudzWorkbookBuffer(buffer, mergePlan),
  );
};
