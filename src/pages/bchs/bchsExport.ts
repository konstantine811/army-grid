import type { CellValue } from "../../excelRoundTrip";
import { getColumnLabel, valueToDisplay } from "../../excelRoundTrip";
import {
  BCHS_PERCENT_COLUMNS,
  bchsToNumber,
  columnLetterToIndex,
  createBchsComparisonRow,
  emptyBchsSupplementRow,
  normalizeBchsText,
} from "./bchsCalc";
import type {
  BchsAnalyticsSnapshot,
  BchsAnalyticsTableRow,
  BchsComparisonRow,
  BchsSupplementRow,
  BchsSupplementSnapshot,
  GeneratedSheet,
  GeneratedWorkbook,
} from "./bchsTypes";

export const setWorkbookCellValue = (
  sheet: any,
  rowNumber: number,
  columnLetter: string,
  value: unknown,
  styles: Record<string, unknown> = {},
) => {
  const normalized =
    typeof value === "number" && value === 0 ? null : (value ?? null);
  const cell = sheet.cell(rowNumber, columnLetterToIndex(columnLetter) + 1);
  cell.value(normalized);
  if (cell && typeof cell.style === "function") {
    Object.entries(styles).forEach(([key, styleValue]) => {
      cell.style(key, styleValue);
    });
  }
  return cell;
};

export const BCHS_TEXT_EXPORT_COLUMNS = new Set(["B", "P", "AO", "AT"]);
export const BCHS_EXPORT_DATA_START_ROW = 11;
export const BCHS_EXPORT_DATA_END_ROW = 28;
export const BCHS_EXPORT_START_COLUMN = columnLetterToIndex("B");
export const BCHS_EXPORT_END_COLUMN = columnLetterToIndex("BL");
/** Whole percents in Excel: 0.845 → 85%. */
export const BCHS_PERCENT_NUMBER_FORMAT = "0%";

export const BCHS_MAIN_EXPORT_UNIT_ORDER = [
  "Командування",
  "1 піхотна рота",
  "2 піхотна рота",
  "3 піхотна рота",
  "Взвод зв'язку",
  "Взвод логістично-евакуаційних безпілотних наземних систем",
  "Взвод матеріально-технічного забезпечення",
  "Взвод перехоплювачів безпілотних літальних апаратів",
  "Взвод протитанкових ракетних комплексів",
  "Відділення радіоелектронної боротьби",
  "Гранатометний взвод",
  "Група безпілотних систем",
  "Зенітно-ракетний взвод",
  "Інженерно-саперне відділення",
  "Кулеметний взвод",
  "Медичний пункт",
  "Мінометний взвод",
  "Розвідувальний взвод",
  "Рота безпілотних авіаційних комплексів",
] as const;

const BCHS_MAIN_DATA_COLUMN_FILLS: Array<[string, string, string]> = [
  ["B", "B", "ddebf7"],
  ["C", "F", "b7dee8"],
  ["G", "J", "d8f0d0"],
  ["K", "K", "ffc000"],
  ["L", "O", "8ed973"],
  ["P", "P", "ffffff"],
  ["Q", "T", "fce4d6"],
  ["U", "U", "44b3d5"],
  ["V", "Y", "b7e1a1"],
  ["Z", "AA", "d9eaf7"],
  ["AB", "AI", "ead0e9"],
  ["AK", "AT", "ffffff"],
  ["AU", "AU", "b7e1a1"],
  ["AZ", "AZ", "b7e1a1"],
  ["BA", "BA", "ffc000"],
  ["BB", "BE", "d9eaf7"],
  ["BF", "BL", "8ed973"],
];

export const styleBchsMainDataRows = (
  sheet: any,
  rowNumbers: number[],
) => {
  rowNumbers
    .filter((rowNumber) => Number.isFinite(rowNumber) && rowNumber > 0)
    .forEach((rowNumber) => {
      const rowRange = sheet.range(`B${rowNumber}`, `BL${rowNumber}`);
      rowRange.style({
        border: true,
        fontFamily: "Times New Roman",
        fontSize: 20,
        bold: true,
        verticalAlignment: "center",
      });
      BCHS_MAIN_DATA_COLUMN_FILLS.forEach(([start, end, fill]) => {
        sheet.range(`${start}${rowNumber}`, `${end}${rowNumber}`).style("fill", fill);
      });
      ["K", "U", "BA"].forEach((letter) => {
        sheet
          .cell(rowNumber, columnLetterToIndex(letter) + 1)
          .style("numberFormat", BCHS_PERCENT_NUMBER_FORMAT)
          .style("bold", true);
      });
    });
};

/**
 * Normalize any percent cell (ratio, "84,5%", 84.5) to a whole-percent ratio.
 * Example: 0.845 / "84.5%" → 0.85 (displays as 85% with format "0%").
 */
export const toBchsExportPercentRatio = (raw: unknown) => {
  const percent = bchsToNumber(raw);
  if (!Number.isFinite(percent) || percent === 0) return null;
  return Math.round(percent * 100) / 100;
};

export const looksLikeExcelFormulaText = (text: string) =>
  /^=/.test(text) ||
  /\b(COUNTIFS?|SUMIFS?|AVERAGEIFS?|IFERROR|VLOOKUP|XLOOKUP|SUM)\s*\(/i.test(
    text,
  );

export const resolveBchsExportCellValue = (letter: string, raw: unknown) => {
  if (raw == null) return null;
  if (BCHS_TEXT_EXPORT_COLUMNS.has(letter)) {
    const text = valueToDisplay(raw as CellValue).trim();
    if (!text || looksLikeExcelFormulaText(text)) return null;
    return text;
  }
  if (BCHS_PERCENT_COLUMNS.has(letter)) return toBchsExportPercentRatio(raw);

  if (typeof raw === "number") return raw === 0 ? null : raw;
  if (typeof raw === "boolean") return raw;
  const text = valueToDisplay(raw as CellValue).trim();
  if (!text || looksLikeExcelFormulaText(text)) return null;
  // Keep non-numeric text (rare labels) as-is; otherwise store a real number
  // so Windows/preview don't depend on COUNTIFS formulas.
  if (/[A-Za-zА-Яа-яІіЇїЄєҐґ]/.test(text) && !/^-?\d+([.,]\d+)?%?$/.test(text)) {
    return text;
  }
  const numeric = bchsToNumber(text);
  if (!Number.isFinite(numeric) || numeric === 0) return null;
  return numeric;
};

/**
 * xlsx-populate saves formula cells WITHOUT cached <v> values.
 * Mac Excel recalculates on open; messenger preview / many Windows opens do not,
 * so C/D/E etc. look empty. Convert formula+cached-value cells into plain values.
 */
export const materializeBchsSheetFormulas = (sheet: any) => {
  for (
    let rowNumber = BCHS_EXPORT_DATA_START_ROW;
    rowNumber <= BCHS_EXPORT_DATA_END_ROW;
    rowNumber += 1
  ) {
    for (
      let columnIndex = BCHS_EXPORT_START_COLUMN;
      columnIndex <= BCHS_EXPORT_END_COLUMN;
      columnIndex += 1
    ) {
      const cell = sheet.cell(rowNumber, columnIndex + 1);
      const formula =
        typeof cell.formula === "function" ? cell.formula() : undefined;
      if (!formula || formula === "SHARED") continue;

      const cached = typeof cell.value === "function" ? cell.value() : undefined;
      if (cached == null || cached === "") continue;

      // Drop explicit zeros so export cells stay blank instead of "0".
      if (cached === 0) {
        cell.value(null);
        continue;
      }

      // .value() clears the formula and keeps the number/text for consumers
      // that do not recalculate (chat preview, Windows Excel without calc).
      cell.value(cached);
    }
  }
};

/**
 * Overlay concrete analytics values (including recalculated AK–AT) on top of
 * formula-materialized cells. Skip empty values so we don't wipe cached numbers.
 */
export const materializeBchsAnalyticsTableRow = (
  sheet: any,
  tableRow: BchsAnalyticsTableRow,
) => {
  Object.entries(tableRow.values).forEach(([letter, raw]) => {
    const value = resolveBchsExportCellValue(letter, raw);
    if (value == null || value === "") {
      // Explicit zeros must clear any previously materialized/cached value.
      const asNumber =
        typeof raw === "number"
          ? raw
          : typeof raw === "string" && /^-?\d+([.,]\d+)?%?$/.test(raw.trim())
            ? bchsToNumber(raw)
            : null;
      if (asNumber === 0) {
        setWorkbookCellValue(sheet, tableRow.rowNumber, letter, null);
      }
      return;
    }
    setWorkbookCellValue(
      sheet,
      tableRow.rowNumber,
      letter,
      value,
      BCHS_PERCENT_COLUMNS.has(letter)
        ? { numberFormat: BCHS_PERCENT_NUMBER_FORMAT }
        : {},
    );
  });
};

export const writeBchsLegacyRow = (sheet: any, row: BchsComparisonRow) => {
  const targetRow = row.rowNumber || 11;
  const shortageRatio =
    row.staff > 0 ? row.shortage / row.staff : 0;
  const actualRatio =
    row.staff > 0 ? row.inRanksActually / row.staff : 0;
  const values: Array<[string, unknown]> = [
    ["B", row.unit],
    ["C", row.staffOfficers],
    ["D", row.staffSergeants],
    ["E", row.staffSoldiers],
    ["F", row.staff],
    ["G", row.listedOfficers],
    ["H", row.listedSergeants],
    ["I", row.listedSoldiers],
    ["J", row.listed],
    ["K", toBchsExportPercentRatio(row.staffedPercent)],
    ["L", row.availableOfficers],
    ["M", row.availableSergeants],
    ["N", row.availableSoldiers],
    ["O", row.available],
    ["T", row.shortage],
    [
      "U",
      toBchsExportPercentRatio(row.shortagePercent) ??
        toBchsExportPercentRatio(shortageRatio),
    ],
    ["Y", row.absent],
    ["Z", row.businessTrip],
    ["AA", row.training],
    ["AB", row.hospitalWounded],
    ["AC", row.hospitalIllness],
    ["AD", row.vacation],
    ["AE", row.awol],
    ["AF", row.missing],
    ["AG", row.killed],
    ["AH", row.medWounded],
    ["AI", row.medIllness],
    ["AK", row.awayOfficers],
    ["AL", row.awaySergeants],
    ["AM", row.awaySoldiers],
    ["AN", row.awayInOtherUnits],
    ["AO", row.awayDestinationsText || null],
    ["AP", row.attachedOfficers],
    ["AQ", row.attachedSergeants],
    ["AR", row.attachedSoldiers],
    ["AS", row.attachedFromOtherUnits],
    ["AT", row.attachedSourcesText || null],
    ["AU", row.unassignedNewcomers],
    ["AZ", row.inRanksActually],
    [
      "BA",
      toBchsExportPercentRatio(row.actualPercent) ??
        toBchsExportPercentRatio(actualRatio),
    ],
    ["BB", row.assaultReady],
    ["BC", row.assaultRecovery],
    ["BD", row.assaultExecution],
    ["BE", row.noBzvp],
    ["BF", row.assaultTotal],
    ["BG", row.vehicleCrew],
    ["BH", row.droneCrew],
    ["BI", row.crewServedWeapons],
    ["BJ", row.commandCombat],
    ["BK", row.supportCombat],
    ["BL", row.combatComponent],
  ];

  values.forEach(([letter, value]) =>
    setWorkbookCellValue(
      sheet,
      targetRow,
      letter,
      value,
      BCHS_PERCENT_COLUMNS.has(letter)
        ? { numberFormat: BCHS_PERCENT_NUMBER_FORMAT }
        : {},
    ),
  );

  // Force whole-percent format even if the template had 0.0%.
  ["K", "U", "BA"].forEach((letter) => {
    const cell = sheet.cell(
      targetRow,
      columnLetterToIndex(letter) + 1,
    );
    if (cell && typeof cell.style === "function") {
      cell.style("numberFormat", BCHS_PERCENT_NUMBER_FORMAT);
    }
  });
};

export const writeBchsAppendixRow = (
  sheet: any,
  rowNumber: number,
  row: BchsComparisonRow,
) => {
  const values: Array<[string, unknown]> = [
    ["C", row.staff],
    ["D", row.listed],
    ["E", row.absent],
    ["F", row.businessTrip],
    ["G", row.training],
    ["H", row.hospitalWounded],
    ["I", row.hospitalIllness],
    ["J", row.vacation],
    ["K", row.awol],
    ["L", row.missing],
    ["M", row.killed],
    ["N", row.medWounded],
    ["O", row.medIllness],
    ["P", row.awayInOtherUnits],
    ["Q", row.attachedFromOtherUnits],
    ["R", row.unassignedNewcomers],
    ["S", row.inRanksActually],
    ["T", row.assaultReady],
    ["U", row.assaultRecovery],
    ["V", row.assaultExecution],
    ["W", row.noBzvp],
    ["X", row.assaultTotal],
    ["Y", row.vehicleCrew],
    ["Z", row.droneCrew],
    ["AA", row.crewServedWeapons],
    ["AB", row.commandCombat],
    ["AC", row.supportCombat],
  ];

  values.forEach(([letter, value]) => setWorkbookCellValue(sheet, rowNumber, letter, value));
};

export const writeBchsPersonnelBzvpRow = (sheet: any, row: BchsSupplementRow) => {
  const values: Array<[string, unknown]> = [
    ["C", row.staff],
    ["D", row.listed],
    ["E", row.available || null],
    ["F", toBchsExportPercentRatio(row.staffedPercent)],
    ["G", row.combatTask],
    ["H", row.replacementReserve],
    ["I", row.taskReserve],
    ["J", row.commanderReserve],
  ];

  values.forEach(([letter, value]) =>
    setWorkbookCellValue(
      sheet,
      row.rowNumber,
      letter,
      value,
      letter === "F" ? { numberFormat: BCHS_PERCENT_NUMBER_FORMAT } : {},
    ),
  );
  row.bzvpBuckets.forEach((bucket, index) => {
    setWorkbookCellValue(
      sheet,
      row.rowNumber,
      getColumnLabel(columnLetterToIndex("K") + index),
      bucket.value,
    );
  });
};

const BCHS_PERSONNEL_1PB_TEMPLATE_ROWS: Array<[number, string]> = [
  [81, "Управління"],
  [82, "Штаб"],
  [83, "1 піхотна рота"],
  [84, "2 піхотна рота"],
  [85, "3 піхотна рота"],
  [86, "Взвод зв'язку"],
  [87, "Взвод логістично-евакуаційних безпілотних наземних систем"],
  [88, "Взвод матеріально-технічного забезпечення"],
  [89, "Взвод перехоплювачів безпілотних літальних апаратів"],
  [90, "Взвод протитанкових ракетних комплексів"],
  [91, "Відділення радіоелектронної боротьби"],
  [92, "Гранатометний взвод"],
  [93, "Група безпілотних систем"],
  [94, "Зенітно-ракетний взвод"],
  [95, "Інженерно-саперне відділення"],
  [96, "Кулеметний взвод"],
  [97, "Мінометний взвод"],
  [98, "Розвідувальний взвод"],
  [99, "Рота безпілотних авіаційних комплексів"],
  [100, "Медичний пункт"],
  [101, "Всього:"],
];

const normalizeBchsUnitKey = (value: string) =>
  normalizeBchsText(value)
    .replace(/[’ʼ`]/g, "'")
    .replace(/\s*-\s*/g, "-");

const findBchsComparisonRowByUnit = (
  rows: BchsComparisonRow[],
  unitName: string,
) => {
  const target = normalizeBchsUnitKey(unitName);
  const aliases =
    target === "штаб" || target === "командування"
      ? ["ж", "упр", "управління", "штаб", "командування"]
      : target === "відділення радіоелектронної боротьби"
        ? ["відділення радіоелектронної боротьби", "відділення реб", "брез"]
        : target === "взвод матеріально-технічного забезпечення"
          ? ["взвод матеріально-технічного забезпечення", "вмтз"]
          : target === "інженерно-саперне відділення"
            ? ["інженерно-саперне відділення", "ісв"]
            : [target];

  return (
    aliases
      .map((alias) =>
        rows.find((row) => normalizeBchsUnitKey(row.unit) === alias),
      )
      .find(Boolean) ??
    rows.find((row) => normalizeBchsUnitKey(row.unit) === target) ??
    rows.find((row) => {
      const unit = normalizeBchsUnitKey(row.unit);
      return aliases.some((alias) => unit.includes(alias) || alias.includes(unit));
    })
  );
};

const BCHS_SUMMABLE_COMPARISON_FIELDS: Array<
  keyof Pick<
    BchsComparisonRow,
    | "staff"
    | "staffOfficers"
    | "staffSergeants"
    | "staffSoldiers"
    | "listed"
    | "listedOfficers"
    | "listedSergeants"
    | "listedSoldiers"
    | "available"
    | "availableOfficers"
    | "availableSergeants"
    | "availableSoldiers"
    | "shortage"
    | "absent"
    | "businessTrip"
    | "training"
    | "hospitalWounded"
    | "hospitalIllness"
    | "vacation"
    | "awol"
    | "missing"
    | "killed"
    | "medWounded"
    | "medIllness"
    | "inRanksActually"
    | "combatComponent"
    | "actualOfficers"
    | "actualSergeants"
    | "actualSoldiers"
    | "awayInOtherUnits"
    | "awayOfficers"
    | "awaySergeants"
    | "awaySoldiers"
    | "attachedFromOtherUnits"
    | "attachedOfficers"
    | "attachedSergeants"
    | "attachedSoldiers"
    | "unassignedNewcomers"
    | "noBzvp"
    | "assaultReady"
    | "assaultRecovery"
    | "assaultExecution"
    | "assaultTotal"
    | "vehicleCrew"
    | "droneCrew"
    | "crewServedWeapons"
    | "commandCombat"
    | "supportCombat"
  >
> = [
  "staff", "staffOfficers", "staffSergeants", "staffSoldiers",
  "listed", "listedOfficers", "listedSergeants", "listedSoldiers",
  "available", "availableOfficers", "availableSergeants", "availableSoldiers",
  "shortage", "absent", "businessTrip", "training", "hospitalWounded",
  "hospitalIllness", "vacation", "awol", "missing", "killed", "medWounded",
  "medIllness", "inRanksActually", "combatComponent", "actualOfficers",
  "actualSergeants", "actualSoldiers", "awayInOtherUnits", "awayOfficers",
  "awaySergeants", "awaySoldiers", "attachedFromOtherUnits",
  "attachedOfficers", "attachedSergeants", "attachedSoldiers",
  "unassignedNewcomers", "noBzvp", "assaultReady", "assaultRecovery",
  "assaultExecution", "assaultTotal", "vehicleCrew", "droneCrew",
  "crewServedWeapons", "commandCombat", "supportCombat",
];

const findBchsComparisonRowsByUnitAliases = (
  rows: BchsComparisonRow[],
  aliases: string[],
) => {
  const normalizedAliases = aliases.map(normalizeBchsUnitKey);
  return rows.filter((row) => {
    const unit = normalizeBchsUnitKey(row.unit);
    return normalizedAliases.some((alias) => unit === alias);
  });
};

const combineBchsComparisonRows = (
  unit: string,
  rows: BchsComparisonRow[],
) => {
  const combined = createBchsComparisonRow({ unit });
  BCHS_SUMMABLE_COMPARISON_FIELDS.forEach((field) => {
    combined[field] = rows.reduce((sum, row) => sum + bchsToNumber(row[field]), 0);
  });
  combined.staffedPercent = combined.staff ? combined.listed / combined.staff : 0;
  combined.shortagePercent = combined.staff ? combined.shortage / combined.staff : 0;
  combined.actualPercent = combined.staff
    ? combined.inRanksActually / combined.staff
    : 0;
  combined.levelPercent = combined.staff
    ? combined.inRanksActually / combined.staff
    : 0;
  combined.balanceActual =
    combined.listed -
    combined.absent -
    combined.awayInOtherUnits +
    combined.attachedFromOtherUnits;
  combined.awayDestinationsText = rows
    .map((row) => row.awayDestinationsText)
    .filter(Boolean)
    .join("\n");
  combined.attachedSourcesText = rows
    .map((row) => row.attachedSourcesText)
    .filter(Boolean)
    .join("\n");
  return combined;
};

export const buildBchsMainExportRows = (
  analytics: BchsAnalyticsSnapshot,
): BchsComparisonRow[] => {
  const total =
    analytics.comparisonRows.find((row) => row.rowNumber === 11) ??
    analytics.total;
  const used = new Set<BchsComparisonRow>();
  const detailRows = BCHS_MAIN_EXPORT_UNIT_ORDER.flatMap((unitName) => {
    const matchedRows =
      unitName === "Командування"
        ? findBchsComparisonRowsByUnitAliases(analytics.comparisonRows, [
            "ж",
            "упр",
            "управління",
            "штаб",
            "командування",
          ])
        : [];
    const row =
      matchedRows.length > 1
        ? combineBchsComparisonRows(unitName, matchedRows)
        : matchedRows[0] ??
          findBchsComparisonRowByUnit(analytics.comparisonRows, unitName);
    if (!row || row === total || used.has(row)) return [];
    matchedRows.forEach((matchedRow) => used.add(matchedRow));
    used.add(row);
    return [{ ...row, unit: unitName }];
  });
  const overflowRows = analytics.comparisonRows.filter(
    (row) =>
      row !== total &&
      row.rowNumber !== 11 &&
      row.unit.trim() &&
      !used.has(row),
  );

  return [total, ...detailRows, ...overflowRows].map((row, index) => ({
    ...row,
    rowNumber: 11 + index,
  }));
};

export const writeBchsPersonnelBzvp1PbTemplate = (
  sheet: any,
  analytics: BchsAnalyticsSnapshot,
) => {
  const total =
    analytics.comparisonRows.find((row) => row.rowNumber === 11) ??
    analytics.total;

  ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"].forEach(
    (letter) => setGeneratedColumnWidth(sheet, letter, letter === "F" ? 9 : 10),
  );

  const taskTotals = {
    combatTask: 0,
    replacementReserve: 0,
    taskReserve: 0,
    commanderReserve: 0,
  };

  BCHS_PERSONNEL_1PB_TEMPLATE_ROWS.forEach(([rowNumber, unitName]) => {
    const isTotal = /всього/i.test(unitName);
    const row = isTotal
      ? total
      : findBchsComparisonRowByUnit(analytics.comparisonRows, unitName);

    if (rowNumber === 81) setWorkbookCellValue(sheet, rowNumber, "A", "1 ПБ");
    setWorkbookCellValue(sheet, rowNumber, "B", isTotal ? null : unitName);
    setWorkbookCellValue(sheet, rowNumber, "C", row?.staff ?? null);
    setWorkbookCellValue(sheet, rowNumber, "D", row?.listed ?? null);
    setWorkbookCellValue(sheet, rowNumber, "E", row?.available ?? null);
    setWorkbookCellValue(
      sheet,
      rowNumber,
      "F",
      row
        ? toBchsExportPercentRatio(row.staff > 0 ? row.listed / row.staff : 0)
        : null,
      { numberFormat: BCHS_PERCENT_NUMBER_FORMAT },
    );

    ["G", "H", "I", "J", "K", "L", "M", "N"].forEach((letter) =>
      setWorkbookCellValue(sheet, rowNumber, letter, null),
    );

    if (isTotal) {
      setWorkbookCellValue(sheet, rowNumber, "A", "Всього:");
      setWorkbookCellValue(sheet, rowNumber, "G", taskTotals.combatTask);
      setWorkbookCellValue(
        sheet,
        rowNumber,
        "H",
        taskTotals.replacementReserve,
      );
      setWorkbookCellValue(sheet, rowNumber, "I", taskTotals.taskReserve);
      setWorkbookCellValue(sheet, rowNumber, "J", taskTotals.commanderReserve);
      return;
    }
    if (!row) return;

    const combatTask = row.assaultExecution;
    const replacementReserve = 0;
    const taskReserve = row.noBzvp;
    const commanderReserve = row.attachedFromOtherUnits;

    taskTotals.combatTask += combatTask;
    taskTotals.replacementReserve += replacementReserve;
    taskTotals.taskReserve += taskReserve;
    taskTotals.commanderReserve += commanderReserve;

    setWorkbookCellValue(sheet, rowNumber, "G", combatTask);
    setWorkbookCellValue(sheet, rowNumber, "H", replacementReserve);
    setWorkbookCellValue(sheet, rowNumber, "I", taskReserve);
    setWorkbookCellValue(sheet, rowNumber, "J", commanderReserve);
  });
};

export const applyExcelStyles = (
  target: { style: (name: string, value: unknown) => unknown },
  styles: Record<string, unknown>,
) => {
  Object.entries(styles).forEach(([name, value]) => {
    if (value !== undefined) target.style(name, value);
  });
};

export const mergeExcelRange = (
  sheet: GeneratedSheet,
  startCell: string,
  endCell: string,
) => {
  const range = sheet.range(startCell, endCell);

  if (typeof range.merged === "function") range.merged(true);
};

export const setGeneratedCell = (
  sheet: GeneratedSheet,
  rowNumber: number,
  columnLetter: string,
  value: unknown,
  styles: Record<string, unknown> = {},
) => {
  const normalized =
    typeof value === "number" && value === 0 ? null : (value ?? null);
  const cell = sheet
    .cell(rowNumber, columnLetterToIndex(columnLetter) + 1)
    .value(normalized);
  applyExcelStyles(cell, styles);

  return cell;
};

export const styleGeneratedArea = (
  sheet: GeneratedSheet,
  startCell: string,
  endCell: string,
) => {
  applyExcelStyles(sheet.range(startCell, endCell), {
    fontFamily: "Times New Roman",
    fontSize: 12,
    border: true,
    verticalAlignment: "center",
  });
};

export const setGeneratedColumnWidth = (
  sheet: GeneratedSheet,
  columnLetter: string,
  width: number,
) => {
  if (typeof sheet.column === "function") sheet.column(columnLetter).width(width);
};

export const estimateExcelTextColumnWidth = (
  text: unknown,
  minWidth = 12,
  maxWidth = 72,
) => {
  const value = valueToDisplay(text as CellValue).trim();
  if (!value) return minWidth;

  const longestLine = value
    .split(/\r?\n/)
    .reduce((max, line) => Math.max(max, line.length), 0);

  // Cyrillic in Times New Roman needs a bit more than Latin per character.
  return Math.min(maxWidth, Math.max(minWidth, Math.ceil(longestLine * 1.2) + 2));
};

export const fitSheetColumnToText = (
  sheet: { column?: (letter: string) => { width: (value: number) => unknown }; cell: (row: number, column: number) => { value?: () => unknown } },
  columnLetter: string,
  rowNumbers: number[],
  options?: {
    minWidth?: number;
    maxWidth?: number;
    headerTexts?: string[];
  },
) => {
  const minWidth = options?.minWidth ?? 12;
  const maxWidth = options?.maxWidth ?? 72;
  let width = minWidth;

  for (const header of options?.headerTexts ?? []) {
    width = Math.max(
      width,
      estimateExcelTextColumnWidth(header, minWidth, maxWidth),
    );
  }

  const columnIndex = columnLetterToIndex(columnLetter) + 1;
  for (const rowNumber of rowNumbers) {
    const cell = sheet.cell(rowNumber, columnIndex);
    const value = typeof cell.value === "function" ? cell.value() : null;
    width = Math.max(
      width,
      estimateExcelTextColumnWidth(value, minWidth, maxWidth),
    );
  }

  if (typeof sheet.column === "function") {
    sheet.column(columnLetter).width(width);
  }
};

/** Auto-size text-heavy BCHS columns (unit, location, destinations, sources). */
export const fitBchsTextExportColumns = (
  sheet: {
    column?: (letter: string) => { width: (value: number) => unknown };
    cell: (row: number, column: number) => { value?: () => unknown };
  },
  rowNumbers: number[],
) => {
  if (rowNumbers.length === 0) return;

  fitSheetColumnToText(sheet, "B", rowNumbers, {
    minWidth: 18,
    maxWidth: 42,
    headerTexts: ["Батальйон (підрозділ)"],
  });
  fitSheetColumnToText(sheet, "P", rowNumbers, {
    minWidth: 14,
    maxWidth: 40,
    headerTexts: ["Місцезнаходження"],
  });
  fitSheetColumnToText(sheet, "AO", rowNumbers, {
    minWidth: 22,
    maxWidth: 80,
    headerTexts: ["Куди відкомандировані"],
  });
  fitSheetColumnToText(sheet, "AT", rowNumbers, {
    minWidth: 22,
    maxWidth: 80,
    headerTexts: ["Звідки прикомандировані"],
  });
};

export const setGeneratedRowHeight = (
  sheet: GeneratedSheet,
  rowNumber: number,
  height: number,
) => {
  if (typeof sheet.row === "function") sheet.row(rowNumber).height(height);
};

export const replaceBchsDateInText = (value: unknown, reportDate: string) => {
  const text = valueToDisplay(value as CellValue);
  if (!text.trim()) return null;

  const withDate = text.replace(/\b\d{2}\.\d{2}\.\d{4}\b/g, reportDate);
  if (withDate !== text) return withDate;
  if (/бчс/i.test(text) && !text.includes(reportDate)) {
    return `${text.replace(/\s+$/, "")} ${reportDate}`;
  }

  return null;
};

export const updateBchsSheetHeaderDate = (
  sheet: GeneratedSheet,
  reportDate: string,
) => {
  for (let rowNumber = 1; rowNumber <= 3; rowNumber += 1) {
    for (let columnNumber = 1; columnNumber <= 30; columnNumber += 1) {
      const cell = sheet.cell(rowNumber, columnNumber);
      const current = typeof cell.value === "function" ? cell.value() : null;
      const next = replaceBchsDateInText(current, reportDate);
      if (next) cell.value(next);
    }
  }
};

export const writeGeneratedBchsWorkbook = (
  workbook: GeneratedWorkbook,
  analytics: BchsAnalyticsSnapshot,
  personnelSupplementOverride?: BchsSupplementSnapshot,
  reportDateOverride?: string,
  sheetName = "Аркуш1",
) => {
  const reportDate = reportDateOverride ?? analytics.reportDate;
  const reportSheet = workbook.sheet(0);
  reportSheet.name(sheetName);
  setGeneratedCell(
    reportSheet,
    1,
    "B",
    `БЧС 1 піхотного батальйону ${reportDate || ""}`.trim(),
    { bold: true, fontSize: 16, horizontalAlignment: "center" },
  );
  mergeExcelRange(reportSheet, "B1", "BL1");
  setGeneratedColumnWidth(reportSheet, "B", 28);
  [
    "C",
    "D",
    "E",
    "G",
    "H",
    "I",
    "L",
    "M",
    "N",
    "Q",
    "R",
    "S",
    "V",
    "W",
    "X",
    "AK",
    "AL",
    "AM",
    "AP",
    "AQ",
    "AR",
  ].forEach((letter) => setGeneratedColumnWidth(reportSheet, letter, 5));
  [
    "F",
    "J",
    "K",
    "O",
    "T",
    "U",
    "Y",
    "Z",
    "AA",
    "AB",
    "AC",
    "AD",
    "AE",
    "AF",
    "AG",
    "AH",
    "AI",
    "AN",
    "AS",
    "AU",
    "AZ",
    "BA",
    "BB",
    "BC",
    "BD",
    "BE",
    "BF",
    "BG",
    "BH",
    "BI",
    "BJ",
    "BK",
    "BL",
  ].forEach((letter) => setGeneratedColumnWidth(reportSheet, letter, 7));
  [4, 5, 6, 7].forEach((rowNumber) =>
    setGeneratedRowHeight(reportSheet, rowNumber, rowNumber === 4 ? 44 : 74),
  );

  const groupHeaders: Array<[string, string, string, string]> = [
    ["B", "B", "Батальйон (підрозділ)", "ddebf7"],
    ["C", "F", "За штатом", "b7dee8"],
    ["G", "J", "За списком", "d8f0d0"],
    ["K", "K", "% укомплектованості", "ffc000"],
    ["L", "O", "Наявність", "8ed973"],
    ["P", "P", "Місцезнаходження", "ffffff"],
    ["Q", "T", "Некомплект", "fce4d6"],
    ["U", "U", "% не укомплектованості", "44b3d5"],
    ["V", "Y", "Відсутні", "b7e1a1"],
    ["Z", "Z", "Відрядження", "d9eaf7"],
    ["AA", "AA", "Навчання", "d9eaf7"],
    ["AB", "AC", "Шпиталь", "ead0e9"],
    ["AD", "AD", "Відпустка", "ead0e9"],
    ["AE", "AE", "СЗЧ", "ead0e9"],
    ["AF", "AF", "Зниклі безвісти", "ead0e9"],
    ["AG", "AG", "Загиблі", "ead0e9"],
    ["AH", "AI", "Мед.рота", "ead0e9"],
    ["AK", "AN", "Виконують завдання в інших підрозділах полку", "ffffff"],
    ["AO", "AO", "Куди відкомандировані", "ffffff"],
    ["AP", "AS", "В розташуванні з інших підрозділів полку", "ffffff"],
    ["AT", "AT", "Звідки прикомандировані", "ffffff"],
    ["AU", "AU", "Новоприбулі (без посади)", "b7e1a1"],
    ["AZ", "AZ", "В строю фактично", "b7e1a1"],
    ["BA", "BA", "% фактичної укомплектованості", "ffc000"],
    ["BB", "BE", "Піхота", "d9eaf7"],
    ["BF", "BL", "Бойова складова", "8ed973"],
  ];

  groupHeaders.forEach(([start, end, label, fill]) => {
    setGeneratedCell(reportSheet, 4, start, label, {
      bold: true,
      fill,
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
    });
    if (start !== end) mergeExcelRange(reportSheet, `${start}4`, `${end}4`);
  });

  const subHeaders: Record<string, string> = {
    C: "Офіцери",
    D: "Сержанти",
    E: "Солдати",
    F: "Всього",
    G: "Офіцери",
    H: "Сержанти",
    I: "Солдати",
    J: "Всього",
    L: "Офіцери",
    M: "Сержанти",
    N: "Солдати",
    O: "Всього",
    Q: "Офіцери",
    R: "Сержанти",
    S: "Солдати",
    T: "Всього",
    V: "Офіцери",
    W: "Сержанти",
    X: "Солдати",
    Y: "Всього",
    AB: "На лікуванні по пораненню",
    AC: "На лікуванні по хворобі",
    AH: "На лікуванні по пораненню",
    AI: "На лікуванні по хворобі",
    AK: "Офіцери",
    AL: "Сержанти",
    AM: "Солдати",
    AN: "Всього",
    AP: "Офіцери",
    AQ: "Сержанти",
    AR: "Солдати",
    AS: "Всього",
    BB: "Готові",
    BC: "На відновленні",
    BD: "На виконанні",
    BE: "Без БЗВП",
    BF: "Штурмовики",
    BG: "Екіпажі техніки",
    BH: "Екіпажі БПЛА",
    BI: "Колективне озброєння",
    BJ: "Управління",
    BK: "Забезпечення",
    BL: "Всього",
  };
  Object.entries(subHeaders).forEach(([letter, label]) =>
    setGeneratedCell(reportSheet, 5, letter, label, {
      bold: true,
      horizontalAlignment: "center",
      textRotation: 90,
      verticalAlignment: "center",
      wrapText: true,
    }),
  );

  const mainExportRows = buildBchsMainExportRows(analytics);

  mainExportRows.forEach((row, index) => {
    const rowNumber = 11 + index;
    const tableRow = analytics.table?.rows.find(
      (item) =>
        normalizeBchsUnitKey(valueToDisplay(item.values.B as CellValue)) ===
        normalizeBchsUnitKey(row.unit),
    );
    const values: Record<string, unknown> = {
      B: row.unit,
      C: tableRow ? resolveBchsExportCellValue("C", tableRow.values.C) : row.staffOfficers,
      D: tableRow ? resolveBchsExportCellValue("D", tableRow.values.D) : row.staffSergeants,
      E: tableRow ? resolveBchsExportCellValue("E", tableRow.values.E) : row.staffSoldiers,
      F: row.staff,
      G: tableRow ? resolveBchsExportCellValue("G", tableRow.values.G) : row.listedOfficers,
      H: tableRow ? resolveBchsExportCellValue("H", tableRow.values.H) : row.listedSergeants,
      I: tableRow ? resolveBchsExportCellValue("I", tableRow.values.I) : row.listedSoldiers,
      J: row.listed,
      K: toBchsExportPercentRatio(row.staffedPercent),
      L: tableRow ? resolveBchsExportCellValue("L", tableRow.values.L) : row.availableOfficers,
      M: tableRow ? resolveBchsExportCellValue("M", tableRow.values.M) : row.availableSergeants,
      N: tableRow ? resolveBchsExportCellValue("N", tableRow.values.N) : row.availableSoldiers,
      O: row.available,
      T: row.shortage,
      U:
        toBchsExportPercentRatio(row.shortagePercent) ??
        toBchsExportPercentRatio(
          row.staff > 0 ? row.shortage / row.staff : 0,
        ),
      V: tableRow ? resolveBchsExportCellValue("V", tableRow.values.V) : null,
      W: tableRow ? resolveBchsExportCellValue("W", tableRow.values.W) : null,
      X: tableRow ? resolveBchsExportCellValue("X", tableRow.values.X) : null,
      Y: row.absent,
      Z: row.businessTrip,
      AA: row.training,
      AB: row.hospitalWounded,
      AC: row.hospitalIllness,
      AD: row.vacation,
      AE: row.awol,
      AF: row.missing,
      AG: row.killed,
      AH: row.medWounded,
      AI: row.medIllness,
      AK: row.awayOfficers,
      AL: row.awaySergeants,
      AM: row.awaySoldiers,
      AN: row.awayInOtherUnits,
      AO: row.awayDestinationsText || null,
      AP: row.attachedOfficers,
      AQ: row.attachedSergeants,
      AR: row.attachedSoldiers,
      AS: row.attachedFromOtherUnits,
      AT: row.attachedSourcesText || null,
      AU: row.unassignedNewcomers,
      AW: row.actualOfficers,
      AX: row.actualSergeants,
      AY: row.actualSoldiers,
      AZ: row.inRanksActually,
      BA:
        toBchsExportPercentRatio(row.actualPercent) ??
        toBchsExportPercentRatio(
          row.staff > 0 ? row.inRanksActually / row.staff : 0,
        ),
      BB: row.assaultReady,
      BC: row.assaultRecovery,
      BD: row.assaultExecution,
      BE: row.noBzvp,
      BF: row.assaultTotal,
      BG: row.vehicleCrew,
      BH: row.droneCrew,
      BI: row.crewServedWeapons,
      BJ: row.commandCombat,
      BK: row.supportCombat,
      BL: row.combatComponent,
    };

    Object.entries(values).forEach(([letter, value]) =>
      setGeneratedCell(reportSheet, rowNumber, letter, value, {
        bold: index === 0,
        horizontalAlignment:
          letter === "B" || letter === "AO" || letter === "AT"
            ? "left"
            : "center",
        wrapText: letter === "B" || letter === "AO" || letter === "AT",
      }),
    );
    ["K", "U", "BA"].forEach((letter) =>
      setGeneratedCell(reportSheet, rowNumber, letter, values[letter], {
        numberFormat: BCHS_PERCENT_NUMBER_FORMAT,
        bold: true,
        horizontalAlignment: "center",
      }),
    );
  });
  styleGeneratedArea(reportSheet, "B4", `BL${10 + mainExportRows.length}`);
  styleBchsMainDataRows(
    reportSheet,
    mainExportRows.map((_, index) => 11 + index),
  );
  fitBchsTextExportColumns(
    reportSheet,
    mainExportRows.map((_, index) => 11 + index),
  );

  const supplementSheet = workbook.addSheet("Особовий склад + БЗВП");
  supplementSheet.name?.("Особовий склад + БЗВП");
  const personnelSupplement =
    analytics.supplement?.kind === "personnel-bzvp"
      ? analytics.supplement
      : personnelSupplementOverride;
  const bzvpHeaders =
    personnelSupplement?.bzvpBuckets.map((item) => item.label) ?? [
      "01.07 - 07.07",
      "08.07 - 14.07",
      "16.07 - 22.07",
      "22.07 - 31.07",
    ];
  const bzvpEndLetter = getColumnLabel(9 + bzvpHeaders.length);

  setGeneratedCell(supplementSheet, 1, "A", "БАТАЛЬЙОН", {
    bold: true,
    fill: "b4c6e7",
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  });
  setGeneratedCell(supplementSheet, 1, "B", "ПІДРОЗДІЛ", {
    bold: true,
    fill: "b4c6e7",
    horizontalAlignment: "center",
    verticalAlignment: "center",
  });
  [
    ["C", "Штатна кількість О/С"],
    ["D", "Кількість за списком"],
    ["E", "Наявність"],
    ["F", "О/С %"],
  ].forEach(([letter, label]) =>
    setGeneratedCell(supplementSheet, 1, letter, label, {
      bold: true,
      fill: letter === "F" ? "d9eaf7" : "ffffff",
      horizontalAlignment: "center",
      verticalAlignment: "center",
      textRotation: 90,
      wrapText: true,
    }),
  );
  setGeneratedCell(
    supplementSheet,
    1,
    "G",
    "О/с на виконанні бойового завдання",
    {
      bold: true,
      fill: "d9eaf7",
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
    },
  );
  mergeExcelRange(supplementSheet, "G1", "J1");
  setGeneratedCell(supplementSheet, 1, "K", "БЗВП", {
    bold: true,
    fill: "b4c6e7",
    horizontalAlignment: "center",
    verticalAlignment: "center",
  });
  mergeExcelRange(supplementSheet, "K1", `${bzvpEndLetter}1`);
  [
    ["G", "На виконанні б/з"],
    ["H", "Резерв згідно графіку заміни"],
    ["I", "Резерв о/с виконання бойових завдань"],
    ["J", "Резерв о/с командира полку"],
  ].forEach(([letter, label]) =>
    setGeneratedCell(supplementSheet, 3, letter, label, {
      bold: true,
      fill: "d9eaf7",
      horizontalAlignment: "center",
      verticalAlignment: "center",
      textRotation: 90,
      wrapText: true,
    }),
  );
  bzvpHeaders.forEach((label, index) =>
    setGeneratedCell(supplementSheet, 3, getColumnLabel(10 + index), label, {
      bold: true,
      fill: "b4c6e7",
      horizontalAlignment: "center",
      verticalAlignment: "center",
      textRotation: 90,
      wrapText: true,
    }),
  );
  ["A", "B", "C", "D", "E", "F"].forEach((letter) =>
    mergeExcelRange(supplementSheet, `${letter}1`, `${letter}3`),
  );
  setGeneratedColumnWidth(supplementSheet, "A", 18);
  setGeneratedColumnWidth(supplementSheet, "B", 34);
  ["C", "D", "E", "F", "G", "H", "I", "J"].forEach((letter) =>
    setGeneratedColumnWidth(supplementSheet, letter, 9),
  );
  bzvpHeaders.forEach((_, index) =>
    setGeneratedColumnWidth(supplementSheet, getColumnLabel(10 + index), 8),
  );
  [1, 2, 3].forEach((rowNumber) =>
    setGeneratedRowHeight(supplementSheet, rowNumber, rowNumber === 1 ? 44 : 58),
  );
  const fallbackSupplementRows = [
    ...analytics.comparisonRows
      .filter((row) => row.rowNumber !== analytics.total.rowNumber)
      .map((row, index) =>
        emptyBchsSupplementRow({
          rowNumber: row.rowNumber,
          battalion: index === 0 ? "1 ПБ" : "",
          unit: row.unit,
          staff: row.staff,
          listed: row.listed,
          available: row.available,
          staffedPercent: bchsToNumber(row.staffedPercent),
          absent: row.absent,
          inRanks: row.inRanksActually,
          detached: row.awayInOtherUnits,
          attached: row.attachedFromOtherUnits,
          newcomers: row.unassignedNewcomers,
          noBzvp: row.noBzvp,
        }),
      ),
    emptyBchsSupplementRow({
      battalion: "Всього:",
      unit: "",
      staff: analytics.total.staff,
      listed: analytics.total.listed,
      available: analytics.total.available,
      staffedPercent: bchsToNumber(analytics.total.staffedPercent),
      combatTask: analytics.total.assaultExecution,
      taskReserve: analytics.total.noBzvp,
      commanderReserve: analytics.total.attachedFromOtherUnits,
    }),
  ];
  const supplementRows = personnelSupplement?.rows ?? fallbackSupplementRows;
  let previousSupplementBattalion = "";

  supplementRows.forEach((row, index) => {
    const rowNumber = index + 4;
    const isTotalRow = /всього/i.test(row.unit) || /всього/i.test(row.battalion);
    const visibleBattalion =
      isTotalRow || row.battalion !== previousSupplementBattalion
        ? row.battalion
        : "";

    if (row.battalion && !isTotalRow) previousSupplementBattalion = row.battalion;

    const values: Record<string, unknown> = {
      A: visibleBattalion,
      B: row.unit,
      C: row.staff,
      D: row.listed,
      E: row.available || null,
      F: row.staffedPercent,
      G: row.combatTask,
      H: row.replacementReserve,
      I: row.taskReserve,
      J: row.commanderReserve,
    };

    Object.entries(values).forEach(([letter, value]) =>
      setGeneratedCell(supplementSheet, rowNumber, letter, value, {
        bold: isTotalRow,
        fill: letter === "A" || letter === "B" ? "b4c6e7" : "ffffff",
        horizontalAlignment: letter === "B" ? "left" : "center",
        wrapText: letter === "B",
      }),
    );
    setGeneratedCell(supplementSheet, rowNumber, "F", toBchsExportPercentRatio(row.staffedPercent), {
      numberFormat: BCHS_PERCENT_NUMBER_FORMAT,
      horizontalAlignment: "center",
      bold: isTotalRow,
    });
    bzvpHeaders.forEach((_, bucketIndex) => {
      const bucketValue = row.bzvpBuckets[bucketIndex]?.value ?? 0;
      setGeneratedCell(
        supplementSheet,
        rowNumber,
        getColumnLabel(10 + bucketIndex),
        bucketValue,
        { horizontalAlignment: "center" },
      );
    });
  });
  styleGeneratedArea(supplementSheet, "A1", `${bzvpEndLetter}${supplementRows.length + 3}`);
  supplementSheet.range("A1", `${bzvpEndLetter}3`).style("bold", true);
  fitSheetColumnToText(
    supplementSheet,
    "A",
    supplementRows.map((_, index) => index + 4),
    { minWidth: 12, maxWidth: 24, headerTexts: ["БАТАЛЬЙОН"] },
  );
  fitSheetColumnToText(
    supplementSheet,
    "B",
    supplementRows.map((_, index) => index + 4),
    { minWidth: 18, maxWidth: 48, headerTexts: ["ПІДРОЗДІЛ"] },
  );

  const appendixSheet = workbook.addSheet("БЧС додаток");
  appendixSheet.name?.("БЧС додаток");
  setGeneratedCell(
    appendixSheet,
    1,
    "B",
    `БЧС полку ${reportDate || ""}`.trim(),
    {
      bold: true,
      fontFamily: "Times New Roman",
      fontSize: 40,
      horizontalAlignment: "center",
      verticalAlignment: "center",
      wrapText: true,
    },
  );
  mergeExcelRange(appendixSheet, "B1", "S2");

  const appendixHeaderStyles = {
    bold: true,
    fontFamily: "Times New Roman",
    fontSize: 30,
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
  };
  const appendixColumnFill = (letter: string, rowNumber?: number) => {
    if (letter === "B") {
      if (rowNumber === 11) return "92d050";
      if (rowNumber === 12) return "b4c6e7";
      return "ddebf7";
    }
    if (letter === "C") return "ddebf7";
    if (letter === "D") return "e2f0d9";
    if (letter === "E") return "c6e0b4";
    if (
      [
        "F",
        "G",
        "H",
        "I",
        "J",
        "K",
        "L",
        "M",
        "N",
        "O",
      ].includes(letter)
    ) {
      return "d9e2f3";
    }
    if (letter === "P" || letter === "S") return "fff2cc";
    if (letter === "Q" || letter === "R") return "c6e0b4";
    return "e2f0d9";
  };
  const appendixSingleHeaders: Array<[string, string, string]> = [
    ["B", "Батальйон (підрозділ)", "ddebf7"],
    ["C", "За штатом", "ddebf7"],
    ["D", "за списком", "e2f0d9"],
    ["E", "Відсутні", "c6e0b4"],
    ["F", "Відрядження", "d9e2f3"],
    ["G", "Навчання", "d9e2f3"],
    ["J", "Відпустка", "d9e2f3"],
    ["K", "СЗЧ", "d9e2f3"],
    ["L", "Зниклі безвісти", "d9e2f3"],
    ["M", "Загиблі", "d9e2f3"],
    ["P", "Відкомандировані", "fff2cc"],
    ["Q", "Прикомандировані", "c6e0b4"],
    ["R", "Новоприбулі (без посади)", "c6e0b4"],
    ["S", "В строю", "a9d18e"],
  ];
  appendixSingleHeaders.forEach(([letter, label, fill]) => {
    setGeneratedCell(appendixSheet, 4, letter, label, {
      ...appendixHeaderStyles,
      fill,
      textRotation: ["B", "S"].includes(letter) ? 0 : 90,
    });
    mergeExcelRange(appendixSheet, `${letter}4`, `${letter}10`);
  });

  [
    ["H", "I", "Шпиталь"],
    ["N", "O", "Мед.рота"],
  ].forEach(([start, end, label]) => {
    setGeneratedCell(appendixSheet, 4, start, label, {
      ...appendixHeaderStyles,
      fill: "d9e2f3",
    });
    mergeExcelRange(appendixSheet, `${start}4`, `${end}4`);
  });
  setGeneratedCell(appendixSheet, 4, "T", "із них:", {
    ...appendixHeaderStyles,
    fill: "a9d18e",
  });
  mergeExcelRange(appendixSheet, "T4", "AC4");

  [
    ["H", "На лікуванні по пораненню"],
    ["I", "На лікуванні по                  хворобі"],
    ["N", "На лікуванні по пораненню"],
    ["O", "На лікуванні по хворобі"],
    ["T", "Штурмовики готові до виконання"],
    ["U", "Штурмовики на відновленні"],
    ["V", "Штурмовики на виконанні"],
    ["W", "Штурмовики без БЗВП"],
    ["X", "Загалом штурмовиків"],
    ["Y", "Єкіпажі техніки"],
    ["Z", "Єкіпажі БПЛА"],
    ["AA", "Розрахунки колективного озброєння"],
    ["AB", "Управління"],
    ["AC", "Забезпечення"],
  ].forEach(([letter, label]) => {
    setGeneratedCell(appendixSheet, 5, letter, label, {
      ...appendixHeaderStyles,
      fill: ["H", "I", "N", "O"].includes(letter) ? "d9e2f3" : "e2f0d9",
      textRotation: 90,
    });
    mergeExcelRange(appendixSheet, `${letter}5`, `${letter}10`);
  });

  setGeneratedColumnWidth(appendixSheet, "B", 73.33);
  [
    "C",
    "D",
    "E",
    "F",
    "G",
    "H",
    "I",
    "J",
    "K",
    "L",
    "M",
    "N",
    "O",
    "P",
    "Q",
    "R",
    "S",
  ].forEach((letter) => setGeneratedColumnWidth(appendixSheet, letter, 34.66));
  setGeneratedColumnWidth(appendixSheet, "T", 25.5);
  ["U", "V", "X", "Y", "Z", "AA", "AB", "AC"].forEach((letter) =>
    setGeneratedColumnWidth(appendixSheet, letter, 17.5),
  );
  setGeneratedColumnWidth(appendixSheet, "W", 22.5);
  [
    [1, 23],
    [2, 51],
    [4, 117.75],
    [5, 39.75],
    [6, 24.75],
    [7, 24.75],
    [8, 24.75],
    [9, 77.25],
    [10, 32],
    [11, 101.25],
    [12, 210],
  ].forEach(([rowNumber, height]) =>
    setGeneratedRowHeight(appendixSheet, rowNumber, height),
  );

  const writeAppendixGeneratedRow = (
    rowNumber: number,
    label: string,
    row: BchsComparisonRow,
    fontSize: number,
    horizontalAlignment: "center" | "right",
  ) => {
    const values: Record<string, unknown> = {
      B: label,
      C: row.staff,
      D: row.listed,
      E: row.absent,
      F: row.businessTrip,
      G: row.training,
      H: row.hospitalWounded,
      I: row.hospitalIllness,
      J: row.vacation,
      K: row.awol,
      L: row.missing,
      M: row.killed,
      N: row.medWounded,
      O: row.medIllness,
      P: row.awayInOtherUnits,
      Q: row.attachedFromOtherUnits,
      R: row.unassignedNewcomers,
      S: row.inRanksActually,
      T: row.assaultReady,
      U: row.assaultRecovery,
      V: row.assaultExecution,
      W: row.noBzvp,
      X: row.assaultTotal,
      Y: row.vehicleCrew,
      Z: row.droneCrew,
      AA: row.crewServedWeapons,
      AB: row.commandCombat,
      AC: row.supportCombat,
    };

    Object.entries(values).forEach(([letter, value]) =>
      setGeneratedCell(appendixSheet, rowNumber, letter, value, {
        bold: true,
        fill: appendixColumnFill(letter, rowNumber),
        fontFamily: "Times New Roman",
        fontSize,
        horizontalAlignment: letter === "B" ? horizontalAlignment : "center",
        verticalAlignment: "center",
        wrapText: true,
      }),
    );
  };

  const appendixTotal =
    analytics.comparisonRows.find((row) => row.rowNumber === 11) ??
    analytics.total;
  writeAppendixGeneratedRow(11, "Загалом:", appendixTotal, 48, "center");
  writeAppendixGeneratedRow(
    12,
    "1 піхотний батальйон",
    appendixTotal,
    36,
    "right",
  );
  styleGeneratedArea(appendixSheet, "B4", "AC12");
  appendixSheet.range("B4", "AC12").style({
    fontFamily: "Times New Roman",
    border: true,
    verticalAlignment: "center",
    wrapText: true,
  });
  appendixSheet.range("B4", "AC10").style({
    bold: true,
    fontSize: 30,
  });
  appendixSheet.range("B11", "AC11").style({
    bold: true,
    fontSize: 48,
  });
  appendixSheet.range("B12", "AC12").style({
    bold: true,
    fontSize: 36,
  });
};
