import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  extractBchsAwayPeopleFromSheet,
  isBchsPersonnelGeneralListSheet,
  normalizeBchsText,
} from "./bchsCalc";
import type { BchsPersonnelAwayPerson } from "./bchsTypes";
import {
  createBchsLabConsoleApi,
  logBchsLabAwayReport,
  type BchsLabConsoleApi,
} from "./bchsLabAway";

export type BchsLabPersonRow = BchsPersonnelAwayPerson & {
  excelRowNumber: number;
  normalized: {
    battalion: string;
    rosterUnit: string;
    status: string;
    rankCategory: string;
    roleType: string;
    combatReadiness: string;
    medicalPlace: string;
  };
};

export type BchsLabParseResult = {
  fileName: string;
  parsedAt: string;
  sheets: Array<{ sheetIndex: number; sheetName: string; rowCount: number }>;
  rosterSheet: {
    sheetIndex: number;
    sheetName: string;
  } | null;
  calculationSheet: {
    sheetIndex: number;
    sheetName: string;
  } | null;
  people: BchsLabPersonRow[];
  novaPeople: BchsLabPersonRow[];
  summary: {
    totalPeople: number;
    novaCount: number;
    withFullName: number;
    statusCounts: Record<string, number>;
    battalionCounts: Record<string, number>;
    unitCounts: Record<string, number>;
  };
};

const toLabPerson = (
  person: BchsPersonnelAwayPerson,
  excelRowNumber: number,
): BchsLabPersonRow => ({
  ...person,
  excelRowNumber,
  normalized: {
    battalion: normalizeBchsText(person.battalion),
    rosterUnit: normalizeBchsText(person.rosterUnit),
    status: normalizeBchsText(person.status),
    rankCategory: normalizeBchsText(person.rankCategory),
    roleType: normalizeBchsText(person.roleType),
    combatReadiness: normalizeBchsText(person.combatReadiness),
    medicalPlace: normalizeBchsText(person.medicalPlace),
  },
});

const countBy = (items: string[]) =>
  items.reduce<Record<string, number>>((acc, key) => {
    const label = key.trim() || "(порожньо)";
    acc[label] = (acc[label] ?? 0) + 1;
    return acc;
  }, {});

const findRosterSheet = (workbook: ExcelWorkbookSnapshot) =>
  workbook.sheets.find(isBchsPersonnelGeneralListSheet) ??
  workbook.sheets.find((sheet) =>
    /аркуш\s*2|загальн.*спис|1\.?\s*ос/i.test(sheet.sheetName),
  ) ??
  null;

const findCalculationSheet = (workbook: ExcelWorkbookSnapshot) =>
  workbook.sheets.find(
    (sheet) => sheet.sheetName.trim().toLowerCase() === "аркуш1",
  ) ?? null;

export const parseBchsLabWorkbook = (
  workbook: ExcelWorkbookSnapshot,
): BchsLabParseResult => {
  const rosterSheet = findRosterSheet(workbook);
  const calculationSheet = findCalculationSheet(workbook);

  const rawPeople = rosterSheet ? extractBchsAwayPeopleFromSheet(rosterSheet) : [];
  const people = rawPeople.map((person, index) => toLabPerson(person, index + 2));
  // Перший фільтр pipeline — battalion «нова»; Lab API працює лише з novaPeople.
  const novaPeople = people.filter(
    (person) => person.normalized.battalion === "нова",
  );

  return {
    fileName: workbook.fileName,
    parsedAt: new Date().toISOString(),
    sheets: workbook.sheets.map((sheet) => ({
      sheetIndex: sheet.sheetIndex,
      sheetName: sheet.sheetName,
      rowCount: sheet.rows.length,
    })),
    rosterSheet: rosterSheet
      ? {
          sheetIndex: rosterSheet.sheetIndex,
          sheetName: rosterSheet.sheetName,
        }
      : null,
    calculationSheet: calculationSheet
      ? {
          sheetIndex: calculationSheet.sheetIndex,
          sheetName: calculationSheet.sheetName,
        }
      : null,
    people,
    novaPeople,
    summary: {
      totalPeople: people.length,
      novaCount: novaPeople.length,
      withFullName: people.filter((person) => person.fullName.trim()).length,
      statusCounts: countBy(people.map((person) => person.status)),
      battalionCounts: countBy(people.map((person) => person.battalion)),
      unitCounts: countBy(novaPeople.map((person) => person.rosterUnit)),
    },
  };
};

export const logBchsLabParseResult = (result: BchsLabParseResult) => {
  console.groupCollapsed(`[BCHS Lab] ${result.fileName}`);
  console.log("parseResult", result);
  console.log("summary", result.summary);
  console.log("novaPeople", result.novaPeople);
  console.log("statusCounts", result.summary.statusCounts);
  console.log("unitCounts (нова)", result.summary.unitCounts);
  console.groupEnd();
};

declare global {
  interface Window {
    __BCHS_LAB__?: {
      snapshot: ExcelWorkbookSnapshot | null;
      parseResult: BchsLabParseResult | null;
      people: BchsLabPersonRow[];
      novaPeople: BchsLabPersonRow[];
      log: () => void;
      countAway: BchsLabConsoleApi["countAway"];
      awayCommand: BchsLabConsoleApi["awayCommand"];
      logAway: BchsLabConsoleApi["logAway"];
    };
  }
}

export const publishBchsLabToConsole = (
  snapshot: ExcelWorkbookSnapshot | null,
  parseResult: BchsLabParseResult | null,
) => {
  const people = parseResult?.people ?? [];
  const novaPeople = parseResult?.novaPeople ?? [];
  const api = createBchsLabConsoleApi(novaPeople);

  window.__BCHS_LAB__ = {
    snapshot,
    parseResult,
    people,
    novaPeople,
    log: () => {
      if (parseResult) logBchsLabParseResult(parseResult);
      else console.info("[BCHS Lab] Немає parseResult — спочатку імпортуйте файл.");
    },
    countAway: api.countAway,
    awayCommand: api.awayCommand,
    logAway: api.logAway,
  };

  if (parseResult?.novaPeople.length) {
    logBchsLabAwayReport(parseResult.novaPeople, "Командування");
  }

  console.info(
    "[BCHS Lab] window.__BCHS_LAB__ · завжди A=нова · logAway('Командування') · awayCommand()",
  );
};
