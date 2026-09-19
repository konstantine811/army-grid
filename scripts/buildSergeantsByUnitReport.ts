/**
 * Збирає «2 а по підрозділам.xlsx» для сержантського складу за місяць.
 *
 * npx jiti scripts/buildSergeantsByUnitReport.ts [month] [year] [outputPath]
 *
 * За замовчуванням: вересень 2026, файли з KINGSTON і запис у шаблон Сержанти.
 */
import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import { readFileSync } from "node:fs";
import XlsxPopulate from "xlsx-populate";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
  type ReadWorkbookOptions,
} from "../src/excelRoundTrip";
import {
  buildSergeantsByUnitReport,
  type SergeantsUnitMetrics,
} from "../src/pages/sergeants/sergeantsByUnitReport";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const DEFAULT_PB =
  "/Volumes/KINGSTON/army_work/ЕЖООС/08.09.2026/1ПБ_07092026.xlsx";
const DEFAULT_EJOOS =
  "/Volumes/KINGSTON/army_work/ЕЖООС/15.09.2026/ЄЖООС_станом_на_07-09-2026.xlsx";
const DEFAULT_STAFF =
  "/Volumes/KINGSTON/army_work/Ранковий звіт/10.09.2026/Штатка.xlsx";
const DEFAULT_OUTPUT =
  "/Volumes/KINGSTON/army_work/Сержанти/2 а по підрозділам.xlsx";

const month = Number(process.argv[2] || 9);
const year = Number(process.argv[3] || 2026);
const outputPath = process.argv[4] || DEFAULT_OUTPUT;

const pbPath = process.env.SERGEANTS_PB || DEFAULT_PB;
const ejoosPath = process.env.SERGEANTS_EJOOS || DEFAULT_EJOOS;
const staffPath = process.env.SERGEANTS_STAFF || DEFAULT_STAFF;

const STAFF_READ_OPTIONS: ReadWorkbookOptions = {
  sheetFilter: (name) => /загальний\s+список/i.test(name.trim()),
  maxColumns: 45,
  skipStyleFills: true,
  preserveLeadingColumns: true,
};

const pb = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);
const ejoos = await readWorkbookSnapshot(
  asFile(ejoosPath),
  EJOOS_SYNC_READ_OPTIONS,
);
const staff = await readWorkbookSnapshot(asFile(staffPath), STAFF_READ_OPTIONS);

const report = buildSergeantsByUnitReport({ pb, ejoos, staff, month, year });

const fillSummarySheet = (
  sheet: XlsxPopulate.Sheet,
  totals: SergeantsUnitMetrics,
  reportMonth: number,
) => {
  // Шаблон має merge A2:B5 / A6:B7; кількості — у колонці D (D2:D9).
  sheet.cell(1, 2).value(reportMonth);
  for (const row of [2, 3, 4, 5, 6, 7, 8, 9]) {
    sheet.cell(row, 2).value(undefined);
  }
  sheet.cell(3, 1).value("2");
  sheet.cell(4, 1).value("3");
  sheet.cell(5, 1).value("4");
  sheet.cell(7, 1).value("2");
  sheet.cell(2, 4).value(totals.combatKilled);
  sheet.cell(3, 4).value(totals.combatSanitary);
  sheet.cell(4, 4).value(totals.combatMissing);
  sheet.cell(5, 4).value(totals.combatPow);
  sheet.cell(6, 4).value(totals.nonCombatIrreversible);
  sheet.cell(7, 4).value(totals.nonCombatSanitary);
  sheet.cell(8, 4).value(totals.onTreatment);
  sheet.cell(9, 4).value(totals.szchMonth);
};

const fillByUnitSheet = (
  workbook: XlsxPopulate.Workbook,
  rows: Array<{ unit: string } & SergeantsUnitMetrics>,
) => {
  const sheetName = "По підрозділах";
  const existing = workbook.sheet(sheetName);
  if (existing) existing.delete();
  const sheet = workbook.addSheet(sheetName);
  const headers = [
    "Підрозділ",
    "Безповоротні",
    "Санітарні",
    "Безвісті",
    "Полон",
    "Небойові безповоротні",
    "Небойові санітарні",
    "На лікуванні",
    "СЗЧ за місяць",
  ];
  headers.forEach((header, index) => {
    sheet.cell(1, index + 1).value(header);
  });
  rows.forEach((row, rowIndex) => {
    const excelRow = rowIndex + 2;
    sheet.cell(excelRow, 1).value(row.unit);
    sheet.cell(excelRow, 2).value(row.combatKilled);
    sheet.cell(excelRow, 3).value(row.combatSanitary);
    sheet.cell(excelRow, 4).value(row.combatMissing);
    sheet.cell(excelRow, 5).value(row.combatPow);
    sheet.cell(excelRow, 6).value(row.nonCombatIrreversible);
    sheet.cell(excelRow, 7).value(row.nonCombatSanitary);
    sheet.cell(excelRow, 8).value(row.onTreatment);
    sheet.cell(excelRow, 9).value(row.szchMonth);
  });
};

mkdirSync(dirname(outputPath), { recursive: true });
const workbook = await XlsxPopulate.fromFileAsync(outputPath);
fillSummarySheet(workbook.sheet(0), report.totals, report.month);
fillByUnitSheet(workbook, report.byUnit);
await workbook.toFileAsync(outputPath);

console.log(`Записано: ${outputPath}`);
console.log(`Місяць: ${report.month}.${report.year}`);
console.log("Підсумок:", report.totals);
console.log(`Підрозділів: ${report.byUnit.length}`);
