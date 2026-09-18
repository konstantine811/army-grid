import XlsxPopulate from "xlsx-populate";
import { buildStaffSheetRosterImportPayload } from "../src/pages/excel-fill/staffSheet";
import {
  buildStaffOverviewRowsFromRoster,
  fillDownRosterUnitRows,
} from "../src/pages/overview/overviewRosterMerge";
import {
  buildOverviewPpdLocationExportSheets,
  filterOverviewPpdVyshneveRows,
  filterOverviewPolygonRows,
  isOverviewPpdVyshneveLocation,
  isOverviewPolygonLocation,
  matchesOverviewRotaUnitFilter,
} from "../src/pages/overview/overviewPpdLocationExport";

const staffPath =
  process.argv[2] ??
  "/Volumes/KINGSTON/army_work/Ранковий звіт/10.09.2026/Штатка.xlsx";

const colLetter = (columnNumber: number) => {
  let column = columnNumber;
  let letters = "";
  while (column) {
    const mod = (column - 1) % 26;
    letters = String.fromCharCode(65 + mod) + letters;
    column = Math.floor((column - 1) / 26);
  }
  return letters;
};

const workbook = await XlsxPopulate.fromFileAsync(staffPath);
const sheet = workbook.sheet(0);
const used = sheet.usedRange();
const endRow = used?.endCell()?.rowNumber?.() ?? 500;
const endCol = used?.endCell()?.columnNumber?.() ?? 45;
const headerRow: string[] = [];
for (let column = 1; column <= endCol; column += 1) {
  headerRow.push(String(sheet.cell(`${colLetter(column)}1`).value() ?? "").trim());
}
console.log("col35 header:", headerRow[34]);

const dataRows: string[][] = [];
for (let rowNumber = 2; rowNumber <= endRow; rowNumber += 1) {
  const row: string[] = [];
  for (let column = 1; column <= endCol; column += 1) {
    row.push(
      String(sheet.cell(`${colLetter(column)}${rowNumber}`).value() ?? "").trim(),
    );
  }
  dataRows.push(row);
}

const payload = buildStaffSheetRosterImportPayload(
  [headerRow, ...dataRows],
  { source: "gviz", sourceLabel: staffPath, includeAllRows: true },
);
const col35 = payload.sheets[0]?.columns.find(
  (column) => column.rosterColumn === 35,
);
console.log("mapped col35:", col35?.label, col35?.key);

const rosterRows = fillDownRosterUnitRows(
  payload.sheets[0]!.rows.map((row, index) => ({
    __dbRowId: `local:${row.excelRowNumber || index + 2}`,
    __rowNumber: row.excelRowNumber,
    ...row.values,
  })),
);
const rosterLabels = Object.fromEntries(
  payload.sheets[0]!.columns.map((column) => [column.key, column.label]),
);
const overviewRows = buildStaffOverviewRowsFromRoster(
  rosterRows,
  rosterLabels,
  payload.sheets[0]!.columns,
);
const unitRows = overviewRows.filter((row) =>
  matchesOverviewRotaUnitFilter(row.unit, "3 рота"),
);
const sample = unitRows.find((row) => row.name.includes("ГНІЦЕВИЧ"));
console.log("unit rows", unitRows.length);
console.log(
  "sample",
  sample?.staffSheetColumns?.staff_31,
  sample?.staffSheetColumns?.staff_35,
);
const ppdRows = filterOverviewPpdVyshneveRows(unitRows, undefined);
const polyRows = filterOverviewPolygonRows(unitRows, undefined);
const ppdWithRoster = filterOverviewPpdVyshneveRows(unitRows, (row) => {
  const name = row.name.split(/\s+/)[0] ?? "";
  return rosterRows.find((item) =>
    String(item.column_14 ?? "").includes(name),
  );
});
const polyWithRoster = filterOverviewPolygonRows(unitRows, (row) => {
  const name = row.name.split(/\s+/)[0] ?? "";
  return rosterRows.find((item) =>
    String(item.column_14 ?? "").includes(name),
  );
});
const sheets = buildOverviewPpdLocationExportSheets(unitRows, rosterRows);
console.log(
  "export sheet rows",
  (sheets[0]?.data.length ?? 0) - 2,
  (sheets[1]?.data.length ?? 0) - 2,
);
console.log("with roster lookup PPD/POLY", ppdWithRoster.length, polyWithRoster.length);
console.log("PPD", ppdRows.length, "POLY", polyRows.length);
console.log(
  "PPD names",
  ppdRows.map((row) => row.name.split(/\s+/)[0]),
);
console.log(
  "POLY names",
  polyRows.map((row) => row.name.split(/\s+/)[0]),
);

const refPath =
  "/Volumes/KINGSTON/army_work/Ранковий звіт/10.09.2026/3 Рота ППД.xlsx";
const refWorkbook = await XlsxPopulate.fromFileAsync(refPath);
const refSheet = refWorkbook.sheet(0);
const refUsed = refSheet.usedRange();
const refEndCol = refUsed?.endCell()?.columnNumber?.() ?? 20;
const refNames: string[] = [];
for (let column = 1; column <= refEndCol; column += 1) {
  for (let rowNumber = 1; rowNumber <= 12; rowNumber += 1) {
    const value = String(
      refSheet.cell(`${colLetter(column)}${rowNumber}`).value() ?? "",
    ).trim();
    if (value.length > 8 && /[А-ЯІЇЄҐ]/u.test(value) && value.includes(" ")) {
      refNames.push(value);
    }
  }
}
console.log("reference names", refNames.length, refNames);

let currentUnit = "";
const rawRows: Array<{
  name: string;
  ae: string;
  ai: string;
  status: string;
}> = [];
for (let rowNumber = 2; rowNumber <= endRow; rowNumber += 1) {
  const unit = String(sheet.cell(`B${rowNumber}`).value() ?? "").trim();
  if (unit) currentUnit = unit;
  if (!matchesOverviewRotaUnitFilter(currentUnit, "3 рота")) continue;
  const name = String(sheet.cell(`N${rowNumber}`).value() ?? "").trim();
  if (!name) continue;
  rawRows.push({
    name,
    ae: String(sheet.cell(`AE${rowNumber}`).value() ?? "").trim(),
    ai: String(sheet.cell(`AI${rowNumber}`).value() ?? "").trim(),
    status: String(sheet.cell(`U${rowNumber}`).value() ?? "").trim(),
  });
}
const rawPpd = rawRows.filter(
  (row) => row.ae === "ППД Вишневе" && (!row.ai || row.ai === "ППД"),
);
const rawPoly = rawRows.filter(
  (row) => row.ae === "ППД Вишневе" && row.ai === "ПОЛІГОН",
);
const rawPpdDuty = rawPpd.filter((row) => row.status === "В строю");
const rawPolyDuty = rawPoly.filter((row) => row.status === "В строю");
console.log("raw PPD all/duty", rawPpd.length, rawPpdDuty.length);
console.log("raw POLY all/duty", rawPoly.length, rawPolyDuty.length);
console.log(
  "raw PPD duty",
  rawPpdDuty.map((row) => row.name.split(/\s+/)[0]),
);
console.log(
  "raw POLY duty",
  rawPolyDuty.map((row) => row.name.split(/\s+/)[0]),
);

const gukRoster = rosterRows.find((row) =>
  String(row.column_14 ?? "").includes("ГУК"),
);
console.log("guk roster keys location", Object.fromEntries(
  Object.entries(gukRoster ?? {}).filter(([key]) =>
    /переб|column_3|column_4|31|35|40/u.test(key),
  ),
));

for (const surname of ["ГУК", "ЖИХАРЄВ", "СВЯТОХА", "ГНІЦЕВИЧ"]) {
  const overview = unitRows.find((row) => row.name.includes(surname));
  const raw = rawPpdDuty.find((row) => row.name.includes(surname)) ??
    rawPolyDuty.find((row) => row.name.includes(surname));
  console.log(surname, {
    raw,
    staff31: overview?.staffSheetColumns?.staff_31,
    staff35: overview?.staffSheetColumns?.staff_35,
    ppd: overview ? isOverviewPpdVyshneveLocation(overview) : null,
    poly: overview ? isOverviewPolygonLocation(overview) : null,
  });
}

const col31 = payload.sheets[0]?.columns.find((column) =>
  column.label.includes("перебування") && !column.label.includes("уточн"),
);
const col35entry = payload.sheets[0]?.columns.filter((column) =>
  column.label.toLowerCase().includes("уточн"),
);
console.log(
  "payload columns 31-ish",
  payload.sheets[0]?.columns.filter((c) => c.key.includes("31") || c.key.includes("35") || c.label.includes("переб")),
);
