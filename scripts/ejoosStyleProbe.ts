/**
 * Перевірка шрифту/розміру/bold у клітинках .xlsx після застосування ЄЖООС.
 *
 * Приклади:
 *   npm run ejoos:style-probe -- .debug-tmp/EJOOS_latest.xlsx --sheet шпо --rows 95-99
 *   npm run ejoos:style-probe -- out.xlsx --sheet табель --rows 120 --cols 6,7,8
 */
import { readFileSync } from "node:fs";
import { EJOOS_WRITE_STYLE } from "../src/pages/ejournal/ejoosStyleConfig";
import {
  inspectWorkbookCellStyles,
  parseRowRange,
} from "../src/pages/ejournal/ejoosStyleInspect";

const args = process.argv.slice(2);
const filePath = args.find((arg) => !arg.startsWith("--"));
if (!filePath) {
  console.error(
    "Використання: npm run ejoos:style-probe -- <file.xlsx> --sheet <ім'я> --rows 95-99 [--cols 1,2,6,7,8]",
  );
  process.exit(1);
}

const readFlag = (name: string) => {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : undefined;
};

const sheet = readFlag("--sheet") ?? "шпо";
const rowsSpec = readFlag("--rows") ?? "95-99";
const colsSpec = readFlag("--cols");
const columns = colsSpec
  ? colsSpec.split(",").map((part) => Number(part.trim())).filter(Boolean)
  : undefined;

const buffer = readFileSync(filePath);
const rows = parseRowRange(rowsSpec);
const cells = await inspectWorkbookCellStyles({
  file: buffer,
  sheetQuery: sheet,
  rows,
  columns,
});

console.log(`Файл: ${filePath}`);
console.log(
  `Очікується: ${EJOOS_WRITE_STYLE.fontName} ${EJOOS_WRITE_STYLE.fontSize}, bold=${EJOOS_WRITE_STYLE.bold}`,
);
console.log(`Аркуш «${sheet}», рядки ${rowsSpec}\n`);

let bad = 0;
for (const cell of cells) {
  const okFont = cell.fontName === EJOOS_WRITE_STYLE.fontName;
  const okSize = cell.fontSize === EJOOS_WRITE_STYLE.fontSize;
  const okBold = cell.bold === EJOOS_WRITE_STYLE.bold;
  const status = okFont && okSize && okBold ? "OK" : "!!";
  if (status === "!!") bad += 1;
  const preview =
    cell.value.length > 36 ? `${cell.value.slice(0, 33)}…` : cell.value;
  console.log(
    `${status} ${cell.ref.padEnd(5)} | ${cell.fontName.padEnd(16)} ${String(cell.fontSize).padStart(2)} | bold=${cell.bold ? "yes" : "no "} | ${preview}`,
  );
}

if (!cells.length) {
  console.log("(клітинок не знайдено — перевірте аркуш і діапазон рядків)");
  process.exit(2);
}

console.log(`\nПідсумок: ${bad} клітинок не відповідають ejoosStyleConfig.ts`);
process.exit(bad > 0 ? 1 : 0);
