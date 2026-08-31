/**
 * Показує «сирі» рядки аркушів ЕЖООС, де згадується особа.
 * Запуск: npx jiti scripts/debugEjoosRawRows.ts <ejoos.xlsx> <підрядок>
 */
import { readFileSync } from "node:fs";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../src/excelRoundTrip";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const [ejoosPath, needle] = process.argv.slice(2);
const ejoos = await readWorkbookSnapshot(asFile(ejoosPath), EJOOS_SYNC_READ_OPTIONS);

const cellText = (value: unknown) => {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "object") {
    const rich = value as { text?: () => string; value?: () => unknown };
    if (typeof rich.text === "function") return rich.text();
    if (typeof rich.value === "function") return cellText(rich.value());
    return "";
  }
  return String(value);
};

for (const sheet of ejoos.sheets) {
  sheet.rawRows.forEach((row, index) => {
    const cells = row.map(cellText);
    if (!cells.some((text) => text.toLowerCase().includes(needle.toLowerCase()))) {
      return;
    }
    console.log(`\n=== ${sheet.sheetName} R${index + 1}`);
    cells.forEach((text, column) => {
      if (!text.trim()) return;
      console.log(`  C${column + 1}: ${text}`);
    });
  });
}
