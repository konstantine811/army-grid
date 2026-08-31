/**
 * Показує, які рядки аркуша дають збіг за ID або ПІБ під час текстових звірок.
 * Запуск: npx jiti scripts/debugEjoosScanMatch.ts <ejoos.xlsx> <аркуш> <ID> <ПІБ>
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

const [ejoosPath, sheetPattern, personId, fullName] = process.argv.slice(2);
const ejoos = await readWorkbookSnapshot(asFile(ejoosPath), EJOOS_SYNC_READ_OPTIONS);
const sheet = ejoos.sheets.find((candidate) =>
  new RegExp(sheetPattern, "i").test(candidate.sheetName),
);
if (!sheet) throw new Error("Аркуш не знайдено");

const cellText = (value: unknown): string => {
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

const ABSENCE_RE = /СЗЧ|САМОВІЛ|БЕЗВІСТ|(?:^|[^А-ЯІЇЄҐа-яіїєґ])ЗБ(?:$|[^А-ЯІЇЄҐа-яіїєґ])/iu;
const nameKey = (fullName || "").toLowerCase().replace(/\s+/g, " ").trim();

sheet.rawRows.forEach((row, index) => {
  const cells = row.map(cellText);
  const text = cells.join(" ");
  if (!ABSENCE_RE.test(text)) return;
  const byId = personId
    ? cells.some((value) => value.replace(/\D/g, "") === personId)
    : false;
  const byName = nameKey
    ? text.toLowerCase().replace(/\s+/g, " ").includes(nameKey)
    : false;
  if (!byId && !byName) return;
  console.log(
    `R${index + 1} ${byId ? "[ID]" : ""}${byName ? "[ПІБ]" : ""}: ${cells
      .map((value, column) => (value.trim() ? `C${column + 1}=${value}` : ""))
      .filter(Boolean)
      .join(" | ")}`,
  );
});
