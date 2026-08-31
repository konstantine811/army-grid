/**
 * Локальна перевірка плану синхронізації по одній особі.
 * Запуск: npx jiti scripts/debugEjoosPerson.ts <ejoos.xlsx> <1пб.xlsx> <ID>
 */
import { readFileSync } from "node:fs";
import { EJOOS_SYNC_READ_OPTIONS, readWorkbookSnapshot } from "../src/excelRoundTrip";
import { DEFAULT_STATUS_RULES } from "../src/pages/ejournal/ejoosRules";
import {
  classifyStaffMove,
  isOutboundStaffMove,
} from "../src/pages/ejournal/ejoosMovementRules";
import {
  buildEjoosSyncPlan,
  findEjoosSheet,
  parseEjoosAbsents,
  parseEjoosOos,
  parseEjoosShpo,
  parsePbArchive,
  parsePbMovements,
  parsePbShPeople,
} from "../src/pages/ejournal/ejoosSyncPlan";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const [ejoosPath, pbPath, personId] = process.argv.slice(2);

const ejoos = await readWorkbookSnapshot(asFile(ejoosPath), EJOOS_SYNC_READ_OPTIONS);
const pb = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);

console.log("Аркуші ЕЖООС:", ejoos.sheets.map((sheet) => sheet.sheetName).join(" | "));
console.log("Аркуші 1ПБ:", pb.sheets.map((sheet) => sheet.sheetName).join(" | "));

const movements = parsePbMovements(pb).filter(
  (event) => event.personId === personId,
);
console.log("\nРУХ:");
for (const event of movements) {
  const scope = classifyStaffMove(event);
  console.log(
    `  R${event.excelRow} ${event.type} | ${event.fullName} | ${event.previousIndex} → ${event.nextIndex} | наказ ${event.orderNumber} від ${event.orderDate} | статус «${event.status}» | куди «${event.destination}» | примітка «${event.note}» | зміна «${event.changeText}» | сфера ${scope}${isOutboundStaffMove(event) ? " (вибуття)" : ""}`,
  );
}

console.log("\narchive:");
for (const period of parsePbArchive(pb).filter((p) => p.personId === personId)) {
  console.log(
    `  R${period.excelRow} ${period.absenceType} | ${period.place} | з ${period.departDate} | повернення «${period.returnDate}» | наказ ${period.orderNumber} від ${period.orderDate}`,
  );
}

console.log("\nsh:");
for (const person of parsePbShPeople(pb).filter((p) => p.personId === personId)) {
  console.log(
    `  R${person.excelRow} ${person.fullName} | інд. ${person.positionIndex} | статус «${person.status}»`,
  );
}

const shpoSheet = findEjoosSheet(ejoos, /шпо|штатно.?посад/i);
const oosSheet = findEjoosSheet(ejoos, /(^|[.\s])оос($|[\s])/i);
const absentSheet = findEjoosSheet(ejoos, /тимчасов.*відсут/i);

console.log("\nЕЖООС ШПО:");
for (const row of parseEjoosShpo(shpoSheet).filter((r) => r.personId === personId)) {
  console.log(`  R${row.excelRow} інд. ${row.positionIndex} | ${row.fullName}`);
}
console.log("ЕЖООС ООС:");
for (const row of parseEjoosOos(oosSheet).filter((r) => r.personId === personId)) {
  console.log(`  R${row.excelRow} інд. ${row.positionIndex} | ${row.fullName}`);
}
console.log("ЕЖООС Тимчасово відсутні:");
for (const row of parseEjoosAbsents(absentSheet).filter(
  (r) => r.personId === personId,
)) {
  console.log(
    `  R${row.excelRow} ${row.ground} | ${row.place} | з ${row.departDate} | факт. повернення «${row.actualReturn}»`,
  );
}

const plan = buildEjoosSyncPlan(ejoos, pb, {
  statusRules: DEFAULT_STATUS_RULES,
});
console.log(`\nОперації плану (усього ${plan.ops.length}):`);
for (const op of plan.ops.filter((candidate) => candidate.personId === personId)) {
  console.log(`  [${op.class}] ${op.kind} | ${op.sheet}`);
  console.log(`    before: ${op.before}`);
  console.log(`    after:  ${op.after}`);
  console.log(`    why:    ${op.why}`);
}
