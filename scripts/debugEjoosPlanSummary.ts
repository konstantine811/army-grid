/**
 * Зведення плану синхронізації: скільки операцій за типом/класом і які саме
 * кроки ланцюга створює детектор.
 * Запуск: npx jiti scripts/debugEjoosPlanSummary.ts <ejoos.xlsx> <1пб.xlsx>
 */
import { readFileSync } from "node:fs";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../src/excelRoundTrip";
import { DEFAULT_STATUS_RULES } from "../src/pages/ejournal/ejoosRules";
import { buildEjoosSyncPlan } from "../src/pages/ejournal/ejoosSyncPlan";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const [ejoosPath, pbPath] = process.argv.slice(2);
const ejoos = await readWorkbookSnapshot(asFile(ejoosPath), EJOOS_SYNC_READ_OPTIONS);
const pb = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);
const plan = buildEjoosSyncPlan(ejoos, pb, { statusRules: DEFAULT_STATUS_RULES });

const byKind = new Map<string, Map<string, number>>();
for (const op of plan.ops) {
  const classes = byKind.get(op.kind) ?? new Map<string, number>();
  classes.set(op.class, (classes.get(op.class) ?? 0) + 1);
  byKind.set(op.kind, classes);
}
console.log(`Усього операцій: ${plan.ops.length}`);
console.log(plan.summary);
for (const [kind, classes] of [...byKind].sort()) {
  console.log(
    `  ${kind}: ${[...classes].map(([cls, count]) => `${cls}=${count}`).join(", ")}`,
  );
}

const dispositions = plan.ops.filter((op) => op.kind === "move_to_disposition");
console.log(`\nУ розпорядження: ${dispositions.length}`);
for (const op of dispositions) {
  console.log(
    `  [${op.class}] ID ${op.personId || "—"} ${op.fullName} · ${op.payload.previousIndex} · ШПО R${op.payload.shpoExcelRow || "—"}${op.payload.skipShpoDisposition === "1" ? " (блок розпорядження вже є)" : ""} · Табель R${op.payload.timesheetExcelRow || "—"} · Відсутні R${op.payload.absenceExcelRow || "новий"} ${op.payload.absenceType} ${op.payload.absenceDate || "—"} · ООС: ${op.payload.remainsInOos === "true" ? "є" : "нема"} · szch=${op.payload.hasSzchContext}/остався=${op.payload.szchRemains}/деінде=${op.payload.szchReflectedElsewhere} · новий запис відсутності=${op.payload.needsAbsenceRecord || "0"}`,
  );
}

const conflicts = plan.ops.filter((op) => op.class === "conflict");
console.log(`\nКонфлікти: ${conflicts.length}`);
for (const op of conflicts) {
  console.log(
    `  ${op.kind} · ID ${op.personId || "—"} ${op.fullName} · ${op.before} → ${op.after}`,
  );
}

const chainOps = plan.ops.filter((op) => op.payload.chainTotal);
console.log(`\nКроки ланцюга: ${chainOps.length}`);
for (const op of chainOps) {
  console.log(
    `  [${op.class}] ${op.kind} · крок ${op.payload.chainStep}/${op.payload.chainTotal} · ID ${op.personId || "—"} ${op.fullName} · ${op.before} → ${op.after}`,
  );
}

const closing = plan.ops.filter((op) => op.payload.closeOldPosition === "1");
console.log(`\nЗакриття старої посади: ${closing.length}`);
for (const op of closing) {
  console.log(
    `  [${op.class}] ID ${op.personId || "—"} ${op.fullName} · ${op.payload.previousIndex} → ${op.payload.nextIndex} · Виключені: «${op.payload.documentsDest}» · ШПО R${op.payload.previousShpoExcelRow || "—"} · Табель R${op.payload.previousIndexTimesheetExcelRow || "—"} → R${op.payload.timesheetExcelRow || "—"}`,
  );
}
