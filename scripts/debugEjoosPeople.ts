/**
 * Коротка перевірка кількох ID: які операції залишає план.
 * Запуск: npx jiti scripts/debugEjoosPeople.ts <ejoos.xlsx> <1пб.xlsx> <ID...>
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

const [ejoosPath, pbPath, ...ids] = process.argv.slice(2);
const ejoos = await readWorkbookSnapshot(asFile(ejoosPath), EJOOS_SYNC_READ_OPTIONS);
const pb = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);
const plan = buildEjoosSyncPlan(ejoos, pb, { statusRules: DEFAULT_STATUS_RULES });

for (const id of ids) {
  const ops = plan.ops.filter((op) => op.personId === id);
  console.log(`\nID ${id}: ${ops.length ? "" : "операцій немає"}`);
  for (const op of ops) {
    console.log(
      `  [${op.class}] ${op.kind} · ${op.fullName} · ${op.before} → ${op.after}`,
    );
  }
}
