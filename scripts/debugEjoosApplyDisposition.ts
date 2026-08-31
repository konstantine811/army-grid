/**
 * Прогін застосування «в розпорядження» на копії книги + перевірка результату.
 * Запуск: npx jiti scripts/debugEjoosApplyDisposition.ts <ejoos.xlsx> <1пб.xlsx> <out.xlsx>
 */
import { readFileSync, writeFileSync } from "node:fs";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../src/excelRoundTrip";
import { applyDispositionWithZip } from "../src/pages/ejournal/ejoosDispositionZip";
import { DEFAULT_STATUS_RULES } from "../src/pages/ejournal/ejoosRules";
import { buildEjoosSyncPlan } from "../src/pages/ejournal/ejoosSyncPlan";

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const [ejoosPath, pbPath, outPath] = process.argv.slice(2);
const ejoos = await readWorkbookSnapshot(asFile(ejoosPath), EJOOS_SYNC_READ_OPTIONS);
const pb = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);
const plan = buildEjoosSyncPlan(ejoos, pb, { statusRules: DEFAULT_STATUS_RULES });
const ops = plan.ops.filter(
  (op) => op.kind === "move_to_disposition" && op.class === "ready",
);
console.log(`Застосовуємо операцій: ${ops.length}`);

// У Node снапшот тримає Buffer, а писар очікує Blob з `arrayBuffer()`.
const ejoosForApply = {
  ...ejoos,
  file: new File([readFileSync(ejoosPath)], ejoos.fileName),
};
const blob = await applyDispositionWithZip({ ejoos: ejoosForApply, plan, ops });
writeFileSync(outPath, Buffer.from(await blob.arrayBuffer()));
console.log(`Записано: ${outPath}`);

const check = await readWorkbookSnapshot(asFile(outPath), EJOOS_SYNC_READ_OPTIONS);
console.log(
  `Перечитано аркушів: ${check.sheets.map((sheet) => `${sheet.sheetName}(${sheet.rawRows.length})`).join(", ")}`,
);
