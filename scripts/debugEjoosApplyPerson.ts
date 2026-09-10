/**
 * Локальний прогін: план → apply → Табель → стилі для однієї особи.
 *
 *   npx jiti scripts/debugEjoosApplyPerson.ts \
 *     ~/Downloads/ЄЖООС.xlsx ~/Downloads/1ПБ.xlsx 11688 \
 *     --as-of 31.08.2026 --out .debug-tmp/volkov-out.xlsx
 *
 * Без --out: лише план і preview, apply не пишемо.
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import {
  EJOOS_SYNC_READ_OPTIONS,
  readWorkbookSnapshot,
} from "../src/excelRoundTrip";
import { applyConfirmedEjoosOps } from "../src/pages/ejournal/ejoosSyncApply";
import {
  buildTimesheetPreview,
  groupOpsIntoPersonChanges,
  isWorkbookApplyOp,
} from "../src/pages/ejournal/ejoosPersonDiff";
import {
  personApplyBlockReason,
  personCanEnterApplyQueue,
} from "../src/pages/ejournal/ejoosOpRequirements";
import { DEFAULT_STATUS_RULES } from "../src/pages/ejournal/ejoosRules";
import { buildEjoosSyncPlan } from "../src/pages/ejournal/ejoosSyncPlan";
import { EJOOS_WRITE_STYLE } from "../src/pages/ejournal/ejoosStyleConfig";
import { inspectWorkbookCellStyles } from "../src/pages/ejournal/ejoosStyleInspect";

const args = process.argv.slice(2);
const readFlag = (name: string) => {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : undefined;
};

const positional = args.filter((arg) => !arg.startsWith("--"));
const [ejoosPath, pbPath, personQuery] = positional;
const asOf = readFlag("--as-of");
const outPath = readFlag("--out");

if (!ejoosPath || !pbPath || !personQuery) {
  console.error(
    "Використання: npx jiti scripts/debugEjoosApplyPerson.ts <ejoos.xlsx> <1пб.xlsx> <ID|ПІБ> [--as-of DD.MM.YYYY] [--out file.xlsx]",
  );
  process.exit(1);
}

const asFile = (path: string) => {
  const buffer = readFileSync(path) as unknown as File & { name: string };
  Object.defineProperty(buffer, "name", {
    value: path.split("/").pop() ?? path,
  });
  return buffer;
};

const needle = personQuery.toLocaleLowerCase("uk-UA");
const matchesPerson = (personId: string, fullName: string) =>
  personId === personQuery ||
  fullName.toLocaleLowerCase("uk-UA").includes(needle);

const ejoosSnap = await readWorkbookSnapshot(
  asFile(ejoosPath),
  EJOOS_SYNC_READ_OPTIONS,
);
const pbSnap = await readWorkbookSnapshot(asFile(pbPath), EJOOS_SYNC_READ_OPTIONS);

const plan = buildEjoosSyncPlan(ejoosSnap, pbSnap, {
  statusRules: DEFAULT_STATUS_RULES,
  sourceAsOfDate: asOf,
});

console.log(`План: ${plan.timesheetDayLabel} (день ${plan.timesheetDay})`);
console.log(`Операцій у плані: ${plan.ops.length}`);

const session = groupOpsIntoPersonChanges(plan, pbSnap);
const person = session.people.find(
  (item) =>
    matchesPerson(item.personId, item.fullName) ||
    item.ops.some(
      (op) =>
        matchesPerson(op.personId, op.fullName) ||
        op.fullName.toLocaleLowerCase("uk-UA").includes(needle),
    ),
);

if (!person) {
  console.error(`Особу «${personQuery}» не знайдено в плані.`);
  process.exit(2);
}

console.log(`\n=== ${person.fullName} · ID ${person.personId} · ${person.severity} ===`);
console.log(`canApply: ${personCanEnterApplyQueue(person)}`);
console.log(`blockReason: ${personApplyBlockReason(person.ops) ?? "(немає)"}`);

const payloadKeys = [
  "returningFromDisposition",
  "timesheetActiveFrom",
  "timesheetAbsenceSpans",
  "timesheetExcelRow",
  "shpoExcelRow",
  "openAbsenceExcelRow",
  "keepOpenSzchTimesheet",
  "orderDate",
  "orderNumber",
];

for (const op of person.ops) {
  console.log(`\n[${op.class}] ${op.kind}`);
  console.log(`  sheet: ${op.sheet}`);
  console.log(`  ${op.before} → ${op.after}`);
  console.log(`  apply: ${isWorkbookApplyOp(op)}`);
  const payload = Object.fromEntries(
    payloadKeys
      .filter((key) => op.payload[key as keyof typeof op.payload])
      .map((key) => [key, op.payload[key as keyof typeof op.payload]]),
  );
  if (Object.keys(payload).length) {
    console.log(`  payload: ${JSON.stringify(payload, null, 2).replaceAll("\n", "\n  ")}`);
  }
}

const preview = buildTimesheetPreview(
  person.ops,
  plan.timesheetDay,
  plan.timesheetDayLabel,
);
if (preview) {
  console.log("\nPreview Табель:");
  console.log(
    preview.runs
      .map((run) =>
        run.from === run.to
          ? `${String(run.from).padStart(2, "0")}: ${run.mark}`
          : `${String(run.from).padStart(2, "0")}–${String(run.to).padStart(2, "0")}: ${run.mark}`,
      )
      .join(" · "),
  );
} else {
  console.log("\nPreview Табель: (null)");
}

const applyOps = person.ops.filter(isWorkbookApplyOp);
if (!outPath) {
  console.log(
    `\nApply ops: ${applyOps.length}. Додайте --out .debug-tmp/out.xlsx щоб записати результат.`,
  );
  process.exit(0);
}

if (personApplyBlockReason(person.ops)) {
  console.error("\nApply заблоковано — див. blockReason вище.");
  process.exit(3);
}

if (!applyOps.length) {
  console.error("\nНемає workbook apply ops.");
  process.exit(4);
}

const ejoosForApply = {
  ...ejoosSnap,
  file: new File([readFileSync(ejoosPath)], ejoosSnap.fileName),
};

console.log(`\nЗастосовуємо ${applyOps.length} ops…`);
const { blob } = await applyConfirmedEjoosOps({
  ejoos: ejoosForApply,
  plan,
  ops: applyOps,
});

mkdirSync(dirname(outPath), { recursive: true });
writeFileSync(outPath, Buffer.from(await blob.arrayBuffer()));
console.log(`Записано: ${outPath}`);

const after = await readWorkbookSnapshot(asFile(outPath), EJOOS_SYNC_READ_OPTIONS);
const ts = after.sheets.find((sheet) => /табель/i.test(sheet.sheetName));
const rows =
  ts?.rawRows
    .map((row, index) => ({
      row: index + 1,
      name: String(row[6] ?? ""),
      id: String(row[7] ?? ""),
      index: String(row[1] ?? ""),
      days: row.slice(8, 39).map((cell) => String(cell ?? "").trim()),
    }))
    .filter(
      (item) =>
        item.id === person.personId ||
        item.name.toLocaleLowerCase("uk-UA").includes(needle),
    ) ?? [];

console.log(`\nРядки Табеля (${rows.length}):`);
for (const item of rows) {
  console.log(`  R${item.row} інд.${item.index} · ${item.name} · ID ${item.id}`);
  const marks = item.days
    .map((mark, dayIndex) =>
      mark
        ? `${String(dayIndex + 1).padStart(2, "0")}=${mark.length > 12 ? `${mark.slice(0, 10)}…` : mark}`
        : "",
    )
    .filter(Boolean);
  console.log(`    ${marks.join("  ")}`);
}

const staffRow = rows.find((row) => /^\d{5,}$/.test(row.index)) ?? rows[0];
if (staffRow) {
  const dayCols = Array.from({ length: 31 }, (_, day) => 9 + day);
  const styles = await inspectWorkbookCellStyles({
    file: readFileSync(outPath),
    sheetQuery: /табель/i,
    rows: [staffRow.row],
    columns: [6, 7, 8, ...dayCols],
  });
  const bad = styles.filter(
    (cell) =>
      cell.fontName !== EJOOS_WRITE_STYLE.fontName ||
      cell.fontSize !== EJOOS_WRITE_STYLE.fontSize ||
      cell.bold !== EJOOS_WRITE_STYLE.bold,
  );
  console.log(
    `\nСтилі R${staffRow.row} (очікується ${EJOOS_WRITE_STYLE.fontName} ${EJOOS_WRITE_STYLE.fontSize}, bold=${EJOOS_WRITE_STYLE.bold}):`,
  );
  for (const cell of styles.slice(0, 12)) {
    console.log(
      `  ${cell.ref} | ${cell.fontName} ${cell.fontSize} bold=${cell.bold ? "yes" : "no"}`,
    );
  }
  if (styles.length > 12) console.log(`  … ще ${styles.length - 12} клітинок`);
  console.log(`  Проблемних: ${bad.length}/${styles.length}`);
  if (bad.length) {
    console.log(
      `\nДетальніше: npm run ejoos:style-probe -- ${outPath} --sheet табель --rows ${staffRow.row}`,
    );
  }
}

console.log(
  `\nСтилі ШПО: npm run ejoos:style-probe -- ${outPath} --sheet шпо --rows <R>`,
);
