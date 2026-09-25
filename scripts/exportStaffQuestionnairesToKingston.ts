/**
 * Експортує підтверджені PDF-анкети для осіб «У штаті» (актуальний «Загальний список»)
 * у /Volumes/KINGSTON/army_work/pre_ankets/Анкети зі штатки/
 * з іменами «ПРІЗВИЩЕ Ім'я По батькові (Позивний).pdf».
 *
 * Той самий набір, що лічильник «З анкетами · N / Без анкет · M» на Особовому складі:
 * 675 осіб зі Штатки + buildQuestionnairePresenceFromPeople.
 *
 * npx jiti scripts/exportStaffQuestionnairesToKingston.ts [outputDir]
 */
import { execFileSync } from "node:child_process";
import { createHash } from "node:crypto";
import {
  mkdirSync,
  readdirSync,
  readFileSync,
  statSync,
  unlinkSync,
  writeFileSync,
} from "node:fs";
import { join, resolve } from "node:path";
import type { BackendPersonnelRosterLatest } from "../src/api";
import type { PersonnelDataset } from "../src/data/personnelDatasetCore";
import {
  buildQuestionnairePresencePeople,
  resolveStoredQuestionnaireExternalId,
} from "../src/pages/personnel/personAttachments";
import { buildStaffRosterQuestionnaireExportTargets } from "../src/pages/personnel/staffRosterQuestionnaireExport";
import {
  buildQuestionnaireExportFileName,
  resolvePersonCallSign,
} from "../src/pages/personnel/personnelUtils";
import { sanitizeFileName } from "../src/shared/browserExport";

const BACKEND_ENV_PATH = resolve(
  import.meta.dirname,
  "../../army-backend/.env",
);
const API_BASE = process.env.ARMY_BACKEND_URL ?? "http://127.0.0.1:4000";
const DISK_ROOT =
  process.env.QUESTIONNAIRES_DISK_ROOT?.trim() ||
  "/Volumes/KINGSTON/army_work";
const PRE_ANKETS_ROOT = join(DISK_ROOT, "pre_ankets");
const DEFAULT_OUTPUT = join(PRE_ANKETS_ROOT, "Анкети зі штатки");

const FORM_ARCHIVE_DIRS = [
  join(DISK_ROOT, "Букля_деловod/Ф5 Ф6 Ф12"),
  join(DISK_ROOT, "Форми 5,6,12"),
];

type QuestionnaireMeta = {
  personExternalId: string;
  fileName: string | null;
};

const parseEnvFile = (path: string) => {
  const out: Record<string, string> = {};
  for (const line of readFileSync(path, "utf8").split("\n")) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const index = trimmed.indexOf("=");
    if (index <= 0) continue;
    const key = trimmed.slice(0, index).trim();
    let value = trimmed.slice(index + 1).trim();
    if (
      (value.startsWith('"') && value.endsWith('"')) ||
      (value.startsWith("'") && value.endsWith("'"))
    ) {
      value = value.slice(1, -1);
    }
    out[key] = value;
  }
  return out;
};

const hashFile = (absolutePath: string) =>
  createHash("md5").update(readFileSync(absolutePath)).digest("hex");

const collectExcludedFormPdfHashes = () => {
  const hashes = new Set<string>();
  const walk = (dir: string) => {
    if (!statSync(dir, { throwIfNoEntry: false })?.isDirectory()) return;
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      if (entry.name.startsWith("._")) continue;
      const fullPath = join(dir, entry.name);
      if (entry.isDirectory()) {
        walk(fullPath);
        continue;
      }
      if (!entry.isFile() || !/\.pdf$/i.test(entry.name)) continue;
      try {
        hashes.add(hashFile(fullPath));
      } catch {
        // ignore unreadable ghost files on USB
      }
    }
  };
  for (const root of FORM_ARCHIVE_DIRS) walk(root);
  return hashes;
};

const login = async () => {
  const env = parseEnvFile(BACKEND_ENV_PATH);
  const email = process.env.ADMIN_EMAIL ?? env.ADMIN_EMAIL;
  const password = process.env.ADMIN_PASSWORD ?? env.ADMIN_PASSWORD;
  if (!email || !password) {
    throw new Error(
      "Не знайдено ADMIN_EMAIL / ADMIN_PASSWORD у army-backend/.env",
    );
  }

  const response = await fetch(`${API_BASE}/auth/login`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email, password }),
  });
  if (!response.ok) {
    throw new Error(`Login failed: ${response.status} ${await response.text()}`);
  }
  const json = (await response.json()) as { accessToken?: string };
  if (!json.accessToken) {
    throw new Error("Login response has no accessToken");
  }
  return json.accessToken;
};

const fetchPersonnelDataset = async (token: string) => {
  const response = await fetch(`${API_BASE}/ejournals/personnel/dataset`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(
      `Personnel dataset failed: ${response.status} ${await response.text()}`,
    );
  }
  return (await response.json()) as PersonnelDataset;
};

const fetchRosterLatest = async (token: string) => {
  const response = await fetch(`${API_BASE}/ejournals/personnel/roster/latest`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(
      `Roster latest failed: ${response.status} ${await response.text()}`,
    );
  }
  return (await response.json()) as BackendPersonnelRosterLatest;
};

const fetchQuestionnaireMeta = async (token: string) => {
  const response = await fetch(`${API_BASE}/ejournals/personnel/questionnaires`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(
      `Questionnaire list failed: ${response.status} ${await response.text()}`,
    );
  }
  return (await response.json()) as QuestionnaireMeta[];
};

const fetchQuestionnairePdf = async (token: string, externalId: string) => {
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      const response = await fetch(
        `${API_BASE}/ejournals/personnel/questionnaires/${encodeURIComponent(externalId)}/file?download=1`,
        { headers: { Authorization: `Bearer ${token}` } },
      );
      if (!response.ok) return null;
      return Buffer.from(await response.arrayBuffer());
    } catch {
      if (attempt === 2) return null;
      await new Promise((resolve) => setTimeout(resolve, 500 * (attempt + 1)));
    }
  }
  return null;
};

const clearTargetPdfs = (outputDir: string) => {
  for (const entry of readdirSync(outputDir, { withFileTypes: true })) {
    if (entry.name.startsWith("._")) continue;
    if (!entry.isFile() || !/\.pdf$/i.test(entry.name)) continue;
    try {
      unlinkSync(join(outputDir, entry.name));
    } catch {
      // exFAT/USB can expose ghost dirents.
    }
  }
};

const outputDir = process.argv[2] || DEFAULT_OUTPUT;

if (!statSync(DISK_ROOT, { throwIfNoEntry: false })?.isDirectory()) {
  throw new Error(`Немає доступу до Kingston: ${DISK_ROOT}`);
}

mkdirSync(outputDir, { recursive: true });
clearTargetPdfs(outputDir);

const token = await login();
const [dataset, rosterLatest, questionnaireMeta] = await Promise.all([
  fetchPersonnelDataset(token),
  fetchRosterLatest(token),
  fetchQuestionnaireMeta(token),
]);

if (!dataset?.rows?.length) {
  throw new Error("Personnel dataset порожній — спочатку імпортуйте Штатку в БД.");
}

const exportTargets = buildStaffRosterQuestionnaireExportTargets({
  personnelRows: dataset.rows,
  rosterLatest,
  questionnaires: questionnaireMeta,
});

const allPeople = buildQuestionnairePresencePeople(dataset.rows);
const expectedWithAnketa = exportTargets.filter(
  (target) => target.hasQuestionnaire,
).length;
const staffTotal = exportTargets.length;
const excludedFormHashes = collectExcludedFormPdfHashes();

const usedSources = new Set<string>();
const usedTargets = new Set<string>();
let copied = 0;
let missing = 0;
let skippedForms = 0;
let unresolved = 0;
const logLines: string[] = [
  `Ціль: ${outputDir}`,
  `Джерело: «Загальний список» Штатки + анкети з БД (як у UI)`,
  `У штаті (Загальний список): ${staffTotal}`,
  `З анкетами (як у UI): ${expectedWithAnketa}`,
  `Анкет у БД: ${questionnaireMeta.length}`,
  `Виключено форм 5/6/12 (хешів): ${excludedFormHashes.size}`,
  "",
];

for (const target of exportTargets) {
  const { record, person } = target;
  const personId = record.summary.externalId;
  const fullName = record.summary.name;
  const callSign = resolvePersonCallSign(record.row) || record.summary.callSign;
  const targetName = sanitizeFileName(
    buildQuestionnaireExportFileName(fullName, callSign),
  );
  const targetPath = join(outputDir, targetName);

  if (!target.hasQuestionnaire) {
    missing += 1;
    logLines.push(`MISS  ${fullName}`);
    continue;
  }

  if (usedTargets.has(targetPath)) {
    missing += 1;
    logLines.push(`CLASH ${fullName} · ${targetName}`);
    continue;
  }

  const dbExternalId = resolveStoredQuestionnaireExternalId(
    person,
    questionnaireMeta,
    allPeople,
  );
  if (!dbExternalId) {
    unresolved += 1;
    logLines.push(`NORES ${fullName} · не знайдено id анкети в БД`);
    continue;
  }

  const sourceKey = `db:${dbExternalId}`;
  if (usedSources.has(sourceKey)) {
    missing += 1;
    logLines.push(`DUPDB ${fullName} · ${dbExternalId}`);
    continue;
  }

  const pdf = await fetchQuestionnairePdf(token, dbExternalId);
  if (!pdf?.length) {
    missing += 1;
    logLines.push(`EMPTY ${fullName} · ${dbExternalId}`);
    continue;
  }

  const pdfHash = createHash("md5").update(pdf).digest("hex");
  if (excludedFormHashes.has(pdfHash)) {
    skippedForms += 1;
    logLines.push(`FORM  ${fullName} · не анкета (форма 5/6/12)`);
    continue;
  }

  writeFileSync(targetPath, pdf);
  usedSources.add(sourceKey);
  usedTargets.add(targetPath);
  copied += 1;
  logLines.push(`OK    ${targetName}`);
}

logLines.push("");
logLines.push(`Скопійовано: ${copied}`);
logLines.push(`Очікувалось (UI «З анкетами»): ${expectedWithAnketa}`);
logLines.push(`Пропущено (форма 5/6/12): ${skippedForms}`);
logLines.push(`Без анкети у штаті: ${missing}`);
if (unresolved) {
  logLines.push(`Не вдалося зіставити id: ${unresolved}`);
}

const reportPath = join(outputDir, "_export-report.txt");
writeFileSync(reportPath, `${logLines.join("\n")}\n`, "utf8");
console.log(logLines.join("\n"));
console.log(`\nЗвіт: ${reportPath}`);

try {
  execFileSync("touch", [outputDir]);
} catch {
  // ignore
}
