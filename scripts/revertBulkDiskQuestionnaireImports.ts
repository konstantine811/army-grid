/**
 * Removes questionnaires and auto-extracted photos from loose bulk disk import.
 *
 * bunx jiti scripts/revertBulkDiskQuestionnaireImports.ts --dry-run --include-today
 * bunx jiti scripts/revertBulkDiskQuestionnaireImports.ts --apply --include-today
 * bunx jiti scripts/revertBulkDiskQuestionnaireImports.ts --apply --photos
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { execFileSync } from "node:child_process";

const BACKEND_ENV_PATH = resolve(
  import.meta.dirname,
  "../../army-backend/.env",
);
const API_BASE = process.env.ARMY_BACKEND_URL ?? "http://127.0.0.1:4000";

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

const login = async () => {
  const env = parseEnvFile(BACKEND_ENV_PATH);
  const email = process.env.ADMIN_EMAIL ?? env.ADMIN_EMAIL;
  const password = process.env.ADMIN_PASSWORD ?? env.ADMIN_PASSWORD;
  if (!email || !password) {
    throw new Error("ADMIN_EMAIL / ADMIN_PASSWORD missing in army-backend/.env");
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
  if (!json.accessToken) throw new Error("Login response has no accessToken");
  return json.accessToken;
};

const args = process.argv.slice(2);
const apply = args.includes("--apply");
const includeToday = args.includes("--include-today");
const includePhotos = args.includes("--photos");
const dryRun = !apply;

const databaseUrl =
  process.env.DATABASE_URL ?? parseEnvFile(BACKEND_ENV_PATH).DATABASE_URL;
if (!databaseUrl) {
  throw new Error("DATABASE_URL missing");
}

const sql = String.raw;
const query = (queryText: string) => {
  const raw = execFileSync(
    "psql",
    [databaseUrl.replace(/\?.*$/, ""), "-t", "-A", "-c", queryText],
    { encoding: "utf8" },
  ).trim();
  return raw ? raw.split("\n").map((line) => line.trim()).filter(Boolean) : [];
};

const windows = [
  {
    label: "2026-09-21 21:00–22:00 (bulk import yesterday evening)",
    from: "2026-09-21 21:00:00",
    to: "2026-09-21 22:00:00",
  },
];
if (includeToday) {
  windows.push({
    label: "2026-09-22 10:00–11:00 (bulk import this morning)",
    from: "2026-09-22 10:00:00",
    to: "2026-09-22 11:00:00",
  });
}

type TargetRow = {
  personExternalId: string;
  fileName: string;
  created: string;
};

const questionnaireTargets: TargetRow[] = [];
for (const window of windows) {
  const rows = query(sql`
    SELECT person_external_id || '|' || COALESCE(file_name, '') || '|' ||
           to_char(created_at AT TIME ZONE 'Europe/Kyiv', 'YYYY-MM-DD HH24:MI:SS')
    FROM person_questionnaires
    WHERE created_at AT TIME ZONE 'Europe/Kyiv' >= '${window.from}'
      AND created_at AT TIME ZONE 'Europe/Kyiv' < '${window.to}'
    ORDER BY created_at
  `).map((line) => {
    const [personExternalId, fileName, created] = line.split("|");
    return { personExternalId, fileName, created };
  });
  console.log(`${window.label}: ${rows.length}`);
  questionnaireTargets.push(...rows);
}

const uniqueQuestionnaireTargets = [
  ...new Map(
    questionnaireTargets.map((row) => [row.personExternalId, row]),
  ).values(),
];

const photoImportFrom = "2026-09-21 21:00:00";
const photoTargets = query(sql`
  WITH bulk_people AS (
    SELECT DISTINCT "entityId" AS person_external_id
    FROM change_logs
    WHERE "entityType" = 'person_questionnaires'
      AND action = 'DELETE_QUESTIONNAIRE'
      AND "createdAt" AT TIME ZONE 'Europe/Kyiv' >= '2026-09-22 07:50:00'
  )
  SELECT p.person_external_id || '|' || COALESCE(p.file_name, '') || '|' ||
         to_char(p.created_at AT TIME ZONE 'Europe/Kyiv', 'YYYY-MM-DD HH24:MI:SS')
  FROM person_photos p
  INNER JOIN bulk_people b ON b.person_external_id = p.person_external_id
  WHERE p.created_at AT TIME ZONE 'Europe/Kyiv' >= '${photoImportFrom}'
  ORDER BY p.created_at
`).map((line) => {
  const [personExternalId, fileName, created] = line.split("|");
  return { personExternalId, fileName, created };
});

console.log(
  `\nQuestionnaires to delete: ${uniqueQuestionnaireTargets.length}`,
);
for (const row of uniqueQuestionnaireTargets) {
  console.log(
    `- ${row.personExternalId} · ${row.fileName || "(no name)"} · ${row.created}`,
  );
}

console.log(`\nAuto-imported photos to delete: ${photoTargets.length}`);
for (const row of photoTargets) {
  console.log(
    `- ${row.personExternalId} · ${row.fileName || "(no name)"} · ${row.created}`,
  );
}

if (dryRun) {
  console.log(
    "\nDRY RUN — use --apply [--include-today] [--photos] to delete.",
  );
  process.exit(0);
}

const token = await login();
let deletedQuestionnaires = 0;
let failedQuestionnaires = 0;

for (const row of uniqueQuestionnaireTargets) {
  const response = await fetch(
    `${API_BASE}/ejournals/personnel/questionnaires/${encodeURIComponent(row.personExternalId)}`,
    {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    },
  );
  if (!response.ok) {
    failedQuestionnaires += 1;
    console.error(
      `QUESTIONNAIRE FAILED ${row.personExternalId}: ${response.status} ${await response.text()}`,
    );
    continue;
  }
  deletedQuestionnaires += 1;
}

console.log(
  `\nDeleted questionnaires: ${deletedQuestionnaires}, failed: ${failedQuestionnaires}`,
);

if (!includePhotos) {
  console.log("Pass --photos to delete auto-imported photos too.");
  process.exit(0);
}

let deletedPhotos = 0;
let failedPhotos = 0;

for (const row of photoTargets) {
  const response = await fetch(
    `${API_BASE}/ejournals/personnel/photos/${encodeURIComponent(row.personExternalId)}`,
    {
      method: "DELETE",
      headers: { Authorization: `Bearer ${token}` },
    },
  );
  if (!response.ok) {
    failedPhotos += 1;
    console.error(
      `PHOTO FAILED ${row.personExternalId}: ${response.status} ${await response.text()}`,
    );
    continue;
  }
  deletedPhotos += 1;
}

console.log(`Deleted photos: ${deletedPhotos}, failed: ${failedPhotos}`);
