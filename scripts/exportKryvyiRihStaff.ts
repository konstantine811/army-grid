/**
 * Excel: в/сл зі штатки, м. Кривий Ріг та Криворізький район.
 *
 * npx jiti scripts/exportKryvyiRihStaff.ts [output.xlsx]
 */
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import XlsxPopulate from "xlsx-populate";
import type { BackendPersonnelRosterLatest } from "../src/api";
import type { PersonnelDataset } from "../src/data/personnelDatasetCore";
import { normalizeRosterMatchText } from "../src/pages/personnel/fighterStatusImport";
import { indexKryvyiRihPlaces } from "../src/pages/personnel-v2/kryvyiRihStaff";
import { buildPersonnelV2StaffSheet } from "../src/pages/personnel-v2/loadPersonnelV2StaffSheet";

const BACKEND_ENV_PATH = resolve(
  import.meta.dirname,
  "../../army-backend/.env",
);
const API_BASE = process.env.ARMY_BACKEND_URL ?? "http://127.0.0.1:4000";
const DEFAULT_OUTPUT =
  "/Volumes/KINGSTON/army_work/exports/в_сл_Кривий_Ріг.xlsx";

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
    throw new Error(`Login failed: ${response.status}`);
  }
  const json = (await response.json()) as { accessToken?: string };
  if (!json.accessToken) throw new Error("Login response has no accessToken");
  return json.accessToken;
};

const apiGet = async <T>(token: string, path: string) => {
  const response = await fetch(`${API_BASE}${path}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(`${path} failed: ${response.status}`);
  }
  return (await response.json()) as T;
};

const staffCell = (
  person: { cells: Array<{ column: number; value: string }> },
  column: number,
) => person.cells.find((cell) => cell.column === column)?.value ?? "";

const outputPath = process.argv[2] || DEFAULT_OUTPUT;
const token = await login();
const [roster, dataset] = await Promise.all([
  apiGet<BackendPersonnelRosterLatest>(
    token,
    "/ejournals/personnel/roster/latest",
  ),
  apiGet<PersonnelDataset>(token, "/ejournals/personnel/dataset"),
]);

const staff = buildPersonnelV2StaffSheet(roster);
const places = indexKryvyiRihPlaces(dataset.rows ?? []);
const people = staff.people.flatMap((person) => {
  if (person.inArchive) return [];
  const place = places[normalizeRosterMatchText(person.name)] ?? "";
  if (!place) return [];
  return [
    {
      name: person.name,
      status: staffCell(person, 21) || staffCell(person, 37),
      combatStatus: staffCell(person, 23) || staffCell(person, 42),
      location: staffCell(person, 31) || staffCell(person, 35) || staffCell(person, 40),
      rank: staffCell(person, 13),
      callSign: staffCell(person, 15),
      unit: staffCell(person, 2),
      place,
    },
  ];
});

const headers = [
  "№",
  "ПІБ",
  "Статус (штатка)",
  "Статус БГ",
  "Місце перебування",
  "Звання",
  "Позивний",
  "Підрозділ",
  "Місце народження / проживання / реєстрації",
];
const workbook = await XlsxPopulate.fromBlankAsync();
const sheet = workbook.sheet(0);
sheet.name("Кривий Ріг");
headers.forEach((header, index) => {
  const cell = sheet.cell(1, index + 1);
  cell.value(header);
  cell.style({
    bold: true,
    fill: "E8EAED",
    border: true,
    wrapText: true,
  });
});
people.forEach((person, index) => {
  const values = [
    index + 1,
    person.name,
    person.status,
    person.combatStatus,
    person.location,
    person.rank,
    person.callSign,
    person.unit,
    person.place,
  ];
  values.forEach((value, column) => {
    const cell = sheet.cell(index + 2, column + 1);
    cell.value(value);
    cell.style({ border: true, wrapText: true });
  });
});
sheet.column(1).width(6);
sheet.column(2).width(36);
sheet.column(3).width(22);
sheet.column(4).width(16);
sheet.column(5).width(28);
sheet.column(6).width(22);
sheet.column(7).width(16);
sheet.column(8).width(22);
sheet.column(9).width(55);
sheet.row(1).height(22);
sheet.freezePanes(1, 1);

await workbook.toFileAsync(outputPath);
console.log(`${people.length} рядків → ${outputPath}`);
