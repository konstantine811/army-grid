/**
 * Знайти однакові номери військових квитків у різних людей (Анкетні дані ООС).
 * Запуск: npx jiti scripts/debugAnketaMilitaryIdDuplicates.ts
 */
import {
  anketaSheetCsvUrls,
  parseAnketaCsv,
} from "../src/pages/anketa-data/anketaSheet";
import { filterAnketaRowsByEjoosSource } from "../src/pages/anketa-data/anketaEjoosPeople";
import { normalizeAnketaNameKey } from "../src/pages/anketa-data/anketaPersonMatch";
import { extractMilitaryIdFromText } from "../src/pages/personnel/vkTpvDovidkyImport";

const fetchCsvText = async () => {
  let lastError: Error | null = null;
  for (const url of anketaSheetCsvUrls()) {
    try {
      const response = await fetch(url, { credentials: "omit" });
      if (!response.ok) {
        lastError = new Error(`HTTP ${response.status} · ${url}`);
        continue;
      }
      const text = await response.text();
      if (!text.trim() || text.trimStart().startsWith("<!DOCTYPE")) {
        lastError = new Error("HTML замість CSV");
        continue;
      }
      return text;
    } catch (error) {
      lastError =
        error instanceof Error ? error : new Error("Невідома помилка мережі");
    }
  }
  throw lastError ?? new Error("Не вдалося завантажити Google Sheet.");
};

const isSkippableMilitaryId = (value: string) => {
  const text = value.replace(/\s+/g, " ").trim();
  if (!text) return true;
  if (/^[-–—−‑‒―_]+$/.test(text)) return true;
  return /^(?:відсутн|дані відсутн|немає|забув)$/iu.test(text);
};

const csv = await fetchCsvText();
const snapshot = parseAnketaCsv(csv, {
  source: "google",
  sourceLabel: "Google Sheets · Анкети",
});
const rows = filterAnketaRowsByEjoosSource(snapshot.rows, "oos");

type PersonHit = { name: string; id: string; raw: string; rowNumber: number };

const byMilitaryId = new Map<string, PersonHit[]>();

for (const row of rows) {
  const raw = String(row.militaryId ?? "").trim();
  if (isSkippableMilitaryId(raw)) continue;
  const key = extractMilitaryIdFromText(raw) || raw.replace(/\s+/g, " ").trim();
  if (!/[\d]/u.test(key)) continue;
  const list = byMilitaryId.get(key) ?? [];
  list.push({
    name: String(row.fullName ?? "").trim(),
    id: String(row.externalId ?? "").trim(),
    raw,
    rowNumber: row.__rowNumber,
  });
  byMilitaryId.set(key, list);
}

const duplicates = [...byMilitaryId.entries()]
  .filter(([, list]) => {
    const names = new Set(
      list.map((item) => normalizeAnketaNameKey(item.name)).filter(Boolean),
    );
    return names.size > 1;
  })
  .sort((left, right) => right[1].length - left[1].length);

console.log(`Анкетні дані ООС: ${rows.length} осіб`);
console.log(`Унікальних номерів ВК/ТПВ з цифрами: ${byMilitaryId.size}`);
console.log(
  `Дублікати номера у різних ПІБ: ${duplicates.length} номер(ів), ${duplicates.reduce((sum, [, list]) => sum + list.length, 0)} рядків`,
);

for (const [militaryId, list] of duplicates.slice(0, 30)) {
  console.log(`\n${militaryId} (${list.length} рядків):`);
  for (const person of list) {
    console.log(
      `  R${person.rowNumber} · ${person.name} · ID ${person.id || "—"} · «${person.raw}»`,
    );
  }
}

if (duplicates.length > 30) {
  console.log(`\n… і ще ${duplicates.length - 30} номер(ів) з дублями.`);
}

// Один ID — різні ПІБ
const byExternalId = new Map<string, PersonHit[]>();
for (const row of rows) {
  const id = String(row.externalId ?? "")
    .trim()
    .replace(/\.0+$/, "");
  if (!id || !/^\d+$/.test(id)) continue;
  const list = byExternalId.get(id) ?? [];
  list.push({
    name: String(row.fullName ?? "").trim(),
    id,
    raw: String(row.militaryId ?? "").trim(),
    rowNumber: row.__rowNumber,
  });
  byExternalId.set(id, list);
}
const duplicateIds = [...byExternalId.entries()].filter(([, list]) => {
  const names = new Set(
    list.map((item) => normalizeAnketaNameKey(item.name)).filter(Boolean),
  );
  return names.size > 1;
});

console.log(`\nОдин ID у різних ПІБ: ${duplicateIds.length} ID`);
for (const [externalId, list] of duplicateIds.slice(0, 15)) {
  console.log(`\nID ${externalId} (${list.length} рядків):`);
  for (const person of list) {
    console.log(
      `  R${person.rowNumber} · ${person.name} · «${person.raw || "—"}»`,
    );
  }
}

// Один ПІБ — різні військові квитки
type NameHit = PersonHit;
const byName = new Map<string, NameHit[]>();
for (const row of rows) {
  const key = normalizeAnketaNameKey(row.fullName);
  if (!key) continue;
  const list = byName.get(key) ?? [];
  list.push({
    name: String(row.fullName ?? "").trim(),
    id: String(row.externalId ?? "").trim(),
    raw: String(row.militaryId ?? "").trim(),
    rowNumber: row.__rowNumber,
  });
  byName.set(key, list);
}
const duplicateNames = [...byName.entries()]
  .filter(([, list]) => {
    if (list.length < 2) return false;
    const values = new Set(
      list.map((item) => item.raw.trim().toLocaleLowerCase("uk-UA")),
    );
    return values.size > 1;
  })
  .sort((left, right) => right[1].length - left[1].length);

console.log(
  `\nОдин ПІБ — різні значення військового квитка: ${duplicateNames.length} осіб`,
);
for (const [, list] of duplicateNames.slice(0, 20)) {
  console.log(`\n${list[0]!.name} (${list.length} рядків):`);
  for (const person of list) {
    console.log(
      `  R${person.rowNumber} · ID ${person.id || "—"} · «${person.raw || "—"}»`,
    );
  }
}

const query = process.argv[2]?.trim().toLocaleLowerCase("uk-UA");
if (query) {
  const hits = snapshot.rows.filter((row) =>
    String(row.fullName ?? "")
      .toLocaleLowerCase("uk-UA")
      .includes(query),
  );
  console.log(`\nПошук «${process.argv[2]}»: ${hits.length} рядків у всій анкеті`);
  for (const row of hits) {
    console.log(
      `  R${row.__rowNumber} · source=${row.__ejoosSource ?? "oos"} · ID ${row.externalId || "—"} · «${row.militaryId || "—"}» · ${row.fullName}`,
    );
  }
}

