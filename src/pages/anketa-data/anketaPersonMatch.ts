import { api, type BackendEjournalImport } from "../../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
} from "../../data/idbDataCache";
import type {
  EjournalPreviewRow,
} from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  findEjournalPersonnelSheet,
  getPersonExternalId,
  loadAllEjournalSheetRows,
  normalizePersonBirthKey,
} from "../personnel/personnelUtils";
import { mergeRosterRowsIntoPreview } from "../personnel/personnelRosterMerge";
import { mapRosterLatestToPreviewRows } from "../excel-fill/rosterSourceSnapshot";
import type { AnketaRow } from "./anketaSheet";

export const normalizeAnketaNameKey = (value: unknown) =>
  String(value ?? "")
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

/** Повний ключ і «прізвище + імʼя», щоб ловити списки без по батькові. */
export const anketaNameKeyVariants = (value: unknown) => {
  const key = normalizeAnketaNameKey(value);
  const keys = new Set<string>();
  if (!key) return keys;
  keys.add(key);
  const parts = key.split(" ").filter(Boolean);
  if (parts.length >= 2) keys.add(`${parts[0]} ${parts[1]}`);
  return keys;
};

export const expandAnketaNameKeySet = (values: Iterable<string>) => {
  const expanded = new Set<string>();
  for (const value of values) {
    for (const key of anketaNameKeyVariants(value)) expanded.add(key);
  }
  return expanded;
};

const normalizeNameKey = normalizeAnketaNameKey;

export const normalizeAnketaExternalIdKey = (value: unknown) => {
  const text = String(value ?? "").trim();
  const compact = text
    .replace(/[\s'`(),+-]/g, "")
    .replaceAll("[", "")
    .replaceAll("]", "");
  const numeric = compact.match(/^(\d+)(?:\.0+)?$/);
  return numeric?.[1] ?? text;
};

export type AnketaPersonnelMatch = {
  row: EjournalPreviewRow;
  summary: ReturnType<typeof buildPersonSummary>;
  matchBy: "externalId" | "rnokpp" | "nameBirth" | "name";
};

type PersonnelIndex = {
  byExternalId: Map<string, AnketaPersonnelMatch>;
  byRnokpp: Map<string, AnketaPersonnelMatch>;
  byNameBirth: Map<string, AnketaPersonnelMatch>;
  byName: Map<string, AnketaPersonnelMatch[]>;
};

let cachedIndex: PersonnelIndex | null = null;
let cachedIndexAt = 0;

const buildPersonnelIndex = (rows: EjournalPreviewRow[]): PersonnelIndex => {
  const byExternalId = new Map<string, AnketaPersonnelMatch>();
  const byRnokpp = new Map<string, AnketaPersonnelMatch>();
  const byNameBirth = new Map<string, AnketaPersonnelMatch>();
  const byName = new Map<string, AnketaPersonnelMatch[]>();

  for (const row of rows) {
    const summary = buildPersonSummary(row);
    const base = { row, summary, matchBy: "name" as const };
    const spreadsheetId = getPersonExternalId(row).trim();
    if (spreadsheetId) {
      const match = { ...base, matchBy: "externalId" as const };
      byExternalId.set(spreadsheetId, match);
      byExternalId.set(normalizeAnketaExternalIdKey(spreadsheetId), match);
    }
    if (summary.externalId) {
      const match = {
        ...base,
        matchBy: "externalId" as const,
      };
      byExternalId.set(summary.externalId, match);
      byExternalId.set(normalizeAnketaExternalIdKey(summary.externalId), match);
    }
    const rnokpp = summary.rnokpp.trim();
    if (rnokpp) {
      byRnokpp.set(rnokpp, { ...base, matchBy: "rnokpp" });
    }
    const nameKey = normalizeNameKey(summary.name);
    if (!nameKey || nameKey === "особа не вибрана") continue;
    const birthKey = normalizePersonBirthKey(summary.birthDate);
    if (birthKey) {
      byNameBirth.set(`${nameKey}|${birthKey}`, {
        ...base,
        matchBy: "nameBirth",
      });
    }
    const list = byName.get(nameKey) ?? [];
    list.push({ ...base, matchBy: "name" });
    byName.set(nameKey, list);
  }

  return { byExternalId, byRnokpp, byNameBirth, byName };
};

const loadMergedPersonnelRows = async (): Promise<EjournalPreviewRow[]> => {
  const [imports, latestRoster] = await Promise.all([
    fetchWithCache<BackendEjournalImport[]>({
      key: CacheKeys.ejournalImports,
      fetcher: () => api.listEjournalImports(),
      isChanged: jsonChanged,
    }).catch(() => []),
    fetchWithCache({
      key: CacheKeys.rosterLatest,
      fetcher: () => api.getLatestPersonnelRoster(),
      isChanged: jsonChanged,
    }).catch(() => null),
  ]);

  const rosterRows = mapRosterLatestToPreviewRows(latestRoster);
  const orderedImports = [...imports].sort(
    (left, right) =>
      new Date(left.updatedAt).getTime() - new Date(right.updatedAt).getTime(),
  );
  // Повна історія може містити десятки великих імпортів. Для доповнення
  // достатньо найближчих версій; найновіша картка все одно має пріоритет.
  const oosSheets = orderedImports
    .map((item) => findEjournalPersonnelSheet([item]))
    .filter((sheet): sheet is NonNullable<typeof sheet> => Boolean(sheet))
    .slice(-8);

  if (!oosSheets.length) {
    return rosterRows.length
      ? mergeRosterRowsIntoPreview({ rows: [] }, rosterRows)
      : [];
  }

  // Історичні ООС потрібні для «Виключені»: у найновішому ООС людини вже може не бути.
  // Імпорти йдуть від старого до нового, тому новіша картка перезаписує старішу.
  const previews = await Promise.all(
    oosSheets.map((sheet) => loadAllEjournalSheetRows(sheet).catch(() => null)),
  );
  const uniqueRows = new Map<string, EjournalPreviewRow>();
  let fallbackIndex = 0;
  for (const preview of previews) {
    for (const row of preview?.rows ?? []) {
      const summary = buildPersonSummary(row);
      const id = getPersonExternalId(row).trim();
      const nameKey = normalizeNameKey(summary.name);
      const birthKey = normalizePersonBirthKey(summary.birthDate);
      const key =
        (id && `id:${id}`) ||
        (nameKey && birthKey && `name-birth:${nameKey}|${birthKey}`) ||
        (nameKey && `name:${nameKey}`) ||
        `row:${fallbackIndex++}`;
      uniqueRows.set(key, row);
    }
  }
  const oosRows = [...uniqueRows.values()];
  return mergeRosterRowsIntoPreview({ rows: oosRows }, rosterRows);
};

export const loadPersonnelIndexForAnketa = async (options?: {
  force?: boolean;
}): Promise<PersonnelIndex> => {
  const now = Date.now();
  if (!options?.force && cachedIndex && now - cachedIndexAt < 60_000) {
    return cachedIndex;
  }

  const rows = await loadMergedPersonnelRows();
  const index = buildPersonnelIndex(rows);
  cachedIndex = index;
  cachedIndexAt = now;
  return index;
};

export const loadPersonnelRowsForAnketa = async (): Promise<EjournalPreviewRow[]> =>
  loadMergedPersonnelRows();

export type AnketaPersonnelMatchResult = {
  match: AnketaPersonnelMatch | null;
  ambiguous: AnketaPersonnelMatch[];
};

export const matchAnketaRowToPersonnelDetailed = (
  anketaRow: AnketaRow | null | undefined,
  index: PersonnelIndex | null,
): AnketaPersonnelMatchResult => {
  if (!anketaRow || !index) {
    return { match: null, ambiguous: [] };
  }

  const externalId = String(anketaRow.externalId ?? "").trim();
  if (externalId) {
    const hit = index.byExternalId.get(externalId);
    if (hit) return { match: hit, ambiguous: [] };
  }

  const rnokpp = String(anketaRow.rnokpp ?? "").trim();
  if (rnokpp) {
    const hit = index.byRnokpp.get(rnokpp);
    if (hit) return { match: hit, ambiguous: [] };
  }

  const nameKey = normalizeNameKey(anketaRow.fullName);
  if (!nameKey) return { match: null, ambiguous: [] };

  const birthKey = normalizePersonBirthKey(String(anketaRow.birthDate ?? ""));
  if (birthKey) {
    const hit = index.byNameBirth.get(`${nameKey}|${birthKey}`);
    if (hit) return { match: hit, ambiguous: [] };
  }

  const byName = index.byName.get(nameKey) ?? [];
  if (byName.length === 1) {
    return { match: byName[0] ?? null, ambiguous: [] };
  }
  if (byName.length > 1) {
    return { match: null, ambiguous: byName };
  }
  return { match: null, ambiguous: [] };
};

export const matchAnketaRowToPersonnel = (
  anketaRow: AnketaRow | null | undefined,
  index: PersonnelIndex | null,
): AnketaPersonnelMatch | null => matchAnketaRowToPersonnelDetailed(anketaRow, index).match;

export const matchLabel = (matchBy: AnketaPersonnelMatch["matchBy"]) => {
  switch (matchBy) {
    case "externalId":
      return "за ID";
    case "rnokpp":
      return "за РНОКПП";
    case "nameBirth":
      return "за ПІБ + датою народження";
    case "name":
      return "за ПІБ";
    default:
      return "зв’язок знайдено";
  }
};
