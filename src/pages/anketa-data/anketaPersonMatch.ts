import { api, type BackendEjournalImport } from "../../api";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  readDataCache,
} from "../../data/idbDataCache";
import type {
  EjournalPreviewRow,
} from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  findEjournalPersonnelSheet,
  getPersonExternalId,
  isLikelyPersonnelRow,
  loadAllEjournalSheetRows,
  normalizePersonBirthKey,
} from "../personnel/personnelUtils";
import { mergeRosterRowsIntoPreview } from "../personnel/personnelRosterMerge";
import { mapRosterLatestToPreviewRows, readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
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
  /** Прізвище + імʼя (без по батькові) — для підказок, не для автозвʼязку. */
  byShortName: Map<string, AnketaPersonnelMatch[]>;
};

export type AnketaPersonnelIndex = PersonnelIndex;

export const anketaPersonnelNamesMatch = (
  anketaFullName: unknown,
  personnelName: unknown,
) => {
  const left = normalizeAnketaNameKey(anketaFullName);
  const right = normalizeAnketaNameKey(personnelName);
  return Boolean(left && right && left === right);
};

const shortNameKeyFromFull = (nameKey: string) => {
  const parts = nameKey.split(" ").filter(Boolean);
  return parts.length >= 2 ? `${parts[0]} ${parts[1]}` : nameKey;
};

let cachedIndex: PersonnelIndex | null = null;
let cachedIndexAt = 0;

const pushNameList = (
  map: Map<string, AnketaPersonnelMatch[]>,
  key: string,
  match: AnketaPersonnelMatch,
) => {
  if (!key) return;
  const list = map.get(key) ?? [];
  list.push(match);
  map.set(key, list);
};

const buildPersonnelIndex = (rows: EjournalPreviewRow[]): PersonnelIndex => {
  const byExternalId = new Map<string, AnketaPersonnelMatch>();
  const byRnokpp = new Map<string, AnketaPersonnelMatch>();
  const byNameBirth = new Map<string, AnketaPersonnelMatch>();
  const byName = new Map<string, AnketaPersonnelMatch[]>();
  const byShortName = new Map<string, AnketaPersonnelMatch[]>();

  for (const row of rows) {
    if (!isLikelyPersonnelRow(row)) continue;
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
    const rnokppDigits = rnokpp.replace(/\D/g, "");
    if (rnokpp) {
      const match = { ...base, matchBy: "rnokpp" as const };
      byRnokpp.set(rnokpp, match);
      if (rnokppDigits.length >= 8) {
        byRnokpp.set(rnokppDigits, match);
      }
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
    pushNameList(byName, nameKey, { ...base, matchBy: "name" });
    pushNameList(byShortName, shortNameKeyFromFull(nameKey), {
      ...base,
      matchBy: "name",
    });
  }

  return { byExternalId, byRnokpp, byNameBirth, byName, byShortName };
};

const appendAnketaCreatedRows = (
  rows: EjournalPreviewRow[],
  created: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  if (!created.length) return rows;

  const used = new Set<string>();
  for (const row of rows) {
    const summary = buildPersonSummary(row);
    const name = normalizeNameKey(summary.name);
    const birth = normalizePersonBirthKey(summary.birthDate);
    const id = normalizeAnketaExternalIdKey(getPersonExternalId(row) || summary.externalId);
    if (id) used.add(`id:${id}`);
    if (name && birth) used.add(`nb:${name}|${birth}`);
    if (name) used.add(`name:${name}`);
  }

  const extras: EjournalPreviewRow[] = [];
  for (const row of created) {
    if (!isLikelyPersonnelRow(row)) continue;
    const summary = buildPersonSummary(row);
    const name = normalizeNameKey(summary.name);
    const birth = normalizePersonBirthKey(summary.birthDate);
    const id = normalizeAnketaExternalIdKey(getPersonExternalId(row) || summary.externalId);
    const marks = [
      id ? `id:${id}` : "",
      name && birth ? `nb:${name}|${birth}` : "",
      name ? `name:${name}` : "",
    ].filter(Boolean);
    if (marks.some((mark) => used.has(mark))) continue;
    extras.push(row);
    marks.forEach((mark) => used.add(mark));
  }

  return extras.length ? [...rows, ...extras] : rows;
};

/** Той самий склад, що на сторінці «Особовий склад»: останній ООС + штатка + додані з анкет. */
const loadMergedPersonnelRows = async (): Promise<EjournalPreviewRow[]> => {
  const [imports, latestRoster, created] = await Promise.all([
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
    readDataCache<EjournalPreviewRow[]>(CacheKeys.anketaCreatedPersonnel).then(
      (stored) => (Array.isArray(stored) ? stored : []),
    ),
  ]);

  const rosterRows = mapRosterLatestToPreviewRows(latestRoster);
  const sheet = findEjournalPersonnelSheet(imports);
  const preview = sheet
    ? await loadAllEjournalSheetRows(sheet).catch(() => null)
    : null;

  if (!preview?.rows?.length && !rosterRows.length && !created.length) {
    return [];
  }

  const withRoster = mergeRosterRowsIntoPreview(
    { rows: preview?.rows ?? [] },
    rosterRows,
  );
  return appendAnketaCreatedRows(withRoster, created);
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

/** Швидкий пошук рядка ООС для зіставлення зі Штаткою (без лінійного .find). */
export const resolvePersonnelRowForStaffRoster = (
  rosterRow: EjournalPreviewRow,
  index: PersonnelIndex,
): EjournalPreviewRow => {
  const externalId = getPersonExternalId(rosterRow).trim();
  if (externalId) {
    const byId =
      index.byExternalId.get(externalId) ??
      index.byExternalId.get(normalizeAnketaExternalIdKey(externalId));
    if (byId) return byId.row;
  }

  const nameKey = normalizeAnketaNameKey(readRosterColumnValue(rosterRow, 14));
  if (nameKey) {
    const matches = index.byName.get(nameKey);
    if (matches?.length === 1) return matches[0]!.row;
  }

  return rosterRow;
};

export type AnketaPersonnelMatchResult = {
  match: AnketaPersonnelMatch | null;
  ambiguous: AnketaPersonnelMatch[];
  /** Схожі за прізвищем + імʼям (інше по батькові / написання). */
  similar: AnketaPersonnelMatch[];
};

const dedupePersonnelMatches = (items: AnketaPersonnelMatch[]) => {
  const seen = new Set<string>();
  const result: AnketaPersonnelMatch[] = [];
  for (const item of items) {
    const key =
      String(item.row.__dbRowId ?? "") ||
      item.summary.externalId ||
      normalizeNameKey(item.summary.name);
    if (!key || seen.has(key)) continue;
    seen.add(key);
    result.push(item);
  }
  return result;
};

const personnelExternalIdKeys = (match: AnketaPersonnelMatch) => {
  const keys = new Set<string>();
  const spreadsheetId = getPersonExternalId(match.row).trim();
  if (spreadsheetId) {
    keys.add(spreadsheetId);
    keys.add(normalizeAnketaExternalIdKey(spreadsheetId));
  }
  const summaryId = String(match.summary.externalId ?? "").trim();
  if (summaryId) {
    keys.add(summaryId);
    keys.add(normalizeAnketaExternalIdKey(summaryId));
  }
  return keys;
};

const disambiguateNameMatches = (
  anketaRow: AnketaRow,
  candidates: AnketaPersonnelMatch[],
): AnketaPersonnelMatchResult => {
  const unique = dedupePersonnelMatches(candidates);
  if (!unique.length) {
    return { match: null, ambiguous: [], similar: [] };
  }
  if (unique.length === 1) {
    return {
      match: { ...unique[0]!, matchBy: "name" },
      ambiguous: [],
      similar: [],
    };
  }

  const externalId = normalizeAnketaExternalIdKey(anketaRow.externalId);
  const birthKey = normalizePersonBirthKey(String(anketaRow.birthDate ?? ""));

  const matchesExternalId = (item: AnketaPersonnelMatch) =>
    externalId ? personnelExternalIdKeys(item).has(externalId) : false;

  const matchesBirth = (item: AnketaPersonnelMatch) =>
    birthKey
      ? normalizePersonBirthKey(item.summary.birthDate) === birthKey
      : false;

  if (externalId && birthKey) {
    const byBoth = unique.filter(
      (item) => matchesExternalId(item) && matchesBirth(item),
    );
    if (byBoth.length === 1) {
      return {
        match: { ...byBoth[0]!, matchBy: "nameBirth" },
        ambiguous: [],
        similar: [],
      };
    }
    return {
      match: null,
      ambiguous: byBoth.length > 1 ? byBoth : unique,
      similar: [],
    };
  }

  if (externalId) {
    const byId = unique.filter(matchesExternalId);
    if (byId.length === 1) {
      return {
        match: { ...byId[0]!, matchBy: "externalId" },
        ambiguous: [],
        similar: [],
      };
    }
    if (byId.length > 1) {
      return { match: null, ambiguous: byId, similar: [] };
    }
  }

  if (birthKey) {
    const byBirth = unique.filter(matchesBirth);
    if (byBirth.length === 1) {
      return {
        match: { ...byBirth[0]!, matchBy: "nameBirth" },
        ambiguous: [],
        similar: [],
      };
    }
    if (byBirth.length > 1) {
      return { match: null, ambiguous: byBirth, similar: [] };
    }
  }

  return { match: null, ambiguous: unique, similar: [] };
};

export const matchAnketaRowToPersonnelDetailed = (
  anketaRow: AnketaRow | null | undefined,
  index: PersonnelIndex | null,
): AnketaPersonnelMatchResult => {
  if (!anketaRow || !index) {
    return { match: null, ambiguous: [], similar: [] };
  }

  const nameKey = normalizeNameKey(anketaRow.fullName);
  const shortKey = nameKey ? shortNameKeyFromFull(nameKey) : "";
  const similarFallback =
    shortKey && shortKey !== nameKey
      ? dedupePersonnelMatches(index.byShortName.get(shortKey) ?? [])
      : [];

  if (nameKey) {
    const byName = index.byName.get(nameKey) ?? [];
    if (byName.length) {
      const byNameResult = disambiguateNameMatches(anketaRow, byName);
      if (byNameResult.match || byNameResult.ambiguous.length) {
        return byNameResult;
      }
    }

    const birthKey = normalizePersonBirthKey(String(anketaRow.birthDate ?? ""));
    if (birthKey) {
      const hit = index.byNameBirth.get(`${nameKey}|${birthKey}`);
      if (hit) {
        return { match: hit, ambiguous: [], similar: [] };
      }
    }
  }

  const rnokpp = String(anketaRow.rnokpp ?? "")
    .replace(/\D/g, "")
    .trim();
  if (rnokpp.length >= 8) {
    const hit = index.byRnokpp.get(rnokpp);
    if (hit && anketaPersonnelNamesMatch(anketaRow.fullName, hit.summary.name)) {
      return { match: hit, ambiguous: [], similar: [] };
    }
  }

  return { match: null, ambiguous: [], similar: similarFallback };
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
