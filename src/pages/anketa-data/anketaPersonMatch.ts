import type { BackendEjournalImport } from "../../api";
import { CacheKeys, readDataCache } from "../../data/idbDataCache";
import type {
  DbPreviewState,
  EjournalPreviewRow,
} from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  findEjournalPersonnelSheet,
  getPersonExternalId,
  normalizePersonBirthKey,
  sheetRowsCacheKey,
} from "../personnel/personnelUtils";
import type { AnketaRow } from "./anketaSheet";

export const normalizeAnketaNameKey = (value: unknown) =>
  String(value ?? "")
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const normalizeNameKey = normalizeAnketaNameKey;

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
      byExternalId.set(spreadsheetId, { ...base, matchBy: "externalId" });
    }
    if (summary.externalId) {
      byExternalId.set(summary.externalId, {
        ...base,
        matchBy: "externalId",
      });
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

export const loadPersonnelIndexForAnketa = async (options?: {
  force?: boolean;
}): Promise<PersonnelIndex> => {
  const now = Date.now();
  if (!options?.force && cachedIndex && now - cachedIndexAt < 60_000) {
    return cachedIndex;
  }

  const imports = await readDataCache<BackendEjournalImport[]>(
    CacheKeys.ejournalImports,
  );
  if (!imports?.length) {
    cachedIndex = {
      byExternalId: new Map(),
      byRnokpp: new Map(),
      byNameBirth: new Map(),
      byName: new Map(),
    };
    cachedIndexAt = now;
    return cachedIndex;
  }

  const sheet = findEjournalPersonnelSheet(imports);
  if (!sheet) {
    cachedIndex = {
      byExternalId: new Map(),
      byRnokpp: new Map(),
      byNameBirth: new Map(),
      byName: new Map(),
    };
    cachedIndexAt = now;
    return cachedIndex;
  }

  const preview = await readDataCache<DbPreviewState>(sheetRowsCacheKey(sheet));
  const index = buildPersonnelIndex(preview?.rows ?? []);
  cachedIndex = index;
  cachedIndexAt = now;
  return index;
};

export const matchAnketaRowToPersonnel = (
  anketaRow: AnketaRow | null | undefined,
  index: PersonnelIndex | null,
): AnketaPersonnelMatch | null => {
  if (!anketaRow || !index) return null;

  const externalId = String(anketaRow.externalId ?? "").trim();
  if (externalId) {
    const hit = index.byExternalId.get(externalId);
    if (hit) return hit;
  }

  const rnokpp = String(anketaRow.rnokpp ?? "").trim();
  if (rnokpp) {
    const hit = index.byRnokpp.get(rnokpp);
    if (hit) return hit;
  }

  const nameKey = normalizeNameKey(anketaRow.fullName);
  if (!nameKey) return null;

  const birthKey = normalizePersonBirthKey(String(anketaRow.birthDate ?? ""));
  if (birthKey) {
    const hit = index.byNameBirth.get(`${nameKey}|${birthKey}`);
    if (hit) return hit;
  }

  const byName = index.byName.get(nameKey);
  if (byName?.length === 1) return byName[0] ?? null;
  return null;
};

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
