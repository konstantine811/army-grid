import type {
  BackendPersonDocument,
  BackendPersonQuestionnaire,
  BackendPersonQuestionnaireMeta,
} from "../../api";
import { api } from "../../api";
import { sanitizeFileName } from "../../shared/browserExport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import {
  buildPersonIdentityFingerprint,
  buildPersonSummary,
  buildQuestionnaireExportFileName,
  collectPersonExternalIdCandidates,
  getPersonDisplayName,
  getPersonFieldValue,
  isLikelyPersonnelRow,
  normalizePersonBirthKey,
  resolvePersonBirthDate,
  resolvePersonCallSign,
  resolvePersonIdentityKey,
} from "./personnelUtils";

export type PersonAttachmentLookupHints = {
  anketaExternalId?: string;
  anketaFullName?: string;
  anketaBirthDate?: string;
};

const normalizeAttachmentNameKey = (value: unknown) =>
  String(value ?? "")
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const pushLegacyAttachmentLookupIds = (
  ids: Set<string>,
  name: string,
  birthDate = "",
  callSign = "",
  includeLooseKeys = false,
) => {
  const nameKey = normalizeAttachmentNameKey(name);
  if (!nameKey || nameKey === "особа не вибрана") return;
  const birthKey = normalizePersonBirthKey(birthDate);
  if (birthKey) ids.add(`name-birth:${nameKey}:${birthKey}`);
  // Без дати «name:» / «name-call:» спільні для однофамільців (два Шевченки).
  if (!birthKey || includeLooseKeys) {
    ids.add(`name:${nameKey}`);
    const callSignKey = normalizeAttachmentNameKey(callSign);
    if (callSignKey) {
      ids.add(`name-call:${nameKey}:${callSignKey}`);
    }
  }
};

const hasQuestionnaireRecord = (
  questionnaire: BackendPersonQuestionnaire | null | undefined,
) => Boolean(questionnaire?.personExternalId || questionnaire?.id);

const FILE_NAME_NOISE = new Set([
  "pdf",
  "анкета",
  "анкети",
  "questionnaire",
  "опитувальник",
  "скан",
  "scan",
]);

const nameTokensOf = (value: string) =>
  normalizeAttachmentNameKey(String(value ?? "").replace(/\.pdf$/i, ""))
    .split(" ")
    .filter((token) => token.length > 1 && !FILE_NAME_NOISE.has(token));

const nameTokensMatchFileName = (fullName: string, fileName: string) => {
  const personTokens = nameTokensOf(fullName);
  const fileTokens = nameTokensOf(fileName);
  if (personTokens.length < 2 || fileTokens.length < 2) return false;
  // Прізвище + імʼя в тому ж порядку. По батькові в файлі може не бути.
  const shared = Math.min(personTokens.length, fileTokens.length);
  for (let index = 0; index < shared; index += 1) {
    if (personTokens[index] !== fileTokens[index]) return false;
  }
  return true;
};

/** Чи PDF-анкета належить цій людині (за назвою файлу). */
export const questionnaireFileMatchesPerson = (
  fileName: string | null | undefined,
  names: Array<string | null | undefined>,
) => {
  const normalizedNames = names.map((name) => String(name ?? "").trim()).filter(Boolean);
  if (!normalizedNames.length) return true;
  const text = String(fileName ?? "").trim();
  if (!text || /^questionnaire\.pdf$/i.test(text)) return true;
  return normalizedNames.some((name) => nameTokensMatchFileName(name, text));
};

/** Чи є PDF-анкета для рядка ООС / штатки (як у картці персоналу). */
export const rowHasListedQuestionnaire = (
  row: EjournalPreviewRow | null,
  items: BackendPersonQuestionnaireMeta[],
  hints?: PersonAttachmentLookupHints,
) => {
  if (!row || !items.length) return false;

  const lookupIds = new Set(collectPersonAttachmentLookupIds(row, hints));
  for (const meta of items) {
    const id = String(meta.personExternalId ?? "").trim();
    if (id && lookupIds.has(id)) return true;
  }

  const names = [
    getPersonDisplayName(row),
    readRosterColumnValue(row, 14),
    String(hints?.anketaFullName ?? "").trim(),
  ].filter(Boolean);

  for (const meta of items) {
    const fileName = String(meta.fileName ?? "").trim();
    if (!fileName) continue;
    if (names.some((name) => nameTokensMatchFileName(name, fileName))) {
      return true;
    }
  }

  return false;
};

let questionnaireIndexCache:
  | { at: number; items: Awaited<ReturnType<typeof api.listPersonQuestionnaires>> }
  | null = null;

const loadQuestionnaireIndex = async () => {
  const now = Date.now();
  if (questionnaireIndexCache && now - questionnaireIndexCache.at < 30_000) {
    return questionnaireIndexCache.items;
  }
  const items = await api.listPersonQuestionnaires().catch(() => []);
  questionnaireIndexCache = { at: now, items };
  return items;
};

const resolveQuestionnaireViaIndex = async (
  lookupIds: string[],
  names: string[],
  allowFileNameFallback = true,
) => {
  const items = await loadQuestionnaireIndex();
  const lookupSet = new Set(lookupIds);

  for (const meta of items) {
    const id = meta.personExternalId?.trim();
    if (!id || !lookupSet.has(id)) continue;
    const questionnaire = await api.getPersonQuestionnaire(id).catch(() => null);
    if (hasQuestionnaireRecord(questionnaire)) {
      return {
        questionnaire,
        resolvedExternalId: questionnaire!.personExternalId?.trim() || id,
      };
    }
  }

  if (!allowFileNameFallback) return null;

  const fileHits = items.filter((meta) =>
    questionnaireFileMatchesPerson(meta.fileName, names),
  );
  const uniqueFileHit =
    fileHits.length === 1
      ? fileHits[0]
      : fileHits.find((meta) => {
          const id = meta.personExternalId?.trim();
          return Boolean(id && lookupSet.has(id));
        });
  if (uniqueFileHit?.personExternalId?.trim()) {
    const id = uniqueFileHit.personExternalId.trim();
    const questionnaire = await api.getPersonQuestionnaire(id).catch(() => null);
    if (hasQuestionnaireRecord(questionnaire)) {
      return {
        questionnaire,
        resolvedExternalId: questionnaire!.personExternalId?.trim() || id,
      };
    }
  }

  return null;
};

const collectNameLookupVariants = (name: string) => {
  const trimmed = String(name ?? "").trim();
  const tokens = normalizeAttachmentNameKey(trimmed).split(" ").filter(Boolean);
  const variants: string[] = [];
  const push = (value: string) => {
    const text = value.trim();
    if (!text || variants.includes(text)) return;
    variants.push(text);
  };
  if (trimmed) push(trimmed);
  if (tokens.length >= 2) push(tokens.slice(0, 2).join(" "));
  return variants;
};

const pushNameAttachmentLookupIds = (
  ids: Set<string>,
  name: string,
  birthDate = "",
  callSign = "",
  includeLooseKeys = false,
) => {
  for (const variant of collectNameLookupVariants(name)) {
    pushLegacyAttachmentLookupIds(
      ids,
      variant,
      birthDate,
      callSign,
      includeLooseKeys,
    );
    const withBirth = buildPersonIdentityFingerprint(variant, birthDate);
    if (withBirth) ids.add(withBirth);
    if (!birthDate || includeLooseKeys) {
      const withoutBirth = buildPersonIdentityFingerprint(variant);
      if (withoutBirth) ids.add(withoutBirth);
      const withCallSign = buildPersonIdentityFingerprint(variant, "", callSign);
      if (withCallSign) ids.add(withCallSign);
    }
  }
};

const lookupIdsByRow = new WeakMap<EjournalPreviewRow, string[]>();

const hasAttachmentLookupHints = (hints?: PersonAttachmentLookupHints) =>
  Boolean(
    hints?.anketaExternalId?.trim() ||
      hints?.anketaFullName?.trim() ||
      hints?.anketaBirthDate?.trim(),
  );

export type CollectAttachmentLookupOptions = {
  /** Also wipe name-only / позивний keys (видалення анкети в однофамільців). */
  includeLooseKeys?: boolean;
};

/** All plausible keys under which photo / questionnaire may have been saved. */
export const collectPersonAttachmentLookupIds = (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
  options?: CollectAttachmentLookupOptions,
) => {
  const includeLooseKeys = Boolean(options?.includeLooseKeys);
  if (row && !hasAttachmentLookupHints(hints) && !includeLooseKeys) {
    const cached = lookupIdsByRow.get(row);
    if (cached) return cached;
  }

  const ids = new Set<string>();
  for (const candidate of collectPersonExternalIdCandidates(row)) {
    ids.add(candidate);
  }

  const push = (value: unknown) => {
    const text = String(value ?? "").trim();
    if (text && text !== "0") ids.add(text);
  };

  push(hints?.anketaExternalId);

  const personnelName = getPersonDisplayName(row);
  const anketaName = String(hints?.anketaFullName ?? "").trim();
  const personnelBirth = resolvePersonBirthDate(row);
  const anketaBirth = String(hints?.anketaBirthDate ?? "").trim();
  const callSign = resolvePersonCallSign(row);

  for (const name of [personnelName, anketaName]) {
    if (!name) continue;
    const birth = name === personnelName ? personnelBirth : anketaBirth;
    const variantCallSign = name === personnelName ? callSign : "";
    pushNameAttachmentLookupIds(
      ids,
      name,
      birth,
      variantCallSign,
      includeLooseKeys,
    );
    if (includeLooseKeys) {
      const nameKey = normalizeAttachmentNameKey(name);
      if (nameKey) {
        ids.add(`roster:${nameKey}`);
        ids.add(`roster:${name.trim()}`);
      }
    }
  }

  const primary = resolvePersonIdentityKey(row);
  if (primary) ids.add(primary);

  const collected = [...ids];
  if (row && !hasAttachmentLookupHints(hints) && !includeLooseKeys) {
    lookupIdsByRow.set(row, collected);
  }
  return collected;
};

export const loadPersonPhotoForRow = async (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
) => {
  const fallback = resolvePersonIdentityKey(row);
  for (const id of collectPersonAttachmentLookupIds(row, hints)) {
    try {
      const photo = await api.getPersonPhoto(id);
      const photoData = photo?.photoData?.trim() || "";
      if (photoData) {
        return { photoData, resolvedExternalId: id };
      }
    } catch {
      /* try next key */
    }
  }
  return { photoData: "", resolvedExternalId: fallback };
};

export type LoadPersonQuestionnaireOptions = {
  /** Два однофамільці — не підставляти PDF лише за назвою файлу. */
  nameIsAmbiguous?: boolean;
};

export const loadPersonQuestionnaireForRow = async (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
  options?: LoadPersonQuestionnaireOptions,
): Promise<{
  questionnaire: BackendPersonQuestionnaire | null;
  resolvedExternalId: string;
}> => {
  const fallback = resolvePersonIdentityKey(row);
  const lookupIds = collectPersonAttachmentLookupIds(row, hints);
  const expectedNames = [
    getPersonDisplayName(row),
    String(hints?.anketaFullName ?? "").trim(),
  ].filter(Boolean);
  const allowFileNameFallback = !options?.nameIsAmbiguous;

  if (fallback) {
    try {
      const questionnaire = await api.getPersonQuestionnaire(fallback);
      if (hasQuestionnaireRecord(questionnaire)) {
        return {
          questionnaire,
          resolvedExternalId: questionnaire!.personExternalId?.trim() || fallback,
        };
      }
    } catch {
      /* fall through to the questionnaire index */
    }
  }

  const indexed = await resolveQuestionnaireViaIndex(
    lookupIds,
    expectedNames,
    allowFileNameFallback,
  );
  if (indexed) return indexed;

  return { questionnaire: null, resolvedExternalId: fallback };
};

/** PDF is stored under any candidate / legacy identity of this row. */
export const buildQuestionnairePresenceMap = (
  rows: EjournalPreviewRow[],
  items: Array<{ personExternalId?: string | null; fileName?: string | null }>,
) => {
  const stored = new Set(
    items
      .map((item) => String(item.personExternalId ?? "").trim())
      .filter(Boolean),
  );
  const map: Record<string, true> = {};
  for (const id of stored) map[id] = true;
  const unmatched: EjournalPreviewRow[] = [];
  for (const row of rows) {
    if (!isLikelyPersonnelRow(row)) continue;
    const current = resolvePersonIdentityKey(row);
    if (!current) continue;
    if (stored.has(current)) {
      map[current] = true;
      continue;
    }
    if (collectPersonAttachmentLookupIds(row).some((id) => stored.has(id))) {
      map[current] = true;
      continue;
    }
    unmatched.push(row);
  }

  if (unmatched.length && items.length) {
    const rowsByShortName = new Map<string, EjournalPreviewRow[]>();
    for (const row of unmatched) {
      const shortKey = nameTokensOf(getPersonDisplayName(row)).slice(0, 2).join(" ");
      if (!shortKey) continue;
      const list = rowsByShortName.get(shortKey) ?? [];
      list.push(row);
      rowsByShortName.set(shortKey, list);
    }
    for (const item of items) {
      const fileName = String(item.fileName ?? "").trim();
      if (!fileName) continue;
      const shortKey = nameTokensOf(fileName).slice(0, 2).join(" ");
      const candidates = rowsByShortName.get(shortKey) ?? [];
      const hits = candidates.filter((row) =>
        questionnaireFileMatchesPerson(fileName, [getPersonDisplayName(row)]),
      );
      if (hits.length !== 1) continue;
      const current = resolvePersonIdentityKey(hits[0]);
      if (current) map[current] = true;
    }
  }
  return map;
};

export type OrphanAttachmentIdentity = {
  nameKey: string;
  birthKey?: string;
  callKey?: string;
};

export const parseOrphanAttachmentIdentityId = (
  id: string,
): OrphanAttachmentIdentity | null => {
  const raw = String(id ?? "").trim();
  if (!raw) return null;

  const takeNameBirth = (body: string): OrphanAttachmentIdentity | null => {
    const birthMatch = body.match(/:(\d{4}-\d{2}-\d{2})$/);
    if (birthMatch) {
      const nameKey = body.slice(0, -birthMatch[0].length).trim();
      return nameKey
        ? { nameKey, birthKey: birthMatch[1] }
        : null;
    }
    const callIdx = body.lastIndexOf(":c:");
    if (callIdx > 0) {
      const nameKey = body.slice(0, callIdx).trim();
      const callKey = body.slice(callIdx + 3).trim();
      return nameKey ? { nameKey, callKey } : null;
    }
    const nameKey = body.trim();
    return nameKey ? { nameKey } : null;
  };

  if (raw.startsWith("p:")) return takeNameBirth(raw.slice(2));
  if (raw.startsWith("name-birth:")) return takeNameBirth(raw.slice("name-birth:".length));
  if (raw.startsWith("name-call:")) {
    const body = raw.slice("name-call:".length);
    const sep = body.lastIndexOf(":");
    if (sep <= 0) return null;
    const nameKey = body.slice(0, sep).trim();
    const callKey = body.slice(sep + 1).trim();
    return nameKey ? { nameKey, callKey } : null;
  }
  if (raw.startsWith("name:")) {
    const nameKey = raw.slice(5).trim();
    return nameKey ? { nameKey } : null;
  }
  if (raw.startsWith("roster:")) {
    const rest = raw.slice("roster:".length).trim();
    if (!rest) return null;
    if (/^[a-z0-9_-]+$/i.test(rest) && !/[а-яіїєґ]/i.test(rest)) return null;
    const nameKey = normalizeAttachmentNameKey(rest);
    return nameKey ? { nameKey } : null;
  }
  return null;
};

export const personNameMatchesOrphanNameKey = (
  personName: string,
  orphanNameKey: string,
) => {
  const personKey = normalizeAttachmentNameKey(personName);
  const orphanKey = normalizeAttachmentNameKey(orphanNameKey);
  if (!personKey || !orphanKey) return false;
  if (personKey === orphanKey) return true;
  const personTokens = personKey.split(" ").filter(Boolean);
  const orphanTokens = orphanKey.split(" ").filter(Boolean);
  if (personTokens.length < 2 || orphanTokens.length < 2) return false;
  const shared = Math.min(personTokens.length, orphanTokens.length);
  for (let index = 0; index < shared; index += 1) {
    if (personTokens[index] !== orphanTokens[index]) return false;
  }
  return true;
};

const shortAttachmentNameKey = (nameKey: string) => {
  const tokens = normalizeAttachmentNameKey(nameKey).split(" ").filter(Boolean);
  return tokens.length >= 2 ? tokens.slice(0, 2).join(" ") : "";
};

const pushIndexedRows = (
  map: Map<string, EjournalPreviewRow[]>,
  key: string,
  row: EjournalPreviewRow,
) => {
  if (!key) return;
  const list = map.get(key);
  if (list) list.push(row);
  else map.set(key, [row]);
};

const pickUniqueIndexedRow = (rows: EjournalPreviewRow[] | undefined) =>
  rows && rows.length === 1 ? rows[0] : null;

type AttachmentRowIndex = {
  uniqueById: Map<string, EjournalPreviewRow>;
  byNameKey: Map<string, EjournalPreviewRow[]>;
  byNameBirthKey: Map<string, EjournalPreviewRow[]>;
};

const buildAttachmentRowIndex = (
  rows: EjournalPreviewRow[],
): AttachmentRowIndex => {
  const uniqueById = new Map<string, EjournalPreviewRow>();
  const ambiguousIds = new Set<string>();
  const byNameKey = new Map<string, EjournalPreviewRow[]>();
  const byNameBirthKey = new Map<string, EjournalPreviewRow[]>();

  for (const row of rows) {
    if (!isLikelyPersonnelRow(row)) continue;
    for (const id of collectPersonAttachmentLookupIds(row)) {
      if (ambiguousIds.has(id)) continue;
      const existing = uniqueById.get(id);
      if (!existing) {
        uniqueById.set(id, row);
      } else if (existing !== row) {
        uniqueById.delete(id);
        ambiguousIds.add(id);
      }
    }

    const nameKey = normalizeAttachmentNameKey(getPersonDisplayName(row));
    const shortKey = shortAttachmentNameKey(nameKey);
    const birthKey = normalizePersonBirthKey(resolvePersonBirthDate(row));
    pushIndexedRows(byNameKey, nameKey, row);
    pushIndexedRows(byNameKey, shortKey, row);
    if (birthKey) {
      pushIndexedRows(byNameBirthKey, `${nameKey}|${birthKey}`, row);
      if (shortKey) {
        pushIndexedRows(byNameBirthKey, `${shortKey}|${birthKey}`, row);
      }
    }
  }

  return { uniqueById, byNameKey, byNameBirthKey };
};

const matchOrphanIdWithIndex = (
  orphanId: string,
  index: AttachmentRowIndex,
) => {
  const unique = index.uniqueById.get(orphanId);
  if (unique) return unique;

  const parsed = parseOrphanAttachmentIdentityId(orphanId);
  if (!parsed) return null;
  const nameKey = normalizeAttachmentNameKey(parsed.nameKey);
  const shortKey = shortAttachmentNameKey(nameKey);
  if (parsed.birthKey) {
    return (
      pickUniqueIndexedRow(
        index.byNameBirthKey.get(`${nameKey}|${parsed.birthKey}`),
      ) ||
      pickUniqueIndexedRow(
        index.byNameBirthKey.get(`${shortKey}|${parsed.birthKey}`),
      )
    );
  }
  return (
    pickUniqueIndexedRow(index.byNameKey.get(nameKey)) ||
    pickUniqueIndexedRow(index.byNameKey.get(shortKey))
  );
};

export const matchOrphanIdToPersonnelRow = (
  orphanId: string,
  rows: EjournalPreviewRow[],
) => matchOrphanIdWithIndex(orphanId, buildAttachmentRowIndex(rows));

export type PersonAttachmentMigrationPair = {
  name: string;
  fromExternalId: string;
  toExternalId: string;
};

export const dedupePersonAttachmentMigrationPairs = (
  pairs: PersonAttachmentMigrationPair[],
) => {
  const seen = new Set<string>();
  return pairs.filter((pair) => {
    const fromExternalId = pair.fromExternalId.trim();
    const toExternalId = pair.toExternalId.trim();
    if (!fromExternalId || !toExternalId || fromExternalId === toExternalId) {
      return false;
    }
    const key = `${fromExternalId}=>${toExternalId}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

export const buildSelfAttachmentMigrationPairs = (
  rows: EjournalPreviewRow[],
) => {
  const pairs: PersonAttachmentMigrationPair[] = [];
  for (const row of rows) {
    if (!isLikelyPersonnelRow(row)) continue;
    const summary = buildPersonSummary(row);
    const toExternalId = summary.externalId;
    if (!toExternalId) continue;
    for (const fromExternalId of collectPersonExternalIdCandidates(row)) {
      pairs.push({
        name: summary.name,
        fromExternalId,
        toExternalId,
      });
    }
  }
  return dedupePersonAttachmentMigrationPairs(pairs);
};

export const buildOrphanAttachmentMigrationPairs = (
  rows: EjournalPreviewRow[],
  orphanIds: Set<string>,
  questionnaireMetas: Array<{
    personExternalId?: string | null;
    fileName?: string | null;
  }> = [],
) => {
  if (!orphanIds.size) return [];

  const index = buildAttachmentRowIndex(rows);
  const pairs: PersonAttachmentMigrationPair[] = [];
  const usedFrom = new Set<string>();
  const usedTo = new Set<string>();
  const pushPair = (row: EjournalPreviewRow, fromExternalId: string) => {
    const toExternalId = resolvePersonIdentityKey(row);
    if (!toExternalId || toExternalId === fromExternalId || usedFrom.has(fromExternalId)) {
      return;
    }
    usedFrom.add(fromExternalId);
    usedTo.add(toExternalId);
    pairs.push({
      name:
        getPersonFieldValue(row, ["прізвище"]) ||
        getPersonFieldValue(row, ["піб"]),
      fromExternalId,
      toExternalId,
    });
  };

  for (const fromExternalId of orphanIds) {
    const row = matchOrphanIdWithIndex(fromExternalId, index);
    if (row) pushPair(row, fromExternalId);
  }

  if (questionnaireMetas.length) {
    const personnelRows = rows.filter(isLikelyPersonnelRow);
    for (const meta of questionnaireMetas) {
      const fromExternalId = String(meta.personExternalId ?? "").trim();
      const fileName = String(meta.fileName ?? "").trim();
      if (!fromExternalId || !fileName || !orphanIds.has(fromExternalId)) continue;
      if (usedFrom.has(fromExternalId)) continue;
      const hits = personnelRows.filter(
        (row) =>
          !usedTo.has(resolvePersonIdentityKey(row)) &&
          questionnaireFileMatchesPerson(fileName, [getPersonDisplayName(row)]),
      );
      if (hits.length === 1) pushPair(hits[0], fromExternalId);
    }
  }
  return dedupePersonAttachmentMigrationPairs(pairs);
};

export type PersonAttachmentMigrationOptions = {
  includeDocuments?: boolean;
  photos?: Array<{ personExternalId: string; photoData: string }>;
  questionnaires?: Array<{ personExternalId: string }>;
  documents?: BackendPersonDocument[];
};

const migrateTargetedPersonAttachments = async (
  pairs: PersonAttachmentMigrationPair[],
  includeDocuments: boolean,
) => {
  let migrated = 0;
  for (const pair of pairs) {
    try {
      const [oldPhoto, newPhoto] = await Promise.all([
        api.getPersonPhoto(pair.fromExternalId).catch(() => null),
        api.getPersonPhoto(pair.toExternalId).catch(() => null),
      ]);
      if (oldPhoto?.photoData && !newPhoto?.photoData) {
        await api.upsertPersonPhoto(pair.toExternalId, {
          photoData: oldPhoto.photoData,
          fileName: `${pair.name}.jpg`,
        });
        migrated += 1;
      }

      const [oldQuestionnaire, newQuestionnaire] = await Promise.all([
        api.getPersonQuestionnaire(pair.fromExternalId).catch(() => null),
        api.getPersonQuestionnaire(pair.toExternalId).catch(() => null),
      ]);
      if (oldQuestionnaire?.fileData && !newQuestionnaire?.fileData) {
        await api.upsertPersonQuestionnaire(pair.toExternalId, {
          fileData: oldQuestionnaire.fileData,
          fileName:
            oldQuestionnaire.fileName?.trim() ||
            sanitizeFileName(buildQuestionnaireExportFileName(pair.name)),
          mimeType: oldQuestionnaire.mimeType ?? "application/pdf",
        });
        migrated += 1;
      }

      if (!includeDocuments) continue;
      const [oldDocs, newDocs] = await Promise.all([
        api.listPersonDocuments(pair.fromExternalId).catch(() => []),
        api.listPersonDocuments(pair.toExternalId).catch(() => []),
      ]);
      const existing = [...newDocs];
      for (const document of oldDocs) {
        const already = existing.some(
          (item) => item.type === document.type && item.title === document.title,
        );
        if (already) continue;
        const created = await api.createPersonDocument(pair.toExternalId, {
          type: document.type,
          title: document.title,
          ...(document.status ? { status: document.status } : {}),
          ...(document.fields ? { fields: document.fields } : {}),
          ...(document.workflow ? { workflow: document.workflow } : {}),
          ...(document.files ? { files: document.files } : {}),
        });
        existing.push(created);
        migrated += 1;
      }
    } catch {
      // One broken ID must not block the rest.
    }
  }
  return migrated;
};

export const migratePersonAttachmentsBetweenIds = async (
  pairs: PersonAttachmentMigrationPair[],
  options: PersonAttachmentMigrationOptions = {},
) => {
  const unique = dedupePersonAttachmentMigrationPairs(pairs);
  if (!unique.length) return 0;

  const includeDocuments = options.includeDocuments !== false;
  if (
    unique.length <= 16 &&
    !options.photos &&
    !options.questionnaires &&
    !options.documents
  ) {
    return migrateTargetedPersonAttachments(unique, includeDocuments);
  }

  const [photos, questionnaires, documents] = await Promise.all([
    options.photos
      ? Promise.resolve(options.photos)
      : api.listPersonPhotos().catch(() => []),
    options.questionnaires
      ? Promise.resolve(options.questionnaires)
      : api.listPersonQuestionnaires().catch(() => []),
    includeDocuments
      ? options.documents
        ? Promise.resolve(options.documents)
        : api.listAllPersonDocuments().catch(() => [])
      : Promise.resolve([] as BackendPersonDocument[]),
  ]);

  const photoById = new Map(
    photos
      .filter((item) => item.personExternalId && item.photoData)
      .map((item) => [item.personExternalId, item.photoData]),
  );
  const questionnaireIds = new Set(
    questionnaires
      .map((item) => item.personExternalId)
      .filter((item): item is string => Boolean(item)),
  );
  const documentsById = new Map<string, BackendPersonDocument[]>();
  for (const document of documents) {
    const id = document.personExternalId?.trim();
    if (!id) continue;
    documentsById.set(id, [...(documentsById.get(id) ?? []), document]);
  }

  let migrated = 0;
  for (const pair of unique) {
    try {
      const hasWork =
        photoById.has(pair.fromExternalId) ||
        questionnaireIds.has(pair.fromExternalId) ||
        (documentsById.get(pair.fromExternalId)?.length ?? 0) > 0;
      if (!hasWork) continue;

      const oldPhoto = photoById.get(pair.fromExternalId);
      if (oldPhoto && !photoById.has(pair.toExternalId)) {
        await api.upsertPersonPhoto(pair.toExternalId, {
          photoData: oldPhoto,
          fileName: `${pair.name}.jpg`,
        });
        photoById.set(pair.toExternalId, oldPhoto);
        migrated += 1;
      }

      if (
        questionnaireIds.has(pair.fromExternalId) &&
        !questionnaireIds.has(pair.toExternalId)
      ) {
        const oldQuestionnaire = await api.getPersonQuestionnaire(
          pair.fromExternalId,
        );
        if (oldQuestionnaire?.fileData) {
          await api.upsertPersonQuestionnaire(pair.toExternalId, {
            fileData: oldQuestionnaire.fileData,
            fileName:
              oldQuestionnaire.fileName?.trim() ||
              sanitizeFileName(buildQuestionnaireExportFileName(pair.name)),
            mimeType: oldQuestionnaire.mimeType ?? "application/pdf",
          });
          questionnaireIds.add(pair.toExternalId);
          migrated += 1;
        }
      }

      if (!includeDocuments) continue;
      const oldDocs = documentsById.get(pair.fromExternalId) ?? [];
      const newDocs = [...(documentsById.get(pair.toExternalId) ?? [])];
      for (const document of oldDocs) {
        const already = newDocs.some(
          (item) => item.type === document.type && item.title === document.title,
        );
        if (already) continue;
        const created = await api.createPersonDocument(pair.toExternalId, {
          type: document.type,
          title: document.title,
          ...(document.status ? { status: document.status } : {}),
          ...(document.fields ? { fields: document.fields } : {}),
          ...(document.workflow ? { workflow: document.workflow } : {}),
          ...(document.files ? { files: document.files } : {}),
        });
        newDocs.push(created);
        migrated += 1;
      }
      documentsById.set(pair.toExternalId, newDocs);
    } catch {
      // Keep going so one broken attachment cannot empty the personnel list.
    }
  }

  return migrated;
};
