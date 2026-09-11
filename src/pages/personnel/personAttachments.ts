import type {
  BackendPersonDocument,
  BackendPersonQuestionnaire,
  BackendPersonQuestionnaireMeta,
} from "../../api";
import { api } from "../../api";
import {
  CacheKeys,
  deleteDataCache,
  fetchWithCache,
  jsonChanged,
  peekDataCache,
  readDataCache,
} from "../../data/idbDataCache";
import { loadSharedQuestionnairesMeta } from "../../data/personnelBootstrap";
import { sanitizeFileName } from "../../shared/browserExport";
import { normalizeAnketaExternalIdKey } from "../anketa-data/anketaPersonMatch";
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

const questionnaireMetaToStub = (
  meta: BackendPersonQuestionnaireMeta,
): BackendPersonQuestionnaire => ({
  id: meta.personExternalId,
  personExternalId: meta.personExternalId,
  fileName: meta.fileName ?? null,
  mimeType: "application/pdf",
  fileData: "",
  createdAt: "",
  updatedAt: "",
});

/** Full questionnaire payload including base64 PDF — use only when file bytes are required. */
export const loadPersonQuestionnaireFull = async (
  personExternalId: string,
) => {
  const id = personExternalId.trim();
  if (!id) return null;
  return api.getPersonQuestionnaire(id).catch(() => null);
};

const FILE_NAME_NOISE = new Set([
  "pdf",
  "анкета",
  "анкети",
  "questionnaire",
  "опитувальник",
  "скан",
  "scan",
]);

const stripQuestionnaireFileNamePrefixTokens = (tokens: string[]) => {
  let start = 0;
  while (start < tokens.length) {
    const token = tokens[start]!;
    if (/^\d+$/.test(token) || FILE_NAME_NOISE.has(token)) {
      start += 1;
      continue;
    }
    break;
  }
  return tokens.slice(start);
};

const nameTokensOf = (value: string) =>
  stripQuestionnaireFileNamePrefixTokens(
    normalizeAttachmentNameKey(String(value ?? "").replace(/\.pdf$/i, ""))
      .split(" ")
      .filter((token) => token.length > 1 && !FILE_NAME_NOISE.has(token)),
  );

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

export const attachmentIdsMatch = (left: string, right: string) => {
  const a = String(left ?? "").trim();
  const b = String(right ?? "").trim();
  if (!a || !b) return false;
  if (a === b) return true;
  const normalizedLeft = normalizeAnketaExternalIdKey(a);
  const normalizedRight = normalizeAnketaExternalIdKey(b);
  return Boolean(
    normalizedLeft &&
      normalizedRight &&
      normalizedLeft === normalizedRight,
  );
};

const lookupSetMatchesId = (lookupIds: readonly string[], id: string) =>
  lookupIds.some((candidate) => attachmentIdsMatch(candidate, id));

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

/** Чи прийняти знайдену анкету для цих lookup-id / імен. */
export const shouldAcceptQuestionnaireAttachment = (
  questionnaire: BackendPersonQuestionnaire | null | undefined,
  lookupIds: readonly string[],
  names: Array<string | null | undefined>,
) => {
  if (!questionnaire) return false;
  if (!String(questionnaire.fileName ?? "").trim()) return true;
  const storedId = String(questionnaire.personExternalId ?? "").trim();
  if (storedId && lookupSetMatchesId(lookupIds, storedId)) return true;
  return questionnaireFileMatchesPerson(questionnaire.fileName, names);
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

const loadQuestionnaireIndex = async (force = false) =>
  loadSharedQuestionnairesMeta({ force }).catch(
    () => [] as BackendPersonQuestionnaireMeta[],
  );

const tryDirectQuestionnaireLookup = async (lookupIds: readonly string[]) => {
  const tried = new Set<string>();
  for (const candidate of lookupIds) {
    const id = normalizeAnketaExternalIdKey(candidate);
    if (!id || !/^\d+$/.test(id) || tried.has(id)) continue;
    tried.add(id);
    const questionnaire = await api.getPersonQuestionnaire(id).catch(() => null);
    if (
      questionnaire?.personExternalId?.trim() ||
      questionnaire?.fileName?.trim()
    ) {
      return {
        questionnaire,
        resolvedExternalId:
          questionnaire.personExternalId?.trim() || id,
      };
    }
  }
  return null;
};

const resolveQuestionnaireViaIndex = async (
  lookupIds: string[],
  names: string[],
  allowFileNameFallback = true,
  forceIndex = false,
) => {
  const items = await loadQuestionnaireIndex(forceIndex);

  for (const meta of items) {
    const id = meta.personExternalId?.trim();
    if (!id || !lookupSetMatchesId(lookupIds, id)) continue;
    return {
      questionnaire: questionnaireMetaToStub(meta),
      resolvedExternalId: id,
    };
  }

  if (!allowFileNameFallback) return null;

  const fileHits = items.filter((meta) =>
    questionnaireFileMatchesPerson(meta.fileName, names),
  );
  if (!fileHits.length) return null;

  const byLookupId = fileHits.find((meta) => {
    const id = meta.personExternalId?.trim();
    return Boolean(id && lookupSetMatchesId(lookupIds, id));
  });
  if (byLookupId?.personExternalId?.trim()) {
    const id = byLookupId.personExternalId.trim();
    return {
      questionnaire: questionnaireMetaToStub(byLookupId),
      resolvedExternalId: id,
    };
  }

  if (fileHits.length === 1 && fileHits[0]?.personExternalId?.trim()) {
    const id = fileHits[0].personExternalId.trim();
    return {
      questionnaire: questionnaireMetaToStub(fileHits[0]),
      resolvedExternalId: id,
    };
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

let availablePhotoIdsCache:
  | { at: number; ids: Set<string> }
  | null = null;
let availablePhotoIdsPromise: Promise<Set<string>> | null = null;

const photoIdsFromList = (
  items: Array<{
    personExternalId?: string;
    hasFile?: boolean;
    hasThumbnail?: boolean;
    photoData?: string;
  }>,
) =>
  new Set(
    items
      .filter(
        (item) =>
          item.personExternalId &&
          (item.hasFile || item.hasThumbnail || Boolean(item.photoData)),
      )
      .map((item) => item.personExternalId!.trim())
      .filter(Boolean),
  );

const rememberPhotoIds = (ids: Set<string>) => {
  availablePhotoIdsCache = { at: Date.now(), ids };
  return ids;
};

export const clearAvailablePersonPhotoIdsCache = () => {
  availablePhotoIdsCache = null;
  availablePhotoIdsPromise = null;
  void deleteDataCache(CacheKeys.personnelPhotoIndex);
};

export const loadAvailablePersonPhotoIds = async (options?: {
  force?: boolean;
}) => {
  const now = Date.now();
  if (
    !options?.force &&
    availablePhotoIdsCache &&
    now - availablePhotoIdsCache.at < 30_000
  ) {
    return availablePhotoIdsCache.ids;
  }

  if (!availablePhotoIdsCache && !options?.force) {
    const persisted = await readDataCache<string[]>(
      CacheKeys.personnelPhotoIndex,
    );
    if (persisted?.length) {
      rememberPhotoIds(new Set(persisted));
    }
  }

  availablePhotoIdsPromise ??= fetchWithCache<string[]>({
    key: CacheKeys.personnelPhotoIndex,
    force: options?.force,
    fetcher: async () => {
      const items = await api.listPersonPhotos();
      return [...photoIdsFromList(items)];
    },
    isChanged: jsonChanged,
  })
    .then((ids) => rememberPhotoIds(new Set(ids)))
    .catch(() => availablePhotoIdsCache?.ids ?? new Set<string>())
    .finally(() => {
      availablePhotoIdsPromise = null;
    });

  return availablePhotoIdsPromise;
};

export const peekAvailablePersonPhotoIds = () =>
  availablePhotoIdsCache?.ids ?? null;

/** Map list externalIds → thumbnail URLs when the photo index is already known. */
export const collectPersonnelListPhotoUpdates = (
  externalIds: string[],
  rowByExternalId: ReadonlyMap<string, EjournalPreviewRow>,
  availableIds: Set<string> | null | undefined,
  requestedIds: ReadonlySet<string>,
) => {
  const updates: Record<string, string> = {};
  for (const externalId of new Set(externalIds)) {
    if (!externalId || requestedIds.has(externalId)) continue;
    const row = rowByExternalId.get(externalId);
    if (!row) continue;
    const photoUrl = personPhotoThumbnailUrlForRow(row, undefined, availableIds);
    if (!photoUrl) continue;
    updates[externalId] = photoUrl;
  }
  return updates;
};

/** DB / filesystem key under which the photo is stored (may differ from roster externalId). */
export const resolvePersonPhotoStorageIdForRow = (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
  availableIds?: Set<string> | null,
) => {
  const ids = availableIds ?? peekAvailablePersonPhotoIds();
  if (!ids?.size) return "";
  for (const id of collectPersonAttachmentLookupIds(row, hints)) {
    if (ids.has(id)) return id;
  }
  return "";
};

/** Prefer the indexed / fingerprint key so saves survive reload and list lookup. */
export const resolvePersonPhotoStorageIdForSave = (
  row: EjournalPreviewRow | null,
  fallbackExternalId = "",
) => {
  const fromIndex = resolvePersonPhotoStorageIdForRow(row);
  if (fromIndex) return fromIndex;

  const fingerprint = buildPersonIdentityFingerprint(
    getPersonDisplayName(row),
    resolvePersonBirthDate(row),
    resolvePersonCallSign(row),
  );
  if (fingerprint) return fingerprint;

  const fallback = fallbackExternalId.trim() || resolvePersonIdentityKey(row);
  return fallback;
};

export const pruneStalePersonPhotos = async (
  row: EjournalPreviewRow | null,
  keepId: string,
) => {
  if (!row || !keepId) return;
  const staleIds = collectPersonAttachmentLookupIds(row).filter(
    (id) => id && id !== keepId,
  );
  await Promise.all(
    staleIds.map((id) => api.deletePersonPhoto(id).catch(() => undefined)),
  );
};

export const personPhotoThumbnailUrlForRow = (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
  availableIds?: Set<string> | null,
  cacheBust?: number | string,
) => {
  const storageId = resolvePersonPhotoStorageIdForRow(row, hints, availableIds);
  return storageId
    ? api.personPhotoFileUrl(storageId, { thumbnail: true, cacheBust })
    : "";
};

export const personPhotoFullUrlForRow = (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
  availableIds?: Set<string> | null,
  cacheBust?: number | string,
) => {
  const storageId = resolvePersonPhotoStorageIdForRow(row, hints, availableIds);
  return storageId
    ? api.personPhotoFileUrl(storageId, { cacheBust })
    : "";
};

/** Lightweight list/card preview: request the 96×128 thumbnail directly. */
export const loadPersonPhotoThumbnailForRow = async (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
) => {
  const fallback = resolvePersonIdentityKey(row);
  const availableIds = await loadAvailablePersonPhotoIds();
  const candidateIds = collectPersonAttachmentLookupIds(row, hints).filter(
    (id) => availableIds.has(id),
  );
  for (const id of candidateIds) {
    try {
      const photoData = (await api.getPersonPhotoThumbnail(id)).trim();
      if (photoData) return { photoData, resolvedExternalId: id };
    } catch {
      /* try next identity key */
    }
  }
  return { photoData: "", resolvedExternalId: fallback };
};

export type LoadPersonQuestionnaireOptions = {
  /** Два однофамільці — не підставляти PDF лише за назвою файлу. */
  nameIsAmbiguous?: boolean;
  /** Оновити список анкет з API перед пошуком. */
  refreshIndex?: boolean;
};

const readWarmDocumentsCatalog = async (): Promise<BackendPersonDocument[]> => {
  const memory = peekDataCache<BackendPersonDocument[]>(CacheKeys.documentsAll);
  if (Array.isArray(memory) && memory.length) return memory;
  const persisted = await readDataCache<BackendPersonDocument[]>(
    CacheKeys.documentsAll,
  );
  return Array.isArray(persisted) ? persisted : [];
};

export const loadPersonDocumentsForRow = async (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
  options?: LoadPersonQuestionnaireOptions,
): Promise<BackendPersonDocument[]> => {
  const fallback = resolvePersonIdentityKey(row);
  const lookupIds = collectPersonAttachmentLookupIds(row, hints);
  const lookupSet = new Set(lookupIds);
  if (fallback) lookupSet.add(fallback);
  const expectedNames = [
    getPersonDisplayName(row),
    String(hints?.anketaFullName ?? "").trim(),
  ]
    .map(normalizeAttachmentNameKey)
    .filter(Boolean);
  const direct = fallback
    ? await api.listPersonDocuments(fallback).catch(() => [])
    : ([] as BackendPersonDocument[]);
  const all = await readWarmDocumentsCatalog();
  const related = all.filter((document) => {
    if (lookupSet.has(document.personExternalId)) return true;
    if (options?.nameIsAmbiguous) return false;
    const metadataName = normalizeAttachmentNameKey(document.personName || "");
    if (metadataName && expectedNames.includes(metadataName)) return true;
    const parsed = parseOrphanAttachmentIdentityId(document.personExternalId);
    return Boolean(
      parsed &&
        expectedNames.some((name) =>
          personNameMatchesOrphanNameKey(name, parsed.nameKey),
        ),
    );
  });
  const unique = new Map<string, BackendPersonDocument>();
  for (const document of [...direct, ...related]) {
    unique.set(document.id, document);
  }
  return [...unique.values()].sort(
    (left, right) =>
      new Date(right.updatedAt).getTime() - new Date(left.updatedAt).getTime(),
  );
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

  const indexed = await resolveQuestionnaireViaIndex(
    lookupIds,
    expectedNames,
    allowFileNameFallback,
    Boolean(options?.refreshIndex),
  );
  if (indexed) return indexed;

  const direct = await tryDirectQuestionnaireLookup(lookupIds);
  if (direct) return direct;

  if (fallback) {
    const items = await loadQuestionnaireIndex(Boolean(options?.refreshIndex));
    const meta = items.find((item) =>
      lookupSetMatchesId([fallback], item.personExternalId?.trim() ?? ""),
    );
    if (meta) {
      const id = meta.personExternalId?.trim() || fallback;
      return {
        questionnaire: questionnaireMetaToStub(meta),
        resolvedExternalId: id,
      };
    }
  }

  return { questionnaire: null, resolvedExternalId: fallback };
};

/** PDF is stored under any candidate / legacy identity of this row. */
export type QuestionnairePresencePerson = {
  currentId: string;
  lookupIds: string[];
  fullName: string;
};

export const buildQuestionnairePresencePeople = (
  rows: EjournalPreviewRow[],
): QuestionnairePresencePerson[] =>
  rows.flatMap((row) => {
    if (!isLikelyPersonnelRow(row)) return [];
    const currentId = resolvePersonIdentityKey(row);
    if (!currentId) return [];
    return [{
      currentId,
      lookupIds: collectPersonAttachmentLookupIds(row),
      fullName: getPersonDisplayName(row),
    }];
  });

export const buildQuestionnairePresenceFromPeople = (
  people: QuestionnairePresencePerson[],
  items: Array<{ personExternalId?: string | null; fileName?: string | null }>,
) => {
  const stored = new Set(
    items
      .map((item) => String(item.personExternalId ?? "").trim())
      .filter(Boolean),
  );
  const map: Record<string, true> = {};
  for (const id of stored) map[id] = true;
  const unmatched: QuestionnairePresencePerson[] = [];
  for (const person of people) {
    if (stored.has(person.currentId)) {
      map[person.currentId] = true;
      continue;
    }
    if (person.lookupIds.some((id) => stored.has(id))) {
      map[person.currentId] = true;
      continue;
    }
    unmatched.push(person);
  }

  if (unmatched.length && items.length) {
    const rowsByShortName = new Map<string, QuestionnairePresencePerson[]>();
    for (const person of unmatched) {
      const shortKey = nameTokensOf(person.fullName).slice(0, 2).join(" ");
      if (!shortKey) continue;
      const list = rowsByShortName.get(shortKey) ?? [];
      list.push(person);
      rowsByShortName.set(shortKey, list);
    }
    for (const item of items) {
      const fileName = String(item.fileName ?? "").trim();
      if (!fileName) continue;
      const shortKey = nameTokensOf(fileName).slice(0, 2).join(" ");
      const candidates = rowsByShortName.get(shortKey) ?? [];
      const hits = candidates.filter((person) =>
        questionnaireFileMatchesPerson(fileName, [person.fullName]),
      );
      if (hits.length !== 1) continue;
      map[hits[0].currentId] = true;
    }
  }
  return map;
};

export const buildQuestionnairePresenceMap = (
  rows: EjournalPreviewRow[],
  items: Array<{ personExternalId?: string | null; fileName?: string | null }>,
) =>
  buildQuestionnairePresenceFromPeople(
    buildQuestionnairePresencePeople(rows),
    items,
  );

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
  photos?: Array<{
    personExternalId: string;
    photoData: string;
    hasFile?: boolean;
    hasThumbnail?: boolean;
  }>;
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
        api
          .listPersonDocuments(pair.fromExternalId, { full: true })
          .catch(() => []),
        api
          .listPersonDocuments(pair.toExternalId, { full: true })
          .catch(() => []),
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
      .filter(
        (item) =>
          item.personExternalId &&
          (item.photoData || item.hasFile || item.hasThumbnail),
      )
      .map((item) => [item.personExternalId, item.photoData || ""]),
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

      const oldPhotoMarker = photoById.get(pair.fromExternalId);
      if (oldPhotoMarker !== undefined && !photoById.has(pair.toExternalId)) {
        const oldPhoto =
          oldPhotoMarker ||
          (await api.getPersonPhoto(pair.fromExternalId).catch(() => null))
            ?.photoData ||
          "";
        if (oldPhoto) {
          await api.upsertPersonPhoto(pair.toExternalId, {
            photoData: oldPhoto,
            fileName: `${pair.name}.jpg`,
          });
          photoById.set(pair.toExternalId, oldPhoto);
          migrated += 1;
        }
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
      const sourceDocumentMarkers =
        documentsById.get(pair.fromExternalId) ?? [];
      if (!sourceDocumentMarkers.length) continue;
      const [oldDocs, loadedNewDocs] = await Promise.all([
        api
          .listPersonDocuments(pair.fromExternalId, { full: true })
          .catch(() => []),
        api
          .listPersonDocuments(pair.toExternalId, { full: true })
          .catch(() => []),
      ]);
      const newDocs = [...loadedNewDocs];
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
