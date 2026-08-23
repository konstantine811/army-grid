import type { BackendPersonDocument, BackendPersonQuestionnaire } from "../../api";
import { api } from "../../api";
import { sanitizeFileName } from "../../shared/browserExport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
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
) => {
  const nameKey = normalizeAttachmentNameKey(name);
  if (!nameKey || nameKey === "особа не вибрана") return;
  ids.add(`name:${nameKey}`);
  const birthKey = normalizePersonBirthKey(birthDate);
  if (birthKey) ids.add(`name-birth:${nameKey}:${birthKey}`);
  const callSignKey = normalizeAttachmentNameKey(callSign);
  if (callSignKey) {
    ids.add(`call:${callSignKey}`);
    ids.add(`name-call:${nameKey}:${callSignKey}`);
  }
};

const hasQuestionnaireRecord = (
  questionnaire: BackendPersonQuestionnaire | null | undefined,
) => Boolean(questionnaire?.personExternalId || questionnaire?.id);

const nameTokensMatchFileName = (fullName: string, fileName: string) => {
  const nameKey = normalizeAttachmentNameKey(fullName);
  const fileKey = normalizeAttachmentNameKey(
    String(fileName ?? "").replace(/\.pdf$/i, ""),
  );
  if (!nameKey || !fileKey) return false;
  const tokens = nameKey.split(" ").filter((token) => token.length > 1);
  if (tokens.length < 2) return false;
  return tokens.every((token) => fileKey.includes(token));
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

  for (const meta of items) {
    const id = meta.personExternalId?.trim();
    if (!id) continue;
    if (!names.some((name) => nameTokensMatchFileName(name, meta.fileName ?? ""))) {
      continue;
    }
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

/** All plausible keys under which photo / questionnaire may have been saved. */
export const collectPersonAttachmentLookupIds = (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
) => {
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
    pushLegacyAttachmentLookupIds(
      ids,
      name,
      name === personnelName ? personnelBirth : anketaBirth,
      name === personnelName ? callSign : "",
    );
    for (const birth of [personnelBirth, anketaBirth, ""]) {
      const withBirth = birth
        ? buildPersonIdentityFingerprint(name, birth)
        : "";
      if (withBirth) ids.add(withBirth);
    }
    const withCallSign = buildPersonIdentityFingerprint(name, "", callSign);
    const nameOnly = buildPersonIdentityFingerprint(name);
    if (withCallSign) ids.add(withCallSign);
    if (nameOnly) ids.add(nameOnly);
  }

  const primary = resolvePersonIdentityKey(row);
  if (primary) ids.add(primary);

  return [...ids];
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

export const loadPersonQuestionnaireForRow = async (
  row: EjournalPreviewRow | null,
  hints?: PersonAttachmentLookupHints,
): Promise<{
  questionnaire: BackendPersonQuestionnaire | null;
  resolvedExternalId: string;
}> => {
  const fallback = resolvePersonIdentityKey(row);
  const lookupIds = collectPersonAttachmentLookupIds(row, hints);
  for (const id of lookupIds) {
    try {
      const questionnaire = await api.getPersonQuestionnaire(id);
      if (hasQuestionnaireRecord(questionnaire)) {
        return {
          questionnaire,
          resolvedExternalId:
            questionnaire!.personExternalId?.trim() || id,
        };
      }
    } catch {
      /* try next key */
    }
  }

  const personnelName = getPersonDisplayName(row);
  const anketaName = String(hints?.anketaFullName ?? "").trim();
  const indexed = await resolveQuestionnaireViaIndex(lookupIds, [
    personnelName,
    anketaName,
  ]);
  if (indexed) return indexed;

  return { questionnaire: null, resolvedExternalId: fallback };
};

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
) => {
  if (!orphanIds.size) return [];

  const pairs: PersonAttachmentMigrationPair[] = [];
  for (const row of rows) {
    if (!isLikelyPersonnelRow(row)) continue;
    const toExternalId = resolvePersonIdentityKey(row);
    if (!toExternalId) continue;
    const name =
      getPersonFieldValue(row, ["прізвище"]) ||
      getPersonFieldValue(row, ["піб"]);
    for (const fromExternalId of collectPersonExternalIdCandidates(row)) {
      if (!orphanIds.has(fromExternalId) || fromExternalId === toExternalId) {
        continue;
      }
      pairs.push({
        name,
        fromExternalId,
        toExternalId,
      });
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
          fileName: sanitizeFileName(
            buildQuestionnaireExportFileName(pair.name),
          ),
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
            fileName: sanitizeFileName(
              buildQuestionnaireExportFileName(pair.name),
            ),
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
