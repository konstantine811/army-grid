import type {
  BackendEjournalImport,
  BackendPersonDocument,
  BackendPersonnelOverviewRow,
  BackendPersonQuestionnaireMeta,
} from "../../api";
import { CacheKeys, readDataCache } from "../../data/idbDataCache";
import { loadSharedEjournalImports } from "../../data/sharedAppData";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  collectPersonAttachmentLookupIds,
  parseOrphanAttachmentIdentityId,
  personNameMatchesOrphanNameKey,
  questionnaireFileMatchesPerson,
} from "../personnel/personAttachments";
import { PERSON_PHONES_DOCUMENT_TYPE } from "../personnel/personPhonesStore";
import { getRosterPersonName, mergeRosterRowsIntoPreview } from "../personnel/personnelRosterMerge";
import {
  findEjournalPersonnelSheet,
  getPersonDisplayName,
  loadAllEjournalSheetRows,
  normalizePersonBirthKey,
  resolvePersonBirthDate,
  resolvePersonIdentityKey,
  sheetRowsCacheKey,
} from "../personnel/personnelUtils";
import { normalizeOverviewName } from "./overviewNameSearch";
import type { OverviewPersonDocumentSummary } from "./OverviewVirtualTable";

export type OverviewPersonnelAssets = {
  photos: Record<string, string>;
  questionnairePresence: Record<string, true>;
  questionnaireSourceIds: Record<string, string>;
  documents: Record<string, OverviewPersonDocumentSummary>;
};

const stripNameNoise = (value: unknown) =>
  String(value ?? "")
    .replace(/\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4}\s*(?:р\.?\s*н\.?)?/gi, " ")
    .replace(/\([^)]*\)/g, " ");

const nameKeyOf = (value: unknown) => normalizeOverviewName(stripNameNoise(value));

const birthFromText = (value: unknown) => {
  const match = String(value ?? "").match(/(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})/);
  return match ? normalizePersonBirthKey(match[1]) : "";
};

/** Unique join key the user asked for: ПІБ + дата народження. */
export const overviewPersonMatchKey = (name: unknown, birth: unknown = "") => {
  const nameKey = nameKeyOf(name);
  if (!nameKey) return "";
  const birthKey =
    normalizePersonBirthKey(String(birth ?? "")) || birthFromText(name);
  return birthKey ? `${nameKey}|${birthKey}` : nameKey;
};

const overviewDocumentTypeLabel = (type: string) =>
  type === "ubdReport"
    ? "Рапорт на УБД"
    : type === "form6Report"
      ? "Форма 6"
      : type === "form12Report"
        ? "Форма 12"
        : type === "serviceCharacteristic"
          ? "Службова характеристика"
          : type === "zhbdCertificate"
            ? "Довідка ЖБД"
            : type === "ubdRestoreReport"
              ? "Рапорт на відновлення УБД"
              : type === "salaryPowerAttorney"
                ? "Довіреність на зарплату"
                : type === "temporaryMilitaryId"
                  ? "Тимчасовий військовий квиток"
                  : type === "lostMilitaryId"
                    ? "Втрата військового квитка"
                    : type;

const personnelName = (row: EjournalPreviewRow) =>
  getPersonDisplayName(row) || getRosterPersonName(row);

const personnelBirth = (row: EjournalPreviewRow) =>
  normalizePersonBirthKey(resolvePersonBirthDate(row)) ||
  birthFromText(personnelName(row));

const takeLookup = (row: EjournalPreviewRow) => {
  try {
    return [
      resolvePersonIdentityKey(row),
      ...collectPersonAttachmentLookupIds(row),
    ].filter(Boolean);
  } catch {
    return [] as string[];
  }
};

const asList = <T,>(value: unknown): T[] =>
  Array.isArray(value) ? value : [];

const remember = (
  byFull: Record<string, string>,
  byName: Record<string, string>,
  nameAmbiguous: Set<string>,
  nameKey: string,
  birthKey: string,
  value: string,
) => {
  if (!nameKey || !value) return;
  if (birthKey) byFull[`${nameKey}|${birthKey}`] = value;
  if (nameAmbiguous.has(nameKey)) return;
  if (byName[nameKey] && byName[nameKey] !== value) {
    delete byName[nameKey];
    nameAmbiguous.add(nameKey);
    return;
  }
  byName[nameKey] = value;
};

const pickByPerson = (
  byFull: Record<string, string>,
  byName: Record<string, string>,
  nameKey: string,
  birthKey: string,
) =>
  (birthKey && byFull[`${nameKey}|${birthKey}`]) || byName[nameKey] || "";

/** Same people Особовий склад shows: аркуш ООС + штатка. */
export const loadPersonnelRowsForOverview = async (
  rosterRows: EjournalPreviewRow[],
): Promise<EjournalPreviewRow[]> => {
  try {
    const cachedImports = await readDataCache<BackendEjournalImport[]>(
      CacheKeys.ejournalImports,
    );
    const imports = cachedImports?.length
      ? cachedImports
      : await loadSharedEjournalImports();
    const sheet = findEjournalPersonnelSheet(imports);
    if (!sheet) return rosterRows;
    const cachedPreview = await readDataCache<DbPreviewState>(
      sheetRowsCacheKey(sheet),
    );
    const preview = cachedPreview ?? (await loadAllEjournalSheetRows(sheet));
    return mergeRosterRowsIntoPreview(preview, rosterRows);
  } catch {
    return rosterRows;
  }
};

/** Copy Особовий склад анкети / фото / документи onto Overview by ПІБ + дата народження. */
export const applyPersonnelAssetsToOverview = (
  overviewRows: BackendPersonnelOverviewRow[],
  personnelRows: EjournalPreviewRow[],
  questionnaires: BackendPersonQuestionnaireMeta[] = [],
  photoList: Array<{ personExternalId?: string; photoData?: string }> = [],
  documents: BackendPersonDocument[] = [],
): OverviewPersonnelAssets => {
  const people = asList<EjournalPreviewRow>(personnelRows).filter(Boolean);
  const qList = asList<BackendPersonQuestionnaireMeta>(questionnaires);
  const photosIn = asList<{ personExternalId?: string; photoData?: string }>(
    photoList,
  );
  const docsIn = asList<BackendPersonDocument>(documents);

  const photoById: Record<string, string> = {};
  for (const item of photosIn) {
    const id = String(item.personExternalId ?? "").trim();
    const data = String(item.photoData ?? "").trim();
    if (id && data) photoById[id] = data;
  }
  const docsById: Record<string, OverviewPersonDocumentSummary> = {};
  for (const document of docsIn) {
    if (document.type === PERSON_PHONES_DOCUMENT_TYPE) continue;
    const id = String(document.personExternalId ?? "").trim();
    if (!id) continue;
    const current = docsById[id] ?? { count: 0, labels: [] };
    current.count += 1;
    const label = document.title?.trim() || overviewDocumentTypeLabel(document.type);
    if (label && !current.labels.includes(label)) current.labels.push(label);
    docsById[id] = current;
  }

  const qByFull: Record<string, string> = {};
  const qByName: Record<string, string> = {};
  const qNameAmbiguous = new Set<string>();
  const photoByFull: Record<string, string> = {};
  const photoByName: Record<string, string> = {};
  const photoNameAmbiguous = new Set<string>();
  const docIdByFull: Record<string, string> = {};
  const docIdByName: Record<string, string> = {};
  const docNameAmbiguous = new Set<string>();
  const peopleByName = new Map<string, Array<{ birthKey: string; ids: string[] }>>();

  const addPersonName = (nameKey: string, birthKey: string, ids: string[]) => {
    if (!nameKey) return;
    const list = peopleByName.get(nameKey) ?? [];
    list.push({ birthKey, ids });
    peopleByName.set(nameKey, list);
  };

  const rememberAll = (
    nameKey: string,
    birthKey: string,
    qSource: string,
    photo: string,
    docId: string,
  ) => {
    remember(qByFull, qByName, qNameAmbiguous, nameKey, birthKey, qSource);
    remember(photoByFull, photoByName, photoNameAmbiguous, nameKey, birthKey, photo);
    remember(docIdByFull, docIdByName, docNameAmbiguous, nameKey, birthKey, docId);
  };

  const rememberParsedId = (id: string, extraName = "", extraBirth = "") => {
    const parsed = parseOrphanAttachmentIdentityId(id);
    const nameKey = nameKeyOf(extraName || parsed?.nameKey || "");
    const birthKey =
      extraBirth ||
      parsed?.birthKey ||
      birthFromText(extraName || parsed?.nameKey || "");
    if (!nameKey) return { nameKey: "", birthKey: "" };
    return { nameKey, birthKey };
  };

  for (const item of qList) {
    const source = String(item.personExternalId ?? "").trim();
    if (!source) continue;
    const { nameKey, birthKey } = rememberParsedId(source);
    if (nameKey) remember(qByFull, qByName, qNameAmbiguous, nameKey, birthKey, source);
  }

  for (const [id, photo] of Object.entries(photoById)) {
    const { nameKey, birthKey } = rememberParsedId(id);
    if (nameKey) remember(photoByFull, photoByName, photoNameAmbiguous, nameKey, birthKey, photo);
  }

  for (const id of Object.keys(docsById)) {
    const { nameKey, birthKey } = rememberParsedId(id);
    if (nameKey) remember(docIdByFull, docIdByName, docNameAmbiguous, nameKey, birthKey, id);
  }

  const storedQIds = new Set(
    qList.map((item) => String(item.personExternalId ?? "").trim()).filter(Boolean),
  );

  for (const row of people) {
    const name = personnelName(row);
    const nameKey = nameKeyOf(name);
    const birthKey = personnelBirth(row);
    if (!nameKey) continue;
    const lookup = takeLookup(row);
    addPersonName(nameKey, birthKey, lookup);
    const qSource = lookup.find((id) => storedQIds.has(id)) || "";
    const photo = lookup.map((id) => photoById[id]).find(Boolean) || "";
    const docId = lookup.find((id) => docsById[id]) || "";
    rememberAll(nameKey, birthKey, qSource, photo, docId);
  }

  for (const item of qList) {
    const fileName = String(item.fileName ?? "").trim();
    const source = String(item.personExternalId ?? "").trim();
    if (!fileName || !source || /^questionnaire\.pdf$/i.test(fileName)) continue;
    const hits = overviewRows.filter((row) =>
      questionnaireFileMatchesPerson(fileName, [row.name]),
    );
    if (hits.length !== 1) continue;
    const nameKey = nameKeyOf(hits[0].name);
    const birthKey = birthFromText(hits[0].name);
    remember(qByFull, qByName, qNameAmbiguous, nameKey, birthKey, source);
  }

  const resolveOverviewIdentity = (row: BackendPersonnelOverviewRow) => {
    const nameKey = nameKeyOf(row.name);
    const birthFromName = birthFromText(row.name);
    const candidates = peopleByName.get(nameKey) ?? [];
    const byBirth = birthFromName
      ? candidates.find((item) => item.birthKey === birthFromName)
      : undefined;
    const unique = candidates.length === 1 ? candidates[0] : undefined;
    const birthKey = birthFromName || byBirth?.birthKey || unique?.birthKey || "";
    return { nameKey, birthKey };
  };

  const photos: Record<string, string> = { ...photoById };
  const questionnairePresence: Record<string, true> = {};
  const questionnaireSourceIds: Record<string, string> = {};
  const documentsOut: Record<string, OverviewPersonDocumentSummary> = {
    ...docsById,
  };

  const assignKeys = (keys: string[], write: (key: string) => void) => {
    for (const key of keys) {
      if (key) write(key);
    }
  };

  for (const row of overviewRows) {
    const rowId = String(row.externalId ?? "").trim();
    const { nameKey, birthKey } = resolveOverviewIdentity(row);
    if (!nameKey) continue;
    const matchKey = overviewPersonMatchKey(row.name, birthKey);
    const keys = [rowId, row.id, matchKey].filter(Boolean);

    let qSource = pickByPerson(qByFull, qByName, nameKey, birthKey);
    if (!qSource && rowId) {
      const parsed = parseOrphanAttachmentIdentityId(rowId);
      if (parsed?.nameKey && personNameMatchesOrphanNameKey(row.name, parsed.nameKey)) {
        qSource = qList.find((item) => item.personExternalId === rowId)
          ? rowId
          : "";
      }
      if (!qSource && qList.some((item) => item.personExternalId === rowId)) {
        qSource = rowId;
      }
    }
    if (qSource) {
      assignKeys(keys, (key) => {
        questionnairePresence[key] = true;
        questionnaireSourceIds[key] = qSource;
      });
    }

    const photo = pickByPerson(photoByFull, photoByName, nameKey, birthKey);
    if (photo) assignKeys(keys, (key) => {
      photos[key] = photo;
    });

    const docId = pickByPerson(docIdByFull, docIdByName, nameKey, birthKey);
    if (docId && docsById[docId]) {
      assignKeys(keys, (key) => {
        documentsOut[key] = docsById[docId];
      });
    }
  }

  return {
    photos,
    questionnairePresence,
    questionnaireSourceIds,
    documents: documentsOut,
  };
};
