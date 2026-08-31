import type { BackendPersonnelOverviewRow } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  collectPersonAttachmentLookupIds,
  loadPersonPhotoForRow,
} from "../personnel/personAttachments";
import {
  buildPersonIdentityFingerprint,
  cleanPersonDisplayName,
  getPersonDisplayName,
  getPersonExternalId,
  isUnstablePersonExternalId,
  resolvePersonIdentityKey,
} from "../personnel/personnelUtils";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";

const birthFromDisplayName = (name: string) => {
  const match = String(name ?? "").match(
    /(\d{1,2}[.\/-]\d{1,2}[.\/-]\d{2,4})\s*(?:р\.?\s*н\.?)?/i,
  );
  return match?.[1] ?? "";
};

const nameKeyFromPhotoId = (id: string) => {
  const raw = String(id ?? "").trim();
  if (!raw) return "";
  if (raw.startsWith("p:")) {
    const body = raw.slice(2);
    const callIdx = body.indexOf(":c:");
    const cut = callIdx >= 0 ? body.slice(0, callIdx) : body;
    return cut.split(":")[0]?.trim() ?? "";
  }
  if (raw.startsWith("name-birth:") || raw.startsWith("name-call:")) {
    return raw.split(":")[1]?.trim() ?? "";
  }
  if (raw.startsWith("name:")) return raw.slice(5).trim();
  return "";
};

export const overviewPhotoLookupKeys = (row: BackendPersonnelOverviewRow) => {
  const birth = birthFromDisplayName(row.name);
  return [
    row.externalId,
    row.id,
    buildPersonIdentityFingerprint(row.name, birth),
    buildPersonIdentityFingerprint(row.name),
    buildPersonIdentityFingerprint(cleanPersonDisplayName(row.name), birth),
    buildPersonIdentityFingerprint(cleanPersonDisplayName(row.name)),
  ].filter(Boolean);
};

export const resolveOverviewPhoto = (
  row: BackendPersonnelOverviewRow,
  photos: Record<string, string>,
) => {
  for (const key of overviewPhotoLookupKeys(row)) {
    const photo = photos[key];
    if (photo) return photo;
  }

  const nameFingerprint = buildPersonIdentityFingerprint(
    cleanPersonDisplayName(row.name) || row.name,
  );
  if (!nameFingerprint) return "";

  for (const [id, photo] of Object.entries(photos)) {
    if (!photo) continue;
    if (
      id === nameFingerprint ||
      id.startsWith(`${nameFingerprint}:`) ||
      nameFingerprint.startsWith(`${id}:`)
    ) {
      return photo;
    }
  }
  return "";
};

export const findOverviewRosterRow = (
  row: BackendPersonnelOverviewRow,
  rosterById: Map<string, EjournalPreviewRow>,
  rosterByName: Map<string, EjournalPreviewRow>,
) => {
  if (row.externalId && !isUnstablePersonExternalId(row.externalId)) {
    const byId = rosterById.get(row.externalId);
    if (byId) return byId;
  }
  const nameKeys = [
    normalizeRosterMatchText(row.name),
    normalizeRosterMatchText(cleanPersonDisplayName(row.name)),
  ].filter(Boolean);
  for (const key of nameKeys) {
    const byName = rosterByName.get(key);
    if (byName) return byName;
  }
  return undefined;
};

export const indexOverviewRosterRows = (rosterRows: EjournalPreviewRow[]) => {
  const rosterById = new Map<string, EjournalPreviewRow>();
  const rosterByName = new Map<string, EjournalPreviewRow>();

  for (const row of rosterRows) {
    const id = getPersonExternalId(row);
    const identity = resolvePersonIdentityKey(row);
    const displayName = getPersonDisplayName(row);
    if (id) rosterById.set(id, row);
    if (identity) rosterById.set(identity, row);
    for (const name of [displayName, row.ПІБ, row.піб, row.прізвище]) {
      const key = normalizeRosterMatchText(name);
      if (key) rosterByName.set(key, row);
    }
  }

  return { rosterById, rosterByName };
};

const photoIdFromListItem = (item: {
  personExternalId?: string;
  externalId?: string;
  person_external_id?: string;
}) =>
  String(
    item.personExternalId ?? item.externalId ?? item.person_external_id ?? "",
  ).trim();

export const buildOverviewPhotoMap = (
  photoList: Array<{
    personExternalId?: string;
    photoData?: string;
    externalId?: string;
    person_external_id?: string;
  }>,
  rows: BackendPersonnelOverviewRow[],
  rosterRows: EjournalPreviewRow[],
) => {
  const photos: Record<string, string> = {};
  const idsWithoutData: string[] = [];

  for (const item of photoList) {
    const id = photoIdFromListItem(item);
    const data = String(item.photoData ?? "").trim();
    if (id && data) photos[id] = data;
    else if (id) idsWithoutData.push(id);
  }

  const { rosterById, rosterByName } = indexOverviewRosterRows(rosterRows);

  const assign = (keys: string[], photo: string) => {
    if (!photo) return;
    for (const key of keys) {
      if (key) photos[key] = photo;
    }
  };

  const findListedPhoto = (keys: string[]) => {
    for (const key of keys) {
      if (photos[key]) return photos[key];
    }
    const nameKeys = keys
      .map((key) => nameKeyFromPhotoId(key) || key.replace(/^p:/, "").split(":")[0])
      .filter(Boolean);
    if (!nameKeys.length) return "";
    for (const [id, data] of Object.entries(photos)) {
      const photoName = nameKeyFromPhotoId(id);
      if (photoName && nameKeys.includes(photoName)) return data;
    }
    return "";
  };

  for (const row of rows) {
    const rosterRow = findOverviewRosterRow(row, rosterById, rosterByName);
    const keys = [
      ...overviewPhotoLookupKeys(row),
      ...(rosterRow ? collectPersonAttachmentLookupIds(rosterRow) : []),
    ].filter(Boolean);
    const photo = findListedPhoto(keys);
    if (photo) assign(keys, photo);
  }

  return { photos, idsWithoutData };
};

const runPool = async <T,>(
  items: T[],
  limit: number,
  worker: (item: T) => Promise<void>,
) => {
  let index = 0;
  const size = Math.min(Math.max(limit, 1), items.length || 1);
  await Promise.all(
    Array.from({ length: size }, async () => {
      while (index < items.length) {
        const current = items[index];
        index += 1;
        await worker(current);
      }
    }),
  );
};

export const fillMissingOverviewPhotos = async (
  rows: BackendPersonnelOverviewRow[],
  rosterRows: EjournalPreviewRow[],
  photos: Record<string, string>,
  onProgress?: (next: Record<string, string>) => void,
) => {
  const next = { ...photos };
  const { rosterById, rosterByName } = indexOverviewRosterRows(rosterRows);
  const missing = rows.filter((row) => !resolveOverviewPhoto(row, next));
  if (!missing.length) return next;

  await runPool(missing, 6, async (row) => {
    const rosterRow = findOverviewRosterRow(row, rosterById, rosterByName);
    const fallbackRow = {
      ПІБ: cleanPersonDisplayName(row.name) || row.name,
      прізвище: cleanPersonDisplayName(row.name) || row.name,
      id: row.externalId,
    } as EjournalPreviewRow;

    const result = await loadPersonPhotoForRow(rosterRow ?? fallbackRow);
    const photo = result.photoData?.trim() || "";
    if (!photo) return;

    const keys = [
      ...overviewPhotoLookupKeys(row),
      result.resolvedExternalId,
      rosterRow ? resolvePersonIdentityKey(rosterRow) : "",
    ].filter(Boolean);
    for (const key of keys) next[key] = photo;
    onProgress?.({ ...next });
  });

  return next;
};
