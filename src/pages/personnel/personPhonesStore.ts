import { api, type BackendPersonDocument } from "../../api";
import { normalizeUaPhone } from "./personnelUtils";

export const PERSON_PHONES_DOCUMENT_TYPE = "personPhones";
export const PERSON_PHONES_DOCUMENT_TITLE = "Телефони";

const STORAGE_KEY = "army-grid:person-phones";

export const uniqueNormalizedPhones = (values: unknown): string[] => {
  const phones: string[] = [];
  const list = Array.isArray(values) ? values : [values];
  for (const value of list) {
    const normalized = normalizeUaPhone(String(value ?? ""));
    if (normalized && !phones.includes(normalized)) phones.push(normalized);
  }
  return phones;
};

export const readStoredPersonPhones = (): Record<string, string[]> => {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    return Object.fromEntries(
      Object.entries(parsed).map(([id, value]) => [
        id,
        uniqueNormalizedPhones(value),
      ]),
    );
  } catch {
    return {};
  }
};

export const writeStoredPersonPhones = (
  phonesById: Record<string, string[]>,
) => {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(phonesById));
};

export const extractPhonesFromDocuments = (
  documents: BackendPersonDocument[] | undefined,
) => {
  const document =
    documents?.find((item) => item.type === PERSON_PHONES_DOCUMENT_TYPE) ??
    null;
  const fields = document?.fields;
  const rawPhones =
    fields && typeof fields === "object" && !Array.isArray(fields)
      ? (fields as { phones?: unknown }).phones
      : undefined;
  return {
    document,
    phones: uniqueNormalizedPhones(rawPhones),
  };
};

export const migrateStoredPersonPhones = (
  pairs: Array<{ fromExternalId: string; toExternalId: string }>,
) => {
  const store = readStoredPersonPhones();
  let changed = false;
  for (const pair of pairs) {
    const fromId = pair.fromExternalId.trim();
    const toId = pair.toExternalId.trim();
    const from = store[fromId];
    if (!fromId || !toId || fromId === toId || !from?.length) continue;
    store[toId] = uniqueNormalizedPhones([...(store[toId] ?? []), ...from]);
    changed = true;
  }
  if (changed) writeStoredPersonPhones(store);
  return store;
};

export const upsertPersonPhonesDocument = async (
  personExternalId: string,
  phones: string[],
  existingDocument?: BackendPersonDocument | null,
) => {
  const payload = { phones: uniqueNormalizedPhones(phones) };
  if (existingDocument?.id) {
    if (payload.phones.length === 0) {
      await api.deletePersonDocument(personExternalId, existingDocument.id);
      return null;
    }
    return api.updatePersonDocument(personExternalId, existingDocument.id, {
      fields: payload,
    });
  }
  if (!payload.phones.length) return null;
  return api.createPersonDocument(personExternalId, {
    type: PERSON_PHONES_DOCUMENT_TYPE,
    title: PERSON_PHONES_DOCUMENT_TITLE,
    fields: payload,
  });
};
