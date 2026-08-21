import { api, type BackendPersonDocument } from "../../api";

export const PERSON_SIGNATURE_DOCUMENT_TYPE = "personSignature";
export const PERSON_SIGNATURE_DOCUMENT_TITLE = "Підпис службовця";

const STORAGE_KEY = "army-grid:person-signatures";

export type PersonSignatureRecord = {
  signatureData: string;
  signatureFileName: string;
};

const normalizeRecord = (value: unknown): PersonSignatureRecord | null => {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const data = value as Record<string, unknown>;
  const signatureData = String(data.signatureData ?? "").trim();
  if (!signatureData.startsWith("data:image")) return null;
  return {
    signatureData,
    signatureFileName: String(data.signatureFileName ?? "").trim() || "підпис.png",
  };
};

export const readStoredPersonSignatures = (): Record<
  string,
  PersonSignatureRecord
> => {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    return Object.fromEntries(
      Object.entries(parsed)
        .map(([id, value]) => [id, normalizeRecord(value)] as const)
        .filter((entry): entry is [string, PersonSignatureRecord] =>
          Boolean(entry[1]),
        ),
    );
  } catch {
    return {};
  }
};

export const writeStoredPersonSignatures = (
  signaturesById: Record<string, PersonSignatureRecord>,
) => {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(signaturesById));
};

export const getStoredPersonSignature = (personExternalId: string) => {
  const id = personExternalId.trim();
  if (!id) return null;
  return readStoredPersonSignatures()[id] ?? null;
};

export const setStoredPersonSignature = (
  personExternalId: string,
  record: PersonSignatureRecord | null,
) => {
  const id = personExternalId.trim();
  if (!id) return;
  const store = readStoredPersonSignatures();
  if (!record?.signatureData) {
    delete store[id];
  } else {
    store[id] = {
      signatureData: record.signatureData,
      signatureFileName: record.signatureFileName.trim() || "підпис.png",
    };
  }
  writeStoredPersonSignatures(store);
};

export const migrateStoredPersonSignatures = (
  pairs: Array<{ fromExternalId: string; toExternalId: string }>,
) => {
  const store = readStoredPersonSignatures();
  let changed = false;
  for (const pair of pairs) {
    const fromId = pair.fromExternalId.trim();
    const toId = pair.toExternalId.trim();
    const from = store[fromId];
    if (!fromId || !toId || fromId === toId || !from) continue;
    if (!store[toId]) {
      store[toId] = from;
      changed = true;
    }
  }
  if (changed) writeStoredPersonSignatures(store);
  return store;
};

const signatureFromDocumentFields = (fields: unknown) =>
  normalizeRecord(fields);

export const extractPersonSignature = (
  documents: BackendPersonDocument[] | undefined,
  personExternalId?: string,
): {
  document: BackendPersonDocument | null;
  signature: PersonSignatureRecord | null;
} => {
  const list = documents ?? [];
  const dedicated =
    list.find((item) => item.type === PERSON_SIGNATURE_DOCUMENT_TYPE) ?? null;
  const fromDedicated = signatureFromDocumentFields(dedicated?.fields);
  if (fromDedicated) {
    return { document: dedicated, signature: fromDedicated };
  }

  for (const type of ["form12Report", "ubdRestoreReport"] as const) {
    for (const document of list) {
      if (document.type !== type) continue;
      const signature = signatureFromDocumentFields(document.fields);
      if (signature) return { document: dedicated, signature };
    }
  }

  if (personExternalId) {
    const stored = getStoredPersonSignature(personExternalId);
    if (stored) return { document: dedicated, signature: stored };
  }

  return { document: dedicated, signature: null };
};

export const withPersonSignature = <
  T extends { signatureData?: string; signatureFileName?: string },
>(
  fields: T,
  signature: PersonSignatureRecord | null,
): T => {
  if (String(fields.signatureData ?? "").trim()) return fields;
  if (!signature?.signatureData) return fields;
  return {
    ...fields,
    signatureData: signature.signatureData,
    signatureFileName:
      signature.signatureFileName || fields.signatureFileName || "підпис.png",
  };
};

export const upsertPersonSignatureDocument = async (
  personExternalId: string,
  signature: PersonSignatureRecord | null,
  existingDocument?: BackendPersonDocument | null,
) => {
  const id = personExternalId.trim();
  if (!id) return null;

  setStoredPersonSignature(id, signature);

  if (!signature?.signatureData) {
    if (existingDocument?.id) {
      await api.deletePersonDocument(id, existingDocument.id);
    }
    return null;
  }

  const fields = {
    signatureData: signature.signatureData,
    signatureFileName: signature.signatureFileName || "підпис.png",
  };

  if (existingDocument?.id) {
    return api.updatePersonDocument(id, existingDocument.id, { fields });
  }

  return api.createPersonDocument(id, {
    type: PERSON_SIGNATURE_DOCUMENT_TYPE,
    title: PERSON_SIGNATURE_DOCUMENT_TITLE,
    fields,
  });
};

export const persistPersonSignature = async (
  personExternalId: string,
  signature: PersonSignatureRecord | null,
  documents?: BackendPersonDocument[],
) => {
  const existing =
    documents?.find((item) => item.type === PERSON_SIGNATURE_DOCUMENT_TYPE) ??
    extractPersonSignature(documents).document;
  return upsertPersonSignatureDocument(personExternalId, signature, existing);
};
