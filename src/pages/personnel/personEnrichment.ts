import { api, type BackendPersonDocument } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  extractPhones,
  getPersonFieldValue,
  normalizeUaPhone,
  previewValueToDisplay,
  resolvePersonFieldKey,
} from "./personnelUtils";
import {
  uniqueNormalizedPhones,
  upsertPersonPhonesDocument,
  writeStoredPersonPhones,
  readStoredPersonPhones,
} from "./personPhonesStore";

export type PersonEnrichmentPatch = {
  rnokpp?: string;
  address?: string;
  phone?: string;
  phones?: string[];
};

export type PersonEnrichmentResult = {
  phones: string[];
  phonesAdded: string[];
  fieldUpdates: Record<string, string>;
  phoneDocument: BackendPersonDocument | null;
};

const cellText = (value: unknown) => previewValueToDisplay(value).trim();

export const isBlankPersonValue = (value: unknown) => !cellText(value);

/** Union phones; never drop existing numbers. */
export const mergePhonesAppendOnly = (
  ...lists: Array<Iterable<string> | undefined>
) =>
  uniqueNormalizedPhones(
    lists.flatMap((list) => (list ? [...list] : [])),
  );

export const phonesFromPatch = (patch: PersonEnrichmentPatch) =>
  uniqueNormalizedPhones([
    ...(patch.phones ?? []),
    ...(patch.phone ? extractPhones(patch.phone) : []),
    ...(patch.phone ? [normalizeUaPhone(patch.phone) || ""] : []),
  ]);

/**
 * Only fills empty personnel fields. Existing РНОКПП / адреса are never overwritten
 * when the source (document / questionnaire / new import) has a value or is empty.
 */
export const buildEmptyPersonFieldUpdates = (
  row: EjournalPreviewRow | null | undefined,
  patch: PersonEnrichmentPatch,
): Record<string, string> => {
  if (!row) return {};
  const updates: Record<string, string> = {};

  const rnokpp = String(patch.rnokpp ?? "").replace(/\D/g, "").trim();
  if (rnokpp.length >= 8) {
    const current = getPersonFieldValue(row, ["рнокпп_за_наявності"]);
    if (isBlankPersonValue(current)) {
      const key =
        resolvePersonFieldKey(row, ["рнокпп_за_наявності"]) ||
        "рнокпп_за_наявності";
      updates[key] = rnokpp;
    }
  }

  const address = String(patch.address ?? "").trim();
  if (address) {
    const current =
      getPersonFieldValue(row, ["адреса", "проживан"]) ||
      getPersonFieldValue(row, ["зареєстров"]) ||
      getPersonFieldValue(row, ["місце", "проживан"]) ||
      getPersonFieldValue(row, ["адреса"]);
    if (isBlankPersonValue(current)) {
      const key =
        resolvePersonFieldKey(row, ["адреса", "проживан"]) ||
        resolvePersonFieldKey(row, ["адреса"]) ||
        "адреса_проживання";
      updates[key] = address;
    }
  }

  return updates;
};

export const applyEnrichmentToPreviewRow = (
  row: EjournalPreviewRow,
  fieldUpdates: Record<string, string>,
): EjournalPreviewRow => {
  if (!Object.keys(fieldUpdates).length) return row;
  return { ...row, ...fieldUpdates };
};

export const syncEnrichmentToPerson = async (options: {
  personExternalId: string;
  rowId?: string | null;
  row?: EjournalPreviewRow | null;
  patch: PersonEnrichmentPatch;
  existingPhones?: string[];
  phoneDocument?: BackendPersonDocument | null;
}): Promise<PersonEnrichmentResult> => {
  const personExternalId = options.personExternalId.trim();
  const incomingPhones = phonesFromPatch(options.patch);
  const before = uniqueNormalizedPhones(options.existingPhones ?? []);
  const phones = mergePhonesAppendOnly(before, incomingPhones);
  const phonesAdded = phones.filter((phone) => !before.includes(phone));

  let phoneDocument = options.phoneDocument ?? null;
  if (personExternalId && phonesAdded.length) {
    const store = readStoredPersonPhones();
    store[personExternalId] = phones;
    writeStoredPersonPhones(store);
    phoneDocument = await upsertPersonPhonesDocument(
      personExternalId,
      phones,
      phoneDocument,
    );
  }

  const fieldUpdates = buildEmptyPersonFieldUpdates(options.row, options.patch);
  if (options.rowId && Object.keys(fieldUpdates).length) {
    await api.updateEjournalRowValues(options.rowId, fieldUpdates);
  }

  return { phones, phonesAdded, fieldUpdates, phoneDocument };
};
