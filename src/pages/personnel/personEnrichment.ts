import { api, type BackendPersonDocument } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  extractPhones,
  getPersonFieldValue,
  isPersistedEjournalRowId,
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

/** Prefer personnel value when present; otherwise keep document-entered text. */
export const preferPersonnelText = (
  personnelValue: unknown,
  documentValue: unknown,
) => {
  const fromPersonnel = cellText(personnelValue);
  if (fromPersonnel) return fromPersonnel;
  return cellText(documentValue);
};

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
    const current =
      getPersonFieldValue(row, ["рнокпп_за_наявності"]) ||
      getPersonFieldValue(row, ["рнокпп"]);
    if (isBlankPersonValue(current)) {
      const key =
        resolvePersonFieldKey(row, ["рнокпп_за_наявності"]) ||
        resolvePersonFieldKey(row, ["рнокпп"]) ||
        "рнокпп_за_наявності";
      // Avoid writing into «відмова від РНОКПП».
      if (!key.toLocaleLowerCase("uk-UA").includes("відмова")) {
        updates[key] = rnokpp;
      } else {
        updates["рнокпп_за_наявності"] = rnokpp;
      }
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

  let row = options.row ?? null;
  let rowId = options.rowId ? String(options.rowId) : "";

  // Resolve ООС row when documents opened without a full personnel snapshot.
  if ((!row || !rowId) && personExternalId) {
    try {
      const profile = await api.getPersonnelProfile(personExternalId);
      if (profile) {
        const oos = profile.ejournal?.oosRow;
        const roster = profile.roster?.row;
        const source =
          oos && typeof oos === "object"
            ? oos
            : roster && typeof roster === "object"
              ? roster
              : null;
        if (source) {
          const id = String(
            (source as { id?: unknown }).id ??
              (source as { __dbRowId?: unknown }).__dbRowId ??
              "",
          ).trim();
          const values =
            (source as { values?: Record<string, unknown> }).values &&
            typeof (source as { values?: unknown }).values === "object"
              ? (source as { values: Record<string, unknown> }).values
              : (source as Record<string, unknown>);
          const resolved = {
            __dbRowId: id || undefined,
            ...values,
          } as EjournalPreviewRow;
          if (!row) row = resolved;
          if (!rowId && id) rowId = id;
        }
      }
    } catch {
      /* keep local row if profile unavailable */
    }
  }

  const fieldUpdates = buildEmptyPersonFieldUpdates(row, options.patch);
  if (isPersistedEjournalRowId(rowId) && Object.keys(fieldUpdates).length) {
    try {
      await api.updateEjournalRowValues(rowId, fieldUpdates, {
        suppressErrorToast: true,
      });
    } catch {
      /* staging row may be absent after roster-only import */
    }
  }

  return { phones, phonesAdded, fieldUpdates, phoneDocument };
};
