import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  collectPersonExternalIdCandidates,
} from "../personnel/personnelUtils";

export const SELECTED_PERSON_STORAGE_KEY = "army-grid:selected-person";
export const SELECTED_DOCUMENT_MODE_KEY = "army-grid:selected-document-mode";

/** Already copied to IndexedDB; safe to drop when localStorage is full. */
const QUOTA_RELIEF_KEYS = ["army-grid:basis-orders-imported"];
const MAX_FIELD_CHARS = 8_000;

export const compactPersonRowForStorage = (
  row: EjournalPreviewRow,
): EjournalPreviewRow => {
  const next: EjournalPreviewRow = {};
  for (const [key, value] of Object.entries(row)) {
    if (value == null) continue;
    if (typeof value === "string") {
      if (value.length > MAX_FIELD_CHARS) continue;
      next[key] = value;
      continue;
    }
    if (typeof value === "number" || typeof value === "boolean") {
      next[key] = value;
    }
  }
  return next;
};

const parseStoredRow = (raw: string | null): EjournalPreviewRow | null => {
  if (!raw) return null;
  const parsed = JSON.parse(raw) as EjournalPreviewRow;
  return parsed && typeof parsed === "object" ? parsed : null;
};

export const readStoredSelectedPersonRow = (): EjournalPreviewRow | null => {
  if (typeof window === "undefined") return null;
  for (const storage of [window.sessionStorage, window.localStorage]) {
    try {
      const row = parseStoredRow(storage.getItem(SELECTED_PERSON_STORAGE_KEY));
      if (row) return row;
    } catch {
      // Try the other storage.
    }
  }
  return null;
};

export const readStoredSelectedDocumentMode = () => {
  if (typeof window === "undefined") return "";
  for (const storage of [window.sessionStorage, window.localStorage]) {
    try {
      const mode = storage.getItem(SELECTED_DOCUMENT_MODE_KEY)?.trim();
      if (mode) return mode;
    } catch {
      // Try the other storage.
    }
  }
  return "";
};

export const storeSelectedPersonForDocuments = (
  row: EjournalPreviewRow,
  mode: string,
) => {
  if (typeof window === "undefined") return;
  const payload = JSON.stringify(compactPersonRowForStorage(row));
  const writeLocal = () => {
    window.localStorage.setItem(SELECTED_PERSON_STORAGE_KEY, payload);
    window.localStorage.setItem(SELECTED_DOCUMENT_MODE_KEY, mode);
  };

  try {
    writeLocal();
  } catch {
    try {
      window.localStorage.removeItem(SELECTED_PERSON_STORAGE_KEY);
    } catch {
      // ignore
    }
    let saved = false;
    try {
      writeLocal();
      saved = true;
    } catch {
      saved = false;
    }
    if (!saved) {
      for (const key of QUOTA_RELIEF_KEYS) {
        try {
          window.localStorage.removeItem(key);
        } catch {
          // ignore
        }
      }
      try {
        writeLocal();
      } catch {
        // Documents page falls back to the personnel dataset.
      }
    }
  }

  try {
    window.sessionStorage.setItem(SELECTED_PERSON_STORAGE_KEY, payload);
    window.sessionStorage.setItem(SELECTED_DOCUMENT_MODE_KEY, mode);
  } catch {
    // ignore
  }
};

export const findDocumentPersonRow = (
  rows: EjournalPreviewRow[],
  personId: string,
  rowId = "",
) => {
  const wantedRowId = rowId.trim();
  if (wantedRowId) {
    const byRow = rows.find((row) => String(row.__dbRowId ?? "") === wantedRowId);
    if (byRow) return byRow;
  }
  const wantedId = personId.trim();
  if (!wantedId) return null;
  return (
    rows.find((row) => {
      if (String(row.__dbRowId ?? "") === wantedId) return true;
      if (buildPersonSummary(row).externalId === wantedId) return true;
      return collectPersonExternalIdCandidates(row).includes(wantedId);
    }) ?? null
  );
};
