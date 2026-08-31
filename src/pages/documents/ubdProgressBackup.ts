import type { BackendPersonDocument } from "../../api";

export type UbdProgressBackupEntry = {
  personExternalId: string;
  documentId: string;
  title: string;
  status: string;
  workflow: Record<string, unknown>;
};

export type UbdProgressBackup = {
  id: string;
  savedAt: number;
  reason: "journal-load" | "before-bulk-apply" | "manual";
  label: string;
  entries: UbdProgressBackupEntry[];
};

const BACKUPS_STORAGE_KEY = "army-grid:ubd-progress-backups";
const UNDO_STORAGE_KEY = "army-grid:ubd-bulk-progress-undo";
const SESSION_BACKUP_FLAG = "army-grid:ubd-journal-backup-saved";
const MAX_BACKUPS = 15;

const readJson = <T>(key: string): T | null => {
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : null;
  } catch {
    return null;
  }
};

const writeJson = (key: string, value: unknown) => {
  try {
    window.localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* ignore quota / private mode */
  }
};

export const documentToUbdProgressBackupEntry = (
  document: BackendPersonDocument,
  workflow: Record<string, unknown>,
): UbdProgressBackupEntry => ({
  personExternalId: document.personExternalId,
  documentId: document.id,
  title: document.title,
  status: String(document.status || workflow.currentStatus || "document"),
  workflow: JSON.parse(JSON.stringify(workflow)),
});

export const loadUbdProgressBackups = (): UbdProgressBackup[] => {
  const backups = readJson<UbdProgressBackup[]>(BACKUPS_STORAGE_KEY);
  return Array.isArray(backups) ? backups : [];
};

export const saveUbdProgressBackup = (
  input: Omit<UbdProgressBackup, "id" | "savedAt">,
): UbdProgressBackup => {
  const backup: UbdProgressBackup = {
    ...input,
    id: crypto.randomUUID(),
    savedAt: Date.now(),
  };
  const next = [backup, ...loadUbdProgressBackups()].slice(0, MAX_BACKUPS);
  writeJson(BACKUPS_STORAGE_KEY, next);
  return backup;
};

export const deleteUbdProgressBackup = (backupId: string) => {
  writeJson(
    BACKUPS_STORAGE_KEY,
    loadUbdProgressBackups().filter((backup) => backup.id !== backupId),
  );
};

export const loadPersistedUbdBulkProgressUndo =
  (): UbdProgressBackupEntry[] | null => {
    const entries = readJson<UbdProgressBackupEntry[]>(UNDO_STORAGE_KEY);
    return Array.isArray(entries) && entries.length ? entries : null;
  };

export const savePersistedUbdBulkProgressUndo = (
  entries: UbdProgressBackupEntry[] | null,
) => {
  if (!entries?.length) {
    window.localStorage.removeItem(UNDO_STORAGE_KEY);
    return;
  }
  writeJson(UNDO_STORAGE_KEY, entries);
};

export const saveSessionJournalUbdBackup = (
  documents: BackendPersonDocument[],
  toEntry: (document: BackendPersonDocument) => UbdProgressBackupEntry | null,
) => {
  if (sessionStorage.getItem(SESSION_BACKUP_FLAG)) return null;
  const entries = documents
    .filter((document) => document.type === "ubdReport")
    .map(toEntry)
    .filter((entry): entry is UbdProgressBackupEntry => entry != null);
  if (!entries.length) return null;
  sessionStorage.setItem(SESSION_BACKUP_FLAG, "1");
  return saveUbdProgressBackup({
    reason: "journal-load",
    label: `Відкриття журналу · ${entries.length} рапортів`,
    entries,
  });
};

/** Додає кроки, які можна довести з файлів/реквізитів у workflow. */
export const inferUbdWorkflowCompletedFromArtifacts = (
  document: BackendPersonDocument,
): Record<string, boolean> => {
  const workflow =
    document.workflow && typeof document.workflow === "object"
      ? (document.workflow as Record<string, unknown>)
      : {};
  const completed =
    workflow.completed && typeof workflow.completed === "object"
      ? (workflow.completed as Record<string, boolean>)
      : {};
  const files =
    document.files && typeof document.files === "object"
      ? (document.files as { ubdScans?: unknown[] })
      : {};

  const accountFileName = String(workflow.accountFileName ?? "").trim();
  const signedScanFileName = String(workflow.signedScanFileName ?? "").trim();
  const mergedPdfFileName = String(workflow.mergedPdfFileName ?? "").trim();
  const hasScans =
    Array.isArray(files.ubdScans) && files.ubdScans.length > 0;

  return {
    document: true,
    account: completed.account === true || Boolean(accountFileName),
    scan:
      completed.scan === true ||
      Boolean(signedScanFileName) ||
      hasScans,
    ready: completed.ready === true || Boolean(mergedPdfFileName),
    sentReport:
      completed.sentReport === true || completed.sent === true,
    sentScans: completed.sentScans === true,
    received: completed.received === true,
    handed: completed.handed === true,
  };
};

export const formatUbdProgressBackupWhen = (savedAt: number) =>
  new Date(savedAt).toLocaleString("uk-UA");
