import { api, type BackendAnketaCellEdit } from "../../api";
import {
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import {
  ANKETA_COLUMNS,
  ANKETA_SHEET_GID,
  ANKETA_SHEET_ID,
  type AnketaColumnKey,
  type AnketaRow,
  type AnketaSheetSnapshot,
} from "./anketaSheet";

export const ANKETA_EDITS_CACHE_KEY = "anketa:cell-edits:v1";

export type AnketaCellEdit = {
  rowNumber: number;
  columnId: AnketaColumnKey;
  value: string;
  updatedAt: string;
  externalId?: string;
  fullName?: string;
};

export type AnketaEditsMap = Record<string, AnketaCellEdit>;

export type UpsertAnketaCellEditResult = {
  edits: AnketaEditsMap;
  serverSynced: boolean;
  serverError?: string;
};

export const anketaEditKey = (
  rowNumber: number,
  columnId: AnketaColumnKey | string,
) => `${rowNumber}:${columnId}`;

const backendToLocal = (record: BackendAnketaCellEdit): AnketaCellEdit => ({
  rowNumber: record.rowNumber,
  columnId: record.columnId as AnketaColumnKey,
  value: record.value,
  updatedAt: record.updatedAt,
  externalId: record.externalId?.trim() || undefined,
  fullName: record.fullName?.trim() || undefined,
});

const localToUpsertPayload = (edit: AnketaCellEdit) => ({
  rowNumber: edit.rowNumber,
  columnId: edit.columnId,
  value: edit.value,
  externalId: edit.externalId,
  fullName: edit.fullName,
  sheetId: ANKETA_SHEET_ID,
  gid: ANKETA_SHEET_GID,
});

const mergeEditsMaps = (
  local: AnketaEditsMap,
  remote: AnketaEditsMap,
): AnketaEditsMap => {
  const merged = { ...local };
  for (const [key, remoteEdit] of Object.entries(remote)) {
    const localEdit = merged[key];
    if (
      !localEdit ||
      new Date(remoteEdit.updatedAt).getTime() >=
        new Date(localEdit.updatedAt).getTime()
    ) {
      merged[key] = remoteEdit;
    }
  }
  return merged;
};

const loadLocalAnketaEdits = async (): Promise<AnketaEditsMap> => {
  const stored = await readDataCache<AnketaEditsMap>(ANKETA_EDITS_CACHE_KEY);
  if (!stored || typeof stored !== "object") return {};
  return stored;
};

export const saveAnketaEdits = async (edits: AnketaEditsMap) => {
  await writeDataCache(ANKETA_EDITS_CACHE_KEY, edits);
};

/** Завантажити правки: локальний кеш + синхронізація з сервером PostgreSQL. */
export const syncAnketaEdits = async (): Promise<AnketaEditsMap> => {
  const local = await loadLocalAnketaEdits();

  let remoteRecords: BackendAnketaCellEdit[] = [];
  try {
    remoteRecords = await api.listAnketaEdits(ANKETA_SHEET_ID, ANKETA_SHEET_GID);
  } catch {
    return local;
  }

  const remote: AnketaEditsMap = {};
  for (const record of remoteRecords) {
    const edit = backendToLocal(record);
    remote[anketaEditKey(edit.rowNumber, edit.columnId)] = edit;
  }

  const toPush = Object.entries(local)
    .filter(([key, edit]) => {
      const remoteEdit = remote[key];
      return (
        !remoteEdit ||
        new Date(edit.updatedAt).getTime() >
          new Date(remoteEdit.updatedAt).getTime()
      );
    })
    .map(([, edit]) => edit);

  if (toPush.length) {
    try {
      const pushed = await api.bulkUpsertAnketaEdits(
        toPush.map(localToUpsertPayload),
      );
      for (const record of pushed.items) {
        const edit = backendToLocal(record);
        remote[anketaEditKey(edit.rowNumber, edit.columnId)] = edit;
      }
    } catch {
      /* локальні правки залишаються, сервер недоступний */
    }
  }

  const merged = mergeEditsMaps(local, remote);
  await saveAnketaEdits(merged);
  return merged;
};

export const loadAnketaEdits = syncAnketaEdits;

export const upsertAnketaCellEdit = async (input: {
  rowNumber: number;
  columnId: AnketaColumnKey;
  value: string;
  externalId?: string;
  fullName?: string;
}): Promise<UpsertAnketaCellEditResult> => {
  const current = await loadLocalAnketaEdits();
  const key = anketaEditKey(input.rowNumber, input.columnId);
  const next: AnketaEditsMap = {
    ...current,
    [key]: {
      rowNumber: input.rowNumber,
      columnId: input.columnId,
      value: input.value,
      updatedAt: new Date().toISOString(),
      externalId: input.externalId?.trim() || undefined,
      fullName: input.fullName?.trim() || undefined,
    },
  };
  await saveAnketaEdits(next);

  try {
    const saved = await api.upsertAnketaCellEdit({
      rowNumber: input.rowNumber,
      columnId: input.columnId,
      value: input.value,
      externalId: input.externalId,
      fullName: input.fullName,
      sheetId: ANKETA_SHEET_ID,
      gid: ANKETA_SHEET_GID,
    });
    next[key] = backendToLocal(saved);
    await saveAnketaEdits(next);
    return { edits: next, serverSynced: true };
  } catch (error) {
    return {
      edits: next,
      serverSynced: false,
      serverError:
        error instanceof Error ? error.message : "Сервер недоступний",
    };
  }
};

export const applyAnketaEditsToRows = (
  rows: AnketaRow[],
  edits: AnketaEditsMap,
): AnketaRow[] => {
  if (!edits || !Object.keys(edits).length) return rows;

  return rows.map((row) => {
    let next = row;
    let changed = false;
    for (const column of ANKETA_COLUMNS) {
      const edit = edits[anketaEditKey(row.__rowNumber, column.key)];
      if (!edit) continue;
      if (String(next[column.key] ?? "") === edit.value) continue;
      if (!changed) {
        next = { ...row };
        changed = true;
      }
      next[column.key] = edit.value;
    }
    return next;
  });
};

export const applyAnketaEditsToSnapshot = (
  snapshot: AnketaSheetSnapshot,
  edits: AnketaEditsMap,
): AnketaSheetSnapshot => ({
  ...snapshot,
  rows: applyAnketaEditsToRows(snapshot.rows, edits),
});

export const countAnketaEdits = (edits: AnketaEditsMap) =>
  Object.keys(edits).length;
