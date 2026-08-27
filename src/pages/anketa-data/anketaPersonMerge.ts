import { api } from "../../api";
import { invalidatePersonnelCaches } from "../../data/idbDataCache";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  isBlankPersonValue,
  mergePhonesAppendOnly,
} from "../personnel/personEnrichment";
import {
  extractPhones,
  getPersonFieldValue,
  isPositionIndexField,
  PERSON_CARD_FIELDS,
  resolvePersonFieldKey,
  type PersonFieldDef,
} from "../personnel/personnelUtils";
import {
  readStoredPersonPhones,
  upsertPersonPhonesDocument,
  writeStoredPersonPhones,
} from "../personnel/personPhonesStore";
import {
  applyAnketaEditsToRows,
  loadAnketaEdits,
  type AnketaEditsMap,
} from "./anketaEdits";
import {
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnel,
  type AnketaPersonnelMatch,
} from "./anketaPersonMatch";
import {
  isAnketaColumnReadonly,
  loadAnketaSheetPreferCache,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";
import { upsertAnketaCellEdit } from "./anketaEdits";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { resolvePersonBirthDate } from "../personnel/personnelUtils";

type AnketaPersonnelFieldMap = {
  anketaKey: AnketaColumnKey;
  parts: string[];
};

const findPersonFieldDef = (parts: string[]) =>
  PERSON_CARD_FIELDS.find(
    (field) =>
      field.parts.length === parts.length &&
      field.parts.every((part, index) => part === parts[index]),
  );

const ANKETA_PERSONNEL_FIELD_MAP: AnketaPersonnelFieldMap[] = [
  { anketaKey: "rank", parts: ["звання"] },
  { anketaKey: "externalId", parts: ["id"] },
  { anketaKey: "positionIndex", parts: ["індекс", "посади"] },
  { anketaKey: "birthDate", parts: ["дата_народження"] },
  { anketaKey: "birthPlace", parts: ["місце_народження"] },
  { anketaKey: "sex", parts: ["стать"] },
  { anketaKey: "rnokpp", parts: ["рнокпп_за_наявності"] },
  { anketaKey: "rnokppRefuse", parts: ["відмова", "рнокпп"] },
  { anketaKey: "idDocumentName", parts: ["назва", "документа", "посвідчує"] },
  { anketaKey: "idDocumentNumber", parts: ["серія", "номер", "документа", "посвідчує"] },
  { anketaKey: "militaryId", parts: ["військового", "квитка"] },
  { anketaKey: "serviceType", parts: ["вид_служби"] },
  { anketaKey: "location", parts: ["місце_дислокації"] },
  { anketaKey: "arrivedFrom", parts: ["звідки", "прибув"] },
  { anketaKey: "contractFrom", parts: ["укладання", "контракту"] },
  { anketaKey: "contractTo", parts: ["закінчення", "контракту"] },
  { anketaKey: "conscriptedWhen", parts: ["коли", "призваний"] },
  { anketaKey: "conscriptedBy", parts: ["ким", "призваний"] },
  { anketaKey: "education", parts: ["освіта"] },
  { anketaKey: "positionDates", parts: ["дати_прийняття_посади"] },
  { anketaKey: "positionOrderNumber", parts: ["номер_наказу_на_прийняття_посади"] },
  { anketaKey: "enlistDate", parts: ["дата_зарахування_до_списків"] },
  { anketaKey: "enlistOrderDate", parts: ["наказ_про_зарахування_до_списків_-_дата"] },
  { anketaKey: "enlistOrderNumber", parts: ["наказ_про_зарахування_до_списків_-_номер"] },
  { anketaKey: "appointmentOrderDate", parts: ["наказ_на_призначення_на_посаду_-_дата"] },
  { anketaKey: "appointmentOrderNumber", parts: ["наказ_на_призначення_на_посаду_-_номер"] },
  { anketaKey: "appointmentInDate", parts: ["вхідна_дата_наказу_на_призначення_на_посаду"] },
  { anketaKey: "appointmentInNumber", parts: ["вхідний_номер_наказу_на_призначення_на_посаду"] },
  { anketaKey: "rankOrderDate", parts: ["наказ_на_присвоєння_останнього_звання_-_дата"] },
  { anketaKey: "rankOrderNumber", parts: ["наказ_на_присвоєння_останнього_звання_-_номер"] },
  { anketaKey: "rankInDate", parts: ["вхідна_дата_наказу_на_присвоєння_останнього_звання"] },
  { anketaKey: "rankInNumber", parts: ["вхідний_номер_наказу_на_присвоєння_останнього_звання"] },
  { anketaKey: "relatives", parts: ["дані", "родичів"] },
  { anketaKey: "additionalInfo", parts: ["додаткова_інформація"] },
];

const normalizeAnketaValueForPersonnel = (
  rawValue: string,
  field: PersonFieldDef | undefined,
) => {
  const trimmed = rawValue.trim();
  if (!trimmed) return "";

  if (
    field &&
    (field.kind === "positionIndex" || isPositionIndexField(field.parts))
  ) {
    return trimmed
      .split(/\s*[·,;]\s*|\s+/)
      .map((part) => part.trim())
      .filter(Boolean)
      .join("\n");
  }

  if (field?.parts.includes("рнокпп_за_наявності")) {
    const digits = trimmed.replace(/\D/g, "");
    return digits.length >= 8 ? digits : trimmed;
  }

  return trimmed.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
};

export type AnketaPersonnelMergePreview = {
  fieldUpdates: Record<string, string>;
  labels: string[];
};

/** Побудувати оновлення для ООС: лише порожні поля заповнюються з анкети. */
export const buildAnketaPersonFieldUpdates = (
  personnelRow: EjournalPreviewRow,
  anketaRow: AnketaRow,
): AnketaPersonnelMergePreview => {
  const fieldUpdates: Record<string, string> = {};
  const labels: string[] = [];

  for (const { anketaKey, parts } of ANKETA_PERSONNEL_FIELD_MAP) {
    if (isAnketaColumnReadonly(anketaKey)) continue;

    const raw = String(anketaRow[anketaKey] ?? "").trim();
    if (!raw) continue;

    const field = findPersonFieldDef(parts);
    const current = getPersonFieldValue(personnelRow, parts);
    if (!isBlankPersonValue(current)) continue;

    const targetKey = resolvePersonFieldKey(personnelRow, parts) || parts.join("_");
    const nextValue = normalizeAnketaValueForPersonnel(raw, field);
    if (!nextValue) continue;

    fieldUpdates[targetKey] = nextValue;
    labels.push(field?.label ?? parts.join(" "));
  }

  return { fieldUpdates, labels };
};

export type AnketaPersonnelMergeResult = {
  fieldUpdates: Record<string, string>;
  phonesAdded: string[];
  skippedReason?: "no-match" | "no-row-id" | "no-updates";
};

export const syncAnketaRowToPersonnel = async (options: {
  anketaRow: AnketaRow;
  match: AnketaPersonnelMatch;
  skipCacheInvalidation?: boolean;
}): Promise<AnketaPersonnelMergeResult> => {
  const rowId = options.match.row.__dbRowId
    ? String(options.match.row.__dbRowId)
    : "";
  if (!rowId || rowId.startsWith("roster:")) {
    return { fieldUpdates: {}, phonesAdded: [], skippedReason: "no-row-id" };
  }

  const preview = buildAnketaPersonFieldUpdates(
    options.match.row,
    options.anketaRow,
  );
  const additionalInfo = String(options.anketaRow.additionalInfo ?? "").trim();
  const incomingPhones = additionalInfo ? extractPhones(additionalInfo) : [];
  const personExternalId = options.match.summary.externalId.trim();

  let phonesAdded: string[] = [];
  if (personExternalId && incomingPhones.length) {
    const store = readStoredPersonPhones();
    const before = store[personExternalId] ?? [];
    const merged = mergePhonesAppendOnly(before, incomingPhones);
    phonesAdded = merged.filter((phone) => !before.includes(phone));
    if (phonesAdded.length) {
      store[personExternalId] = merged;
      writeStoredPersonPhones(store);
      await upsertPersonPhonesDocument(personExternalId, merged);
    }
  }

  if (!Object.keys(preview.fieldUpdates).length && !phonesAdded.length) {
    return { fieldUpdates: {}, phonesAdded: [], skippedReason: "no-updates" };
  }

  if (Object.keys(preview.fieldUpdates).length) {
    await api.updateEjournalRowValues(rowId, preview.fieldUpdates);
    if (!options.skipCacheInvalidation) {
      await invalidatePersonnelCaches();
    }
  }

  return { fieldUpdates: preview.fieldUpdates, phonesAdded };
};

export type AnketaBulkMergeReport = {
  processed: number;
  matched: number;
  updated: number;
  skippedNoMatch: number;
  skippedNoRowId: number;
  skippedNoUpdates: number;
  fieldCount: number;
  phonesAdded: number;
  errors: Array<{ name: string; message: string }>;
};

export const mergeAnketaRowsToPersonnel = async (options: {
  rows: AnketaRow[];
  edits?: AnketaEditsMap;
  onProgress?: (done: number, total: number) => void;
}): Promise<AnketaBulkMergeReport> => {
  const mergedRows =
    options.edits && Object.keys(options.edits).length
      ? applyAnketaEditsToRows(options.rows, options.edits)
      : options.rows;
  const index = await loadPersonnelIndexForAnketa({ force: true });

  const report: AnketaBulkMergeReport = {
    processed: 0,
    matched: 0,
    updated: 0,
    skippedNoMatch: 0,
    skippedNoRowId: 0,
    skippedNoUpdates: 0,
    fieldCount: 0,
    phonesAdded: 0,
    errors: [],
  };

  for (let indexOffset = 0; indexOffset < mergedRows.length; indexOffset += 1) {
    const anketaRow = mergedRows[indexOffset]!;
    report.processed += 1;
    options.onProgress?.(report.processed, mergedRows.length);

    const match = matchAnketaRowToPersonnel(anketaRow, index);
    if (!match) {
      report.skippedNoMatch += 1;
      continue;
    }
    report.matched += 1;

    try {
      const result = await syncAnketaRowToPersonnel({
        anketaRow,
        match,
        skipCacheInvalidation: true,
      });
      if (result.skippedReason === "no-row-id") {
        report.skippedNoRowId += 1;
        continue;
      }
      if (result.skippedReason === "no-updates") {
        report.skippedNoUpdates += 1;
        continue;
      }
      report.updated += 1;
      report.fieldCount += Object.keys(result.fieldUpdates).length;
      report.phonesAdded += result.phonesAdded.length;
    } catch (error) {
      report.errors.push({
        name: anketaRow.fullName?.trim() || `рядок ${anketaRow.__rowNumber}`,
        message:
          error instanceof Error ? error.message : "Невідома помилка злиття",
      });
    }
  }

  if (report.updated > 0 || report.phonesAdded > 0) {
    await invalidatePersonnelCaches();
  }

  return report;
};

export const formatAnketaBulkMergeReport = (report: AnketaBulkMergeReport) => {
  const parts = [
    `оновлено осіб: ${report.updated}`,
    `полів: ${report.fieldCount}`,
    report.phonesAdded ? `телефонів: ${report.phonesAdded}` : "",
    report.skippedNoMatch ? `без збігу: ${report.skippedNoMatch}` : "",
    report.skippedNoUpdates ? `без нових даних: ${report.skippedNoUpdates}` : "",
    report.skippedNoRowId ? `лише ранковий список: ${report.skippedNoRowId}` : "",
    report.errors.length ? `помилок: ${report.errors.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

/** Завантажує кеш/Google «Анкети» + локальні правки й доповнює порожні поля в ООС. */
export const mergeCachedAnketaToPersonnel = async (options?: {
  onProgress?: (done: number, total: number) => void;
}): Promise<AnketaBulkMergeReport> => {
  const snapshot = await loadAnketaSheetPreferCache();
  if (!snapshot.rows.length) {
    throw new Error(
      "Немає анкетних даних. Відкрийте «Анкетні дані» або оновіть таблицю з Google.",
    );
  }
  const edits = await loadAnketaEdits();
  return mergeAnketaRowsToPersonnel({
    rows: snapshot.rows,
    edits,
    onProgress: options?.onProgress,
  });
};

export const previewAnketaRowPersonnelMerge = (
  anketaRow: AnketaRow,
  match: AnketaPersonnelMatch | null,
) => {
  if (!match) return null;
  return buildAnketaPersonFieldUpdates(match.row, anketaRow);
};

const readPersonnelValueForAnketa = (
  personnelRow: EjournalPreviewRow,
  anketaKey: AnketaColumnKey,
  parts: string[],
) => {
  const fromOos = getPersonFieldValue(personnelRow, parts).trim();
  if (fromOos) return fromOos;

  if (anketaKey === "rnokpp") {
    return readRosterColumnValue(personnelRow, 19).trim();
  }
  if (anketaKey === "birthDate") {
    return (
      resolvePersonBirthDate(personnelRow) ||
      readRosterColumnValue(personnelRow, 16) ||
      readRosterColumnValue(personnelRow, 17)
    ).trim();
  }
  if (anketaKey === "rank") {
    return readRosterColumnValue(personnelRow, 13).trim();
  }
  if (anketaKey === "militaryId") {
    return readRosterColumnValue(personnelRow, 11).trim();
  }
  return "";
};

export type PersonnelToAnketaFieldUpdate = {
  columnId: AnketaColumnKey;
  value: string;
  label: string;
};

/** Доповнити порожні поля анкети з особового складу / ранкового списку. */
export const buildPersonnelToAnketaFieldUpdates = (
  personnelRow: EjournalPreviewRow,
  anketaRow: AnketaRow,
): PersonnelToAnketaFieldUpdate[] => {
  const updates: PersonnelToAnketaFieldUpdate[] = [];

  for (const { anketaKey, parts } of ANKETA_PERSONNEL_FIELD_MAP) {
    if (isAnketaColumnReadonly(anketaKey)) continue;
    const current = String(anketaRow[anketaKey] ?? "").trim();
    if (current) continue;

    const raw = readPersonnelValueForAnketa(personnelRow, anketaKey, parts);
    if (!raw) continue;

    const field = findPersonFieldDef(parts);
    let nextValue = raw.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();
    if (anketaKey === "rnokpp") {
      const digits = nextValue.replace(/\D/g, "");
      if (digits.length >= 8) nextValue = digits;
    }

    updates.push({
      columnId: anketaKey,
      value: nextValue,
      label: field?.label ?? parts.join(" "),
    });
  }

  return updates;
};

export type PersonnelToAnketaBulkMergeReport = {
  processed: number;
  matched: number;
  updated: number;
  fieldCount: number;
  skippedNoMatch: number;
  skippedNoUpdates: number;
  errors: Array<{ name: string; message: string }>;
};

export const mergePersonnelToAnketaRows = async (options: {
  rows: AnketaRow[];
  onProgress?: (done: number, total: number) => void;
}): Promise<PersonnelToAnketaBulkMergeReport> => {
  const index = await loadPersonnelIndexForAnketa({ force: true });
  const report: PersonnelToAnketaBulkMergeReport = {
    processed: 0,
    matched: 0,
    updated: 0,
    fieldCount: 0,
    skippedNoMatch: 0,
    skippedNoUpdates: 0,
    errors: [],
  };

  for (const anketaRow of options.rows) {
    report.processed += 1;
    options.onProgress?.(report.processed, options.rows.length);

    const match = matchAnketaRowToPersonnel(anketaRow, index);
    if (!match) {
      report.skippedNoMatch += 1;
      continue;
    }
    report.matched += 1;

    const updates = buildPersonnelToAnketaFieldUpdates(match.row, anketaRow);
    if (!updates.length) {
      report.skippedNoUpdates += 1;
      continue;
    }

    let wroteAny = false;
    for (const update of updates) {
      try {
        await upsertAnketaCellEdit({
          rowNumber: anketaRow.__rowNumber,
          columnId: update.columnId,
          value: update.value,
          externalId: String(anketaRow.externalId ?? "").trim() || undefined,
          fullName: anketaRow.fullName,
        });
        report.fieldCount += 1;
        wroteAny = true;
      } catch (error) {
        report.errors.push({
          name: anketaRow.fullName?.trim() || `рядок ${anketaRow.__rowNumber}`,
          message:
            error instanceof Error
              ? error.message
              : "Не вдалося записати поле анкети",
        });
      }
    }
    if (wroteAny) report.updated += 1;
  }

  return report;
};

export const formatPersonnelToAnketaMergeReport = (
  report: PersonnelToAnketaBulkMergeReport,
) => {
  const parts = [
    `оновлено анкет: ${report.updated}`,
    `полів: ${report.fieldCount}`,
    report.skippedNoMatch ? `без збігу в ООС: ${report.skippedNoMatch}` : "",
    report.skippedNoUpdates ? `без нових даних: ${report.skippedNoUpdates}` : "",
    report.errors.length ? `помилок: ${report.errors.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

export const mergePersonnelToCachedAnketa = async (options?: {
  onProgress?: (done: number, total: number) => void;
}): Promise<PersonnelToAnketaBulkMergeReport> => {
  const snapshot = await loadAnketaSheetPreferCache();
  if (!snapshot.rows.length) {
    throw new Error(
      "Немає анкетних даних. Оновіть таблицю з Google або імпортуйте CSV.",
    );
  }
  const edits = await loadAnketaEdits();
  const rows =
    edits && Object.keys(edits).length
      ? applyAnketaEditsToRows(snapshot.rows, edits)
      : snapshot.rows;
  return mergePersonnelToAnketaRows({ rows, onProgress: options?.onProgress });
};
