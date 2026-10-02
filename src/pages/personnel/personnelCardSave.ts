import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { isPositionIndexField } from "./personnelUtils";

export const personCardValuesForSave = (
  fields: Array<{ key: string; parts: string[] }>,
  editValues: Record<string, string>,
) =>
  Object.fromEntries(
    fields.map((field) => {
      const raw = editValues[field.key] ?? "";
      if (!isPositionIndexField(field.parts)) return [field.key, raw];
      return [
        field.key,
        raw
          .split(/\s*[·,;]\s*|\s+/)
          .map((part) => part.trim())
          .filter(Boolean)
          .join("\n"),
      ];
    }),
  );

export const cardSavedMessage = (name: string) => `Картку оновлено: ${name}.`;

export const cardSaveError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося зберегти картку особи.";

export const mergeSavedPersonRow = (
  row: EjournalPreviewRow,
  rowId: string,
  values: Record<string, unknown>,
): EjournalPreviewRow =>
  row.__dbRowId === rowId
    ? { ...row, ...values, __dbRowId: rowId }
    : row;

export const replaceSavedPersonnelRows = (
  rows: EjournalPreviewRow[],
  rowId: string,
  values: Record<string, unknown>,
) => rows.map((row) => mergeSavedPersonRow(row, rowId, values));
