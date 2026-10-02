import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import type { EditablePersonField } from "./card/personCardTypes";
import { uniqueNormalizedPhones } from "./personPhonesStore";
import {
  formatPersonFieldValue,
  resolvePersonBirthDate,
} from "./personnelUtils";

export const personEditDraft = (
  fields: EditablePersonField[],
  row: EjournalPreviewRow | null,
) =>
  Object.fromEntries(
    fields.map((field) => {
      const raw = row?.[field.key];
      let text = formatPersonFieldValue(raw, field);
      if (
        field.parts.includes("дата_народження") &&
        !String(text ?? "").trim()
      ) {
        text = resolvePersonBirthDate(row);
      }
      return [field.key, text];
    }),
  );

export const shouldCopyFoundQuestionnaire = (
  canEdit: boolean,
  found: boolean,
  resolvedExternalId: string,
  listExternalId: string,
) =>
  canEdit &&
  found &&
  Boolean(resolvedExternalId) &&
  resolvedExternalId !== listExternalId &&
  !listExternalId.startsWith("p:");

export const mergedDocumentPhones = (
  current: string[] | undefined,
  incoming: string[],
) => {
  const existing = current ?? [];
  const merged = uniqueNormalizedPhones([...existing, ...incoming]);
  const same =
    merged.length === existing.length &&
    merged.every((phone, index) => phone === existing[index]);
  return same ? null : merged;
};
