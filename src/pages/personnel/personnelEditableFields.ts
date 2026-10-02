import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import type { EditablePersonField, EditablePersonSection } from "./card/personCardTypes";
import {
  PERSON_CARD_FIELDS,
  PERSON_SECTION_LABELS,
  formatPersonFieldValue,
  looksLikePersonBirthDate,
  normalizePersonBirthKey,
  resolvePersonFieldKey,
  type PersonFieldDef,
} from "./personnelUtils";

const EDITABLE_SECTIONS: PersonFieldDef["section"][] = [
  "identity",
  "service",
  "orders",
  "contacts",
];

export const resolveEditablePersonFields = (
  row: EjournalPreviewRow | null,
): EditablePersonField[] =>
  PERSON_CARD_FIELDS.map((field) => ({
    ...field,
    key: resolvePersonFieldKey(row, field.parts),
  })).filter((field): field is EditablePersonField => Boolean(field.key));

const hideBecauseBirthId = (
  field: EditablePersonField,
  row: EjournalPreviewRow | null,
  birthDate: string,
) => {
  if (!field.parts.includes("id")) return false;
  const idValue = formatPersonFieldValue(row?.[field.key], field).trim();
  if (looksLikePersonBirthDate(idValue)) return true;
  return (
    Boolean(birthDate) &&
    normalizePersonBirthKey(idValue) === normalizePersonBirthKey(birthDate)
  );
};

export const groupEditablePersonFields = (
  fields: EditablePersonField[],
  row: EjournalPreviewRow | null,
  birthDate: string,
): EditablePersonSection[] => {
  const birth = String(birthDate ?? "").trim();
  return EDITABLE_SECTIONS.map((section) => ({
    section,
    label: PERSON_SECTION_LABELS[section],
    fields: fields.filter((field) => {
      if (field.section !== section) return false;
      if (field.parts.includes("місце_перебування")) return false;
      if (hideBecauseBirthId(field, row, birth)) return false;
      return true;
    }),
  })).filter((group) => group.fields.length > 0);
};
