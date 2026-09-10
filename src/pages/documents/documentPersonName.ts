import type { BackendPersonDocument } from "../../api";
import { personNameFromSyntheticDocumentId } from "./documentPersonIdentity";

const readDocumentFieldText = (
  fields: Record<string, unknown>,
  key: string,
) => {
  const value = fields[key];
  return typeof value === "string" ? value.trim() : "";
};

/** Display name for a journal / person document row. */
export const getDocumentPersonName = (document: BackendPersonDocument) => {
  const metadataName = String(document.personName ?? "").trim();
  if (metadataName) return metadataName;

  const fields = (document.fields || {}) as Record<string, unknown>;
  const fullName =
    readDocumentFieldText(fields, "fullName") ||
    readDocumentFieldText(fields, "pib") ||
    readDocumentFieldText(fields, "name") ||
    readDocumentFieldText(fields, "ПІБ") ||
    readDocumentFieldText(fields, "ФИО");
  if (fullName) return fullName;

  const assembled = [
    readDocumentFieldText(fields, "lastName"),
    readDocumentFieldText(fields, "firstName"),
    readDocumentFieldText(fields, "patronymic"),
  ]
    .filter(Boolean)
    .join(" ");
  if (assembled) return assembled;

  const syntheticName = personNameFromSyntheticDocumentId(
    document.personExternalId,
  );
  if (syntheticName) return syntheticName;

  return document.personExternalId
    ? `ID ${document.personExternalId}`
    : "Без ПІБ";
};
