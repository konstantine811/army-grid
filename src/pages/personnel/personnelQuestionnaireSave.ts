import { api } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { sanitizeFileName } from "../../shared/browserExport";
import {
  collectPersonAttachmentLookupIds,
  resolveCanonicalPersonId,
} from "./personAttachments";
import {
  buildQuestionnaireExportFileName,
  renameQuestionnaireFile,
} from "./personnelUtils";

export const questionnaireDeletedMessage = (name: string) =>
  `Анкету видалено: ${name}.`;

export const questionnaireDeleteError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося видалити анкету з БД.";

export const questionnaireDeleteIds = (
  externalId: string,
  row: EjournalPreviewRow | null,
) =>
  [
    ...new Set([
      externalId,
      ...collectPersonAttachmentLookupIds(row, undefined, {
        includeLooseKeys: true,
      }),
    ]),
  ].filter(Boolean);

export const deleteStoredQuestionnaires = (ids: string[]) =>
  Promise.all(
    ids.map((id) => api.deletePersonQuestionnaire(id).catch(() => undefined)),
  );

export const questionnaireSaveMissingIdMessage =
  "Не вдалося зберегти анкету: у вибраної особи немає ID.";

export const questionnaireSavedMessage = (name: string, fileName: string) =>
  `Анкету збережено в БД: ${name} · ${fileName}.`;

export const questionnaireSaveError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося зберегти анкету в БД.";

export const saveQuestionnaireFile = async ({
  row,
  externalId,
  name,
  callSign,
  file,
}: {
  row: EjournalPreviewRow | null;
  externalId: string;
  name: string;
  callSign: string;
  file: File;
}) => {
  const exportFileName = sanitizeFileName(
    buildQuestionnaireExportFileName(name, callSign),
  );
  const fileToSave = renameQuestionnaireFile(file, exportFileName);
  const storageId =
    resolveCanonicalPersonId(collectPersonAttachmentLookupIds(row)) ||
    externalId;
  const saved = await api.upsertPersonQuestionnaireFile(storageId, fileToSave);
  return { saved, exportFileName };
};
