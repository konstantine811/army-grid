import { api } from "../../api";
import { downloadBlob } from "../../shared/browserExport";
import {
  downloadVisibleQuestionnaire,
  pickDiskQuestionnaireMatch,
  matchesQuestionnairePerson,
} from "./personnelQuestionnaireFile";

export const questionnaireExportError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося експортувати анкету.";

export const questionnaireRevealError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося показати анкету у Finder";

export const exportPersonnelQuestionnaire = async ({
  externalId,
  storedPersonId,
  fileName,
  fileData,
  pendingFile,
  diskFile,
}: {
  externalId: string;
  storedPersonId: string;
  fileName: string;
  fileData?: string | null;
  pendingFile: File | null;
  diskFile: File | null;
}) => {
  if (externalId && fileData && !pendingFile && !diskFile) {
    const blob = await api.fetchPersonQuestionnaireFile(
      storedPersonId,
      fileName,
      true,
    );
    downloadBlob(blob, fileName);
    return `Експортовано: ${fileName}`;
  }

  downloadVisibleQuestionnaire(fileName, {
    pendingFile,
    diskFile,
    fileData,
  });
  if (pendingFile || diskFile || fileData) return `Експортовано: ${fileName}`;
  return "";
};

export const revealPersonnelQuestionnaire = async ({
  externalId,
  fullName,
  callSign,
  storedFileName,
}: {
  externalId: string;
  fullName: string;
  callSign: string;
  storedFileName: string;
}) => {
  const person = { externalId, fullName, callSign };
  let result = await api.searchQuestionnairesOnDisk({
    people: [person],
    refreshIndex: false,
  });
  const exactMatches = (items: typeof result.people) =>
    (items[0]?.matches ?? []).filter((match) =>
      matchesQuestionnairePerson(match, fullName),
    );
  let matches = exactMatches(result.people);
  if (!matches.length) {
    result = await api.searchQuestionnairesOnDisk({
      people: [person],
      refreshIndex: true,
    });
    matches = exactMatches(result.people);
  }

  const match = pickDiskQuestionnaireMatch(matches, storedFileName);
  if (!match) {
    throw new Error(
      "Оригінальний PDF за повним ПІБ вибраної особи не знайдено у папці анкет.",
    );
  }

  await api.revealDiskQuestionnaireInFinder(match.relativePath);
  return `Відкрито у Finder: ${match.fileName}`;
};
