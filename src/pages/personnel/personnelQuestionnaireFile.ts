import { cleanPersonDisplayName, dataUrlToFile, downloadQuestionnairePdf, extractBirthDateFromPersonName, normalizePersonBirthKey } from "./personnelUtils";

export const matchesQuestionnairePerson = (
  match: { fileName: string; relativePath?: string },
  fullName: string,
) => {
  const normalize = (value: string) => value.normalize("NFC")
    .toLocaleLowerCase("uk-UA").replace(/[^\p{L}\p{N}]+/gu, " ").trim();
  const name = normalize(cleanPersonDisplayName(fullName));
  const fileName = normalize(match.fileName);
  if (!name || !(` ${fileName} `).includes(` ${name} `)) return false;
  const dateKey = (value: string) => /^\d{4}-\d{2}-\d{2}$/.test(value)
    ? value : normalizePersonBirthKey(value);
  const birthKey = extractBirthDateFromPersonName(fullName);
  if (!birthKey) return true;
  const dates = `${match.relativePath ?? ""} ${match.fileName}`
    .match(/\d{4}-\d{2}-\d{2}|\d{1,2}[.\/-]\d{1,2}[.\/-]\d{4}/g) ?? [];
  return dates.some((date) => dateKey(date) === birthKey);
};

export const MAX_QUESTIONNAIRE_FILE_BYTES = 350 * 1024 * 1024;

export const formatQuestionnaireFileSize = (bytes: number) => {
  if (!Number.isFinite(bytes)) return "";
  const mb = bytes / 1024 / 1024;
  return `${mb.toFixed(mb >= 10 ? 0 : 1)} MB`;
};

export const questionnaireTooLargeMessage = (bytes: number) =>
  `PDF завеликий для збереження в БД: ${formatQuestionnaireFileSize(bytes)}. Максимум: ${formatQuestionnaireFileSize(MAX_QUESTIONNAIRE_FILE_BYTES)}.`;

export const questionnaireCropFile = ({
  diskPreviewFile,
  pendingQuestionnaireFile,
  questionnairePreviewFile,
  fileData,
  fileName,
}: {
  diskPreviewFile: File | null;
  pendingQuestionnaireFile: File | null;
  questionnairePreviewFile: File | null;
  fileData?: string | null;
  fileName?: string | null;
}) => {
  if (diskPreviewFile) return diskPreviewFile;
  if (pendingQuestionnaireFile) return pendingQuestionnaireFile;
  if (questionnairePreviewFile) return questionnairePreviewFile;
  if (!fileData?.trim()) return null;
  return dataUrlToFile(fileData, fileName || "questionnaire.pdf");
};

export const downloadVisibleQuestionnaire = (
  fileName: string,
  sources: {
    pendingFile?: File | null;
    diskFile?: File | null;
    fileData?: string | null;
  },
) => {
  if (sources.pendingFile) {
    downloadQuestionnairePdf(fileName, { file: sources.pendingFile });
    return true;
  }
  if (sources.diskFile) {
    downloadQuestionnairePdf(fileName, { file: sources.diskFile });
    return true;
  }
  if (sources.fileData) {
    downloadQuestionnairePdf(fileName, { fileData: sources.fileData });
    return true;
  }
  return false;
};

export const pickDiskQuestionnaireMatch = <T extends { fileName: string }>(
  matches: T[],
  storedFileName: string,
) => {
  const stored = storedFileName.normalize("NFC").toLocaleLowerCase("uk-UA");
  return (
    matches.find(
      (item) =>
        item.fileName.normalize("NFC").toLocaleLowerCase("uk-UA") === stored,
    ) ?? matches[0]
  );
};
