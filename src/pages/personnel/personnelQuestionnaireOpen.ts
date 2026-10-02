import { api } from "../../api";
import { MAX_QUESTIONNAIRE_FILE_BYTES, questionnaireTooLargeMessage } from "./personnelQuestionnaireFile";
import { normalizeRosterMatchText } from "./fighterStatusImport";
import { dataUrlToFile, dataUrlToObjectUrl } from "./personnelUtils";

export const questionnairePreviewHeading = (name: string, fileName: string) =>
  `Анкета · ${name}${fileName ? ` · ${fileName}` : ""}`;

export const questionnaireReviewHeading = (name: string, fileName: string) =>
  `Перегляд анкети перед збереженням · ${name} · ${fileName}`;

export const isCallsignPersonId = (externalId: string) =>
  /^p:.+:c:/i.test(externalId);

export const samePersonNameCount = (names: string[], name: string) => {
  const key = normalizeRosterMatchText(name);
  return names.filter((item) => normalizeRosterMatchText(item) === key).length;
};

export const questionnaireFileRejection = (
  file: File,
  externalId: string,
) => {
  if (!externalId) return "Не вдалося додати анкету: у вибраної особи немає ID.";
  if (file.type && file.type !== "application/pdf") {
    return "Анкета має бути у форматі PDF.";
  }
  if (file.size > MAX_QUESTIONNAIRE_FILE_BYTES) {
    return questionnaireTooLargeMessage(file.size);
  }
  return "";
};

export const questionnairePreviewFileFromBlob = (blob: Blob, fileName: string) =>
  new File([blob], fileName, { type: blob.type || "application/pdf" });

export const questionnaireOpenError = (error: unknown) =>
  error instanceof Error
    ? `Не вдалося відкрити анкету: ${error.message}`
    : "Не вдалося відкрити анкету.";

export const questionnairePreviewPersonId = (
  storedId: string | null | undefined,
  listId: string,
) => String(storedId ?? "").trim() || listId;

export const previewFromQuestionnaireDataUrl = (
  fileData: string,
  fileName: string,
) => ({
  url: dataUrlToObjectUrl(fileData),
  file: dataUrlToFile(fileData, fileName),
});

export const loadQuestionnairePreviewObject = async (
  externalId: string,
  fileName: string,
) => {
  const blob = await api.fetchPersonQuestionnaireFile(externalId, fileName);
  return {
    file: questionnairePreviewFileFromBlob(blob, fileName),
    url: URL.createObjectURL(blob),
  };
};

export const personnelListStatusMessage = ({
  focusedName,
  fromCache,
  rowCount,
  sheetName,
}: {
  focusedName: string;
  fromCache?: boolean;
  rowCount: number;
  sheetName: string;
}) => {
  if (focusedName) return `Відкрито картку: ${focusedName}.`;
  if (fromCache) return `Кеш: ${rowCount} записів · ${sheetName}. Оновлюю з БД…`;
  return `Завантажено особовий склад з БД: ${rowCount} записів · ${sheetName}.`;
};
