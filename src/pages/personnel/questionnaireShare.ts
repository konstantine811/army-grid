import {
  dataUrlToUint8Array,
  downloadBlob,
  sanitizeFileName,
} from "../../shared/browserExport";

export type QuestionnairePdfSource = {
  file?: Blob | File;
  fileData?: string;
};

const ensurePdfFileName = (fileName: string) => {
  const safeName = sanitizeFileName(fileName.trim() || "anketa.pdf");
  return safeName.toLowerCase().endsWith(".pdf") ? safeName : `${safeName}.pdf`;
};

export const resolveQuestionnairePdfFile = (
  fileName: string,
  source: QuestionnairePdfSource,
): File | null => {
  const safeName = ensurePdfFileName(fileName);

  if (source.file) {
    if (source.file instanceof File && source.file.name === safeName) {
      return source.file;
    }
    return new File([source.file], safeName, {
      type: source.file.type || "application/pdf",
    });
  }

  if (source.fileData) {
    const bytes = dataUrlToUint8Array(source.fileData);
    return new File([bytes], safeName, { type: "application/pdf" });
  }

  return null;
};

export const buildQuestionnaireShareText = (
  personName: string,
  fileName: string,
) => {
  const label = personName.trim() || "військовослужбовця";
  return `Анкета: ${label}\n${ensurePdfFileName(fileName)}`;
};

export const canNativeShareQuestionnaire = (file: File) =>
  typeof navigator !== "undefined" &&
  typeof navigator.share === "function" &&
  typeof navigator.canShare === "function" &&
  navigator.canShare({ files: [file] });

export const nativeShareQuestionnaire = async (
  file: File,
  personName: string,
) => {
  await navigator.share({
    files: [file],
    title: file.name,
    text: buildQuestionnaireShareText(personName, file.name),
  });
};

export type QuestionnaireMessenger = "whatsapp" | "telegram";

const isMobileDevice = () =>
  typeof navigator !== "undefined" &&
  /iPhone|iPad|iPod|Android/i.test(navigator.userAgent);

const buildWhatsAppShareUrl = (text: string) => {
  const encodedText = encodeURIComponent(text);
  return isMobileDevice()
    ? `https://wa.me/?text=${encodedText}`
    : `https://web.whatsapp.com/send?text=${encodedText}`;
};

export const shareQuestionnaireViaMessenger = (
  messenger: QuestionnaireMessenger,
  fileName: string,
  source: QuestionnairePdfSource,
  personName: string,
) => {
  const file = resolveQuestionnairePdfFile(fileName, source);
  if (!file) {
    throw new Error("Немає PDF для надсилання.");
  }

  downloadBlob(file, file.name);

  const text = `${buildQuestionnaireShareText(personName, file.name)}\n\nPDF завантажено — прикріпіть файл до повідомлення.`;

  const url =
    messenger === "whatsapp"
      ? buildWhatsAppShareUrl(text)
      : `https://t.me/share/url?url=&text=${encodeURIComponent(text)}`;

  window.open(url, "_blank", "noopener,noreferrer");
};
