import { api } from "../../../../api";
import { extractPassportPhotoFromPdf } from "../../autoPassportPhoto";
import { createPhotoThumbnailDataUrl } from "../../photoCompression";
import { dataUrlToFile } from "../../personnelUtils";
import { isStrictDiskQuestionnaireMatch } from "../../questionnaireDiskMatch";

const pause = (delayMs: number) =>
  new Promise<void>((resolve) => globalThis.setTimeout(resolve, delayMs));

export type BulkExtractPhotosTarget = {
  externalId: string;
  fullName: string;
};

export type BulkExtractPhotosSkipReason =
  | "no_pdf"
  | "name_mismatch"
  | "no_face"
  | "save_failed";

export type BulkExtractPhotosRowResult = {
  externalId: string;
  fullName: string;
  status: "saved" | "skipped" | "failed";
  reason?: BulkExtractPhotosSkipReason;
  message?: string;
  photoData?: string;
};

export type BulkExtractPhotosResult = {
  saved: number;
  skipped: number;
  failed: number;
  rows: BulkExtractPhotosRowResult[];
};

const loadQuestionnairePdfFile = async (externalId: string, fullName: string) => {
  const questionnaire = await api.getPersonQuestionnaire(externalId);
  if (!questionnaire?.fileData?.trim()) {
    return { ok: false as const, reason: "no_pdf" as const };
  }
  if (
    questionnaire.fileName &&
    !isStrictDiskQuestionnaireMatch(fullName, questionnaire.fileName)
  ) {
    return { ok: false as const, reason: "name_mismatch" as const };
  }
  return {
    ok: true as const,
    file: dataUrlToFile(
      questionnaire.fileData,
      questionnaire.fileName || "questionnaire.pdf",
    ),
  };
};

export const bulkExtractPhotosFromQuestionnaires = async (input: {
  targets: BulkExtractPhotosTarget[];
  signal?: AbortSignal;
  onProgress?: (done: number, total: number, saved: number) => void;
}): Promise<BulkExtractPhotosResult> => {
  const rows: BulkExtractPhotosRowResult[] = [];
  let saved = 0;
  let skipped = 0;
  let failed = 0;
  const total = input.targets.length;

  for (let index = 0; index < input.targets.length; index += 1) {
    if (input.signal?.aborted) break;
    const target = input.targets[index]!;
    const done = index + 1;

    try {
      const pdf = await loadQuestionnairePdfFile(target.externalId, target.fullName);
      if (!pdf.ok) {
        skipped += 1;
        rows.push({
          externalId: target.externalId,
          fullName: target.fullName,
          status: "skipped",
          reason: pdf.reason,
        });
        input.onProgress?.(done, total, saved);
        await pause(20);
        continue;
      }

      const photo = await extractPassportPhotoFromPdf(pdf.file);
      if (!photo) {
        skipped += 1;
        rows.push({
          externalId: target.externalId,
          fullName: target.fullName,
          status: "skipped",
          reason: "no_face",
        });
        input.onProgress?.(done, total, saved);
        await pause(20);
        continue;
      }

      const thumbnailData = await createPhotoThumbnailDataUrl(photo.dataUrl);
      const savedPhoto = await api.upsertPersonPhoto(target.externalId, {
        photoData: photo.dataUrl,
        thumbnailData,
        fileName: pdf.file.name,
        mimeType: "image/jpeg",
        crop: photo.crop,
      });

      saved += 1;
      rows.push({
        externalId: target.externalId,
        fullName: target.fullName,
        status: "saved",
        photoData: savedPhoto?.photoData || photo.dataUrl,
      });
    } catch (error) {
      failed += 1;
      rows.push({
        externalId: target.externalId,
        fullName: target.fullName,
        status: "failed",
        reason: "save_failed",
        message: error instanceof Error ? error.message : "Unknown error",
      });
    }

    input.onProgress?.(done, total, saved);
    await pause(50);
  }

  return { saved, skipped, failed, rows };
};
