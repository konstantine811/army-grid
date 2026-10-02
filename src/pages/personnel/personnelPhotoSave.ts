import { api } from "../../api";
import type { CropRect } from "./PhotoCropDialog";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  clearAvailablePersonPhotoIdsCache,
  collectPersonAttachmentLookupIds,
  pruneStalePersonPhotos,
  resolvePersonPhotoStorageIdForSave,
} from "./personAttachments";
import {
  compressPhotoDataUrl,
  createPhotoThumbnailDataUrl,
} from "./photoCompression";

export const photoMissingIdMessage =
  "Не вдалося зберегти фото: у вибраної особи немає ID.";

export const photoDeletedMessage = (name: string) => `Фото видалено: ${name}.`;

export const photoDeleteError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося видалити фото з БД.";

export const assignPhotoValues = (
  photos: Record<string, string>,
  keys: string[],
  value: string,
) => {
  const next = { ...photos };
  for (const key of keys) next[key] = value;
  return next;
};

export const removePhotoKeys = (
  photos: Record<string, string>,
  keys: string[],
) => {
  const next = { ...photos };
  for (const key of keys) delete next[key];
  return next;
};

export const personPhotoLookupKeys = (
  externalId: string,
  storageId: string,
  row: EjournalPreviewRow,
) =>
  [
    ...new Set([
      externalId,
      storageId,
      ...collectPersonAttachmentLookupIds(row),
    ]),
  ].filter(Boolean);

export const imageUrlLoads = (url: string) =>
  new Promise<boolean>((resolve) => {
    const image = new Image();
    image.onload = () => resolve(true);
    image.onerror = () => resolve(false);
    image.src = url;
  });

const photoCacheBust = (updatedAt?: string | null) =>
  updatedAt ? Date.parse(updatedAt) || Date.now() : Date.now();

export const storeCroppedPersonPhoto = async ({
  row,
  externalId,
  personName,
  dataUrl,
  crop,
  fileName,
  onLocalPreview,
}: {
  row: EjournalPreviewRow;
  externalId: string;
  personName: string;
  dataUrl: string;
  crop: CropRect;
  fileName?: string;
  onLocalPreview: (keys: string[], dataUrl: string) => void;
}) => {
  const storageId =
    resolvePersonPhotoStorageIdForSave(row, externalId) || externalId;
  const compressedDataUrl = await compressPhotoDataUrl(dataUrl).catch(
    () => dataUrl,
  );
  onLocalPreview(
    personPhotoLookupKeys(externalId, storageId, row),
    compressedDataUrl,
  );

  try {
    const thumbnailData = await createPhotoThumbnailDataUrl(compressedDataUrl);
    const saved = await api.upsertPersonPhoto(storageId, {
      photoData: compressedDataUrl,
      thumbnailData,
      fileName,
      mimeType: "image/jpeg",
      crop,
    });
    await pruneStalePersonPhotos(row, storageId);
    clearAvailablePersonPhotoIdsCache();
    const savedStorageId = saved.personExternalId.trim() || storageId;
    const photoUrl = api.personPhotoFileUrl(savedStorageId, {
      cacheBust: photoCacheBust(saved.updatedAt),
    });
    const resolvedKeys = personPhotoLookupKeys(externalId, savedStorageId, row);
    const displayValue = (await imageUrlLoads(photoUrl))
      ? photoUrl
      : compressedDataUrl;
    return {
      ok: true as const,
      savedStorageId,
      resolvedKeys,
      displayValue,
      message: `Фото збережено в БД: ${personName}.`,
    };
  } catch (error) {
    return {
      ok: false as const,
      message:
        error instanceof Error
          ? `${error.message} (фото показано локально, але не збережено в БД)`
          : "Не вдалося зберегти фото в БД.",
    };
  }
};

export const deleteStoredPersonPhotos = async (
  row: EjournalPreviewRow,
  externalId: string,
) => {
  const storageId =
    resolvePersonPhotoStorageIdForSave(row, externalId) || externalId;
  const deleteIds = personPhotoLookupKeys(externalId, storageId, row);
  await Promise.all(
    deleteIds.map((id) => api.deletePersonPhoto(id).catch(() => undefined)),
  );
  clearAvailablePersonPhotoIdsCache();
  return { storageId, deleteIds };
};
