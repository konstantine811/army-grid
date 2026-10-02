import type { Dispatch, SetStateAction } from "react";
import {
  api,
  type BackendPersonDocument,
  type BackendPersonQuestionnaire,
} from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { notifyPersonnelAttachmentChanged } from "../../shared/personnelAttachmentSync";
import type { CropRect } from "./PhotoCropDialog";
import { loadAvailablePersonPhotoIds } from "./personAttachments";
import {
  personCardValuesForSave,
  replaceSavedPersonnelRows,
  cardSaveError,
  cardSavedMessage,
} from "./personnelCardSave";
import type { EditablePersonField } from "./card/personCardTypes";
import { buildPersonnelListIndex, type PersonnelListIndex } from "./personnelListIndex";
import {
  phoneClipboardText,
  phoneCopiedMessage,
  phoneCopyFailedMessage,
  phoneDraftProblem,
  phoneDuplicateMessage,
  phoneRemovedMessage,
  phoneSavedMessage,
  savePersonPhonesDocument,
  withSavedPersonPhones,
} from "./personnelPhoneActions";
import {
  assignPhotoValues,
  deleteStoredPersonPhotos,
  photoDeleteError,
  photoDeletedMessage,
  photoMissingIdMessage,
  removePhotoKeys,
  storeCroppedPersonPhoto,
} from "./personnelPhotoSave";
import {
  deleteStoredQuestionnaires,
  questionnaireDeleteError,
  questionnaireDeletedMessage,
  questionnaireDeleteIds,
} from "./personnelQuestionnaireSave";
import { normalizeUaPhone } from "./personnelUtils";

type PersonRow = {
  summary: { externalId: string; name: string };
  row: EjournalPreviewRow;
};

export const bindPersonnelCardEdits = ({
  selectedRow,
  selectedSummary,
  selectedPhoto,
  editableFields,
  editValues,
  savedPhones,
  phoneDraft,
  phoneDocByExternalId,
  photoCropFile,
  questionnaire,
  personnelRows,
  sourceRowsRef,
  allRowsRef,
  listIndexRef,
  requestedPhotoIdsRef,
  setIsLoading,
  setMessage,
  setPersonnelDataEpoch,
  setPhotoByExternalId,
  setPhotoIndexReady,
  setPhotoCropFile,
  setIsPhotoCropOpen,
  setIsDiskFloatingCrop,
  setPhonesByExternalId,
  setPhoneDocByExternalId,
  setPhoneDraft,
  setIsSavingPhone,
  setQuestionnaire,
  setSelectedRowId,
  setMobilePane,
  clearQuestionnairePresence,
  releaseQuestionnairePreviewCache,
  closeQuestionnairePreview,
}: {
  selectedRow: EjournalPreviewRow | null;
  selectedSummary: { externalId: string; name: string };
  selectedPhoto: string;
  editableFields: EditablePersonField[];
  editValues: Record<string, string>;
  savedPhones: string[];
  phoneDraft: string;
  phoneDocByExternalId: Record<string, BackendPersonDocument>;
  photoCropFile: File | null;
  questionnaire: BackendPersonQuestionnaire | null;
  personnelRows: PersonRow[];
  sourceRowsRef: { current: EjournalPreviewRow[] };
  allRowsRef: { current: EjournalPreviewRow[] };
  listIndexRef: { current: PersonnelListIndex };
  requestedPhotoIdsRef: { current: Set<string> };
  setIsLoading: (value: boolean) => void;
  setMessage: (value: string) => void;
  setPersonnelDataEpoch: Dispatch<SetStateAction<number>>;
  setPhotoByExternalId: Dispatch<SetStateAction<Record<string, string>>>;
  setPhotoIndexReady: Dispatch<SetStateAction<number>>;
  setPhotoCropFile: (file: File | null) => void;
  setIsPhotoCropOpen: (value: boolean) => void;
  setIsDiskFloatingCrop: (value: boolean) => void;
  setPhonesByExternalId: Dispatch<SetStateAction<Record<string, string[]>>>;
  setPhoneDocByExternalId: Dispatch<
    SetStateAction<Record<string, BackendPersonDocument>>
  >;
  setPhoneDraft: (value: string) => void;
  setIsSavingPhone: (value: boolean) => void;
  setQuestionnaire: (value: BackendPersonQuestionnaire | null) => void;
  setSelectedRowId: (value: string) => void;
  setMobilePane: (value: "list" | "card" | "side") => void;
  clearQuestionnairePresence: (externalIds: string[]) => void;
  releaseQuestionnairePreviewCache: () => void;
  closeQuestionnairePreview: () => void;
}) => {
  const saveSelectedPerson = async () => {
    const rowId = selectedRow?.__dbRowId;
    if (!rowId) return;
    setIsLoading(true);
    try {
      const values = personCardValuesForSave(editableFields, editValues);
      const updatedRow = await api.updateEjournalRowValues(rowId, values);
      sourceRowsRef.current = replaceSavedPersonnelRows(
        sourceRowsRef.current,
        rowId,
        updatedRow.values,
      );
      allRowsRef.current = replaceSavedPersonnelRows(
        allRowsRef.current,
        rowId,
        updatedRow.values,
      );
      listIndexRef.current = buildPersonnelListIndex(allRowsRef.current);
      setPersonnelDataEpoch((value) => value + 1);
      setMessage(cardSavedMessage(selectedSummary.name));
    } catch (error) {
      setMessage(cardSaveError(error));
    } finally {
      setIsLoading(false);
    }
  };

  const openPhotoCrop = (
    file: File | undefined,
    options?: { floating?: boolean },
  ) => {
    if (!file) return;
    setIsDiskFloatingCrop(Boolean(options?.floating));
    setPhotoCropFile(file);
    setIsPhotoCropOpen(true);
  };

  const savePersonPhoto = async (dataUrl: string, crop: CropRect) => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !selectedRow) {
      setMessage(photoMissingIdMessage);
      return;
    }
    const result = await storeCroppedPersonPhoto({
      row: selectedRow,
      externalId,
      personName: selectedSummary.name,
      dataUrl,
      crop,
      fileName: photoCropFile?.name,
      onLocalPreview: (keys, preview) => {
        setPhotoByExternalId((photos) => assignPhotoValues(photos, keys, preview));
      },
    });
    if (!result.ok) {
      setMessage(result.message);
      return;
    }
    setPhotoByExternalId((photos) =>
      assignPhotoValues(photos, result.resolvedKeys, result.displayValue),
    );
    for (const key of result.resolvedKeys) requestedPhotoIdsRef.current.add(key);
    void loadAvailablePersonPhotoIds({ force: true }).then(() => {
      setPhotoIndexReady((value) => value + 1);
    });
    notifyPersonnelAttachmentChanged(result.savedStorageId, "photo");
    setMessage(result.message);
  };

  const deleteSelectedPhoto = async () => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !selectedPhoto || !selectedRow) return;
    if (!window.confirm(`Видалити фото для ${selectedSummary.name || "особи"}?`)) {
      return;
    }
    try {
      const { storageId, deleteIds } = await deleteStoredPersonPhotos(
        selectedRow,
        externalId,
      );
      setPhotoByExternalId((photos) => removePhotoKeys(photos, deleteIds));
      notifyPersonnelAttachmentChanged(storageId, "photo");
      setMessage(photoDeletedMessage(selectedSummary.name));
    } catch (error) {
      setMessage(photoDeleteError(error));
    }
  };

  const persistPersonPhones = async (externalId: string, phones: string[]) => {
    setPhonesByExternalId((current) =>
      withSavedPersonPhones(current, externalId, phones),
    );
    const saved = await savePersonPhonesDocument(
      externalId,
      phones,
      phoneDocByExternalId[externalId] ?? null,
    );
    if (!saved.ok) return false;
    setPhoneDocByExternalId((current) => {
      if (saved.document) return { ...current, [externalId]: saved.document };
      const updated = { ...current };
      delete updated[externalId];
      return updated;
    });
    return true;
  };

  const addSelectedPersonPhone = async () => {
    const externalId = selectedSummary.externalId;
    const normalized = normalizeUaPhone(phoneDraft);
    if (!selectedRow || !externalId) {
      setMessage(phoneDraftProblem(false, normalized));
      return;
    }
    if (!normalized) {
      setMessage(phoneDraftProblem(true, normalized));
      return;
    }
    if (savedPhones.includes(normalized)) {
      setPhoneDraft("");
      setMessage(phoneDuplicateMessage(normalized));
      return;
    }
    setIsSavingPhone(true);
    const savedToDb = await persistPersonPhones(externalId, [
      ...savedPhones,
      normalized,
    ]);
    setIsSavingPhone(false);
    setPhoneDraft("");
    setMessage(phoneSavedMessage(normalized, savedToDb));
  };

  const removeSelectedPersonPhone = async (phone: string) => {
    const externalId = selectedSummary.externalId;
    if (!externalId) return;
    setIsSavingPhone(true);
    const savedToDb = await persistPersonPhones(
      externalId,
      savedPhones.filter((item) => item !== phone),
    );
    setIsSavingPhone(false);
    setMessage(phoneRemovedMessage(phone, savedToDb));
  };

  const copySelectedPersonPhone = async (phone: string) => {
    try {
      await navigator.clipboard.writeText(phoneClipboardText(phone));
      setMessage(phoneCopiedMessage(phone));
    } catch {
      setMessage(phoneCopyFailedMessage());
    }
  };

  const deleteSelectedQuestionnaire = async () => {
    const externalId = selectedSummary.externalId;
    if (!externalId || !questionnaire) return;
    if (!window.confirm(`Видалити анкету для ${selectedSummary.name || "особи"}?`)) {
      return;
    }
    try {
      const deleteIds = questionnaireDeleteIds(externalId, selectedRow);
      await deleteStoredQuestionnaires(deleteIds);
      setQuestionnaire(null);
      clearQuestionnairePresence(deleteIds);
      notifyPersonnelAttachmentChanged(externalId, "questionnaire");
      releaseQuestionnairePreviewCache();
      closeQuestionnairePreview();
      setMessage(questionnaireDeletedMessage(selectedSummary.name));
    } catch (error) {
      setMessage(questionnaireDeleteError(error));
    }
  };

  const focusPersonByExternalId = (externalId: string) => {
    const record = personnelRows.find(
      (item) => item.summary.externalId === externalId,
    );
    if (record?.row.__dbRowId) {
      setSelectedRowId(record.row.__dbRowId);
      setMobilePane("card");
    }
  };

  const openDiskPhotoCrop = (file: File, externalId: string) => {
    focusPersonByExternalId(externalId);
    openPhotoCrop(file, { floating: true });
  };

  return {
    saveSelectedPerson,
    openPhotoCrop,
    savePersonPhoto,
    deleteSelectedPhoto,
    addSelectedPersonPhone,
    removeSelectedPersonPhone,
    copySelectedPersonPhone,
    deleteSelectedQuestionnaire,
    focusPersonByExternalId,
    openDiskPhotoCrop,
  };
};
