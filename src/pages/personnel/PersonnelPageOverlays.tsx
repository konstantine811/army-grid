import type { CropRect } from "./PhotoCropDialog";
import { PhotoCropDialog } from "./PhotoCropDialog";
import { FloatingQuestionnairePreview } from "./FloatingQuestionnairePreview";
import { PersonPhotoLightbox } from "./card/PersonPhotoLightbox";
import { QuestionnaireDiskSearchDialog } from "./QuestionnaireDiskSearchDialog";
import { QuestionnairePhotoExtractDialog } from "./QuestionnairePhotoExtractDialog";
import type { QuestionnairePdfSource } from "./questionnaireShare";

export function PersonnelPageOverlays({
  personName,
  selectedPhoto,
  isPhotoLightboxOpen,
  onClosePhoto,
  photoCropFile,
  isPhotoCropOpen,
  isDiskFloatingCrop,
  onClosePhotoCrop,
  onSavePhoto,
  onMessage,
  isPhotoExtractOpen,
  photoExtractPeople,
  onClosePhotoExtract,
  onPhotoExtractSaved,
  isDiskSearchOpen,
  diskSearchPeople,
  onCloseDiskSearch,
  onDiskSearchConfirmed,
  onDiskSearchPhotoSaved,
  onDiskSearchPreview,
  onDiskSearchCrop,
  isQuestionnairePreviewOpen,
  questionnairePreviewTitle,
  questionnairePreviewUrl,
  pendingQuestionnaireFile,
  isUploadingQuestionnaire,
  isLoadingQuestionnairePreview,
  isDiskFloatingPreview,
  questionnaireCropFile,
  onCloseQuestionnairePreview,
  onOpenQuestionnaireTab,
  onDownloadQuestionnaire,
  onSaveQuestionnaire,
  questionnaireExportFileName,
  questionnaireShareSource,
}: {
  personName: string;
  selectedPhoto: string;
  isPhotoLightboxOpen: boolean;
  onClosePhoto: () => void;
  photoCropFile: File | null;
  isPhotoCropOpen: boolean;
  isDiskFloatingCrop: boolean;
  onClosePhotoCrop: () => void;
  onSavePhoto: (dataUrl: string, crop: CropRect) => void | Promise<void>;
  onMessage: (message: string) => void;
  isPhotoExtractOpen: boolean;
  photoExtractPeople: Array<{ externalId: string; fullName: string }>;
  onClosePhotoExtract: () => void;
  onPhotoExtractSaved: (externalId: string, photoData: string) => void;
  isDiskSearchOpen: boolean;
  diskSearchPeople: Array<{
    rowId: string;
    externalId: string;
    fullName: string;
    callSign: string;
    missingQuestionnaire: boolean;
    missingPhoto: boolean;
  }>;
  onCloseDiskSearch: () => void;
  onDiskSearchConfirmed: (externalId: string) => void;
  onDiskSearchPhotoSaved: (externalId: string, photoData: string) => void;
  onDiskSearchPreview: (file: File, title: string, externalId: string) => void;
  onDiskSearchCrop: (file: File, externalId: string) => void;
  isQuestionnairePreviewOpen: boolean;
  questionnairePreviewTitle: string;
  questionnairePreviewUrl: string;
  pendingQuestionnaireFile: boolean;
  isUploadingQuestionnaire: boolean;
  isLoadingQuestionnairePreview: boolean;
  isDiskFloatingPreview: boolean;
  questionnaireCropFile: File | null;
  onCloseQuestionnairePreview: () => void;
  onOpenQuestionnaireTab: () => void;
  onDownloadQuestionnaire: () => void;
  onSaveQuestionnaire: () => void;
  questionnaireExportFileName: string;
  questionnaireShareSource: QuestionnairePdfSource | null;
}) {
  return (
    <>
      {isPhotoLightboxOpen && selectedPhoto ? (
        <PersonPhotoLightbox
          name={personName}
          photo={selectedPhoto}
          onClose={onClosePhoto}
        />
      ) : null}
      <PhotoCropDialog
        file={photoCropFile}
        open={isPhotoCropOpen}
        floating={isDiskFloatingCrop}
        onClose={onClosePhotoCrop}
        onMessage={onMessage}
        onSave={onSavePhoto}
      />
      <QuestionnairePhotoExtractDialog
        open={isPhotoExtractOpen}
        people={photoExtractPeople}
        onClose={onClosePhotoExtract}
        onPhotoSaved={onPhotoExtractSaved}
      />
      <QuestionnaireDiskSearchDialog
        open={isDiskSearchOpen}
        people={diskSearchPeople}
        onClose={onCloseDiskSearch}
        onConfirmed={onDiskSearchConfirmed}
        onAutoPhotoSaved={onDiskSearchPhotoSaved}
        onPreviewQuestionnaire={onDiskSearchPreview}
        onCropPhoto={onDiskSearchCrop}
      />
      <FloatingQuestionnairePreview
        open={isQuestionnairePreviewOpen}
        title={questionnairePreviewTitle || `Анкета · ${personName}`}
        previewUrl={questionnairePreviewUrl}
        pendingFile={pendingQuestionnaireFile}
        isUploading={isUploadingQuestionnaire || isLoadingQuestionnairePreview}
        placement={isDiskFloatingPreview ? "left" : "center"}
        childrenHint={
          isLoadingQuestionnairePreview ? "Завантажую PDF анкети…" : undefined
        }
        defaultWidth={isDiskFloatingPreview ? 560 : 760}
        defaultHeight={isDiskFloatingPreview ? 720 : 820}
        cropFile={questionnaireCropFile}
        onClose={onCloseQuestionnairePreview}
        onSaveCrop={onSavePhoto}
        onCropMessage={onMessage}
        onOpenTab={onOpenQuestionnaireTab}
        onDownload={onDownloadQuestionnaire}
        onSave={onSaveQuestionnaire}
        shareFileName={questionnaireExportFileName}
        sharePersonName={personName}
        shareSource={questionnaireShareSource}
        onShareNotify={onMessage}
      />
    </>
  );
}
