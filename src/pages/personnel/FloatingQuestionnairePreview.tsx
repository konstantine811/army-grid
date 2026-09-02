import { useEffect, useRef, useState, type ReactNode } from "react";
import { Button, Typography } from "@/components/sci/SciPrimitives";
import { AddPhotoAlternateOutlinedIcon } from "@/components/sci/icons";
import { FloatingWindow, type FloatingPlacement } from "./FloatingWindow";
import {
  QuestionnaireShareButton,
} from "./QuestionnaireShareButton";
import type { QuestionnairePdfSource } from "./questionnaireShare";
import {
  PhotoCropStage,
  type CropRect,
  type PhotoCropStageHandle,
} from "./PhotoCropDialog";

type FloatingQuestionnairePreviewProps = {
  open: boolean;
  title: string;
  previewUrl: string;
  pendingFile: boolean;
  isUploading: boolean;
  placement?: FloatingPlacement;
  defaultWidth?: number;
  defaultHeight?: number;
  minWidth?: number;
  minHeight?: number;
  className?: string;
  shareFileName?: string;
  sharePersonName?: string;
  shareSource?: QuestionnairePdfSource | null;
  cropFile?: File | null;
  onShareNotify?: (message: string) => void;
  onClose: () => void;
  onSaveCrop?: (dataUrl: string, crop: CropRect) => void;
  onCropMessage?: (message: string) => void;
  onOpenTab: () => void;
  onDownload?: () => void;
  onSave?: () => void;
  childrenHint?: ReactNode;
};

export function FloatingQuestionnairePreview({
  open,
  title,
  previewUrl,
  pendingFile,
  isUploading,
  placement = "left",
  defaultWidth = 560,
  defaultHeight = 720,
  minWidth = 360,
  minHeight = 420,
  className = "",
  shareFileName = "",
  sharePersonName = "",
  shareSource = null,
  cropFile = null,
  onShareNotify,
  onClose,
  onSaveCrop,
  onCropMessage,
  onOpenTab,
  onDownload,
  onSave,
  childrenHint,
}: FloatingQuestionnairePreviewProps) {
  const cropRef = useRef<PhotoCropStageHandle>(null);
  const [cropping, setCropping] = useState(false);
  const [cropReady, setCropReady] = useState(false);

  useEffect(() => {
    if (!open) setCropping(false);
  }, [open]);

  const startCrop = () => {
    if (!cropFile || !onSaveCrop) return;
    setCropReady(false);
    setCropping(true);
  };

  const finishCrop = (dataUrl: string, crop: CropRect) => {
    onSaveCrop?.(dataUrl, crop);
    setCropping(false);
  };

  return (
    <FloatingWindow
      open={open}
      title={cropping ? `${title} · фото` : title}
      onClose={onClose}
      placement={placement}
      defaultWidth={defaultWidth}
      defaultHeight={defaultHeight}
      minWidth={minWidth}
      minHeight={minHeight}
      className={["floating-questionnaire-preview", cropping ? "is-cropping" : "", className]
        .filter(Boolean)
        .join(" ")}
      bodyClassName={cropping ? "is-crop-body" : ""}
      footer={
        cropping ? (
          <>
            <Button size="small" variant="outlined" onClick={() => setCropping(false)}>
              Назад до анкети
            </Button>
            <Button
              size="small"
              variant="contained"
              disabled={!cropReady}
              onClick={() => cropRef.current?.save()}
              sx={{ color: "#1a1a14" }}
            >
              Зберегти фото
            </Button>
          </>
        ) : (
          <>
            {onSaveCrop ? (
              <Button
                size="small"
                variant="outlined"
                disabled={!previewUrl || !cropFile}
                startIcon={<AddPhotoAlternateOutlinedIcon />}
                onClick={startCrop}
              >
                Вирізати фото
              </Button>
            ) : null}
            <Button
              size="small"
              variant="outlined"
              disabled={!previewUrl}
              onClick={onOpenTab}
            >
              Нова вкладка
            </Button>
            {onDownload ? (
              <Button
                size="small"
                variant="outlined"
                disabled={!previewUrl}
                onClick={onDownload}
              >
                Експорт PDF
              </Button>
            ) : null}
            {shareSource ? (
              <QuestionnaireShareButton
                disabled={!previewUrl}
                fileName={shareFileName}
                personName={sharePersonName}
                source={shareSource}
                onNotify={onShareNotify}
              />
            ) : null}
            {pendingFile && onSave ? (
              <Button
                size="small"
                variant="contained"
                disabled={isUploading}
                onClick={onSave}
                sx={{ color: "#1a1a14" }}
              >
                {isUploading ? "Збереження…" : "Зберегти анкету"}
              </Button>
            ) : null}
          </>
        )
      }
    >
      {cropping && cropFile ? (
        <PhotoCropStage
          ref={cropRef}
          file={cropFile}
          active={open && cropping}
          compact
          onSave={finishCrop}
          onMessage={onCropMessage ?? onShareNotify ?? (() => undefined)}
          onReadyChange={setCropReady}
        />
      ) : (
        <>
          {pendingFile || childrenHint ? (
            <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
              {childrenHint ??
                "Перевірте анкету. Можна вирізати фото прямо тут, потім зберегти."}
            </Typography>
          ) : null}
          {previewUrl ? (
            <iframe
              className="questionnaire-preview-frame"
              src={previewUrl}
              title="Перегляд анкети PDF"
            />
          ) : (
            <Typography variant="body2" color="text.secondary">
              Немає PDF для перегляду.
            </Typography>
          )}
        </>
      )}
    </FloatingWindow>
  );
}
