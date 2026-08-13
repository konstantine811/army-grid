import { type ReactNode } from "react";
import { Button, Typography } from "@/components/sci/SciPrimitives";
import { AddPhotoAlternateOutlinedIcon } from "@/components/sci/icons";
import { FloatingWindow, type FloatingPlacement } from "./FloatingWindow";

type FloatingQuestionnairePreviewProps = {
  open: boolean;
  title: string;
  previewUrl: string;
  pendingFile: boolean;
  isUploading: boolean;
  placement?: FloatingPlacement;
  onClose: () => void;
  onCrop: () => void;
  onOpenTab: () => void;
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
  onClose,
  onCrop,
  onOpenTab,
  onSave,
  childrenHint,
}: FloatingQuestionnairePreviewProps) {
  return (
    <FloatingWindow
      open={open}
      title={title}
      onClose={onClose}
      placement={placement}
      defaultWidth={560}
      defaultHeight={720}
      minWidth={360}
      minHeight={420}
      className="floating-questionnaire-preview"
      footer={
        <>
          <Button
            size="small"
            variant="outlined"
            disabled={!previewUrl}
            startIcon={<AddPhotoAlternateOutlinedIcon />}
            onClick={onCrop}
          >
            Вирізати фото
          </Button>
          <Button
            size="small"
            variant="outlined"
            disabled={!previewUrl}
            onClick={onOpenTab}
          >
            Нова вкладка
          </Button>
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
      }
    >
      {pendingFile || childrenHint ? (
        <Typography variant="body2" color="text.secondary" sx={{ mb: 1 }}>
          {childrenHint ??
            "Перевірте анкету. Можна вирізати фото з PDF, потім зберегти."}
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
    </FloatingWindow>
  );
}
