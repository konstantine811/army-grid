import {
  Button,
  Checkbox,
  LinearProgress,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import { PictureAsPdfOutlinedIcon } from "@/components/sci/icons";
import { FloatingWindow } from "../../personnel/FloatingWindow";
import type { AnketaColumnKey } from "../anketaSheet";
import type { AnketaOcrProposal } from "../anketaQuestionnaireOcr";

const OCR_TEXTAREA_ROWS: Partial<Record<AnketaColumnKey, number>> = {
  relatives: 8,
  additionalInfo: 4,
};

const isOcrTextareaField = (columnId: AnketaColumnKey, value: string) =>
  columnId in OCR_TEXTAREA_ROWS || value.includes("\n");

const ocrTextareaRows = (columnId: AnketaColumnKey, value: string) =>
  OCR_TEXTAREA_ROWS[columnId] ?? (value.includes("\n") ? 3 : 2);

type AnketaOcrReviewDialogProps = {
  open: boolean;
  proposals: AnketaOcrProposal[];
  focusedColumnId?: AnketaOcrProposal["columnId"] | null;
  isApplying?: boolean;
  onClose: () => void;
  onToggle: (columnId: AnketaOcrProposal["columnId"]) => void;
  onValueChange: (
    columnId: AnketaOcrProposal["columnId"],
    value: string,
  ) => void;
  onSelectAll: (selected: boolean) => void;
  onConfirm: () => void;
  canOpenQuestionnaire?: boolean;
  questionnaireOpen?: boolean;
  onOpenQuestionnaire?: () => void;
};

const confidenceLabel = (value: string) => {
  const text = String(value ?? "").trim().toLowerCase();
  if (text === "high") return "висока";
  if (text === "low") return "низька";
  return "середня";
};

const confidenceClass = (value: string) => {
  const text = String(value ?? "").trim().toLowerCase();
  if (text === "high") return "anketa-ocr-confidence-high";
  if (text === "low") return "anketa-ocr-confidence-low";
  return "anketa-ocr-confidence-medium";
};

export function AnketaOcrReviewDialog({
  open,
  proposals,
  focusedColumnId = null,
  isApplying = false,
  onClose,
  onToggle,
  onValueChange,
  onSelectAll,
  onConfirm,
  canOpenQuestionnaire = false,
  questionnaireOpen = false,
  onOpenQuestionnaire,
}: AnketaOcrReviewDialogProps) {
  const selectedCount = proposals.filter((item) => item.selected).length;
  const focusedProposal = focusedColumnId
    ? proposals.find((item) => item.columnId === focusedColumnId)
    : null;

  const title = focusedProposal
    ? `Підтвердити · ${focusedProposal.label}`
    : "Розпізнавання анкети · перевірка";

  const subtitle = questionnaireOpen
    ? "звірте з PDF поруч · перетягніть вікно"
    : "перетягніть за шапку · розмір за кут";

  return (
    <FloatingWindow
      open={open}
      title={title}
      subtitle={subtitle}
      onClose={() => {
        if (!isApplying) onClose();
      }}
      placement="right"
      defaultWidth={520}
      defaultHeight={560}
      minWidth={400}
      minHeight={360}
      className="anketa-ocr-floating"
      bodyClassName="anketa-ocr-floating-body"
      footer={
        <>
          {onOpenQuestionnaire ? (
            <Button
              size="small"
              variant="outlined"
              disabled={isApplying || !canOpenQuestionnaire}
              startIcon={<PictureAsPdfOutlinedIcon />}
              onClick={onOpenQuestionnaire}
              className="anketa-ocr-open-questionnaire"
            >
              {questionnaireOpen ? "Анкета відкрита" : "Відкрити анкету"}
            </Button>
          ) : null}
          <Button
            size="small"
            variant="outlined"
            disabled={isApplying}
            onClick={onClose}
          >
            Скасувати
          </Button>
          <Button
            size="small"
            variant="contained"
            disabled={isApplying || selectedCount === 0}
            onClick={onConfirm}
            sx={{ color: "#1a1a14" }}
          >
            {isApplying ? "Записую…" : `Підтвердити · ${selectedCount}`}
          </Button>
        </>
      }
    >
      <Typography variant="body2" color="text.secondary" sx={{ mb: 1.5, flex: "0 0 auto" }}>
        {focusedProposal
          ? `Gemini запропонував значення для «${focusedProposal.label}». Звірте з PDF-анкетою поруч і підтвердьте — можна також записати інші порожні обрані поля.`
          : "Звірте запропоновані значення з PDF-анкетою поруч (бланк, паспорт, військовий квиток, РНОКПП) і підтвердьте."}
      </Typography>
      <div className="anketa-ocr-review-actions">
        <Button size="small" variant="text" onClick={() => onSelectAll(true)}>
          Обрати всі
        </Button>
        <Button size="small" variant="text" onClick={() => onSelectAll(false)}>
          Зняти всі
        </Button>
        <span className="anketa-ocr-review-count">
          Обрано: {selectedCount} / {proposals.length}
        </span>
      </div>
      <div className="anketa-ocr-review-list">
        {proposals.map((proposal) => {
          const isTextarea = isOcrTextareaField(
            proposal.columnId,
            proposal.value,
          );
          return (
          <div
            key={proposal.columnId}
            className={[
              "anketa-ocr-review-row",
              isTextarea ? "anketa-ocr-review-row-multiline" : "",
              proposal.columnId === "relatives"
                ? "anketa-ocr-review-row-relatives"
                : "",
              proposal.columnId === focusedColumnId
                ? "anketa-ocr-review-row-focused"
                : "",
            ]
              .filter(Boolean)
              .join(" ")}
          >
            <Checkbox
              checked={proposal.selected}
              onCheckedChange={() => onToggle(proposal.columnId)}
              label={proposal.label}
            />
            <TextField
              size="small"
              fullWidth
              multiline={isTextarea}
              rows={isTextarea ? ocrTextareaRows(proposal.columnId, proposal.value) : undefined}
              className={
                isTextarea ? "anketa-ocr-review-textarea" : undefined
              }
              value={proposal.value}
              onChange={(event) =>
                onValueChange(proposal.columnId, event.target.value)
              }
            />
            <div className="anketa-ocr-review-meta">
              {proposal.source ? (
                <span className="anketa-ocr-source">{proposal.source}</span>
              ) : null}
              <span
                className={`anketa-ocr-confidence ${confidenceClass(proposal.confidence)}`}
              >
                {confidenceLabel(proposal.confidence)}
              </span>
            </div>
          </div>
          );
        })}
      </div>
      {isApplying ? <LinearProgress sx={{ mt: 1.5, flex: "0 0 auto" }} /> : null}
    </FloatingWindow>
  );
}
