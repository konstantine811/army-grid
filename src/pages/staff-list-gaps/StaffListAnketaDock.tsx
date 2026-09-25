import { Button, Typography } from "@/components/sci/SciPrimitives";
import { useAnketaPersonPanel } from "../anketa-data/hooks/useAnketaPersonPanel";
import { matchLabel } from "../anketa-data/anketaPersonMatch";
import type { AnketaRow } from "../anketa-data/anketaSheet";

type StaffListAnketaDockProps = {
  anketaRow: AnketaRow | null;
  personLabel?: string;
  onClose: () => void;
  onMessage?: (message: string) => void;
};

/** Права панель з PDF-анкетою для «Списки Excel». */
export function StaffListAnketaDock({
  anketaRow,
  personLabel,
  onClose,
  onMessage,
}: StaffListAnketaDockProps) {
  const panel = useAnketaPersonPanel(anketaRow, onMessage, {
    autoOpenQuestionnairePreview: true,
  });

  if (!anketaRow) return null;

  const title = personLabel?.trim() || panel.displayName;

  return (
    <aside className="staff-list-questionnaire-dock person-card-panel">
      <div className="staff-list-questionnaire-dock-header">
        <div className="panel-heading">Анкета · {title}</div>
        <Button size="small" variant="text" onClick={onClose}>
          Закрити
        </Button>
      </div>

      <p className="staff-list-questionnaire-dock-meta">
        {panel.matchStatus === "loading"
          ? "Шукаю в особовому складі…"
          : panel.match
            ? `Знайдено в ООС · ${matchLabel(panel.match.matchBy)}`
            : panel.isLoadingAttachments
              ? "Завантажую PDF…"
              : panel.questionnaire
                ? panel.exportFileName
                : "PDF-анкету не знайдено"}
      </p>

      <div className="staff-list-questionnaire-dock-body">
        {panel.isLoadingAttachments ? (
          <Typography variant="body2" color="text.secondary">
            Завантаження анкети…
          </Typography>
        ) : panel.previewUrl ? (
          <iframe
            className="questionnaire-preview-frame"
            src={panel.previewUrl}
            title={`Анкета · ${title}`}
          />
        ) : (
          <Typography variant="body2" color="text.secondary">
            {panel.match
              ? panel.isLoadingAttachments
                ? "Завантаження анкети…"
                : "PDF-анкету не знайдено. Спробуйте «Відкрити анкету» або завантажте PDF в особовому складі."
              : "Особу в особовому складі не знайдено — перевірте ПІБ у списку."}
          </Typography>
        )}
      </div>

      <div className="staff-list-questionnaire-dock-actions">
        <Button
          size="small"
          variant="outlined"
          disabled={!panel.questionnaire && !panel.previewUrl}
          onClick={() => void panel.openQuestionnairePreview()}
        >
          Відкрити анкету
        </Button>
        <Button
          size="small"
          variant="outlined"
          disabled={!panel.previewUrl}
          onClick={() => void panel.openQuestionnaireTab()}
        >
          Нова вкладка
        </Button>
        <Button
          size="small"
          variant="outlined"
          disabled={!panel.questionnaire}
          onClick={() => void panel.downloadQuestionnaire()}
        >
          Експорт PDF
        </Button>
      </div>
    </aside>
  );
}
