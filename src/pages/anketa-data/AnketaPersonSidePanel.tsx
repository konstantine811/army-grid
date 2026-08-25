import { useMemo } from "react";
import { buildPersonnelRoute } from "../../app/navigation";
import {
  Button,
  Chip,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  PersonSearchOutlinedIcon,
  PictureAsPdfOutlinedIcon,
} from "@/components/sci/icons";
import { FloatingQuestionnairePreview } from "../personnel/FloatingQuestionnairePreview";
import { matchLabel } from "./anketaPersonMatch";
import type { AnketaColumnKey, AnketaRow } from "./anketaSheet";
import type { AnketaEmptyCell } from "./anketaGaps";
import {
  ANKETA_MISSING_VALUE_PRESETS,
  listAnketaEmptyCells,
} from "./anketaGaps";
import { useAnketaPersonPanel } from "./hooks/useAnketaPersonPanel";

type AnketaPersonSidePanelProps = {
  anketaRow: AnketaRow | null;
  focusedEmpty: AnketaEmptyCell | null;
  gapColumnKeys: AnketaColumnKey[];
  onClose: () => void;
  onFillMissing?: (value: string) => void;
  onMessage?: (message: string) => void;
};

const FIELD_ROWS: Array<{ key: AnketaColumnKey; label: string }> = [
  { key: "rank", label: "Звання" },
  { key: "externalId", label: "ID" },
  { key: "positionIndex", label: "Індекс посади" },
  { key: "serviceType", label: "Вид служби" },
  { key: "birthDate", label: "Дата народження" },
  { key: "birthPlace", label: "Місце народження" },
  { key: "sex", label: "Стать" },
  { key: "rnokpp", label: "РНОКПП" },
  { key: "idDocumentNumber", label: "Документ" },
  { key: "location", label: "Дислокація" },
  { key: "arrivedFrom", label: "Звідки прибув" },
  { key: "contractFrom", label: "Контракт з" },
  { key: "contractTo", label: "Контракт до" },
  { key: "militaryId", label: "Військовий квиток" },
  { key: "education", label: "Освіта" },
  { key: "relatives", label: "Родичі" },
  { key: "additionalInfo", label: "Додаткова інформація" },
];

export function AnketaPersonSidePanel({
  anketaRow,
  focusedEmpty,
  gapColumnKeys,
  onClose,
  onFillMissing,
  onMessage,
}: AnketaPersonSidePanelProps) {
  const panel = useAnketaPersonPanel(anketaRow, onMessage);

  const rowGaps = useMemo(() => {
    if (!anketaRow) return [];
    return listAnketaEmptyCells([anketaRow], gapColumnKeys);
  }, [anketaRow, gapColumnKeys]);

  if (!anketaRow) return null;

  return (
    <>
      <aside className="anketa-person-side-panel person-card-panel">
        <div className="anketa-person-side-header">
          <div className="panel-heading">Картка службовця</div>
          <Button size="small" variant="text" onClick={onClose}>
            Закрити
          </Button>
        </div>

        <div className="person-card-hero anketa-person-hero">
          <div className="person-avatar">
            {panel.photoData ? (
              <img alt={panel.displayName} src={panel.photoData} />
            ) : (
              <PersonSearchOutlinedIcon />
            )}
          </div>
            <div className="anketa-person-hero-meta">
            {panel.match?.summary.callSign ? (
              <div className="person-callsign-row">
                <span className="person-callsign" title="Позивний">
                  <span className="person-callsign-label">позивний</span>
                  <strong>{panel.match.summary.callSign}</strong>
                </span>
              </div>
            ) : null}
            <Typography component="h2" variant="h5">
              {panel.displayName}
            </Typography>
            <div className="person-action-tags">
              {(anketaRow.rank || panel.match?.summary.rank) && (
                <Chip
                  label={anketaRow.rank || panel.match?.summary.rank}
                  size="small"
                  className="anketa-person-tag"
                />
              )}
              {(anketaRow.positionIndex || panel.match?.summary.positionIndex) && (
                <Chip
                  label={`Посада: ${
                    anketaRow.positionIndex || panel.match?.summary.positionIndex
                  }`}
                  size="small"
                  className="anketa-person-tag"
                />
              )}
              {(anketaRow.serviceType || panel.match?.summary.serviceType) && (
                <Chip
                  label={anketaRow.serviceType || panel.match?.summary.serviceType}
                  size="small"
                  className="anketa-person-tag"
                />
              )}
            </div>
            <Typography variant="caption" color="text.secondary">
              {panel.matchStatus === "loading"
                ? "Шукаю в особовому складі…"
                : panel.match
                  ? `Знайдено в ООС · ${matchLabel(panel.match.matchBy)}`
                  : "У кеші особового складу не знайдено"}
            </Typography>
          </div>
        </div>

        <div className="person-card-scroll anketa-person-scroll">
          {focusedEmpty ? (
            <div className="anketa-person-focus-banner">
              <strong>Порожня комірка</strong>
              <span>
                {focusedEmpty.a1} · {focusedEmpty.header}
              </span>
              {onFillMissing ? (
                <div className="anketa-missing-presets anketa-missing-presets-inline">
                  <div className="anketa-missing-presets-title">
                    Швидко заповнити статус
                  </div>
                  <div className="anketa-missing-presets-list">
                    {ANKETA_MISSING_VALUE_PRESETS.map((preset) => (
                      <button
                        key={preset}
                        type="button"
                        onClick={() => onFillMissing(preset)}
                      >
                        {preset}
                      </button>
                    ))}
                  </div>
                </div>
              ) : null}
            </div>
          ) : null}

          <div className="person-action-fields">
            {FIELD_ROWS.map((field) => {
              const value = String(anketaRow[field.key] ?? "").trim();
              const isGap = gapColumnKeys.includes(field.key) && !value;
              return (
                <span
                  key={field.key}
                  className={isGap ? "anketa-person-field-gap" : undefined}
                >
                  <strong>{field.label}</strong>
                  {value || "—"}
                </span>
              );
            })}
          </div>

          {rowGaps.length ? (
            <div className="person-edit-section">
              <div className="panel-heading">
                Пропуски в вибраних колонках · {rowGaps.length}
              </div>
              <ul className="anketa-person-gap-list">
                {rowGaps.map((gap) => (
                  <li key={`${gap.columnId}-${gap.a1}`}>
                    <span>{gap.header}</span>
                    <small>{gap.a1}</small>
                  </li>
                ))}
              </ul>
            </div>
          ) : null}

          <div className="person-edit-section">
            <div className="panel-heading">Анкета (PDF)</div>
            {panel.isLoadingAttachments ? (
              <div className="person-document-empty">Завантаження…</div>
            ) : panel.questionnaire ? (
              <article className="person-document-shell is-ready">
                <button
                  className="person-document-item is-ready"
                  type="button"
                  onClick={() => void panel.openQuestionnairePreview()}
                >
                  <PictureAsPdfOutlinedIcon />
                  <span>
                    <strong>Анкета (PDF)</strong>
                    <small>{panel.exportFileName} · переглянути</small>
                  </span>
                </button>
              </article>
            ) : (
              <div className="person-document-empty">
                <PictureAsPdfOutlinedIcon />
                <span>
                  {panel.match
                    ? "Анкета ще не додана в БД"
                    : "Потрібен зв’язок з особовим складом"}
                </span>
              </div>
            )}
          </div>

          <div className="anketa-person-side-actions">
            <Button
              variant="contained"
              size="small"
              disabled={!panel.match || panel.isMerging || !panel.mergePreview?.labels.length}
              onClick={() => void panel.mergeToPersonnel()}
            >
              {panel.isMerging
                ? "Переношу…"
                : panel.mergePreview?.labels.length
                  ? `Перенести в ООС · ${panel.mergePreview.labels.length}`
                  : "Перенести в ООС"}
            </Button>
            <Button
              variant="outlined"
              size="small"
              disabled={!panel.match}
              onClick={() => {
                if (!panel.match) return;
                const target = {
                  rowId: panel.match.row.__dbRowId,
                  externalId:
                    panel.personnelExternalId || panel.match.summary.externalId,
                };
                try {
                  window.localStorage.setItem(
                    "army-grid:focus-personnel",
                    JSON.stringify(target),
                  );
                } catch {
                  /* ignore */
                }
                window.open(
                  buildPersonnelRoute(target),
                  "_blank",
                  "noopener,noreferrer",
                );
              }}
            >
              Відкрити в особовому складі
            </Button>
          </div>
        </div>
      </aside>

      <FloatingQuestionnairePreview
        open={panel.previewOpen}
        title={panel.previewTitle}
        previewUrl={panel.previewUrl}
        pendingFile={false}
        isUploading={false}
        placement="right"
        defaultWidth={720}
        defaultHeight={920}
        minWidth={480}
        minHeight={420}
        className="floating-questionnaire-preview is-anketa-edge"
        shareFileName={panel.exportFileName}
        sharePersonName={panel.displayName}
        shareSource={panel.shareSource}
        onClose={panel.closePreview}
        onOpenTab={panel.openQuestionnaireTab}
        onDownload={panel.downloadQuestionnaire}
      />
    </>
  );
}
