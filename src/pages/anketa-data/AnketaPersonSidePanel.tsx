import { useMemo } from "react";
import { openPersonnelPerson } from "../../app/navigation";
import {
  buildPersonnelFocusTargetFromRow,
  type PersonnelFocusTarget,
} from "../personnel/personnelFocus";
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
import { matchLabel, anketaPersonnelNamesMatch } from "./anketaPersonMatch";
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
  onOpenPersonnel?: (target: PersonnelFocusTarget) => void;
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
  onOpenPersonnel,
}: AnketaPersonSidePanelProps) {
  const panel = useAnketaPersonPanel(anketaRow, onMessage);
  const nameMismatch = Boolean(
    anketaRow &&
      panel.match &&
      !panel.questionnaire &&
      !anketaPersonnelNamesMatch(anketaRow.fullName, panel.match.summary.name),
  );

  const rowGaps = useMemo(() => {
    if (!anketaRow) return [];
    return listAnketaEmptyCells([anketaRow], gapColumnKeys);
  }, [anketaRow, gapColumnKeys]);

  const openPersonnelTarget = (target: PersonnelFocusTarget) => {
    if (onOpenPersonnel) {
      onOpenPersonnel(target);
      return;
    }
    openPersonnelPerson(target);
  };

  const openPersonnelForMatch = (match: NonNullable<typeof panel.match>) => {
    const target = buildPersonnelFocusTargetFromRow(match.row);
    const attachmentId = panel.personnelExternalId.trim();
    openPersonnelTarget(
      attachmentId ? { ...target, externalId: attachmentId } : target,
    );
  };

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
                : nameMismatch
                  ? "Знайдено запис з іншим ПІБ — PDF і фото не показуємо"
                : panel.match
                  ? `Знайдено в ООС · ${matchLabel(panel.match.matchBy)}`
                  : panel.ambiguousMatches.length > 1
                    ? `Знайдено ${panel.ambiguousMatches.length} осіб з таким ПІБ — уточніть ID і дату народження в анкеті`
                    : panel.ambiguousMatches.length === 1
                      ? "Є кілька збігів за ПІБ — додайте ID і дату народження"
                    : panel.similarMatches.length
                      ? "Точного збігу немає — нижче схожі за прізвищем + імʼям"
                      : "У кеші особового складу не знайдено"}
            </Typography>
            {panel.similarMatches.length && !panel.match ? (
              <div className="anketa-person-similar-list">
                <Typography variant="caption" color="text.secondary" component="div">
                  Можливо це інша особа або інше написання ПІБ:
                </Typography>
                {panel.similarMatches.slice(0, 6).map((item) => (
                  <button
                    key={item.summary.externalId || item.row.__dbRowId}
                    type="button"
                    className="anketa-person-similar-item"
                    onClick={() =>
                      openPersonnelTarget(
                        buildPersonnelFocusTargetFromRow(item.row),
                      )
                    }
                  >
                    {[
                      item.summary.name,
                      item.summary.rank,
                      item.summary.rnokpp,
                      item.summary.birthDate,
                    ]
                      .filter(Boolean)
                      .join(" · ")}
                  </button>
                ))}
              </div>
            ) : null}
            {panel.ambiguousMatches.length > 1 ? (
              <div className="anketa-person-similar-list">
                <Typography variant="caption" color="text.secondary" component="div">
                  Кілька осіб з таким ПІБ — оберіть або уточніть ID / дату народження:
                </Typography>
                {panel.ambiguousMatches.slice(0, 6).map((item) => (
                  <button
                    key={item.summary.externalId || item.row.__dbRowId}
                    type="button"
                    className="anketa-person-similar-item"
                    onClick={() =>
                      openPersonnelTarget(
                        buildPersonnelFocusTargetFromRow(item.row),
                      )
                    }
                  >
                    {[
                      item.summary.name,
                      item.summary.rank,
                      item.summary.birthDate,
                      item.summary.externalId,
                    ]
                      .filter(Boolean)
                      .join(" · ")}
                  </button>
                ))}
              </div>
            ) : null}
          </div>
        </div>

        <div className="person-card-scroll anketa-person-scroll">
          <div className="anketa-person-side-actions anketa-person-side-actions-top">
            <Button
              variant="outlined"
              size="small"
              startIcon={<PictureAsPdfOutlinedIcon />}
              disabled={
                panel.isLoadingAttachments ||
                nameMismatch ||
                !panel.questionnaire
              }
              onClick={() => void panel.openQuestionnairePreview()}
            >
              {panel.isLoadingAttachments
                ? "Завантаження анкети…"
                : panel.questionnaire
                  ? "Відкрити анкету (PDF)"
                  : panel.match
                    ? "Анкета не додана"
                    : "Анкета недоступна"}
            </Button>
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
                openPersonnelForMatch(panel.match);
              }}
            >
              Відкрити в особовому складі
            </Button>
          </div>

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

          {panel.questionnaire ? (
            <Typography variant="caption" color="text.secondary" sx={{ px: 0.5 }}>
              {panel.exportFileName}
            </Typography>
          ) : null}
        </div>
      </aside>

      <FloatingQuestionnairePreview
        open={panel.previewOpen}
        title={panel.previewTitle}
        previewUrl={panel.previewUrl}
        pendingFile={false}
        isUploading={false}
        placement="left"
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
