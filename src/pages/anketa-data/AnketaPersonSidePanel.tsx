import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  api,
  type BackendPersonQuestionnaire,
} from "../../api";
import { buildPersonnelRoute } from "../../app/navigation";
import {
  subscribePersonnelAttachmentChanges,
} from "../../shared/personnelAttachmentSync";
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
import {
  buildQuestionnaireExportFileName,
  downloadQuestionnairePdf,
  revokeQuestionnairePreviewUrl,
} from "../personnel/personnelUtils";
import type { QuestionnairePdfSource } from "../personnel/questionnaireShare";
import {
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnel,
  matchLabel,
  type AnketaPersonnelMatch,
} from "./anketaPersonMatch";
import type { AnketaColumnKey, AnketaRow } from "./anketaSheet";
import type { AnketaEmptyCell } from "./anketaGaps";
import {
  ANKETA_MISSING_VALUE_PRESETS,
  listAnketaEmptyCells,
} from "./anketaGaps";

type AnketaPersonSidePanelProps = {
  anketaRow: AnketaRow | null;
  focusedEmpty: AnketaEmptyCell | null;
  gapColumnKeys: AnketaColumnKey[];
  onClose: () => void;
  onFillMissing?: (value: string) => void;
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
}: AnketaPersonSidePanelProps) {
  const [match, setMatch] = useState<AnketaPersonnelMatch | null>(null);
  const [matchStatus, setMatchStatus] = useState<"idle" | "loading" | "ready">(
    "idle",
  );
  const [photoData, setPhotoData] = useState("");
  const [questionnaire, setQuestionnaire] =
    useState<BackendPersonQuestionnaire | null>(null);
  const [isLoadingAttachments, setIsLoadingAttachments] = useState(false);
  const [previewOpen, setPreviewOpen] = useState(false);
  const [previewUrl, setPreviewUrl] = useState("");
  const [previewTitle, setPreviewTitle] = useState("");
  const attachmentsEpochRef = useRef(0);

  const reloadAttachments = useCallback(async (externalId: string) => {
    const epoch = ++attachmentsEpochRef.current;
    setIsLoadingAttachments(true);
    try {
      const [photo, nextQuestionnaire] = await Promise.all([
        api.getPersonPhoto(externalId).catch(() => null),
        api.getPersonQuestionnaire(externalId).catch(() => null),
      ]);
      if (epoch !== attachmentsEpochRef.current) return;
      setPhotoData(photo?.photoData?.trim() || "");
      setQuestionnaire(nextQuestionnaire);
    } finally {
      if (epoch === attachmentsEpochRef.current) {
        setIsLoadingAttachments(false);
      }
    }
  }, []);

  const rowGaps = useMemo(() => {
    if (!anketaRow) return [];
    return listAnketaEmptyCells([anketaRow], gapColumnKeys);
  }, [anketaRow, gapColumnKeys]);

  useEffect(() => {
    let cancelled = false;
    if (!anketaRow) {
      setMatch(null);
      setMatchStatus("idle");
      return;
    }
    setMatchStatus("loading");
    void loadPersonnelIndexForAnketa()
      .then((index) => {
        if (cancelled) return;
        setMatch(matchAnketaRowToPersonnel(anketaRow, index));
        setMatchStatus("ready");
      })
      .catch(() => {
        if (cancelled) return;
        setMatch(null);
        setMatchStatus("ready");
      });
    return () => {
      cancelled = true;
    };
  }, [anketaRow]);

  const personnelExternalId = match?.summary.externalId ?? "";
  const displayName =
    anketaRow?.fullName?.trim() || match?.summary.name || "Службовець";
  const exportFileName = useMemo(
    () =>
      buildQuestionnaireExportFileName(
        displayName,
        match?.summary.callSign ?? "",
      ),
    [displayName, match?.summary.callSign],
  );

  const shareSource = useMemo<QuestionnairePdfSource | null>(() => {
    if (!questionnaire?.fileData) return null;
    return { fileData: questionnaire.fileData };
  }, [questionnaire]);

  useEffect(() => {
    setPhotoData("");
    setQuestionnaire(null);
    setPreviewOpen(false);
    setPreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return "";
    });

    if (!personnelExternalId) return;

    void reloadAttachments(personnelExternalId);
  }, [personnelExternalId, reloadAttachments]);

  useEffect(() => {
    if (!personnelExternalId) return;

    return subscribePersonnelAttachmentChanges((change) => {
      if (change.externalId !== personnelExternalId) return;
      void reloadAttachments(personnelExternalId);
    });
  }, [personnelExternalId, reloadAttachments]);

  useEffect(() => {
    if (!personnelExternalId) return;

    const refreshOnFocus = () => {
      if (document.visibilityState !== "visible") return;
      void reloadAttachments(personnelExternalId);
    };

    window.addEventListener("focus", refreshOnFocus);
    document.addEventListener("visibilitychange", refreshOnFocus);
    return () => {
      window.removeEventListener("focus", refreshOnFocus);
      document.removeEventListener("visibilitychange", refreshOnFocus);
    };
  }, [personnelExternalId, reloadAttachments]);

  useEffect(() => {
    return () => {
      setPreviewUrl((current) => {
        if (current) revokeQuestionnairePreviewUrl(current);
        return "";
      });
    };
  }, []);

  const openQuestionnairePreview = () => {
    if (!questionnaire?.fileData || !personnelExternalId) return;
    const nextUrl = api.getPersonQuestionnaireFileUrl(
      personnelExternalId,
      exportFileName,
    );
    setPreviewTitle(`Анкета · ${displayName} · ${exportFileName}`);
    setPreviewUrl((current) => {
      if (current) revokeQuestionnairePreviewUrl(current);
      return nextUrl;
    });
    setPreviewOpen(true);
  };

  const openQuestionnaireTab = () => {
    if (!questionnaire?.fileData || !personnelExternalId) return;
    window.open(
      api.getPersonQuestionnaireFileUrl(personnelExternalId, exportFileName),
      "_blank",
      "noopener,noreferrer",
    );
  };

  const downloadQuestionnaire = () => {
    if (!questionnaire?.fileData) return;
    downloadQuestionnairePdf(exportFileName, {
      fileData: questionnaire.fileData,
    });
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
            {photoData ? (
              <img alt={displayName} src={photoData} />
            ) : (
              <PersonSearchOutlinedIcon />
            )}
          </div>
          <div>
            {match?.summary.callSign ? (
              <div className="person-callsign-row">
                <span className="person-callsign" title="Позивний">
                  <span className="person-callsign-label">позивний</span>
                  <strong>{match.summary.callSign}</strong>
                </span>
              </div>
            ) : null}
            <Typography component="h2" variant="h5">
              {displayName}
            </Typography>
            <div className="person-action-tags">
              {(anketaRow.rank || match?.summary.rank) && (
                <Chip
                  label={anketaRow.rank || match?.summary.rank}
                  size="small"
                />
              )}
              {(anketaRow.positionIndex || match?.summary.positionIndex) && (
                <Chip
                  label={`Посада: ${
                    anketaRow.positionIndex || match?.summary.positionIndex
                  }`}
                  size="small"
                />
              )}
              {(anketaRow.serviceType || match?.summary.serviceType) && (
                <Chip
                  label={anketaRow.serviceType || match?.summary.serviceType}
                  size="small"
                />
              )}
            </div>
            <Typography variant="caption" color="text.secondary">
              {matchStatus === "loading"
                ? "Шукаю в особовому складі…"
                : match
                  ? `Знайдено в ООС · ${matchLabel(match.matchBy)}`
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
            {isLoadingAttachments ? (
              <div className="person-document-empty">Завантаження…</div>
            ) : questionnaire?.fileData ? (
              <article className="person-document-shell is-ready">
                <button
                  className="person-document-item is-ready"
                  type="button"
                  onClick={openQuestionnairePreview}
                >
                  <PictureAsPdfOutlinedIcon />
                  <span>
                    <strong>Анкета (PDF)</strong>
                    <small>{exportFileName} · переглянути</small>
                  </span>
                </button>
              </article>
            ) : (
              <div className="person-document-empty">
                <PictureAsPdfOutlinedIcon />
                <span>
                  {personnelExternalId
                    ? "Анкета ще не додана в БД"
                    : "Потрібен зв’язок з особовим складом"}
                </span>
              </div>
            )}
          </div>

          <div className="anketa-person-side-actions">
            <Button
              variant="outlined"
              size="small"
              disabled={!match}
              onClick={() => {
                if (!match) return;
                const target = {
                  rowId: match.row.__dbRowId,
                  externalId: match.summary.externalId,
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
        open={previewOpen}
        title={previewTitle}
        previewUrl={previewUrl}
        pendingFile={false}
        isUploading={false}
        placement="right"
        shareFileName={exportFileName}
        sharePersonName={displayName}
        shareSource={shareSource}
        onClose={() => {
          setPreviewOpen(false);
          setPreviewUrl((current) => {
            if (current) revokeQuestionnairePreviewUrl(current);
            return "";
          });
        }}
        onOpenTab={openQuestionnaireTab}
        onDownload={downloadQuestionnaire}
      />
    </>
  );
}
