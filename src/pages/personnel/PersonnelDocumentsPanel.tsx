import { Button } from "@/components/sci/SciPrimitives";
import {
  ArticleOutlinedIcon,
  ArrowLeftOutlinedIcon,
  DeleteOutlineOutlinedIcon,
  FileDownloadOutlinedIcon,
  PictureAsPdfOutlinedIcon,
  SearchOutlinedIcon,
} from "@/components/sci/icons";
import type { BackendPersonDocument } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";

export type PersonnelDocumentMode =
  | "default"
  | "salaryPowerAttorney"
  | "ubdReport"
  | "form6Report"
  | "form12Report"
  | "serviceCharacteristic"
  | "zhbdCertificate"
  | "ubdRestoreReport"
  | "temporaryMilitaryId"
  | "lostMilitaryId";

type DocumentAction = {
  mode: Exclude<PersonnelDocumentMode, "default">;
  title: string;
  hint: string;
};

const DOCUMENT_ACTIONS: DocumentAction[] = [
  {
    mode: "salaryPowerAttorney",
    title: "Довіреність зарплати",
    hint: "створити документ і вести прогрес",
  },
  {
    mode: "ubdReport",
    title: "Рапорт на УБД",
    hint: "рапорт, скани документів, статус",
  },
  {
    mode: "ubdRestoreReport",
    title: "Рапорт на відновлення УБД",
    hint: "пошкоджене посвідчення, клопотання, скани",
  },
  {
    mode: "form6Report",
    title: "Форма 6",
    hint: "рапорт для довідки УБД, персональні дані, скани",
  },
  {
    mode: "form12Report",
    title: "Форма 12",
    hint: "рапорт Ф-12, дані бійця, підпис PNG",
  },
  {
    mode: "serviceCharacteristic",
    title: "Службова характеристика",
    hint: "звання, ПІБ, посада, текст, підпис командира",
  },
  {
    mode: "zhbdCertificate",
    title: "Довідка ЖБД",
    hint: "період, посада, підстава, підпис",
  },
  {
    mode: "temporaryMilitaryId",
    title: "Тимчасовий військовий квиток",
    hint: "фото, рядок для замовлення, прогрес",
  },
  {
    mode: "lostMilitaryId",
    title: "Втрата військового квитка",
    hint: "рапорт, наказ, акт розслідування",
  },
];

const DOCUMENT_PLACEHOLDERS = [
  "Довідка про проходження служби",
  "Витяг з наказу",
  "Рапорт",
];

export function PersonnelDocumentsPanel({
  selectedRow,
  externalId,
  questionnaireExists,
  questionnaireFileName,
  relatedDocuments,
  isUploadingQuestionnaire,
  fullPosition,
  onBackToCard,
  onBackToList,
  onAddQuestionnaire,
  onOpenQuestionnaire,
  onExportQuestionnaire,
  onRevealQuestionnaire,
  onDeleteQuestionnaire,
  onOpenDocument,
}: {
  selectedRow: EjournalPreviewRow | null | undefined;
  externalId: string;
  questionnaireExists: boolean;
  questionnaireFileName: string;
  relatedDocuments: BackendPersonDocument[];
  isUploadingQuestionnaire: boolean;
  fullPosition: string;
  onBackToCard: () => void;
  onBackToList: () => void;
  onAddQuestionnaire: (file?: File) => void;
  onOpenQuestionnaire: () => void;
  onExportQuestionnaire: () => void;
  onRevealQuestionnaire: () => void;
  onDeleteQuestionnaire: () => void;
  onOpenDocument: (
    row: EjournalPreviewRow,
    mode: PersonnelDocumentMode,
    meta?: { fullPosition?: string },
  ) => void;
}) {
  const readyTypes = new Set(relatedDocuments.map((document) => document.type));

  return (
    <aside className="person-side-panel">
      <div className="personnel-mobile-card-nav">
        <Button
          size="small"
          variant="outlined"
          startIcon={<ArrowLeftOutlinedIcon fontSize="small" />}
          onClick={onBackToCard}
        >
          До картки
        </Button>
        <Button size="small" variant="outlined" onClick={onBackToList}>
          До списку
        </Button>
      </div>
      <div className="analytics-panel">
        <div className="person-documents-header">
          <div className="panel-heading">Пов’язані документи</div>
          <Button
            component="label"
            disabled={!selectedRow || !externalId || isUploadingQuestionnaire}
            size="small"
            variant="outlined"
            startIcon={<PictureAsPdfOutlinedIcon />}
          >
            Додати анкету
            <input
              hidden
              type="file"
              accept="application/pdf,.pdf"
              onChange={(event) => {
                onAddQuestionnaire(event.target.files?.[0]);
                event.target.value = "";
              }}
            />
          </Button>
        </div>
        <div className="person-document-list">
          {questionnaireExists ? (
            <article className="person-document-shell is-ready">
              <button
                className="person-document-item is-ready"
                type="button"
                onClick={onOpenQuestionnaire}
              >
                <PictureAsPdfOutlinedIcon />
                <span>
                  <strong>Анкета (PDF)</strong>
                  <small>{questionnaireFileName} · переглянути</small>
                </span>
              </button>
              <div className="person-document-actions">
                <button
                  aria-label="Експорт анкети"
                  className="person-document-delete"
                  disabled={!selectedRow}
                  onClick={onExportQuestionnaire}
                  title={`Експорт: ${questionnaireFileName}`}
                  type="button"
                >
                  <FileDownloadOutlinedIcon />
                </button>
                <button
                  aria-label="Показати анкету у Finder"
                  className="person-document-delete person-document-action--finder"
                  disabled={!selectedRow}
                  onClick={onRevealQuestionnaire}
                  title="Показати оригінал у Finder"
                  type="button"
                >
                  <SearchOutlinedIcon />
                </button>
                <button
                  aria-label="Видалити анкету"
                  className="person-document-delete"
                  disabled={!selectedRow}
                  onClick={onDeleteQuestionnaire}
                  title="Видалити анкету"
                  type="button"
                >
                  <DeleteOutlineOutlinedIcon />
                </button>
              </div>
            </article>
          ) : (
            <div className="person-document-empty">
              <PictureAsPdfOutlinedIcon />
              <span>Анкета ще не додана</span>
            </div>
          )}

          {DOCUMENT_PLACEHOLDERS.map((item) => (
            <div key={item}>
              <ArticleOutlinedIcon />
              <span>{item}</span>
            </div>
          ))}

          {DOCUMENT_ACTIONS.map(({ mode, title, hint }) => (
            <button
              className={`person-document-item${readyTypes.has(mode) ? " is-ready" : ""}`}
              disabled={!selectedRow}
              key={mode}
              type="button"
              onClick={() => {
                if (!selectedRow) return;
                onOpenDocument(
                  selectedRow,
                  mode,
                  mode === "zhbdCertificate" ? { fullPosition } : undefined,
                );
              }}
            >
              <ArticleOutlinedIcon />
              <span>
                <strong>{title}</strong>
                <small>{hint}</small>
              </span>
            </button>
          ))}
        </div>
      </div>
    </aside>
  );
}
