import type { Dispatch, SetStateAction } from "react";
import type { BackendPersonQuestionnaire } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { PersonCardAvatar } from "./card/PersonCardAvatar";
import { PersonCardFacts } from "./card/PersonCardFacts";
import { PersonCardIdentity } from "./card/PersonCardIdentity";
import { PersonCardPhones } from "./card/PersonCardPhones";
import { PersonEditableFields } from "./card/PersonEditableFields";
import { PersonFighterStatusSection } from "./card/PersonFighterStatusSection";
import { PersonRosterFieldsSection } from "./card/PersonRosterFieldsSection";
import type {
  AnketaRow,
  EditablePersonSection,
  FighterStatusCardField,
  PersonCardSummary,
  PersonnelRosterCardField,
} from "./card/personCardTypes";

export function PersonnelPersonCard({
  selectedPhoto,
  selectedSummary,
  selectedRowId,
  questionnaire,
  questionnaireByExternalId,
  selectedRow,
  selectedCallSign,
  rosterStatus,
  cardSummary,
  parsedPhones,
  savedPhones,
  isSavingPhone,
  phoneDraft,
  birthDateWithAge,
  selectedFullPosition,
  selectedPersonNote,
  fighterStatusFieldRows,
  rosterFieldRows,
  editableFieldsBySection,
  selectedAnketaRow,
  editValues,
  setIsPhotoLightboxOpen,
  setMessage,
  openQuestionnairePreview,
  deleteSelectedPhoto,
  openPhotoCrop,
  copySelectedPersonPhone,
  removeSelectedPersonPhone,
  setPhoneDraft,
  addSelectedPersonPhone,
  setEditValues,
  saveSelectedPerson,
}: {
  selectedPhoto: string;
  selectedSummary: PersonCardSummary;
  selectedRowId: string;
  questionnaire: BackendPersonQuestionnaire | null;
  questionnaireByExternalId: Record<string, true>;
  selectedRow: EjournalPreviewRow | null;
  selectedCallSign: string;
  rosterStatus: string;
  cardSummary: PersonCardSummary;
  parsedPhones: string[];
  savedPhones: string[];
  isSavingPhone: boolean;
  phoneDraft: string;
  birthDateWithAge: string;
  selectedFullPosition: string;
  selectedPersonNote: string;
  fighterStatusFieldRows: FighterStatusCardField[];
  rosterFieldRows: PersonnelRosterCardField[];
  editableFieldsBySection: EditablePersonSection[];
  selectedAnketaRow: AnketaRow | null;
  editValues: Record<string, string>;
  setIsPhotoLightboxOpen: (open: boolean) => void;
  setMessage: (message: string) => void;
  openQuestionnairePreview: () => void | Promise<void>;
  deleteSelectedPhoto: () => void | Promise<void>;
  openPhotoCrop: (file: File | undefined) => void;
  copySelectedPersonPhone: (phone: string) => void | Promise<void>;
  removeSelectedPersonPhone: (phone: string) => void | Promise<void>;
  setPhoneDraft: (value: string) => void;
  addSelectedPersonPhone: () => void | Promise<void>;
  setEditValues: Dispatch<SetStateAction<Record<string, string>>>;
  saveSelectedPerson: () => void | Promise<void>;
}) {
  const hasQuestionnaire =
    Boolean(questionnaire?.personExternalId) ||
    Boolean(
      selectedSummary.externalId &&
        questionnaireByExternalId[selectedSummary.externalId],
    );

  return (
    <section className="person-card-panel">
      <div className="person-card-hero">
        <PersonCardAvatar
          photo={selectedPhoto}
          name={selectedSummary.name}
          canOpen={Boolean(selectedRowId)}
          hasRow={Boolean(selectedRow)}
          hasQuestionnaire={hasQuestionnaire}
          onOpenPhoto={() => setIsPhotoLightboxOpen(true)}
          onMissingQuestionnaire={() => setMessage("Анкета ще не додана.")}
          onOpenQuestionnaire={openQuestionnairePreview}
          onDeletePhoto={deleteSelectedPhoto}
          onAddPhoto={openPhotoCrop}
        />
        <PersonCardIdentity
          name={selectedSummary.name}
          callSign={selectedCallSign}
          rosterStatus={rosterStatus}
          rank={cardSummary.rank}
          positionIndex={cardSummary.positionIndex}
          serviceType={cardSummary.serviceType}
        />
      </div>
      <div className="person-card-scroll">
        <div className="person-action-fields">
          <PersonCardPhones
            phones={parsedPhones}
            savedPhones={savedPhones}
            draft={phoneDraft}
            disabled={!selectedRow || isSavingPhone}
            onDraftChange={setPhoneDraft}
            onAdd={addSelectedPersonPhone}
            onCopy={copySelectedPersonPhone}
            onRemove={removeSelectedPersonPhone}
          />
          <PersonCardFacts
            rnokpp={cardSummary.rnokpp}
            birthDateWithAge={birthDateWithAge}
            location={cardSummary.location}
            fullPosition={selectedFullPosition}
            note={selectedPersonNote}
            positionTitle={selectedSummary.positionTitle}
            militaryId={cardSummary.militaryId}
            arrivedFrom={cardSummary.arrivedFrom}
          />
        </div>
        <PersonFighterStatusSection fields={fighterStatusFieldRows} />
        <PersonRosterFieldsSection fields={rosterFieldRows} />
        <PersonEditableFields
          sections={editableFieldsBySection}
          anketa={selectedAnketaRow}
          editValues={editValues}
          birthDate={selectedSummary.birthDate}
          onChange={(key, value) =>
            setEditValues((values) => ({ ...values, [key]: value }))
          }
          onSave={saveSelectedPerson}
        />
      </div>
    </section>
  );
}
