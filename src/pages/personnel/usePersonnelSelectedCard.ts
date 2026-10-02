import { useEffect, useMemo, useState } from "react";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  extractFighterStatusFieldRows,
} from "./fighterStatusImport";
import {
  matchPersonnelV2AnketaRow,
  summaryPreferringAnketa,
  type PersonnelV2AnketaIndex,
} from "../personnel-v2/personnelV2AnketaFill";
import {
  buildVisibleRosterFieldRows,
  resolvePersonnelPersonNote,
} from "./personnelCardFields";
import {
  groupEditablePersonFields,
  resolveEditablePersonFields,
} from "./personnelEditableFields";
import { personEditDraft } from "./personnelSelectedCard";
import {
  extractPersonCallSign,
  extractPhones,
  formatPersonBirthDateWithAge,
  getPersonFullPositionTitle,
  pickFullPositionFromPersonRow,
  resolvePersonBirthDate,
  resolvePersonRosterStatus,
  buildPersonSummary,
} from "./personnelUtils";
import {
  isPersonPhotoDisplayOverride,
  personPhotoCacheBustFromUrl,
  personPhotoFullUrlForRow,
} from "./personAttachments";
import { uniqueNormalizedPhones } from "./personPhonesStore";

export const usePersonnelSelectedCard = ({
  selectedRow,
  rosterLabels,
  anketaIndex,
  phonesByExternalId,
  photoByExternalId,
  photoIndexReady,
  questionnaireFileName,
  setPhoneDraft,
}: {
  selectedRow: EjournalPreviewRow | null;
  rosterLabels: Record<string, string>;
  anketaIndex: PersonnelV2AnketaIndex | null;
  phonesByExternalId: Record<string, string[]>;
  photoByExternalId: Record<string, string>;
  photoIndexReady: number;
  questionnaireFileName?: string;
  setPhoneDraft: (value: string) => void;
}) => {
  const [editValues, setEditValues] = useState<Record<string, string>>({});
  const selectedSummary = useMemo(
    () => buildPersonSummary(selectedRow),
    [selectedRow],
  );
  const selectedPhotoLocalUrl = useMemo(() => {
    const externalId = selectedSummary.externalId;
    return externalId ? photoByExternalId[externalId] || "" : "";
  }, [photoByExternalId, selectedSummary.externalId]);
  const selectedPhotoCacheBust = useMemo(
    () => personPhotoCacheBustFromUrl(selectedPhotoLocalUrl),
    [selectedPhotoLocalUrl],
  );
  const selectedPhotoFullUrl = useMemo(
    () =>
      selectedRow
        ? personPhotoFullUrlForRow(
            selectedRow,
            undefined,
            undefined,
            selectedPhotoCacheBust,
          )
        : "",
    [photoIndexReady, selectedRow, selectedPhotoCacheBust],
  );
  const selectedPhoto = useMemo(() => {
    const externalId = selectedSummary.externalId;
    if (!externalId) return "";
    const local = photoByExternalId[externalId] || "";
    if (local && isPersonPhotoDisplayOverride(local)) return local;
    if (selectedPhotoFullUrl) return selectedPhotoFullUrl;
    return local;
  }, [photoByExternalId, selectedPhotoFullUrl, selectedSummary.externalId]);
  const selectedCallSign = useMemo(
    () =>
      selectedSummary.callSign?.trim() ||
      extractPersonCallSign(
        questionnaireFileName,
        selectedSummary.name,
        selectedSummary.additionalInfo,
      ),
    [
      questionnaireFileName,
      selectedSummary.additionalInfo,
      selectedSummary.callSign,
      selectedSummary.name,
    ],
  );
  const editableFields = useMemo(
    () => resolveEditablePersonFields(selectedRow),
    [selectedRow],
  );
  const editableFieldsBySection = useMemo(
    () =>
      groupEditablePersonFields(
        editableFields,
        selectedRow,
        selectedSummary.birthDate,
      ),
    [editableFields, selectedRow, selectedSummary.birthDate],
  );
  const selectedAnketaRow = useMemo(() => {
    if (!anketaIndex || !selectedRow) return null;
    return matchPersonnelV2AnketaRow(
      anketaIndex,
      selectedSummary.name,
      resolvePersonBirthDate(selectedRow),
      selectedSummary.externalId,
    );
  }, [
    anketaIndex,
    selectedRow,
    selectedSummary.externalId,
    selectedSummary.name,
  ]);
  const cardSummary = useMemo(
    () =>
      summaryPreferringAnketa(selectedSummary, selectedAnketaRow) ??
      selectedSummary,
    [selectedAnketaRow, selectedSummary],
  );
  const birthDateWithAge = useMemo(
    () => formatPersonBirthDateWithAge(cardSummary.birthDate),
    [cardSummary.birthDate],
  );
  const rosterFieldRows = useMemo(
    () =>
      buildVisibleRosterFieldRows({
        row: selectedRow,
        rosterLabels,
        personName: selectedSummary.name,
        birthDate: selectedSummary.birthDate,
        cardRnokpp: cardSummary.rnokpp,
      }),
    [
      cardSummary.rnokpp,
      rosterLabels,
      selectedRow,
      selectedSummary.birthDate,
      selectedSummary.name,
    ],
  );
  const rosterStatus = useMemo(
    () => resolvePersonRosterStatus(selectedRow, rosterLabels),
    [rosterLabels, selectedRow],
  );
  const fighterStatusFieldRows = useMemo(
    () =>
      extractFighterStatusFieldRows(selectedRow, rosterLabels).filter(
        (field) => field.key !== "fighter_status_note",
      ),
    [rosterLabels, selectedRow],
  );
  const selectedFullPosition = useMemo(
    () =>
      getPersonFullPositionTitle(selectedRow) ||
      pickFullPositionFromPersonRow(selectedRow) ||
      selectedSummary.positionTitle ||
      "",
    [selectedRow, selectedSummary.positionTitle],
  );
  const selectedPersonNote = useMemo(
    () => resolvePersonnelPersonNote(selectedRow, rosterLabels),
    [rosterLabels, selectedRow],
  );
  const savedPhones = useMemo(
    () =>
      (selectedSummary.externalId &&
        phonesByExternalId[selectedSummary.externalId]) ||
      [],
    [phonesByExternalId, selectedSummary.externalId],
  );
  const parsedPhones = useMemo(
    () =>
      uniqueNormalizedPhones([
        ...extractPhones(selectedAnketaRow?.additionalInfo ?? ""),
        ...savedPhones,
      ]),
    [savedPhones, selectedAnketaRow?.additionalInfo],
  );

  useEffect(() => {
    setEditValues(personEditDraft(editableFields, selectedRow));
    setPhoneDraft("");
  }, [editableFields, selectedRow, setPhoneDraft]);

  return {
    selectedSummary,
    selectedPhoto,
    selectedPhotoFullUrl,
    selectedCallSign,
    editableFields,
    editableFieldsBySection,
    editValues,
    setEditValues,
    selectedAnketaRow,
    cardSummary,
    birthDateWithAge,
    rosterFieldRows,
    rosterStatus,
    fighterStatusFieldRows,
    selectedFullPosition,
    selectedPersonNote,
    savedPhones,
    parsedPhones,
  };
};
