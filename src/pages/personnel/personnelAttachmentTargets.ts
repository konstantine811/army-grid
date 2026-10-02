export type PersonnelAttachmentPerson = {
  inStaff: boolean;
  rowId: string;
  externalId: string;
  fullName: string;
  callSign: string;
};

export const collectMissingQuestionnairePeople = (
  people: PersonnelAttachmentPerson[],
  questionnaireByExternalId: Record<string, true>,
  photoByExternalId: Record<string, string>,
) =>
  people.flatMap((person) => {
    if (!person.rowId || !person.externalId) return [];
    if (questionnaireByExternalId[person.externalId]) return [];
    return [
      {
        rowId: person.rowId,
        externalId: person.externalId,
        fullName: person.fullName,
        callSign: person.callSign,
        missingQuestionnaire: true,
        missingPhoto: !photoByExternalId[person.externalId],
      },
    ];
  });

export const collectPhotoExtractTargets = (
  people: PersonnelAttachmentPerson[],
  questionnaireByExternalId: Record<string, true>,
  photoByExternalId: Record<string, string>,
) =>
  people.flatMap((person) => {
    if (!person.inStaff || !person.externalId) return [];
    if (!questionnaireByExternalId[person.externalId]) return [];
    if (photoByExternalId[person.externalId]) return [];
    return [{ externalId: person.externalId, fullName: person.fullName }];
  });
