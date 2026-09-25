import type { BackendPersonnelRosterLatest } from "../../api";
import { normalizeAnketaNameKey } from "../anketa-data/anketaPersonMatch";
import { countStaffSheetPersonsInRoster } from "../anketa-data/staffSheetPreview";
import { mapRosterLatestToPreviewRows } from "../excel-fill/rosterSourceSnapshot";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildQuestionnairePresenceFromPeople,
  buildQuestionnairePresencePeople,
  type QuestionnairePresencePerson,
} from "./personAttachments";
import {
  buildPersonnelListIndex,
  type PersonnelListRecord,
} from "./personnelListIndex";
import { fillDownRosterUnitRows } from "./rosterRowFill";
import { getRosterPersonName } from "./personnelRosterMerge";

export type StaffRosterQuestionnaireExportTarget = {
  rosterRow: EjournalPreviewRow;
  record: PersonnelListRecord;
  person: QuestionnairePresencePerson;
  hasQuestionnaire: boolean;
};

const pickStaffRecordForRosterRow = (
  rosterRow: EjournalPreviewRow,
  inStaffById: Map<string, PersonnelListRecord>,
  inStaffByName: Map<string, PersonnelListRecord[]>,
  presence: Record<string, true>,
): PersonnelListRecord | null => {
  const rowId = String(rosterRow.__dbRowId ?? "");
  const byId = rowId ? inStaffById.get(rowId) : undefined;
  if (byId) return byId;

  const nameKey = normalizeAnketaNameKey(getRosterPersonName(rosterRow));
  const candidates = nameKey ? (inStaffByName.get(nameKey) ?? []) : [];
  if (candidates.length === 1) return candidates[0]!;

  const withQuestionnaire = candidates.filter(
    (record) => presence[record.summary.externalId],
  );
  if (withQuestionnaire.length === 1) return withQuestionnaire[0]!;

  const nonRoster = candidates.filter(
    (record) => !/^roster:/i.test(String(record.row.__dbRowId ?? "")),
  );
  if (nonRoster.length === 1) return nonRoster[0]!;
  if (nonRoster.length) {
    return (
      nonRoster.find((record) => presence[record.summary.externalId]) ??
      nonRoster[0]!
    );
  }

  return candidates[0] ?? null;
};

/** Один рядок на людину з актуального «Загального списку» Штатки (675), як у UI. */
export const buildStaffRosterQuestionnaireExportTargets = (input: {
  personnelRows: EjournalPreviewRow[];
  rosterLatest: BackendPersonnelRosterLatest | null | undefined;
  questionnaires: Array<{
    personExternalId?: string | null;
    fileName?: string | null;
  }>;
}): StaffRosterQuestionnaireExportTarget[] => {
  const rosterStaffRows = fillDownRosterUnitRows(
    mapRosterLatestToPreviewRows(input.rosterLatest),
  ).filter((row) => countStaffSheetPersonsInRoster([row]));

  const listIndex = buildPersonnelListIndex(input.personnelRows);
  const people = buildQuestionnairePresencePeople(input.personnelRows);
  const peopleById = new Map(people.map((person) => [person.currentId, person]));
  const presence = buildQuestionnairePresenceFromPeople(
    people,
    input.questionnaires,
  );

  const inStaffById = new Map<string, PersonnelListRecord>();
  const inStaffByName = new Map<string, PersonnelListRecord[]>();
  for (const record of listIndex.records) {
    if (!record.inStaff) continue;
    inStaffById.set(String(record.row.__dbRowId ?? ""), record);
    const nameKey = normalizeAnketaNameKey(record.summary.name);
    if (!nameKey) continue;
    const list = inStaffByName.get(nameKey) ?? [];
    list.push(record);
    inStaffByName.set(nameKey, list);
  }

  const targets: StaffRosterQuestionnaireExportTarget[] = [];
  for (const rosterRow of rosterStaffRows) {
    const record = pickStaffRecordForRosterRow(
      rosterRow,
      inStaffById,
      inStaffByName,
      presence,
    );
    if (!record) continue;
    const person = peopleById.get(record.summary.externalId);
    if (!person) continue;
    targets.push({
      rosterRow,
      record,
      person,
      hasQuestionnaire: Boolean(presence[record.summary.externalId]),
    });
  }

  return targets;
};
