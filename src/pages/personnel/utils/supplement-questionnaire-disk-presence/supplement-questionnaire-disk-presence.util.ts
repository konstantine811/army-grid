import { api } from "../../../../api";
import {
  readDataCache,
  questionnaireDiskPresenceCacheKey,
  writeDataCache,
} from "../../../../data/idbDataCache";
import {
  mergeQuestionnaireDiskPresence,
  type QuestionnaireDiskScanPerson,
  type QuestionnairePresencePerson,
} from "../../personAttachments";
import { searchQuestionnairesOnDiskInBatches } from "../../questionnaireDiskSearchBatch";

export const supplementQuestionnaireDiskPresence = async (input: {
  dbPresence: Record<string, true>;
  people: QuestionnairePresencePerson[];
  inStaffIds: ReadonlySet<string>;
  datasetFingerprint: string;
  signal?: AbortSignal;
  force?: boolean;
  onProgress?: (done: number, total: number) => void;
}): Promise<Record<string, true>> => {
  const scanPeople = input.people.filter(
    (person) =>
      input.inStaffIds.has(person.currentId) && !input.dbPresence[person.currentId],
  );
  if (!scanPeople.length) return input.dbPresence;

  const cacheKey = questionnaireDiskPresenceCacheKey(
    input.datasetFingerprint,
    scanPeople.length,
  );
  if (!input.force) {
    const cached = await readDataCache<Record<string, true>>(cacheKey);
    if (cached) {
      return { ...input.dbPresence, ...cached };
    }
  }

  const result = await searchQuestionnairesOnDiskInBatches({
    people: scanPeople.map((person) => ({
      externalId: person.currentId,
      fullName: person.fullName,
      callSign: person.callSign,
      missingQuestionnaire: true,
    })),
    search: (payload) => api.searchQuestionnairesOnDisk(payload),
    isCancelled: () => Boolean(input.signal?.aborted),
    onProgress: input.onProgress,
  });

  const diskResults: QuestionnaireDiskScanPerson[] = result.people.map((person) => ({
    externalId: person.externalId,
    fullName: person.fullName,
    callSign: person.callSign,
    matches: person.matches.map((match) => ({
      fileName: match.fileName,
      relativePath: match.relativePath,
    })),
  }));

  const peopleById = new Map(input.people.map((person) => [person.currentId, person]));
  const merged = mergeQuestionnaireDiskPresence(
    input.dbPresence,
    diskResults,
    peopleById,
  );

  const diskOnly: Record<string, true> = {};
  for (const person of scanPeople) {
    if (merged[person.currentId] && !input.dbPresence[person.currentId]) {
      diskOnly[person.currentId] = true;
    }
  }
  void writeDataCache(cacheKey, diskOnly);

  return merged;
};
