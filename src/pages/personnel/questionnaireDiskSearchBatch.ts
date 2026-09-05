import type { DiskQuestionnaireSearchResult } from "../../api";

export type QuestionnaireDiskSearchInput = {
  externalId: string;
  fullName: string;
  callSign?: string;
  missingQuestionnaire?: boolean;
  missingPhoto?: boolean;
};

export const QUESTIONNAIRE_DISK_SEARCH_BATCH_SIZE = 60;
export const QUESTIONNAIRE_DISK_SEARCH_PAUSE_MS = 75;

const pause = (delayMs: number) =>
  new Promise<void>((resolve) => globalThis.setTimeout(resolve, delayMs));

/** Keep large disk searches bounded so one response cannot exhaust the browser. */
export const searchQuestionnairesOnDiskInBatches = async (input: {
  people: QuestionnaireDiskSearchInput[];
  search: (payload: {
    people: QuestionnaireDiskSearchInput[];
    refreshIndex?: boolean;
  }) => Promise<DiskQuestionnaireSearchResult>;
  isCancelled?: () => boolean;
  onProgress?: (done: number, total: number) => void;
  batchSize?: number;
  pauseMs?: number;
  /** Force a new USB directory scan instead of using the backend TTL cache. */
  refreshIndex?: boolean;
}): Promise<DiskQuestionnaireSearchResult> => {
  const batchSize = Math.max(1, input.batchSize ?? QUESTIONNAIRE_DISK_SEARCH_BATCH_SIZE);
  const pauseMs = Math.max(0, input.pauseMs ?? QUESTIONNAIRE_DISK_SEARCH_PAUSE_MS);
  let root = "";
  let scannedFiles = 0;
  const people: DiskQuestionnaireSearchResult["people"] = [];

  for (let offset = 0; offset < input.people.length; offset += batchSize) {
    if (input.isCancelled?.()) break;
    const batch = input.people.slice(offset, offset + batchSize);
    const result = await input.search({
      people: batch,
      // At most one forced USB scan. Following batches reuse the backend cache.
      refreshIndex: offset === 0 && Boolean(input.refreshIndex),
    });
    root ||= result.root;
    scannedFiles = Math.max(scannedFiles, result.scannedFiles);
    people.push(...result.people);
    input.onProgress?.(people.length, input.people.length);
    if (offset + batchSize < input.people.length && pauseMs > 0) {
      await pause(pauseMs);
    }
  }

  return {
    root,
    scannedFiles,
    matchedPeople: people.filter((person) => person.matches.length > 0).length,
    people,
  };
};
