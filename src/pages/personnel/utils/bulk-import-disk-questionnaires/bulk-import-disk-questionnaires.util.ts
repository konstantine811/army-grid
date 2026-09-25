import { api } from "../../../../api";
import { sanitizeFileName } from "../../../../shared/browserExport";
import {
  pickQuestionnaireDiskImportMatch,
  type QuestionnairePresencePerson,
} from "../../personAttachments";
import { buildQuestionnaireExportFileName } from "../../personnelUtils";
import { searchQuestionnairesOnDiskInBatches } from "../../questionnaireDiskSearchBatch";

const pause = (delayMs: number) =>
  new Promise<void>((resolve) => globalThis.setTimeout(resolve, delayMs));

export type BulkImportDiskQuestionnairesResult = {
  saved: number;
  skipped: number;
  failed: number;
};

export const bulkImportDiskQuestionnaires = async (input: {
  people: QuestionnairePresencePerson[];
  dbPresence: Record<string, true>;
  inStaffIds: ReadonlySet<string>;
  signal?: AbortSignal;
  onProgress?: (done: number, total: number, saved: number) => void;
}): Promise<BulkImportDiskQuestionnairesResult> => {
  const targets = input.people.filter(
    (person) =>
      input.inStaffIds.has(person.currentId) && !input.dbPresence[person.currentId],
  );
  if (!targets.length) {
    return { saved: 0, skipped: 0, failed: 0 };
  }

  const result = await searchQuestionnairesOnDiskInBatches({
    people: targets.map((person) => ({
      externalId: person.currentId,
      fullName: person.fullName,
      callSign: person.callSign,
      missingQuestionnaire: true,
    })),
    search: (payload) => api.searchQuestionnairesOnDisk(payload),
    isCancelled: () => Boolean(input.signal?.aborted),
  });

  const peopleById = new Map(input.people.map((person) => [person.currentId, person]));
  let saved = 0;
  let skipped = 0;
  let failed = 0;
  let done = 0;

  for (const row of result.people) {
    if (input.signal?.aborted) break;
    done += 1;
    const person = peopleById.get(row.externalId);
    if (!person) {
      skipped += 1;
      input.onProgress?.(done, result.people.length, saved);
      continue;
    }

    const match = pickQuestionnaireDiskImportMatch(
      person.fullName,
      person.callSign,
      row.matches.map((item) => ({
        fileName: item.fileName,
        relativePath: item.relativePath,
      })),
    );
    if (!match) {
      skipped += 1;
      input.onProgress?.(done, result.people.length, saved);
      continue;
    }

    try {
      await api.confirmDiskQuestionnaire(
        person.currentId,
        match.relativePath,
        sanitizeFileName(
          buildQuestionnaireExportFileName(person.fullName, person.callSign),
        ),
        { suppressErrorToast: true, poolPriority: "normal" },
      );
      saved += 1;
    } catch {
      failed += 1;
    }

    input.onProgress?.(done, result.people.length, saved);
    await pause(40);
  }

  return { saved, skipped, failed };
};
