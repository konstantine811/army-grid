import { api } from "../../api";
import {
  questionnairePresenceCacheKey,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { runHeavyJob } from "../../workers/runHeavyJob";
import {
  buildQuestionnairePresencePeople,
  loadAvailablePersonPhotoIds,
  narrowQuestionnairePresenceForPeople,
  peekAvailablePersonPhotoIds,
  rememberQuestionnairePhotoAliases,
} from "./personAttachments";
import { loadPersonnelIdentityLinks } from "./personnelIdentity";
import { isLikelyPersonnelRow } from "./personnelUtils";

type QuestionnaireItem = {
  personExternalId: string;
  fileName?: string | null;
};

const peopleStamp = (
  people: ReturnType<typeof buildQuestionnairePresencePeople>,
) =>
  `${people.length}:${people[0]?.currentId ?? ""}:${people[people.length - 1]?.currentId ?? ""}`;

export const loadQuestionnairePresence = async ({
  rows,
  signal,
  prefetchedItems,
  force,
  datasetFingerprint,
  onPhotoAliasesReady,
  onPeople,
}: {
  rows: EjournalPreviewRow[];
  signal?: AbortSignal;
  prefetchedItems?: Promise<QuestionnaireItem[]>;
  force?: boolean;
  datasetFingerprint: string;
  onPhotoAliasesReady?: () => void;
  onPeople?: (
    people: ReturnType<typeof buildQuestionnairePresencePeople>,
  ) => void;
}) => {
  const items = (await prefetchedItems) ?? (await api.listPersonQuestionnaires({ signal }));
  const sourceRows = rows.filter(isLikelyPersonnelRow);
  if (!signal?.aborted) {
    await loadPersonnelIdentityLinks(sourceRows, items, signal).catch(
      () => undefined,
    );
  }
  const photoIds =
    peekAvailablePersonPhotoIds() ?? (await loadAvailablePersonPhotoIds());
  rememberQuestionnairePhotoAliases(items, photoIds);
  onPhotoAliasesReady?.();
  const people = buildQuestionnairePresencePeople(sourceRows);
  onPeople?.(people);
  const cacheKey = `${questionnairePresenceCacheKey(
    datasetFingerprint,
    items,
  )}:${peopleStamp(people)}:v5`;

  if (!force) {
    const cached = await readDataCache<Record<string, true>>(cacheKey);
    if (cached) {
      return {
        items,
        people,
        uiPresence: narrowQuestionnairePresenceForPeople(cached, people),
        aborted: false,
      };
    }
  }

  const presence = await runHeavyJob({
    type: "buildQuestionnairePresence",
    people,
    questionnaires: items.map(({ personExternalId, fileName }) => ({
      personExternalId,
      fileName,
    })),
  });
  if (signal?.aborted) {
    return { items, people, uiPresence: null, aborted: true };
  }
  const uiPresence = narrowQuestionnairePresenceForPeople(presence, people);
  void writeDataCache(cacheKey, uiPresence);
  return { items, people, uiPresence, aborted: false };
};
