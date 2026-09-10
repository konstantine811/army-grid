import { api } from "../api";
import type { BackendPersonQuestionnaireMeta } from "../api";
import { loadAvailablePersonPhotoIds } from "../pages/personnel/personAttachments";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
} from "./idbDataCache";
import {
  loadPersonnelDataset,
  type PersonnelDataset,
} from "./personnelDataset";

export type PersonnelBootstrapOptions = {
  force?: boolean;
  signal?: AbortSignal;
  onCached?: (dataset: PersonnelDataset) => void | Promise<void>;
  /** Prefetch questionnaire index (shared between Overview and Personnel). */
  questionnaires?: boolean;
  /** Prefetch photo id index. */
  photoIndex?: boolean;
};

export type PersonnelBootstrapResult = {
  dataset: PersonnelDataset;
  questionnaires?: BackendPersonQuestionnaireMeta[];
  questionnairesPromise?: Promise<BackendPersonQuestionnaireMeta[]>;
};

let questionnairesInFlight: Promise<BackendPersonQuestionnaireMeta[]> | null =
  null;
let questionnairesInFlightForce = false;

/** Single in-flight fetch for `/personnel/questionnaires` meta list. */
export const loadSharedQuestionnairesMeta = async (options?: {
  force?: boolean;
  signal?: AbortSignal;
}): Promise<BackendPersonQuestionnaireMeta[]> => {
  if (
    questionnairesInFlight &&
    !options?.force &&
    !questionnairesInFlightForce
  ) {
    return questionnairesInFlight;
  }

  const promise = fetchWithCache({
    key: CacheKeys.questionnairesMeta,
    force: options?.force,
    signal: options?.signal,
    fetcher: () => api.listPersonQuestionnaires({ signal: options?.signal }),
    isChanged: jsonChanged,
  }).catch(() => [] as BackendPersonQuestionnaireMeta[]);

  questionnairesInFlight = promise;
  questionnairesInFlightForce = Boolean(options?.force);
  try {
    return await promise;
  } finally {
    if (questionnairesInFlight === promise) {
      questionnairesInFlight = null;
      questionnairesInFlightForce = false;
    }
  }
};

/** Coordinated bootstrap for pages that need personnel dataset + shared meta. */
export const bootstrapPersonnelAppData = async (
  options: PersonnelBootstrapOptions = {},
): Promise<PersonnelBootstrapResult> => {
  const dataset = await loadPersonnelDataset({
    force: options.force,
    signal: options.signal,
    onCached: options.onCached,
  });

  const questionnairesPromise = options.questionnaires
    ? loadSharedQuestionnairesMeta({
        force: options.force,
        signal: options.signal,
      })
    : undefined;

  if (options.photoIndex) {
    const prefetchPhotos = () => {
      void loadAvailablePersonPhotoIds({ force: options.force }).catch(
        () => new Set<string>(),
      );
    };
    if ("requestIdleCallback" in window) {
      window.requestIdleCallback(prefetchPhotos, { timeout: 8_000 });
    } else {
      globalThis.setTimeout(prefetchPhotos, 2_000);
    }
  }

  return { dataset, questionnairesPromise };
};
