import type { BackendPersonQuestionnaireMeta } from "../../api";
import type { BackendPersonnelOverviewRow } from "../../api";
import { loadSharedDocumentsAll } from "../../data/sharedAppData";
import { loadSharedQuestionnairesMeta } from "../../data/personnelBootstrap";
import type { PersonnelDataset } from "../../data/personnelDataset";
import {
  overviewAssetsCacheKey,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import { runHeavyJob } from "../../workers/runHeavyJob";
import {
  buildOverviewPersonnelIdentities,
  type OverviewPersonnelAssets,
} from "./overviewPersonnelAssets";
import { fetchOverviewAssetsSnapshot } from "./overviewServerSnapshots";

export const directQuestionnairePresence = (
  items: BackendPersonQuestionnaireMeta[] | null | undefined,
) =>
  Object.fromEntries(
    (Array.isArray(items) ? items : [])
      .map((item) => item.personExternalId?.trim())
      .filter(Boolean)
      .map((id) => [id, true] as const),
  ) as Record<string, true>;

export async function fetchOverviewPersonnelAssets(options: {
  overviewRows: BackendPersonnelOverviewRow[];
  dataset: PersonnelDataset;
  signal?: AbortSignal;
  force?: boolean;
}): Promise<OverviewPersonnelAssets> {
  const { overviewRows, dataset, signal, force = false } = options;
  const alive = () => !signal?.aborted;

  const serverAssets = await fetchOverviewAssetsSnapshot({
    fingerprint: dataset.fingerprint,
    signal,
  });
  if (serverAssets && alive()) {
    return {
      photos: {},
      questionnairePresence: serverAssets.questionnairePresence,
      questionnaireSourceIds: serverAssets.questionnaireSourceIds,
      documents: serverAssets.documents,
    };
  }

  const [questionnaireList, documentList] = await Promise.all([
    loadSharedQuestionnairesMeta({ force, signal }),
    loadSharedDocumentsAll({ force, signal }).catch(() => []),
  ]);
  if (!alive()) {
    return {
      questionnairePresence: directQuestionnairePresence(questionnaireList),
      questionnaireSourceIds: {},
      documents: {},
      photos: {},
    };
  }

  const cacheKey = overviewAssetsCacheKey(
    dataset.fingerprint,
    overviewRows,
    Array.isArray(questionnaireList) ? questionnaireList : [],
    Array.isArray(documentList) ? documentList : [],
  );

  if (!force) {
    const cached = await readDataCache<OverviewPersonnelAssets>(cacheKey);
    if (cached && alive()) return cached;
  }

  const assets = await runHeavyJob({
    type: "applyOverviewAssets",
    overviewRows,
    personnelIdentities: buildOverviewPersonnelIdentities(dataset.rows),
    questionnaires: Array.isArray(questionnaireList)
      ? questionnaireList.map(({ personExternalId, fileName }) => ({
          personExternalId,
          fileName,
        }))
      : [],
    documents: Array.isArray(documentList)
      ? documentList.map(({ id, personExternalId, type, title }) => ({
          id,
          personExternalId,
          type,
          title,
        }))
      : [],
  });

  if (alive()) void writeDataCache(cacheKey, assets);
  return assets;
}
