import {
  CacheKeys,
  overviewStaffCacheKey,
  readDataCache,
} from "./idbDataCache";
import type { PersonnelDataset } from "./personnelDatasetCore";

let warmStarted = false;

/** Hydrate L1 memory from IndexedDB so pages can peekDataCache instantly. */
export const warmPersonnelAppCaches = () => {
  if (warmStarted) return;
  warmStarted = true;

  void readDataCache<PersonnelDataset>(CacheKeys.personnelDataset).then(
    (dataset) => {
      if (!dataset?.fingerprint) return;
      void readDataCache(overviewStaffCacheKey(dataset.fingerprint));
    },
  );
  void readDataCache(CacheKeys.questionnairesMeta);
  void readDataCache(CacheKeys.overview);
};

/** Preload heavy route chunks after auth so Suspense is shorter on first navigation. */
export const preloadPersonnelAppChunks = () => {
  void import("../pages/overview/OverviewPage");
  void import("../pages/personnel/PersonnelPage");
};

export const bootstrapPersonnelAppShell = () => {
  warmPersonnelAppCaches();
  preloadPersonnelAppChunks();
};
