import {
  startTransition,
  useCallback,
  useDeferredValue,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import writeXlsxFile, { type SheetData } from "write-excel-file/browser";
import { MagneticTapePreloader } from "@/components/sci/MagneticTapePreloader";
import {
  Alert,
  Box,
  Button,
  Chip,
  LinearProgress,
  MenuItem,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import { BusinessCenterOutlinedIcon } from "@/components/sci/icons";
import { LocalHospitalOutlinedIcon } from "@/components/sci/icons";
import { BeachAccessOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import {
  api,
  type BackendPersonQuestionnaireMeta,
  type BackendPersonnelOverview,
  type BackendPersonnelOverviewRow,
} from "../../api";
import {
  openPersonnelFromOverview,
} from "../../app/navigation";
import {
  loadPersonnelDataset,
  type PersonnelDataset,
} from "../../data/personnelDataset";
import {
  loadPersonnelBootstrapMeta,
  loadPersonnelVersionProbe,
} from "../../data/personnelVersion";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  overviewStaffCacheKey,
  peekDataCache,
  readDataCache,
  writeDataCache,
} from "../../data/idbDataCache";
import { STAFF_SHEET_SYNCED_EVENT } from "../../data/staffSheetAutoSync";
import {
  createStoredZipBlob,
  dataUrlToUint8Array,
  downloadBlob,
  sanitizeFileName,
} from "../../shared/browserExport";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  buildPersonSummary,
  buildQuestionnaireExportFileName,
} from "../personnel/personnelUtils";
import {
  fillMissingOverviewPhotos,
  resolveOverviewPhoto,
} from "./overviewPhotos";
import {
  directQuestionnairePresence,
  fetchOverviewPersonnelAssets,
} from "./overviewAssetsLoad";
import { fetchMergedOverviewSnapshot, fetchOverviewStaffSnapshot } from "./overviewServerSnapshots";
import {
  canSkipOverviewDatasetReload,
  canUseOverviewWarmCacheOnly,
  persistOverviewStaffCache,
  readBootstrapStaffPayload,
} from "./overviewWarmLoad";
import { OVERVIEW_DEFERRED_ASSET_COLUMN_IDS } from "./overviewStaffSheetColumns";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import {
  overviewNameMatchesQuery,
  parseOverviewNameQueries,
  buildOverviewRowSearchText,
} from "./overviewNameSearch";
import { overviewMergeCacheKey, overviewMergeFingerprint } from "./overviewMergeCache";
import {
  OverviewVirtualTable,
  type OverviewPersonDocumentSummary,
  type OverviewQuestionnaireTarget,
} from "./OverviewVirtualTable";
import { loadPersonnelOverviewInBatches } from "./overviewBatchLoad";
import {
  buildOverviewMetrics,
  buildPersonnelStaffOverview,
  buildRosterOnlyOverview,
  buildStaffOverviewRowsFromPersonnel,
  buildStaffOverviewRowsFromRoster,
  fillDownRosterUnitRows,
  summarizeStaffFromRoster,
  summarizeNovaStaffForUnits,
} from "./overviewRosterMerge";
import {
  buildImportantOverviewExportFileName,
  buildImportantOverviewExportSheets,
  buildOverviewExportSheetData,
  buildOverviewExportSheetOptions,
} from "./overviewExport";
import { runHeavyJob } from "../../workers/runHeavyJob";
import { buildOverviewSideStats } from "./overviewSideStats";
import { pullStaffSheetRosterImportPayload } from "../excel-fill/staffSheet";
import { loadPbWorkbookFromDb } from "../ejournal/loadEjournalWorkbooksFromDb";
import { parsePbArchive } from "../ejournal/ejoosParsers";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import {
  buildOverviewRotaCopyText,
  buildOverviewWhatsAppCopyText,
} from "./overviewCopyText";
import { exportOverviewPpdLocationReport } from "./overviewPpdLocationExport";
import { exportOverviewRotaBchsMorningReport } from "./overviewRotaBchsMorningExport";
import { exportOverviewRotaGudzReport } from "./overviewRotaGudzExport";

const SOURCE_FILTERS = [
  { value: "staff", label: "Штатка" },
  { value: "ejoos", label: "ЕЖООС" },
  { value: "all", label: "Усі джерела" },
] as const;

type OverviewSourceFilter = (typeof SOURCE_FILTERS)[number]["value"];

type OverviewLoadMode = "full" | "dataset-only" | "ejoos-only";

const scheduleIdleTask = (task: () => void, timeout = 5000) => {
  if (typeof requestIdleCallback === "function") {
    return requestIdleCallback(task, { timeout });
  }
  return window.setTimeout(task, Math.min(timeout, 3000));
};

const normalizeRosterText = normalizeRosterMatchText;

const withStaffOverviewStatus = (
  row: BackendPersonnelOverviewRow,
): BackendPersonnelOverviewRow => {
  if (!row.staffStatus) return row;
  return {
    ...row,
    status: row.staffStatus,
    statusLabel: row.staffStatusLabel || row.statusLabel,
  };
};

const formatSourceTimestamp = (iso: string | null | undefined) =>
  iso ? new Date(iso).toLocaleString("uk-UA") : null;

const formatStaffSourceLabel = (updatedAt: string | null) => {
  const stamp = formatSourceTimestamp(updatedAt);
  return stamp ? `Штатка · ${stamp}` : "Штатка";
};

const staffOverviewDatasetMeta = (dataset: PersonnelDataset) => ({
  importId: dataset.version.rosterImportId,
  importName: "Особовий склад · спільний dataset",
});

const resolveStaffOverviewForDataset = async (
  dataset: PersonnelDataset,
): Promise<BackendPersonnelOverview> => {
  const staffKey = overviewStaffCacheKey(dataset.fingerprint);
  const peeked = peekDataCache<BackendPersonnelOverview>(staffKey);
  if (peeked?.rows?.length) return peeked;
  const cached = await readDataCache<BackendPersonnelOverview>(staffKey);
  if (cached?.rows?.length) return cached;
  return buildPersonnelStaffOverview(
    dataset.rows,
    dataset.rosterLabels,
    staffOverviewDatasetMeta(dataset),
  );
};

const buildCallSignByExternalId = (rosterRows: EjournalPreviewRow[]) => {
  const map: Record<string, string> = {};
  for (const row of rosterRows) {
    const summary = buildPersonSummary(row);
    if (summary.externalId && summary.callSign) {
      map[summary.externalId] = summary.callSign;
    }
  }
  return map;
};

export function OverviewPage({ active = true }: { active?: boolean }) {
  const cachedQuestionnaires = peekDataCache<BackendPersonQuestionnaireMeta[]>(
    CacheKeys.questionnairesMeta,
  );
  const [data, setData] = useState<BackendPersonnelOverview | null>(null);
  const [photos, setPhotos] = useState<Record<string, string>>({});
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >(() => directQuestionnairePresence(cachedQuestionnaires));
  const [questionnairePresenceStatus, setQuestionnairePresenceStatus] =
    useState<"idle" | "loading" | "ready">("idle");
  const [
    questionnaireSourceIdByExternalId,
    setQuestionnaireSourceIdByExternalId,
  ] = useState<Record<string, string>>({});
  const [documentsByExternalId, setDocumentsByExternalId] = useState<
    Record<string, OverviewPersonDocumentSummary>
  >({});
  const [callSignByExternalId, setCallSignByExternalId] = useState<
    Record<string, string>
  >({});
  const [query, setQuery] = useState("");
  const [source, setSource] = useState<OverviewSourceFilter>("staff");
  const sourceRef = useRef<OverviewSourceFilter>("staff");
  sourceRef.current = source;
  const [rosterUpdatedAt, setRosterUpdatedAt] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const ejoosLoadedRef = useRef(false);
  const datasetFingerprintRef = useRef("");
  const personnelOverviewRowsCacheRef = useRef<{
    key: string;
    rows: BackendPersonnelOverviewRow[];
  }>({ key: "", rows: [] });
  const serverStaffRowsRef = useRef<BackendPersonnelOverviewRow[]>([]);
  const [staffOverviewEpoch, setStaffOverviewEpoch] = useState(0);
  const rosterRowsRef = useRef<EjournalPreviewRow[]>([]);
  const personnelRowsRef = useRef<EjournalPreviewRow[]>([]);
  const rosterColumnsRef = useRef<
    Array<{ key: string; letter?: string; originalIndex?: number }>
  >([]);
  const rosterLabelsRef = useRef<Record<string, string>>({});
  const photosRef = useRef<Record<string, string>>({});
  const requestedPhotoKeysRef = useRef(new Set<string>());
  const loadControllerRef = useRef<AbortController | null>(null);
  const loadSeqRef = useRef(0);
  const staffSyncReloadTimerRef = useRef(0);
  const ejoosControllerRef = useRef<AbortController | null>(null);
  const assetsStartedForFingerprintRef = useRef("");
  const staffAssetsContextRef = useRef<{
    rows: BackendPersonnelOverviewRow[];
    dataset: PersonnelDataset;
  } | null>(null);
  const assetsControllerRef = useRef<AbortController | null>(null);
  const assetsLoadSeqRef = useRef(0);
  const [rosterEpoch, setRosterEpoch] = useState(0);
  const [personnelEpoch, setPersonnelEpoch] = useState(0);
  photosRef.current = photos;

  const resolveLoadMode = (): OverviewLoadMode =>
    sourceRef.current === "staff" ? "dataset-only" : "full";

  const questionnairePresenceStatusRef = useRef(questionnairePresenceStatus);
  questionnairePresenceStatusRef.current = questionnairePresenceStatus;

  const runOverviewAssetsLoad = useCallback(
    async (
      overviewRows: BackendPersonnelOverviewRow[],
      dataset: PersonnelDataset,
      options?: { force?: boolean },
    ) => {
      assetsControllerRef.current?.abort();
      const controller = new AbortController();
      assetsControllerRef.current = controller;
      const seq = ++assetsLoadSeqRef.current;
      const alive = () =>
        seq === assetsLoadSeqRef.current && !controller.signal.aborted;

      assetsStartedForFingerprintRef.current = dataset.fingerprint;
      setQuestionnairePresenceStatus("loading");
      try {
        const assets = await fetchOverviewPersonnelAssets({
          overviewRows,
          dataset,
          signal: controller.signal,
          force: options?.force,
        });
        if (!alive()) return;
        startTransition(() => {
          setQuestionnaireByExternalId((current) => ({
            ...current,
            ...assets.questionnairePresence,
          }));
          setQuestionnaireSourceIdByExternalId(assets.questionnaireSourceIds);
          setDocumentsByExternalId(assets.documents);
          setQuestionnairePresenceStatus("ready");
        });
      } catch (error) {
        console.warn("[Огляд] Не вдалося підставити дані з Особового складу", error);
        if (alive()) setQuestionnairePresenceStatus("ready");
      }
    },
    [],
  );

  const ensureOverviewAssets = useCallback(
    (options?: { force?: boolean }) => {
      const ctx = staffAssetsContextRef.current;
      if (!ctx) return;
      if (
        !options?.force &&
        assetsStartedForFingerprintRef.current === ctx.dataset.fingerprint &&
        questionnairePresenceStatusRef.current === "ready"
      ) {
        return;
      }
      void runOverviewAssetsLoad(ctx.rows, ctx.dataset, options);
    },
    [runOverviewAssetsLoad],
  );

  const handleColumnVisibilityChange = useCallback(
    (visibility: Record<string, boolean>) => {
      if (sourceRef.current !== "staff") return;
      const needsAssets = OVERVIEW_DEFERRED_ASSET_COLUMN_IDS.some(
        (columnId) => visibility[columnId] !== false,
      );
      if (needsAssets) ensureOverviewAssets();
    },
    [ensureOverviewAssets],
  );

  const load = async (
    signal?: AbortSignal,
    options?: {
      force?: boolean;
      refreshStaffServer?: boolean;
      mode?: OverviewLoadMode;
      seq?: number;
    },
  ) => {
    const seq = options?.seq ?? loadSeqRef.current;
    const alive = () =>
      !signal?.aborted && seq === loadSeqRef.current;
    const force = Boolean(options?.force);
    const refreshStaffServer = Boolean(options?.refreshStaffServer);
    const mode = options?.mode ?? resolveLoadMode();
    const loadDataset = mode !== "ejoos-only";
    const loadEjoos = mode !== "dataset-only";
    if (force) {
      assetsStartedForFingerprintRef.current = "";
      serverStaffRowsRef.current = [];
      personnelOverviewRowsCacheRef.current = { key: "", rows: [] };
      staffAssetsContextRef.current = null;
      assetsControllerRef.current?.abort();
      if (sourceRef.current === "staff") {
        setQuestionnairePresenceStatus("idle");
      }
      setStaffOverviewEpoch((value) => value + 1);
    }
    const hasWarmStaffRows =
      sourceRef.current === "staff" && serverStaffRowsRef.current.length > 0;
    if (!hasWarmStaffRows) setIsLoading(true);
    try {
      let rosterFingerprint = datasetFingerprintRef.current;
      const applyOverview = async (
        overview: BackendPersonnelOverview,
        rosterRows: EjournalPreviewRow[],
        rosterLabels: Record<string, string> = {},
        rosterColumns: Array<{
          key: string;
          letter?: string;
          originalIndex?: number;
        }> = [],
        fromCache = false,
        skipRosterMerge = false,
      ) => {
        let mergedOverview = overview;
        if (!skipRosterMerge) {
          const mergeCacheKey = overviewMergeCacheKey(
            overview,
            rosterFingerprint,
            rosterRows,
            rosterColumns,
          );
          const mergeFingerprint = overviewMergeFingerprint(
            overview,
            rosterFingerprint,
            rosterRows,
          );
          const serverMerged = await fetchMergedOverviewSnapshot({
            mergeFingerprint,
            signal,
          });
          if (serverMerged?.rows?.length) {
            mergedOverview = serverMerged;
            void writeDataCache(mergeCacheKey, serverMerged);
          }
          if (!force && mergedOverview === overview) {
            const cachedMerge =
              await readDataCache<BackendPersonnelOverview>(mergeCacheKey);
            if (cachedMerge?.rows?.length) {
              mergedOverview = cachedMerge;
            }
          }
          if (mergedOverview === overview) {
            try {
              mergedOverview = await runHeavyJob({
                type: "mergeOverview",
                overview,
                rosterRows,
                rosterLabels,
                columns: rosterColumns,
              });
              void writeDataCache(mergeCacheKey, mergedOverview);
            } catch {
              mergedOverview =
                source === "staff" && rosterRows.length
                  ? buildRosterOnlyOverview(
                      rosterRows,
                      rosterLabels,
                      rosterColumns,
                    )
                  : overview;
            }
          }
        }
        if (!alive()) return mergedOverview;
        rosterRowsRef.current = rosterRows;
        rosterLabelsRef.current = rosterLabels;
        rosterColumnsRef.current = rosterColumns;
        setRosterEpoch((value) => value + 1);
        setData(mergedOverview);
        try {
          setCallSignByExternalId(buildCallSignByExternalId(rosterRows));
        } catch {
          setCallSignByExternalId({});
        }
        setMessage(
          fromCache
            ? `Кеш огляду · оновлюю з БД…`
            : mergedOverview.importName
              ? source === "staff"
                ? `Штатка: є · ${buildStaffOverviewRowsFromRoster(rosterRows, rosterLabels, rosterColumns).length} осіб`
                : `ЕЖООС: ${mergedOverview.importName} · ${mergedOverview.rows.length} записів`
              : rosterRows.length
                ? `Штатка: ${rosterRows.length} осіб · імпорту ЕЖООС немає`
                : "Немає імпорту ЕЖООС і Штатки",
        );
        return mergedOverview;
      };

      const loadServerStaffOverview = async (
        dataset: PersonnelDataset,
        bootstrap?: import("../../api").BackendPersonnelBootstrap | null,
      ) => {
        const staff = await fetchOverviewStaffSnapshot({
          fingerprint: dataset.fingerprint,
          signal,
          force: true,
          bootstrap,
          allowRebuild: true,
        });
        if (!alive() || !staff?.rows?.length) return;
        serverStaffRowsRef.current = staff.rows.map(withStaffOverviewStatus);
        personnelOverviewRowsCacheRef.current = {
          key: `${dataset.fingerprint}:${rosterEpoch}`,
          rows: serverStaffRowsRef.current,
        };
        void persistOverviewStaffCache(dataset.fingerprint, staff, bootstrap);
        setStaffOverviewEpoch((value) => value + 1);
      };

      const bindDatasetStaffContext = (dataset: PersonnelDataset) => {
        const staffRows = resolveStaffOverviewRows(dataset);
        staffAssetsContextRef.current = { rows: staffRows, dataset };
        if (sourceRef.current !== "staff") {
          void runOverviewAssetsLoad(staffRows, dataset, { force });
        }
      };

      const ensureLocalStaffSnapshot = async (dataset: PersonnelDataset) => {
        if (serverStaffRowsRef.current.length) return;
        const staffOverview = await resolveStaffOverviewForDataset(dataset);
        if (!alive()) return;
        rememberStaffSnapshotRows(staffOverview, dataset.fingerprint);
      };

  const rememberStaffSnapshotRows = (
    staff: BackendPersonnelOverview,
    fingerprint: string,
  ) => {
    serverStaffRowsRef.current = staff.rows.map(withStaffOverviewStatus);
    personnelOverviewRowsCacheRef.current = {
      key: `${fingerprint}:${rosterEpoch}`,
      rows: serverStaffRowsRef.current,
    };
    setStaffOverviewEpoch((value) => value + 1);
    void persistOverviewStaffCache(fingerprint, staff, null);
  };

      const resolveStaffOverviewRows = (dataset: PersonnelDataset) => {
        const cacheKey = `${dataset.fingerprint}:${rosterEpoch}`;
        const cachedStaffRows =
          serverStaffRowsRef.current.length &&
          personnelOverviewRowsCacheRef.current.key === cacheKey
            ? serverStaffRowsRef.current
            : personnelOverviewRowsCacheRef.current.key === cacheKey
              ? personnelOverviewRowsCacheRef.current.rows
              : null;
        return (
          cachedStaffRows ??
          buildStaffOverviewRowsFromPersonnel(
            dataset.rows,
            dataset.rosterLabels,
          )
        );
      };

      const startDatasetStaff = async (
        dataset: PersonnelDataset,
        bootstrap?: import("../../api").BackendPersonnelBootstrap | null,
      ) => {
        await ensureLocalStaffSnapshot(dataset);
        if (!alive()) return;
        bindDatasetStaffContext(dataset);
        if (refreshStaffServer) {
          await loadServerStaffOverview(dataset, bootstrap);
        }
      };

      const rememberDatasetFingerprint = (fingerprint: string) => {
        if (
          datasetFingerprintRef.current &&
          datasetFingerprintRef.current !== fingerprint
        ) {
          ejoosLoadedRef.current = false;
          serverStaffRowsRef.current = [];
        }
        datasetFingerprintRef.current = fingerprint;
      };

      rosterFingerprint = datasetFingerprintRef.current;
      let rosterRows: EjournalPreviewRow[] = rosterRowsRef.current;
      let rosterLabels: Record<string, string> = rosterLabelsRef.current;
      let rosterColumns: Array<{
        key: string;
        letter?: string;
        originalIndex?: number;
      }> = rosterColumnsRef.current;
      let paintedFromCache = false;
      let datasetPromise: Promise<PersonnelDataset | null | undefined>;

      if (loadDataset) {
        const datasetOnly = !loadEjoos;
        const cachedDataset = await readDataCache<PersonnelDataset>(
          CacheKeys.personnelDataset,
        );
        if (loadEjoos) {
          void readDataCache<BackendPersonnelOverview>(CacheKeys.overview);
        }
        rosterFingerprint = cachedDataset?.fingerprint ?? "";
        if (rosterFingerprint) rememberDatasetFingerprint(rosterFingerprint);
        rosterRows = cachedDataset?.rosterRows ?? [];
        rosterLabels = cachedDataset?.rosterLabels ?? {};
        rosterColumns = cachedDataset?.rosterColumns ?? [];

        if (cachedDataset?.rows?.length) {
          personnelRowsRef.current = cachedDataset.rows;
          setPersonnelEpoch((value) => value + 1);
          setRosterUpdatedAt(cachedDataset.rosterUpdatedAt);
          const staffOverview =
            await resolveStaffOverviewForDataset(cachedDataset);
          if (!alive()) return;
          rememberStaffSnapshotRows(staffOverview, cachedDataset.fingerprint);
          staffAssetsContextRef.current = {
            rows: serverStaffRowsRef.current,
            dataset: cachedDataset,
          };
          await applyOverview(
            staffOverview,
            rosterRows,
            rosterLabels,
            rosterColumns,
            true,
            true,
          );
          paintedFromCache = true;
          setIsLoading(false);
        }

        if (
          canUseOverviewWarmCacheOnly({ force, datasetOnly, cachedDataset }) &&
          cachedDataset
        ) {
          scheduleIdleTask(() => {
            void (async () => {
              const probe = await loadPersonnelVersionProbe().catch(() => null);
              if (
                !probe?.snapshot?.matchesLive ||
                probe.fingerprint !== cachedDataset.fingerprint
              ) {
                startLoad(true);
              }
            })();
          }, 6_000);
          return;
        }

        const [bootstrap, versionProbe] = await Promise.all([
          loadPersonnelBootstrapMeta({ force, signal }),
          loadPersonnelVersionProbe({ force, signal }),
        ]);
        if (!alive()) return;

        const skipDatasetReload = canSkipOverviewDatasetReload({
          force,
          datasetOnly,
          cachedDataset,
          versionProbe,
        });

        if (skipDatasetReload && cachedDataset) {
          const bootstrapStaff = readBootstrapStaffPayload(
            bootstrap,
            cachedDataset.fingerprint,
          );
          if (bootstrapStaff) {
            rememberStaffSnapshotRows(
              bootstrapStaff,
              cachedDataset.fingerprint,
            );
            void persistOverviewStaffCache(
              cachedDataset.fingerprint,
              bootstrapStaff,
              bootstrap,
            );
          }
          await startDatasetStaff(cachedDataset, bootstrap);
          if (datasetOnly) return;
          datasetPromise = Promise.resolve(cachedDataset);
        } else {
          datasetPromise = loadPersonnelDataset({
            force,
            signal,
            versionProbe,
            bootstrapMeta: bootstrap,
            onCached: async (nextDataset) => {
              if (!alive()) return;
              rosterFingerprint = nextDataset.fingerprint;
              rememberDatasetFingerprint(nextDataset.fingerprint);
              rosterRows = nextDataset.rosterRows;
              rosterLabels = nextDataset.rosterLabels;
              rosterColumns = nextDataset.rosterColumns;
              personnelRowsRef.current = nextDataset.rows;
              setPersonnelEpoch((value) => value + 1);
              setRosterUpdatedAt(nextDataset.rosterUpdatedAt);
              const staffOverview =
                await resolveStaffOverviewForDataset(nextDataset);
              if (!alive()) return;
              rememberStaffSnapshotRows(staffOverview, nextDataset.fingerprint);
              await applyOverview(
                staffOverview,
                rosterRows,
                rosterLabels,
                rosterColumns,
                true,
                true,
              );
              paintedFromCache = true;
              setIsLoading(false);
            },
          })
            .then(async (nextDataset) => {
              if (!alive()) return nextDataset;
              rosterFingerprint = nextDataset.fingerprint;
              rememberDatasetFingerprint(nextDataset.fingerprint);
              rosterRows = nextDataset.rosterRows;
              rosterLabels = nextDataset.rosterLabels;
              rosterColumns = nextDataset.rosterColumns;
              personnelRowsRef.current = nextDataset.rows;
              setPersonnelEpoch((value) => value + 1);
              setRosterUpdatedAt(nextDataset.rosterUpdatedAt);
              await startDatasetStaff(nextDataset, bootstrap);
              return nextDataset;
            })
            .catch(() => cachedDataset ?? null);
        }

        if (!loadEjoos) {
          const dataset = await datasetPromise;
          if (!alive()) return;
          if (!dataset?.rows.length || !rosterRows.length) {
            setMessage(
              "Не вдалося завантажити Штатку. Огляд ООС не показано, щоб не відображати неправильні 2000+ записів.",
            );
          }
          return;
        }
      } else {
        rosterFingerprint = datasetFingerprintRef.current;
        rosterRows = rosterRowsRef.current;
        rosterLabels = rosterLabelsRef.current;
        rosterColumns = rosterColumnsRef.current;
        datasetPromise = readDataCache<PersonnelDataset>(
          CacheKeys.personnelDataset,
        );
        if (mode === "ejoos-only" && !personnelRowsRef.current.length) {
          const cachedDataset = await datasetPromise;
          if (cachedDataset?.rows.length) {
            rememberDatasetFingerprint(cachedDataset.fingerprint);
            personnelRowsRef.current = cachedDataset.rows;
            rosterRows = cachedDataset.rosterRows;
            rosterLabels = cachedDataset.rosterLabels;
            rosterColumns = cachedDataset.rosterColumns;
            rosterFingerprint = cachedDataset.fingerprint;
          }
        }
      }

      const overview = await fetchWithCache({
        key: CacheKeys.overview,
        force,
        signal,
        fetcher: () =>
          loadPersonnelOverviewInBatches({
            force,
            signal,
            onPage: paintedFromCache
              ? undefined
              : async (partial, meta) => {
                  await datasetPromise;
                  if (!alive() || !rosterRows.length) return;
                  await applyOverview(
                    partial,
                    rosterRows,
                    rosterLabels,
                    rosterColumns,
                    false,
                    true,
                  );
                  if (meta.complete) return;
                  setMessage(
                    partial.importName
                      ? `Джерело: ${partial.importName} · ${meta.done} з ${meta.total}. Довантажую…`
                      : `Огляд: ${meta.done} з ${meta.total}. Довантажую…`,
                  );
                },
          }),
        isChanged: jsonChanged,
      });
      const dataset = await datasetPromise;
      if (!alive()) return;
      if (!dataset?.rows.length || !rosterRows.length) {
        setMessage(
          "Не вдалося завантажити Штатку. Огляд ООС не показано, щоб не відображати неправильні 2000+ записів.",
        );
        return;
      }
      const mergedOverview = await applyOverview(
        overview,
        rosterRows,
        rosterLabels,
        rosterColumns,
      );
      ejoosLoadedRef.current = true;

      staffAssetsContextRef.current = { rows: mergedOverview.rows, dataset };
      if (sourceRef.current !== "staff") {
        void runOverviewAssetsLoad(mergedOverview.rows, dataset, { force });
      }
    } catch (error) {
      if (!alive()) return;
      setMessage(
        error instanceof Error ? error.message : "Не вдалося завантажити огляд",
      );
    } finally {
      if (alive()) setIsLoading(false);
    }
  };

  const startLoad = (
    force = false,
    options?: { refreshStaffServer?: boolean },
  ) => {
    const seq = ++loadSeqRef.current;
    loadControllerRef.current?.abort();
    const controller = new AbortController();
    loadControllerRef.current = controller;
    void load(controller.signal, {
      force,
      refreshStaffServer: options?.refreshStaffServer,
      mode: resolveLoadMode(),
      seq,
    });
  };

  const startEjoosLoad = () => {
    ejoosControllerRef.current?.abort();
    const controller = new AbortController();
    ejoosControllerRef.current = controller;
    void load(controller.signal, { mode: "ejoos-only" });
  };

  useLayoutEffect(() => {
    if (!active) return;
    const dataset = peekDataCache<PersonnelDataset>(CacheKeys.personnelDataset);
    if (!dataset?.rows?.length) return;

    datasetFingerprintRef.current = dataset.fingerprint;
    personnelRowsRef.current = dataset.rows;
    rosterRowsRef.current = dataset.rosterRows ?? [];
    rosterLabelsRef.current = dataset.rosterLabels ?? {};
    rosterColumnsRef.current = dataset.rosterColumns ?? [];
    setRosterUpdatedAt(dataset.rosterUpdatedAt);
    setPersonnelEpoch((value) => value + 1);
    setRosterEpoch((value) => value + 1);

    const staff = peekDataCache<BackendPersonnelOverview>(
      overviewStaffCacheKey(dataset.fingerprint),
    );
    if (!staff?.rows?.length) return;

    serverStaffRowsRef.current = staff.rows.map(withStaffOverviewStatus);
    personnelOverviewRowsCacheRef.current = {
      key: `${dataset.fingerprint}:0`,
      rows: serverStaffRowsRef.current,
    };
    setStaffOverviewEpoch((value) => value + 1);
    setData(staff);
    setMessage("Кеш огляду · оновлюю з БД…");
    staffAssetsContextRef.current = {
      rows: serverStaffRowsRef.current,
      dataset,
    };
  }, [active]);

  useEffect(() => {
    if (!active) {
      loadSeqRef.current += 1;
      loadControllerRef.current?.abort();
      ejoosControllerRef.current?.abort();
      return;
    }
    startLoad();
    return () => {
      loadSeqRef.current += 1;
    };
  }, [active]);

  useEffect(() => {
    if (!active || source === "staff" || ejoosLoadedRef.current) return;
    startEjoosLoad();
  }, [active, source]);

  useEffect(() => {
    if (!active || source === "staff") return;
    ensureOverviewAssets();
  }, [active, source, ensureOverviewAssets]);

  useEffect(() => {
    if (!active) return;
    const onSynced = () => {
      window.clearTimeout(staffSyncReloadTimerRef.current);
      staffSyncReloadTimerRef.current = window.setTimeout(() => {
        startLoad(true);
      }, 800);
    };
    window.addEventListener(STAFF_SHEET_SYNCED_EVENT, onSynced);
    return () => {
      window.removeEventListener(STAFF_SHEET_SYNCED_EVENT, onSynced);
      window.clearTimeout(staffSyncReloadTimerRef.current);
    };
  }, [active]);

  const deferredQuery = useDeferredValue(query);
  const nameQueries = useMemo(
    () => parseOverviewNameQueries(deferredQuery),
    [deferredQuery],
  );
  const isNameListSearch = nameQueries.length > 1;

  const personnelOverviewRows = useMemo(() => {
    if (serverStaffRowsRef.current.length) {
      return serverStaffRowsRef.current;
    }
    if (!personnelRowsRef.current.length) {
      return [] as BackendPersonnelOverviewRow[];
    }
    const cacheKey = `${datasetFingerprintRef.current}:${rosterEpoch}`;
    const cached = personnelOverviewRowsCacheRef.current;
    if (cached.key === cacheKey) return cached.rows;
    const rows = buildStaffOverviewRowsFromPersonnel(
      personnelRowsRef.current,
      rosterLabelsRef.current,
    ).map(withStaffOverviewStatus);
    personnelOverviewRowsCacheRef.current = { key: cacheKey, rows };
    return rows;
  }, [personnelEpoch, rosterEpoch, staffOverviewEpoch]);

  const sourceRows = useMemo(() => {
    if (!data) return [] as BackendPersonnelOverviewRow[];
    if (source === "staff") return personnelOverviewRows;
    return source === "ejoos"
      ? data.rows.filter((row) => row.fromEjoos)
      : data.rows;
  }, [data, personnelOverviewRows, source]);

  const staffSummary = useMemo(
    () =>
      summarizeStaffFromRoster(
        rosterRowsRef.current,
        rosterColumnsRef.current,
      ),
    [rosterEpoch],
  );

  /** Pasted FIO list looks through every loaded person, not only the current source. */
  const nameSearchRows = useMemo(() => {
    if (!isNameListSearch || source === "staff") return sourceRows;
    if (!data) return sourceRows;
    return data.rows.map((row) =>
      row.inNovaStaff ? withStaffOverviewStatus(row) : row,
    );
  }, [data, isNameListSearch, source, sourceRows]);

  const rowSearchTextCacheRef = useRef(new Map<string, string>());
  const documentsByExternalIdRef = useRef(documentsByExternalId);
  documentsByExternalIdRef.current = documentsByExternalId;

  useEffect(() => {
    rowSearchTextCacheRef.current.clear();
  }, [documentsByExternalId, nameSearchRows]);

  const getRowSearchText = useCallback((row: BackendPersonnelOverviewRow) => {
    const cached = rowSearchTextCacheRef.current.get(row.id);
    if (cached !== undefined) return cached;
    const documentLabels =
      row.externalId && documentsByExternalIdRef.current[row.externalId]
        ? documentsByExternalIdRef.current[row.externalId].labels.join(" ")
        : "";
    const text = buildOverviewRowSearchText(row, documentLabels);
    rowSearchTextCacheRef.current.set(row.id, text);
    return text;
  }, []);

  const filteredRows = useMemo(() => {
    return nameSearchRows.filter((row) => {
      if (!nameQueries.length) return true;

      if (isNameListSearch) {
        return nameQueries.some((nameQuery) =>
          overviewNameMatchesQuery(row.name, nameQuery),
        );
      }

      const normalizedQuery = normalizeRosterText(nameQueries[0] ?? "");
      if (!normalizedQuery) return true;
      if (overviewNameMatchesQuery(row.name, nameQueries[0] ?? "")) return true;

      return getRowSearchText(row).includes(normalizedQuery);
    });
  }, [
    documentsByExternalId,
    getRowSearchText,
    isNameListSearch,
    nameQueries,
    nameSearchRows,
  ]);

  const metrics = useMemo(() => {
    const metricRows =
      source === "staff"
        ? filteredRows.filter((row) => row.inStaff)
        : filteredRows;
    return buildOverviewMetrics(metricRows);
  }, [filteredRows, source]);

  const sideStats = useMemo(
    () => buildOverviewSideStats(filteredRows, data),
    [data, filteredRows],
  );

  const nameListMatchStats = useMemo(() => {
    if (!isNameListSearch || !data) return null;
    const matched: string[] = [];
    const missing: string[] = [];
    for (const nameQuery of nameQueries) {
      const hit = nameSearchRows.some((row) =>
        overviewNameMatchesQuery(row.name, nameQuery),
      );
      if (hit) matched.push(nameQuery);
      else missing.push(nameQuery);
    }
    return { matched, missing, total: nameQueries.length };
  }, [isNameListSearch, nameQueries, nameSearchRows]);

  const openQuestionnaire = useCallback(async (target: OverviewQuestionnaireTarget) => {
    if (!target.externalId) return;
    if (!target.hasQuestionnaire) {
      openPersonnelFromOverview({ externalId: target.externalId });
      setMessage(`Відкрито картку ${target.name} — можна додати анкету.`);
      return;
    }

    try {
      setMessage(`Відкриваю анкету: ${target.name}…`);
      const url = await api.createPersonQuestionnairePreviewUrl(
        questionnaireSourceIdByExternalId[target.externalId] ??
          target.externalId,
        buildQuestionnaireExportFileName(target.name),
      );
      window.open(url, "_blank", "noopener,noreferrer");
      setMessage(`Анкета відкрита: ${target.name}`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося відкрити анкету: ${error.message}`
          : "Не вдалося відкрити анкету.",
      );
    }
  }, [questionnaireSourceIdByExternalId]);

  const onNeedPhoto = useCallback((row: BackendPersonnelOverviewRow) => {
    if (resolveOverviewPhoto(row, photosRef.current)) return;
    const requestKey = row.externalId || row.id || row.name;
    if (!requestKey || requestedPhotoKeysRef.current.has(requestKey)) return;
    requestedPhotoKeysRef.current.add(requestKey);
    void fillMissingOverviewPhotos(
      [row],
      personnelRowsRef.current.length
        ? personnelRowsRef.current
        : rosterRowsRef.current,
      photosRef.current,
    ).then((next) => {
      setPhotos((current) => {
        const merged = { ...current, ...next };
        photosRef.current = merged;
        return merged;
      });
      if (
        !resolveOverviewPhoto(row, photosRef.current) &&
        !personnelRowsRef.current.length &&
        !rosterRowsRef.current.length
      ) {
        requestedPhotoKeysRef.current.delete(requestKey);
      }
    });
  }, []);

  const prefetchOverviewPhotos = useCallback((rows: BackendPersonnelOverviewRow[]) => {
    if (!rows.length) return;
    const missing = rows.filter((row) => {
      const requestKey = row.externalId || row.id || row.name;
      if (!requestKey || requestedPhotoKeysRef.current.has(requestKey)) {
        return false;
      }
      return !resolveOverviewPhoto(row, photosRef.current);
    });
    if (!missing.length) return;
    for (const row of missing) {
      requestedPhotoKeysRef.current.add(row.externalId || row.id || row.name);
    }
    void fillMissingOverviewPhotos(
      missing,
      personnelRowsRef.current.length
        ? personnelRowsRef.current
        : rosterRowsRef.current,
      photosRef.current,
    ).then((next) => {
      setPhotos((current) => {
        const merged = { ...current, ...next };
        photosRef.current = merged;
        return merged;
      });
    });
  }, []);

  const onVisibleRowsChange = useCallback(
    (visibleRows: BackendPersonnelOverviewRow[]) => {
      prefetchOverviewPhotos(visibleRows);
    },
    [prefetchOverviewPhotos],
  );

  const exportOverviewTable = async (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => {
    const exportedAt = new Date().toISOString().slice(0, 10);
    const visibleColumns = context.columns;
    const hasQuestionnaireColumn = visibleColumns.some(
      (column) => column.id === "questionnaire",
    );

    const sheetData: SheetData = buildOverviewExportSheetData(context, {
      activeFilters: context.filters,
    });

    await writeXlsxFile(
      sheetData,
      buildOverviewExportSheetOptions(context),
      { fontFamily: "Arial", fontSize: 10 },
    ).toFile(
      `Огляд особового складу ${exportedAt}.xlsx`,
    );

    if (!hasQuestionnaireColumn) {
      setMessage(`Експортовано таблицю: ${context.rows.length} рядків.`);
      return;
    }

    ensureOverviewAssets();
    let presenceMap = questionnaireByExternalId;
    const ctx = staffAssetsContextRef.current;
    if (ctx && questionnairePresenceStatus !== "ready") {
      const assets = await fetchOverviewPersonnelAssets({
        overviewRows: ctx.rows,
        dataset: ctx.dataset,
      });
      presenceMap = assets.questionnairePresence;
      startTransition(() => {
        setQuestionnaireByExternalId((current) => ({
          ...current,
          ...assets.questionnairePresence,
        }));
        setQuestionnaireSourceIdByExternalId(assets.questionnaireSourceIds);
        setDocumentsByExternalId(assets.documents);
        setQuestionnairePresenceStatus("ready");
      });
    }

    const questionnaireRows = context.rows.filter(
      (row) => row.externalId && presenceMap[row.externalId],
    );
    const files: Array<{ name: string; data: Uint8Array }> = [];

    for (const row of questionnaireRows) {
      const questionnaire = await api.getPersonQuestionnaire(
        questionnaireSourceIdByExternalId[row.externalId] ?? row.externalId,
      );
      if (!questionnaire?.fileData) continue;
      const fileName = sanitizeFileName(
        buildQuestionnaireExportFileName(
          row.name,
          row.externalId ? callSignByExternalId[row.externalId] : undefined,
        ),
      );
      files.push({
        name: fileName.toLocaleLowerCase("uk-UA").endsWith(".pdf")
          ? fileName
          : `${fileName}.pdf`,
        data: dataUrlToUint8Array(questionnaire.fileData),
      });
    }

    if (files.length) {
      downloadBlob(
        createStoredZipBlob(files),
        `Анкети огляд ${exportedAt}.zip`,
      );
    }
    setMessage(
      `Експортовано таблицю: ${context.rows.length} рядків · анкет у ZIP: ${files.length}.`,
    );
  };

  const exportImportantOverviewColumns = async (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => {
    const exportedAt = new Date();
    let exportRows = context.allRows ?? context.rows;
    let exportRosterRows = rosterRowsRef.current;
    let exportRosterColumns:
      | Array<{ key: string; letter?: string; originalIndex?: number }>
      | undefined;
    let archivePeriods: ReturnType<typeof parsePbArchive> = [];
    try {
      setMessage("Оновлюю «Загальний список» і archive 1ПБ для експорту…");
      const payload = await pullStaffSheetRosterImportPayload({
        source: "gviz",
      });
      const sourceSheet = payload.sheets[0];
      if (sourceSheet) {
        exportRosterColumns = sourceSheet.columns;
        exportRosterRows = fillDownRosterUnitRows(
          sourceSheet.rows.map((row, index) => ({
            __dbRowId: `google:${row.excelRowNumber || index + 2}`,
            __rowNumber: row.excelRowNumber,
            ...row.values,
          })),
        );
        const rosterLabels = Object.fromEntries(
          sourceSheet.columns.map((column) => [column.key, column.label]),
        );
        const freshRows = buildStaffOverviewRowsFromRoster(
          exportRosterRows,
          rosterLabels,
          sourceSheet.columns,
        ).map(withStaffOverviewStatus);
        if (freshRows.length) exportRows = freshRows;
      }
    } catch (error) {
      console.warn(
        "[Огляд] Не вдалося оновити Google перед експортом, використовую поточні дані",
        error,
      );
    }
    try {
      const pbWorkbook = await loadPbWorkbookFromDb();
      archivePeriods = pbWorkbook ? parsePbArchive(pbWorkbook) : [];
    } catch (error) {
      console.warn(
        "[Огляд] Не вдалося прочитати archive поточного 1ПБ",
        error,
      );
    }
    const sheets = buildImportantOverviewExportSheets(
      exportRows,
      { activeFilters: context.filters },
      exportRows,
      exportRosterRows,
      archivePeriods,
      exportRosterColumns,
    );

    await writeXlsxFile(sheets, {
      fontFamily: "Arial",
      fontSize: 10,
    }).toFile(
      buildImportantOverviewExportFileName(
        { activeFilters: context.filters },
        exportedAt,
      ),
    );
    setMessage(
      `Експортовано важливі колонки: ${context.rows.length} рядків.`,
    );
  };

  const exportRotaGudzReport = async (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => {
    try {
      await exportOverviewRotaGudzReport(context);
      const unit =
        context.filters?.find((filter) => filter.id === "unit")?.values[0] ??
        "рота";
      setMessage(`Експортовано звіт роти (ГУД): ${unit}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати звіт роти (ГУД).",
      );
    }
  };

  const exportPpdLocationReport = async (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => {
    let exportRows = context.allRows ?? context.rows;
    let exportRosterRows = rosterRowsRef.current;
    try {
      setMessage("Оновлюю «Загальний список» для експорту ППД / Полігон…");
      const payload = await pullStaffSheetRosterImportPayload({
        source: "gviz",
      });
      const sourceSheet = payload.sheets[0];
      if (sourceSheet) {
        exportRosterRows = fillDownRosterUnitRows(
          sourceSheet.rows.map((row, index) => ({
            __dbRowId: `google:${row.excelRowNumber || index + 2}`,
            __rowNumber: row.excelRowNumber,
            ...row.values,
          })),
        );
        const rosterLabels = Object.fromEntries(
          sourceSheet.columns.map((column) => [column.key, column.label]),
        );
        const freshRows = buildStaffOverviewRowsFromRoster(
          exportRosterRows,
          rosterLabels,
          sourceSheet.columns,
        ).map(withStaffOverviewStatus);
        if (freshRows.length) {
          exportRows = freshRows;
        }
      }
    } catch (error) {
      console.warn(
        "[Огляд] Не вдалося оновити Google перед експортом ППД / Полігон",
        error,
      );
    }

    try {
      const { ppdCount, polygonCount } = await exportOverviewPpdLocationReport(
        context,
        exportRows,
        exportRosterRows,
      );
      setMessage(
        `Експортовано ППД / Полігон: ${ppdCount} на ППД, ${polygonCount} на полігоні.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати ППД / Полігон.",
      );
    }
  };

  const exportRotaBchsMorningReport = async (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => {
    let exportRows = context.allRows ?? context.rows;
    let exportRosterRows = rosterRowsRef.current;
    const selectedUnits = context.filters?.find((filter) => filter.id === "unit")?.values ?? [];
    let staffCount = summarizeNovaStaffForUnits(
      rosterRowsRef.current, selectedUnits, rosterColumnsRef.current,
    ).staff;
    try {
      setMessage("Оновлюю «Загальний список» для експорту БЧС…");
      const payload = await pullStaffSheetRosterImportPayload({
        source: "gviz",
      });
      const sourceSheet = payload.sheets[0];
      if (sourceSheet) {
        exportRosterRows = fillDownRosterUnitRows(
          sourceSheet.rows.map((row, index) => ({
            __dbRowId: `google:${row.excelRowNumber || index + 2}`,
            __rowNumber: row.excelRowNumber,
            ...row.values,
          })),
        );
        const rosterLabels = Object.fromEntries(
          sourceSheet.columns.map((column) => [column.key, column.label]),
        );
        const freshRows = buildStaffOverviewRowsFromRoster(
          exportRosterRows,
          rosterLabels,
          sourceSheet.columns,
        ).map(withStaffOverviewStatus);
        if (freshRows.length) {
          exportRows = freshRows;
          staffCount = summarizeNovaStaffForUnits(
            exportRosterRows, selectedUnits, sourceSheet.columns,
          ).staff;
        }
      }
    } catch (error) {
      console.warn(
        "[Огляд] Не вдалося оновити Google перед експортом БЧС",
        error,
      );
    }

    try {
      await exportOverviewRotaBchsMorningReport({
        ...context,
        allRows: exportRows,
      }, staffCount, exportRosterRows);
      const unit =
        context.filters?.find((filter) => filter.id === "unit")?.values[0] ??
        "рота";
      setMessage(`Експортовано БЧС (ранковий ПБ): ${unit}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати БЧС (ранковий ПБ).",
      );
    }
  };

  return (
    <main className="main-panel overview-page">
      <div className="overview-screen">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Огляд
          </Typography>
          <Typography variant="body2" color="text.secondary">
            {message}
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button variant="outlined" onClick={() => startLoad(true, { refreshStaffServer: true })}>
            Оновити
          </Button>
        </Stack>
      </header>

      {isLoading && data ? <LinearProgress color="primary" /> : null}

      <section className="overview-metrics">
        <article className="overview-metric-card">
          <span>
            <PersonOutlinedIcon fontSize="small" />
            {source === "staff" ? "У штаті" : "Усього"}
          </span>
          <strong>{metrics?.total ?? "—"}</strong>
          {source === "staff" && staffSummary.positions > 0 ? (
            <em className="overview-metric-note">
              штат {staffSummary.positions}
              {staffSummary.vacant > 0
                ? ` · вакант ${staffSummary.vacant}`
                : ""}
            </em>
          ) : null}
        </article>
        <article className="overview-metric-card tone-ok">
          <span>
            <ShieldOutlinedIcon fontSize="small" /> На службі
          </span>
          <strong>{metrics?.onDuty ?? "—"}</strong>
        </article>
        <article className="overview-metric-card tone-trip">
          <span>
            <BusinessCenterOutlinedIcon fontSize="small" /> Відрядження
          </span>
          <strong>{metrics?.businessTrip ?? "—"}</strong>
        </article>
        <article className="overview-metric-card tone-leave">
          <span>
            <BeachAccessOutlinedIcon fontSize="small" /> Відпустка
          </span>
          <strong>{metrics?.leave ?? "—"}</strong>
        </article>
        <article className="overview-metric-card tone-medical">
          <span>
            <LocalHospitalOutlinedIcon fontSize="small" /> Лікування
          </span>
          <strong>{metrics?.medical ?? "—"}</strong>
        </article>
      </section>

      <section className="overview-toolbar">
        <label
          className={`overview-search${isNameListSearch ? " is-multiline" : ""}`}
        >
          <SearchOutlinedIcon fontSize="small" />
          <textarea
            value={query}
            rows={1}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="ПІБ або список (по одному в рядок) — Ctrl+V"
            spellCheck={false}
          />
          {query.trim() ? (
            <button
              type="button"
              className="overview-search-clear"
              aria-label="Очистити пошук"
              onClick={() => setQuery("")}
            >
              ✕
            </button>
          ) : null}
        </label>
        <TextField
          select
          size="small"
          className="overview-filter"
          label="Джерело"
          value={source}
          onChange={(event) => {
            const next = event.target.value as OverviewSourceFilter;
            setSource(next);
          }}
        >
          {SOURCE_FILTERS.map((item) => (
            <MenuItem key={item.value} value={item.value}>
              {item.value === "staff"
                ? formatStaffSourceLabel(rosterUpdatedAt)
                : item.value === "ejoos"
                  ? data?.importName
                    ? `ЕЖООС · ${data.importName}`
                    : "ЕЖООС"
                  : item.label}
            </MenuItem>
          ))}
        </TextField>
      </section>

      {nameListMatchStats ? (
        <Alert
          severity={nameListMatchStats.missing.length ? "warning" : "success"}
          className="overview-name-list-alert"
        >
          Список: знайдено {nameListMatchStats.matched.length} з{" "}
          {nameListMatchStats.total}
          {nameListMatchStats.missing.length
            ? ` · не знайдено: ${nameListMatchStats.missing.join("; ")}`
            : ""}
        </Alert>
      ) : null}

      <div className="overview-table-panel">
        {isLoading && !data ? (
          <div className="overview-table-preloader">
            <MagneticTapePreloader
              status="ЗАВАНТАЖЕННЯ ОГЛЯДУ"
              hint="Читаю Штатку та готую таблицю особового складу."
            />
          </div>
        ) : null}
        <OverviewVirtualTable
          rows={filteredRows}
          photos={photos}
          onNeedPhoto={onNeedPhoto}
          onVisibleRowsChange={onVisibleRowsChange}
          questionnaireByExternalId={questionnaireByExternalId}
          questionnaireLoading={questionnairePresenceStatus === "loading"}
          questionnairePresenceStatus={questionnairePresenceStatus}
          documentsByExternalId={documentsByExternalId}
          onColumnVisibilityChange={handleColumnVisibilityChange}
          onOpenQuestionnaire={openQuestionnaire}
          emptyMessage="Немає записів за поточними фільтрами."
          onExport={(context) => void exportOverviewTable(context)}
          onImportantExport={(context) =>
            void exportImportantOverviewColumns(context)
          }
          onRotaGudzExport={(context) => void exportRotaGudzReport(context)}
          onRotaBchsMorningExport={(context) =>
            void exportRotaBchsMorningReport(context)
          }
          onPpdLocationExport={(context) =>
            void exportPpdLocationReport(context)
          }
          copyTextBuilder={buildOverviewWhatsAppCopyText}
          rotaCopyTextBuilder={buildOverviewRotaCopyText}
        />
        <footer className="overview-table-footer">
          <span>
            Показано всі {filteredRows.length} з {sourceRows.length}
          </span>
        </footer>
      </div>
      </div>

      <aside className="overview-side">
        <section className="overview-side-card overview-critical-card">
          <div className="panel-heading">Критичні терміни</div>
          <ul className="overview-critical-list">
            {sideStats.critical.map((item) => {
              const splitAt = item.text.indexOf(":");
              const name =
                splitAt >= 0 ? item.text.slice(0, splitAt).trim() : item.text;
              const meta = splitAt >= 0 ? item.text.slice(splitAt + 1).trim() : "";
              return (
                <li key={item.id} className={`tone-${item.severity}`}>
                  <span className="overview-critical-name">{name}</span>
                  {meta ? (
                    <span className="overview-critical-meta">{meta}</span>
                  ) : null}
                </li>
              );
            })}
            {sideStats.critical.length === 0 && (
              <li className="tone-info">Критичних термінів немає</li>
            )}
          </ul>
        </section>

        <section className="overview-side-card">
          <div className="panel-heading">Зміни сьогодні</div>
          <div className="overview-today-changes">
            <Chip
              className="overview-status-chip tone-ok"
              label={`+${sideStats.todayChanges.onDuty}`}
              size="small"
            />
            <Chip
              className="overview-status-chip tone-trip"
              label={`+${sideStats.todayChanges.businessTrip}`}
              size="small"
            />
            <Chip
              className="overview-status-chip tone-leave"
              label={`+${sideStats.todayChanges.leave}`}
              size="small"
            />
            <Chip
              className="overview-status-chip tone-medical"
              label={`+${sideStats.todayChanges.medical}`}
              size="small"
            />
          </div>
          <Typography variant="body2" color="text.secondary">
            Усього змін: {sideStats.todayChanges.total}
          </Typography>
        </section>

        <section className="overview-side-card overview-updates-card">
          <div className="panel-heading">Оновлення сьогодні</div>
          <strong>{String(sideStats.todayUpdates).padStart(2, "0")}</strong>
          <span>записів оновлено</span>
        </section>
      </aside>
    </main>
  );
}
