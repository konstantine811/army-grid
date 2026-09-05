import {
  useCallback,
  useDeferredValue,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import writeXlsxFile, { type SheetData } from "write-excel-file/browser";
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
  openPersonnelInNewTab,
} from "../../app/navigation";
import {
  loadPersonnelDataset,
  type PersonnelDataset,
} from "../../data/personnelDataset";
import {
  CacheKeys,
  fetchWithCache,
  jsonChanged,
  peekDataCache,
  readDataCache,
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
  buildOverviewPersonnelIdentities,
} from "./overviewPersonnelAssets";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import {
  overviewNameMatchesQuery,
  parseOverviewNameQueries,
} from "./overviewNameSearch";
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

const SOURCE_FILTERS = [
  { value: "staff", label: "Штатка" },
  { value: "ejoos", label: "ЕЖООС" },
  { value: "all", label: "Усі джерела" },
] as const;

type OverviewSourceFilter = (typeof SOURCE_FILTERS)[number]["value"];

const normalizeRosterText = normalizeRosterMatchText;

const directQuestionnairePresence = (
  items: BackendPersonQuestionnaireMeta[] | null | undefined,
) =>
  Object.fromEntries(
    (Array.isArray(items) ? items : [])
      .map((item) => item.personExternalId?.trim())
      .filter(Boolean)
      .map((id) => [id, true] as const),
  ) as Record<string, true>;

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

export function OverviewPage() {
  const cachedQuestionnaires = peekDataCache<BackendPersonQuestionnaireMeta[]>(
    CacheKeys.questionnairesMeta,
  );
  const [data, setData] = useState<BackendPersonnelOverview | null>(null);
  const [photos, setPhotos] = useState<Record<string, string>>({});
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >(() => directQuestionnairePresence(cachedQuestionnaires));
  const [questionnairePresenceStatus, setQuestionnairePresenceStatus] =
    useState<"loading" | "ready">("loading");
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
  const [rosterUpdatedAt, setRosterUpdatedAt] = useState<string | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const rosterRowsRef = useRef<EjournalPreviewRow[]>([]);
  const personnelRowsRef = useRef<EjournalPreviewRow[]>([]);
  const rosterColumnsRef = useRef<
    Array<{ key: string; letter?: string; originalIndex?: number }>
  >([]);
  const rosterLabelsRef = useRef<Record<string, string>>({});
  const photosRef = useRef<Record<string, string>>({});
  const requestedPhotoKeysRef = useRef(new Set<string>());
  const loadControllerRef = useRef<AbortController | null>(null);
  const [rosterEpoch, setRosterEpoch] = useState(0);
  const [personnelEpoch, setPersonnelEpoch] = useState(0);
  photosRef.current = photos;

  const load = async (
    signal?: AbortSignal,
    options?: { force?: boolean },
  ) => {
    const alive = () => !signal?.aborted;
    const force = Boolean(options?.force);
    setIsLoading(true);
    try {
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
          try {
            mergedOverview = await runHeavyJob({
              type: "mergeOverview",
              overview,
              rosterRows,
              rosterLabels,
              columns: rosterColumns,
            });
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

      let assetsSeq = 0;
      const applyPersonnelAssets = async (
        overviewRows: BackendPersonnelOverviewRow[],
        dataset: PersonnelDataset,
      ) => {
        const seq = ++assetsSeq;
        setQuestionnairePresenceStatus("loading");
        const questionnairePromise = fetchWithCache({
            key: CacheKeys.questionnairesMeta,
            force,
            signal,
            fetcher: () => api.listPersonQuestionnaires({ signal }),
            isChanged: jsonChanged,
          }).catch(() => []);
        const documentPromise = fetchWithCache({
            key: CacheKeys.documentsAll,
            force,
            signal,
            fetcher: () => api.listAllPersonDocuments({ signal }),
            isChanged: jsonChanged,
          }).catch(() => []);
        const questionnaireList = await questionnairePromise;
        if (!alive() || seq !== assetsSeq) return;
        setQuestionnaireByExternalId((current) => ({
          ...current,
          ...directQuestionnairePresence(questionnaireList),
        }));

        const paintAssets = async (
          people: EjournalPreviewRow[],
          documentList: Awaited<typeof documentPromise>,
        ) => {
          try {
            const assets = await runHeavyJob({
              type: "applyOverviewAssets",
              overviewRows,
              personnelIdentities: buildOverviewPersonnelIdentities(people),
              questionnaires: Array.isArray(questionnaireList)
                ? questionnaireList.map(({ personExternalId, fileName }) => ({
                    personExternalId,
                    fileName,
                  }))
                : [],
              documents: Array.isArray(documentList)
                ? documentList.map(
                    ({ id, personExternalId, type, title }) => ({
                      id,
                      personExternalId,
                      type,
                      title,
                    }),
                  )
                : [],
            });
            if (!alive() || seq !== assetsSeq) return;
            setQuestionnaireByExternalId((current) => ({
              ...current,
              ...assets.questionnairePresence,
            }));
            setQuestionnaireSourceIdByExternalId(assets.questionnaireSourceIds);
            setDocumentsByExternalId(assets.documents);
            setQuestionnairePresenceStatus("ready");
          } catch (error) {
            console.warn("[Огляд] Не вдалося підставити дані з Особового складу", error);
            if (alive() && seq === assetsSeq) {
              setQuestionnairePresenceStatus("ready");
            }
          }
        };

        const personnelRows = dataset.rows;
        if (!alive() || seq !== assetsSeq) return;
        personnelRowsRef.current = personnelRows;
        setPersonnelEpoch((value) => value + 1);
        // Questionnaire badges must not wait for the document metadata request.
        await paintAssets(personnelRowsRef.current, []);
        const documentList = await documentPromise;
        if (!alive() || seq !== assetsSeq) return;
        await paintAssets(personnelRowsRef.current, documentList);
      };
      const startDatasetAssets = (dataset: PersonnelDataset) => {
        const staffRows = buildStaffOverviewRowsFromPersonnel(
          dataset.rows,
          dataset.rosterLabels,
        );
        void applyPersonnelAssets(staffRows, dataset);
      };

      const [cachedDataset] = await Promise.all([
        readDataCache<PersonnelDataset>(CacheKeys.personnelDataset),
        readDataCache<BackendPersonnelOverview>(CacheKeys.overview),
      ]);
      let rosterRows: EjournalPreviewRow[] = [];
      let rosterLabels: Record<string, string> = {};
      let rosterColumns: Array<{
        key: string;
        letter?: string;
        originalIndex?: number;
      }> = [];
      rosterRows = cachedDataset?.rosterRows ?? [];
      rosterLabels = cachedDataset?.rosterLabels ?? {};
      rosterColumns = cachedDataset?.rosterColumns ?? [];
      let paintedFromCache = false;
      if (cachedDataset?.rows?.length) {
        personnelRowsRef.current = cachedDataset.rows;
        setPersonnelEpoch((value) => value + 1);
        setRosterUpdatedAt(cachedDataset.rosterUpdatedAt);
        await applyOverview(
          buildPersonnelStaffOverview(
            cachedDataset.rows,
            cachedDataset.rosterLabels,
            {
              importId: cachedDataset.version.rosterImportId,
              importName: "Особовий склад · спільний dataset",
            },
          ),
          rosterRows,
          rosterLabels,
          rosterColumns,
          true,
          true,
        );
        paintedFromCache = true;
        setIsLoading(false);
        startDatasetAssets(cachedDataset);
      }

      const datasetPromise = loadPersonnelDataset({
        force,
        signal,
        onCached: async (nextDataset) => {
          if (!alive()) return;
          rosterRows = nextDataset.rosterRows;
          rosterLabels = nextDataset.rosterLabels;
          rosterColumns = nextDataset.rosterColumns;
          personnelRowsRef.current = nextDataset.rows;
          setPersonnelEpoch((value) => value + 1);
          setRosterUpdatedAt(nextDataset.rosterUpdatedAt);
          await applyOverview(
            buildPersonnelStaffOverview(
              nextDataset.rows,
              nextDataset.rosterLabels,
              {
                importId: nextDataset.version.rosterImportId,
                importName: "Особовий склад · спільний dataset",
              },
            ),
            rosterRows,
            rosterLabels,
            rosterColumns,
            true,
            true,
          );
          paintedFromCache = true;
          setIsLoading(false);
          startDatasetAssets(nextDataset);
        },
      })
        .then((nextDataset) => {
          rosterRows = nextDataset.rosterRows;
          rosterLabels = nextDataset.rosterLabels;
          rosterColumns = nextDataset.rosterColumns;
          personnelRowsRef.current = nextDataset.rows;
          setPersonnelEpoch((value) => value + 1);
          setRosterUpdatedAt(nextDataset.rosterUpdatedAt);
          startDatasetAssets(nextDataset);
          return nextDataset;
        })
        .catch(() => cachedDataset);

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
                    !meta.complete,
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

      if (source !== "staff") {
        void applyPersonnelAssets(mergedOverview.rows, dataset);
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

  const startLoad = (force = false) => {
    loadControllerRef.current?.abort();
    const controller = new AbortController();
    loadControllerRef.current = controller;
    void load(controller.signal, { force });
  };

  useEffect(() => {
    startLoad();
    return () => {
      loadControllerRef.current?.abort();
    };
  }, []);

  useEffect(() => {
    const onSynced = () => {
      startLoad(true);
    };
    window.addEventListener(STAFF_SHEET_SYNCED_EVENT, onSynced);
    return () => window.removeEventListener(STAFF_SHEET_SYNCED_EVENT, onSynced);
  }, []);

  const deferredQuery = useDeferredValue(query);
  const nameQueries = useMemo(
    () => parseOverviewNameQueries(deferredQuery),
    [deferredQuery],
  );
  const isNameListSearch = nameQueries.length > 1;

  const personnelOverviewRows = useMemo(() => {
    if (!personnelRowsRef.current.length) {
      return [] as BackendPersonnelOverviewRow[];
    }
    return buildStaffOverviewRowsFromPersonnel(
      personnelRowsRef.current,
      rosterLabelsRef.current,
    ).map(withStaffOverviewStatus);
  }, [personnelEpoch, rosterEpoch]);

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

      return [
        row.name,
        row.externalId,
        row.rank,
        row.positionTitle,
        row.unit,
        row.statusLabel,
        row.fighterDirection,
        row.fighterExitDate,
        row.fighterReturnDate,
        row.fighterTotalDays,
        row.fighterStatus,
        row.externalId && documentsByExternalId[row.externalId]
          ? documentsByExternalId[row.externalId].labels.join(" ")
          : "",
      ]
        .join(" ")
        .split(" ")
        .map(normalizeRosterText)
        .join(" ")
        .includes(normalizedQuery);
    });
  }, [
    documentsByExternalId,
    isNameListSearch,
    nameQueries,
    source,
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
      openPersonnelInNewTab({
        rowId: target.rowId,
        externalId: target.externalId,
      });
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

  useEffect(() => {
    if (!filteredRows.length || filteredRows.length > 40) return;
    const missing = filteredRows.filter((row) => {
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
  }, [filteredRows, rosterEpoch]);

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

    const questionnaireRows = context.rows.filter(
      (row) => row.externalId && questionnaireByExternalId[row.externalId],
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
    let archivePeriods: ReturnType<typeof parsePbArchive> = [];
    try {
      setMessage("Оновлюю «Загальний список» і archive 1ПБ для експорту…");
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
          <Button variant="outlined" onClick={() => startLoad(true)}>
            Оновити
          </Button>
        </Stack>
      </header>

      {isLoading && <LinearProgress color="primary" />}

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
        <OverviewVirtualTable
          rows={filteredRows}
          photos={photos}
          onNeedPhoto={onNeedPhoto}
          questionnaireByExternalId={questionnaireByExternalId}
          questionnaireLoading={questionnairePresenceStatus === "loading"}
          documentsByExternalId={documentsByExternalId}
          onOpenQuestionnaire={openQuestionnaire}
          emptyMessage={
            isLoading && !data
              ? "Завантаження огляду..."
              : "Немає записів за поточними фільтрами."
          }
          onExport={(context) => void exportOverviewTable(context)}
          onImportantExport={(context) =>
            void exportImportantOverviewColumns(context)
          }
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
