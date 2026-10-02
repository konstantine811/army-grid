import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from "react";
import {
  Alert,
  LinearProgress,
} from "@/components/sci/SciPrimitives";
import { PersonnelMobileTabs } from "./PersonnelMobileTabs";
import {
  api,
  type BackendPersonDocument,
  type BackendPersonQuestionnaire,
} from "../../api";
import { useAuth } from "../../auth/AuthProvider";
import {
  CacheKeys,
  peekDataCache,
  subscribeDataCache,
} from "../../data/idbDataCache";
import { STAFF_SHEET_SYNCED_EVENT } from "../../data/staffSheetAutoSync";
import {
  personnelDatasetToPreview,
  type PersonnelDataset,
} from "../../data/personnelDataset";
import type {
  DbPreviewState,
  EjournalPreviewRow,
} from "../ejournal/ejournalTypes";
import {
  PersonnelListPanel,
  type PersonnelListPanelHandle,
} from "./PersonnelListPanel";
import {
  PersonnelDocumentsPanel,
  type PersonnelDocumentMode,
} from "./PersonnelDocumentsPanel";
import {
  clearPersonnelFocusTarget,
  findPersonnelRowByFocusTarget,
  normalizePersonnelFocusTarget,
  readPersonnelFocusTarget,
  type PersonnelFocusTarget,
} from "./personnelFocus";
import {
  indexPersonnelV2AnketaRows,
  loadPersonnelV2AnketaRows,
  type PersonnelV2AnketaIndex,
} from "../personnel-v2/personnelV2AnketaFill";
import {
  buildQuestionnaireExportFileName,
  getPersonDisplayName,
  isLikelyPersonnelRow,
} from "./personnelUtils";
import { sanitizeFileName } from "../../shared/browserExport";
import { notifyPersonnelAttachmentChanged } from "../../shared/personnelAttachmentSync";
import { readStoredPersonPhones } from "./personPhonesStore";
import {
  buildQuestionnairePresencePeople,
  clearAvailablePersonPhotoIdsCache,
  collectPersonnelListPhotoUpdates,
  loadAvailablePersonPhotoIds,
  peekAvailablePersonPhotoIds,
  questionnaireFileMatchesPerson,
} from "./personAttachments";
import { PersonnelPersonCard } from "./PersonnelPersonCard";
import { staffFilterScopeChange } from "./personnelListScope";
import {
  collectMissingQuestionnairePeople,
  collectPhotoExtractTargets,
} from "./personnelAttachmentTargets";
import { PersonnelPageOverlays } from "./PersonnelPageOverlays";
import { PersonnelPageToolbar } from "./PersonnelPageToolbar";
import { loadQuestionnairePresence } from "./loadQuestionnairePresence";
import {
  applyPersonnelListPreview,
  loadFullPersonnelList,
  loadStaffPersonnelList,
  publishedPersonnelScope,
} from "./personnelListLoad";
import { useSelectedPersonDetails } from "./useSelectedPersonDetails";
import { usePersonnelWorkbookImports } from "./usePersonnelWorkbookImports";
import { bindPersonnelCardEdits } from "./personnelCardEdits";
import { healOrphanPersonnelAttachments } from "./healOrphanPersonnelAttachments";
import { usePersonnelSelectedCard } from "./usePersonnelSelectedCard";
import { usePersonnelQuestionnairePreview } from "./usePersonnelQuestionnairePreview";
import {
  type PersonnelListIndex,
} from "./personnelListIndex";

export function PersonnelPage({
  onOpenDocuments,
}: {
  onOpenDocuments: (
    row: EjournalPreviewRow,
    mode?: PersonnelDocumentMode,
    meta?: { fullPosition?: string },
  ) => void;
}) {
  const { canEditArea } = useAuth();
  const canEdit = canEditArea("personnel");
  const personnelAllRowsRef = useRef<EjournalPreviewRow[]>([]);
  const personnelScopeRef = useRef({ all: false, archive: false });
  const personnelSourceRowsRef = useRef<EjournalPreviewRow[]>([]);
  const personnelListIndexRef = useRef<PersonnelListIndex>({
    records: [],
    staffCounts: { all: 0, in: 0, archive: 0 },
  });
  const questionnairePresencePeopleRef = useRef<
    ReturnType<typeof buildQuestionnairePresencePeople>
  >([]);
  const [personnelDataEpoch, setPersonnelDataEpoch] = useState(0);
  const [rosterLabels, setRosterLabels] = useState<Record<string, string>>({});
  const [selectedRowId, setSelectedRowId] = useState("");
  const [mobilePane, setMobilePane] = useState<"list" | "card" | "side">(
    "list",
  );
  const personnelFocusLockRef = useRef<PersonnelFocusTarget | null>(null);
  const personnelLoadGenerationRef = useRef(0);
  const personnelLoadControllerRef = useRef<AbortController | null>(null);
  const requestedPhotoIdsRef = useRef(new Set<string>());
  const personnelRowByExternalIdRef = useRef(
    new Map<string, EjournalPreviewRow>(),
  );
  const personnelDatasetFingerprintRef = useRef("");
  const holdStatusUntilRef = useRef(0);
  const listPanelRef = useRef<PersonnelListPanelHandle | null>(null);
  const questionnairePresenceLoadRef = useRef<Promise<
    Array<{ personExternalId: string; fileName?: string | null }>
  > | null>(null);
  const [personnelListInitialSearch] = useState(
    () => readPersonnelFocusTarget().search,
  );
  const [isMobilePersonnelLayout, setIsMobilePersonnelLayout] = useState(() =>
    typeof window !== "undefined"
      ? window.matchMedia("(max-width: 980px)").matches
      : false,
  );
  const [photoByExternalId, setPhotoByExternalId] = useState<
    Record<string, string>
  >({});
  const [anketaIndex, setAnketaIndex] = useState<PersonnelV2AnketaIndex | null>(
    null,
  );

  useEffect(() => {
    let cancelled = false;
    void loadPersonnelV2AnketaRows()
      .then((rows) => {
        if (!cancelled) setAnketaIndex(indexPersonnelV2AnketaRows(rows));
      })
      .catch(() => {
        if (!cancelled) setAnketaIndex(indexPersonnelV2AnketaRows([]));
      });
    return () => {
      cancelled = true;
    };
  }, []);
  const [phonesByExternalId, setPhonesByExternalId] = useState<
    Record<string, string[]>
  >(() => readStoredPersonPhones());
  const [phoneDocByExternalId, setPhoneDocByExternalId] = useState<
    Record<string, BackendPersonDocument>
  >({});
  const [phoneDraft, setPhoneDraft] = useState("");
  const [isSavingPhone, setIsSavingPhone] = useState(false);
  const [questionnaireByExternalId, setQuestionnaireByExternalId] = useState<
    Record<string, true>
  >({});
  const [questionnairePresenceStatus, setQuestionnairePresenceStatus] =
    useState<"loading" | "ready" | "error">("loading");
  const [staffFilter, setStaffFilter] = useState<"all" | "in" | "archive">(
    "in",
  );
  const [photoCropFile, setPhotoCropFile] = useState<File | null>(null);
  const [isPhotoCropOpen, setIsPhotoCropOpen] = useState(false);
  const [questionnaire, setQuestionnaire] =
    useState<BackendPersonQuestionnaire | null>(null);
  const [personRelatedDocuments, setPersonRelatedDocuments] = useState<
    BackendPersonDocument[]
  >([]);
  const [isDiskFloatingCrop, setIsDiskFloatingCrop] = useState(false);
  const [isDiskSearchOpen, setIsDiskSearchOpen] = useState(false);
  const [isPhotoExtractOpen, setIsPhotoExtractOpen] = useState(false);
  const [isPhotoLightboxOpen, setIsPhotoLightboxOpen] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const [isLoading, setIsLoading] = useState(false);
  const [photoIndexReady, setPhotoIndexReady] = useState(0);
  const personnelRows = useMemo(() => {
    void personnelDataEpoch;
    return personnelListIndexRef.current.records;
  }, [personnelDataEpoch]);
  useEffect(() => {
    const next = new Map<string, EjournalPreviewRow>();
    for (const record of personnelRows) {
      const externalId = record.summary.externalId;
      if (externalId) next.set(externalId, record.row);
    }
    personnelRowByExternalIdRef.current = next;
  }, [personnelRows]);

  useEffect(() => {
    const media = window.matchMedia("(max-width: 980px)");
    const sync = () => setIsMobilePersonnelLayout(media.matches);
    sync();
    media.addEventListener("change", sync);
    return () => media.removeEventListener("change", sync);
  }, []);

  const markQuestionnaireInDb = useCallback((externalId: string) => {
    if (!externalId) return;
    setQuestionnaireByExternalId((current) => ({
      ...current,
      [externalId]: true,
    }));
  }, []);

  const clearQuestionnairePresence = useCallback((externalIds: string[]) => {
    if (!externalIds.length) return;
    setQuestionnaireByExternalId((current) => {
      const next = { ...current };
      for (const id of externalIds) delete next[id];
      return next;
    });
  }, []);

  const staffCounts = useMemo(() => {
    void personnelDataEpoch;
    return personnelListIndexRef.current.staffCounts;
  }, [personnelDataEpoch]);
  const selectedRecord = useMemo(
    () =>
      personnelRows.find((record) => record.row.__dbRowId === selectedRowId) ??
      null,
    [personnelRows, selectedRowId],
  );
  const selectedRow = selectedRecord?.row ?? null;
  const shouldLoadSelectedPersonDetails = Boolean(
    selectedRowId &&
      selectedRow &&
      (!isMobilePersonnelLayout || mobilePane !== "list"),
  );
  const {
    selectedSummary,
    selectedPhoto,
    selectedPhotoFullUrl,
    selectedCallSign,
    editableFields,
    editableFieldsBySection,
    editValues,
    setEditValues,
    selectedAnketaRow,
    cardSummary,
    birthDateWithAge,
    rosterFieldRows,
    rosterStatus,
    fighterStatusFieldRows,
    selectedFullPosition,
    selectedPersonNote,
    savedPhones,
    parsedPhones,
  } = usePersonnelSelectedCard({
    selectedRow,
    rosterLabels,
    anketaIndex,
    phonesByExternalId,
    photoByExternalId,
    photoIndexReady,
    questionnaireFileName: questionnaire?.fileName ?? undefined,
    setPhoneDraft,
  });
  const questionnaireExportFileName = useMemo(() => {
    const stored = String(questionnaire?.fileName ?? "").trim();
    if (
      stored &&
      questionnaireFileMatchesPerson(stored, [selectedSummary.name])
    ) {
      return sanitizeFileName(stored);
    }
    return sanitizeFileName(
      buildQuestionnaireExportFileName(selectedSummary.name, selectedCallSign),
    );
  }, [questionnaire?.fileName, selectedCallSign, selectedSummary.name]);
  const {
    pendingQuestionnaireFile,
    isQuestionnairePreviewOpen,
    isDiskFloatingPreview,
    isUploadingQuestionnaire,
    isLoadingQuestionnairePreview,
    questionnairePreviewTitle,
    questionnairePreviewUrl,
    shareSource: currentQuestionnaireShareSource,
    cropFile: questionnaireCropFileSource,
    releaseQuestionnairePreviewCache,
    closeQuestionnairePreview,
    openDiskQuestionnairePreview,
    openQuestionnairePreview,
    openQuestionnaireInNewTab,
    beginQuestionnaireReview,
    exportCurrentQuestionnaire,
    revealCurrentQuestionnaireInFinder,
    confirmPendingQuestionnaire,
  } = usePersonnelQuestionnairePreview({
    questionnaire,
    setQuestionnaire,
    selectedRow,
    personName: selectedSummary.name,
    externalId: selectedSummary.externalId,
    callSign: selectedCallSign,
    exportFileName: questionnaireExportFileName,
    personNames: personnelRows.map((item) => item.summary.name),
    setMessage,
    markQuestionnaireInDb,
  });

  useEffect(() => {
    setIsPhotoLightboxOpen(false);
  }, [selectedRowId]);

  useEffect(() => {
    if (!isPhotoLightboxOpen) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setIsPhotoLightboxOpen(false);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [isPhotoLightboxOpen]);

  useSelectedPersonDetails({
    enabled: shouldLoadSelectedPersonDetails,
    canEdit,
    personnelRows,
    selectedRow,
    externalId: selectedSummary.externalId,
    personName: selectedSummary.name,
    setQuestionnaire,
    setDocuments: setPersonRelatedDocuments,
    markQuestionnairePresent: markQuestionnaireInDb,
    setPhoneDocument: (id, document) => {
      setPhoneDocByExternalId((current) => ({ ...current, [id]: document }));
    },
    setPhonesByExternalId,
  });

  const resetPersonnelPhotoRequests = () => {
    requestedPhotoIdsRef.current.clear();
    clearAvailablePersonPhotoIdsCache();
  };

  const applyPhotoUrlUpdates = useCallback((updates: Record<string, string>) => {
    if (!Object.keys(updates).length) return;
    for (const externalId of Object.keys(updates)) {
      requestedPhotoIdsRef.current.add(externalId);
    }
    setPhotoByExternalId((current) => ({ ...current, ...updates }));
  }, []);

  const handleListPhotoLoadError = useCallback((externalId: string) => {
    // Keep the id requested so a broken thumbnail cannot re-enter the list
    // and restart the visible-photo effect.
    requestedPhotoIdsRef.current.add(externalId);
    setPhotoByExternalId((photos) => {
      if (!photos[externalId]) return photos;
      const next = { ...photos };
      delete next[externalId];
      return next;
    });
  }, []);

  const loadVisiblePersonnelPhotos = useCallback(
    (externalIds: string[]) => {
      const peekIds = peekAvailablePersonPhotoIds();
      if (peekIds?.size) {
        applyPhotoUrlUpdates(
          collectPersonnelListPhotoUpdates(
            externalIds,
            personnelRowByExternalIdRef.current,
            peekIds,
            requestedPhotoIdsRef.current,
          ),
        );
      }

      void loadAvailablePersonPhotoIds().then((availableIds) => {
        setPhotoIndexReady((value) => (value > 0 ? value : 1));
        applyPhotoUrlUpdates(
          collectPersonnelListPhotoUpdates(
            externalIds,
            personnelRowByExternalIdRef.current,
            availableIds,
            requestedPhotoIdsRef.current,
          ),
        );
      });
    },
    [applyPhotoUrlUpdates],
  );

  const loadPersonnelQuestionnaireIds = async (
    rows?: EjournalPreviewRow[],
    signal?: AbortSignal,
    prefetchedItems?: Promise<
      Array<{ personExternalId: string; fileName?: string | null }>
    >,
    options?: { force?: boolean },
  ) => {
    if (!options?.force && questionnairePresenceLoadRef.current) {
      return questionnairePresenceLoadRef.current;
    }

    const task = (async () => {
      setQuestionnairePresenceStatus("loading");
      try {
        const loaded = await loadQuestionnairePresence({
          rows: rows ?? personnelAllRowsRef.current ?? [],
          signal,
          prefetchedItems,
          force: options?.force,
          datasetFingerprint: personnelDatasetFingerprintRef.current,
          onPhotoAliasesReady: () =>
            setPhotoIndexReady((value) => (value > 0 ? value + 1 : 1)),
          onPeople: (people) => {
            questionnairePresencePeopleRef.current = people;
          },
        });
        if (loaded.uiPresence && !loaded.aborted) {
          setQuestionnaireByExternalId(loaded.uiPresence);
          setQuestionnairePresenceStatus("ready");
        }
        return loaded.items;
      } catch {
        setQuestionnairePresenceStatus("error");
        return [];
      }
    })();

    questionnairePresenceLoadRef.current = task;
    try {
      return await task;
    } finally {
      if (questionnairePresenceLoadRef.current === task) {
        questionnairePresenceLoadRef.current = null;
      }
    }
  };

  const personnelPreviewRefs = {
    focusLockRef: personnelFocusLockRef,
    fingerprintRef: personnelDatasetFingerprintRef,
    sourceRowsRef: personnelSourceRowsRef,
    scopeRef: personnelScopeRef,
    allRowsRef: personnelAllRowsRef,
    listIndexRef: personnelListIndexRef,
    presencePeopleRef: questionnairePresencePeopleRef,
  };

  const applyPersonnelPreview = (
    preview: DbPreviewState,
    options?: {
      fromCache?: boolean;
      isCancelled?: () => boolean;
      datasetFingerprint?: string;
    },
  ) =>
    applyPersonnelListPreview({
      preview,
      options,
      refs: personnelPreviewRefs,
      resetPhotoRequests: resetPersonnelPhotoRequests,
      setSelectedRowId,
      setMobilePane,
      setPersonnelDataEpoch,
      setMessage,
    });

  const loadPersonnel = (signal?: AbortSignal, options?: { force?: boolean }) =>
    loadFullPersonnelList({
      signal,
      force: options?.force,
      generationRef: personnelLoadGenerationRef,
      allRowsRef: personnelAllRowsRef,
      setIsLoading,
      setMessage,
      setRosterLabels,
      applyPreview: applyPersonnelPreview,
      loadQuestionnaireIds: (rows, loadSignal, prefetched, loadOptions) =>
        loadPersonnelQuestionnaireIds(rows, loadSignal, prefetched, loadOptions),
      heal: healOrphanAttachmentsInBackground,
    });

  const publishOpenPersonnelScope = () => {
    const published = publishedPersonnelScope(
      personnelSourceRowsRef.current,
      personnelScopeRef.current,
    );
    personnelAllRowsRef.current = published.visibleRows;
    personnelListIndexRef.current = published.index;
    setPersonnelDataEpoch((value) => value + 1);
  };

  const loadRosterStaff = (signal?: AbortSignal, options?: { force?: boolean }) =>
    loadStaffPersonnelList({
      signal,
      force: options?.force,
      hasPaintedRows: personnelAllRowsRef.current.length > 0,
      setIsLoading,
      setMessage,
      applyPreview: (preview) => applyPersonnelPreview(preview),
      setPhotoIndexReady,
      loadQuestionnaireIds: (rows, loadSignal) =>
        loadPersonnelQuestionnaireIds(rows, loadSignal),
    });

  const startPersonnelLoad = (options?: { force?: boolean }) => {
    personnelLoadControllerRef.current?.abort();
    const controller = new AbortController();
    personnelLoadControllerRef.current = controller;
    return loadPersonnel(controller.signal, options);
  };

  const healOrphanAttachmentsInBackground = (
    rows: EjournalPreviewRow[],
    isCancelled: (() => boolean) | undefined,
    photosPromise: Promise<
      Array<{ personExternalId: string; photoData: string }>
    >,
    questionnairesPromise: Promise<Array<{ personExternalId: string }>>,
  ) =>
    healOrphanPersonnelAttachments({
      rows,
      isCancelled,
      photosPromise,
      questionnairesPromise,
      onPhonesReplaced: setPhonesByExternalId,
      reloadAttachments: async () => {
        resetPersonnelPhotoRequests();
        await Promise.all([
          loadAvailablePersonPhotoIds(),
          loadPersonnelQuestionnaireIds(rows),
        ]);
      },
    });

  const missingDiskSearchPeople = useMemo(
    () =>
      collectMissingQuestionnairePeople(
        personnelRows.map((record) => ({
          inStaff: record.inStaff,
          rowId: record.row.__dbRowId ?? "",
          externalId: record.summary.externalId,
          fullName: record.summary.name,
          callSign: record.summary.callSign,
        })),
        questionnaireByExternalId,
        photoByExternalId,
      ),
    [personnelRows, photoByExternalId, questionnaireByExternalId],
  );

  const questionnairePhotoExtractTargets = useMemo(
    () =>
      collectPhotoExtractTargets(
        personnelRows.map((record) => ({
          inStaff: record.inStaff,
          rowId: record.row.__dbRowId ?? "",
          externalId: record.summary.externalId,
          fullName: record.summary.name,
          callSign: record.summary.callSign,
        })),
        questionnaireByExternalId,
        photoByExternalId,
      ),
    [personnelRows, photoByExternalId, questionnaireByExternalId],
  );

  const {
    isMergingAnketaData,
    isMergingVkTpvDovidky,
    importVkTpvDovidkyWorkbook,
    mergeMissingFieldsFromAnketaData,
    importPersonnelRosterWorkbook,
  } = usePersonnelWorkbookImports({
    reload: (options) => startPersonnelLoad(options),
    holdStatusUntilRef,
    setIsLoading,
    setMessage,
  });

  useEffect(() => {
    return subscribeDataCache(CacheKeys.personnelDataset, () => {
      if (!personnelScopeRef.current.all) return;
      const cached = peekDataCache<PersonnelDataset>(CacheKeys.personnelDataset);
      const preview = cached ? personnelDatasetToPreview(cached) : null;
      if (!preview?.rows?.length) return;
      if (
        cached!.fingerprint === personnelDatasetFingerprintRef.current &&
        personnelListIndexRef.current.records.length > 0
      ) {
        return;
      }
      setRosterLabels(cached!.rosterLabels);
      void applyPersonnelPreview(preview, {
        fromCache: true,
        datasetFingerprint: cached!.fingerprint,
      });
    });
  }, []);

  useEffect(() => {
    let mounted = true;
    void loadRosterStaff();
    const refreshAfterStaffSync = () => {
      if (!mounted) return;
      if (personnelScopeRef.current.all) {
        void startPersonnelLoad({ force: true });
        return;
      }
      void loadRosterStaff(undefined, { force: true });
    };
    window.addEventListener(STAFF_SHEET_SYNCED_EVENT, refreshAfterStaffSync);
    return () => {
      mounted = false;
      personnelLoadControllerRef.current?.abort();
      window.removeEventListener(
        STAFF_SHEET_SYNCED_EVENT,
        refreshAfterStaffSync,
      );
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    const handleOpenPersonnel = (event: Event) => {
      const detail = normalizePersonnelFocusTarget(
        (event as CustomEvent<PersonnelFocusTarget>).detail ?? {},
      );
      if (!detail.rowId && !detail.externalId && !detail.search) return;
      personnelFocusLockRef.current = detail;
      if (detail.search) listPanelRef.current?.setSearch(detail.search);
      const rows = personnelAllRowsRef.current.filter(isLikelyPersonnelRow);
      if (!rows.length) return;
      const focusedRow = findPersonnelRowByFocusTarget(rows, detail);
      if (!focusedRow?.__dbRowId) return;
      setSelectedRowId(focusedRow.__dbRowId);
      setMobilePane("card");
      setMessage(
        `Відкрито картку: ${getPersonDisplayName(focusedRow) || "особу"}.`,
      );
      personnelFocusLockRef.current = null;
      clearPersonnelFocusTarget();
    };

    window.addEventListener("army-grid:open-personnel", handleOpenPersonnel);
    return () =>
      window.removeEventListener(
        "army-grid:open-personnel",
        handleOpenPersonnel,
      );
  }, []);

  const {
    saveSelectedPerson,
    openPhotoCrop,
    savePersonPhoto,
    deleteSelectedPhoto,
    addSelectedPersonPhone,
    removeSelectedPersonPhone,
    copySelectedPersonPhone,
    deleteSelectedQuestionnaire,
    focusPersonByExternalId,
    openDiskPhotoCrop,
  } = bindPersonnelCardEdits({
    selectedRow,
    selectedSummary,
    selectedPhoto,
    editableFields,
    editValues,
    savedPhones,
    phoneDraft,
    phoneDocByExternalId,
    photoCropFile,
    questionnaire,
    personnelRows,
    sourceRowsRef: personnelSourceRowsRef,
    allRowsRef: personnelAllRowsRef,
    listIndexRef: personnelListIndexRef,
    requestedPhotoIdsRef,
    setIsLoading,
    setMessage,
    setPersonnelDataEpoch,
    setPhotoByExternalId,
    setPhotoIndexReady,
    setPhotoCropFile,
    setIsPhotoCropOpen,
    setIsDiskFloatingCrop,
    setPhonesByExternalId,
    setPhoneDocByExternalId,
    setPhoneDraft,
    setIsSavingPhone,
    setQuestionnaire,
    setSelectedRowId,
    setMobilePane,
    clearQuestionnairePresence,
    releaseQuestionnairePreviewCache,
    closeQuestionnairePreview,
  });


  return (
    <main className="main-panel personnel-page">
      <PersonnelPageToolbar
        canEdit={canEdit}
        isLoading={isLoading}
        isMergingAnketaData={isMergingAnketaData}
        isMergingVkTpvDovidky={isMergingVkTpvDovidky}
        photoExtractCount={questionnairePhotoExtractTargets.length}
        diskSearchCount={missingDiskSearchPeople.length}
        onMergeAnketa={() => void mergeMissingFieldsFromAnketaData()}
        onImportVkTpv={(file) => void importVkTpvDovidkyWorkbook(file)}
        onOpenPhotoExtract={() => setIsPhotoExtractOpen(true)}
        onOpenDiskSearch={() => setIsDiskSearchOpen(true)}
        onImportRoster={(file) => void importPersonnelRosterWorkbook(file)}
        onRefresh={() => void startPersonnelLoad({ force: true })}
      />
      {isLoading && personnelRows.length > 0 ? (
        <LinearProgress color="primary" />
      ) : isMergingAnketaData || isMergingVkTpvDovidky ? (
        <LinearProgress color="primary" />
      ) : null}
      <Alert
        severity="info"
        variant="outlined"
        className="personnel-page-alert"
      >
        {message}
      </Alert>

      <PersonnelMobileTabs
        pane={mobilePane}
        canOpenPerson={Boolean(selectedRowId)}
        onChange={setMobilePane}
      />

      <section className={`personnel-layout mobile-pane-${mobilePane}`}>
        <PersonnelListPanel
          ref={listPanelRef}
          initialSearch={personnelListInitialSearch}
          personnelRows={personnelRows}
          phonesByExternalId={phonesByExternalId}
          staffFilter={staffFilter}
          staffCounts={staffCounts}
          questionnaireByExternalId={questionnaireByExternalId}
          questionnairePresenceStatus={questionnairePresenceStatus}
          isLoading={isLoading}
      selectedRowId={selectedRowId}
      photoIndexReady={photoIndexReady}
      photoByExternalId={photoByExternalId}
          selectedPhotoFullUrl={
            isMobilePersonnelLayout && mobilePane === "list"
              ? ""
              : selectedPhotoFullUrl
          }
          onStaffFilterChange={(value) => {
            setStaffFilter(value);
            setSelectedRowId("");
            const change = staffFilterScopeChange(
              value,
              personnelSourceRowsRef.current,
            );
            if (change.all) personnelScopeRef.current.all = true;
            if (change.archive) personnelScopeRef.current.archive = true;
            if (change.reload) {
              void startPersonnelLoad();
              return;
            }
            if (change.publish) publishOpenPersonnelScope();
          }}
          onNeedPhotos={loadVisiblePersonnelPhotos}
          onPhotoLoadError={handleListPhotoLoadError}
          onSelect={(rowId) => {
            setSelectedRowId(rowId);
            setMobilePane("card");
          }}
          keyboardEnabled={
            !isPhotoCropOpen &&
            !isQuestionnairePreviewOpen &&
            !isPhotoLightboxOpen &&
            !isDiskSearchOpen &&
            !isPhotoExtractOpen
          }
        />

        <PersonnelPersonCard
          selectedPhoto={selectedPhoto}
          selectedSummary={selectedSummary}
          selectedRowId={selectedRowId}
          questionnaire={questionnaire}
          questionnaireByExternalId={questionnaireByExternalId}
          selectedRow={selectedRow}
          selectedCallSign={selectedCallSign}
          rosterStatus={rosterStatus}
          cardSummary={cardSummary}
          parsedPhones={parsedPhones}
          savedPhones={savedPhones}
          isSavingPhone={isSavingPhone}
          phoneDraft={phoneDraft}
          birthDateWithAge={birthDateWithAge}
          selectedFullPosition={selectedFullPosition}
          selectedPersonNote={selectedPersonNote}
          fighterStatusFieldRows={fighterStatusFieldRows}
          rosterFieldRows={rosterFieldRows}
          editableFieldsBySection={editableFieldsBySection}
          selectedAnketaRow={selectedAnketaRow}
          editValues={editValues}
          setIsPhotoLightboxOpen={setIsPhotoLightboxOpen}
          setMessage={setMessage}
          openQuestionnairePreview={openQuestionnairePreview}
          deleteSelectedPhoto={deleteSelectedPhoto}
          openPhotoCrop={openPhotoCrop}
          copySelectedPersonPhone={copySelectedPersonPhone}
          removeSelectedPersonPhone={removeSelectedPersonPhone}
          setPhoneDraft={setPhoneDraft}
          addSelectedPersonPhone={addSelectedPersonPhone}
          setEditValues={setEditValues}
          saveSelectedPerson={saveSelectedPerson}
        />
        <PersonnelDocumentsPanel
          selectedRow={selectedRow}
          externalId={selectedSummary.externalId}
          questionnaireExists={Boolean(questionnaire)}
          questionnaireFileName={questionnaireExportFileName}
          relatedDocuments={personRelatedDocuments}
          isUploadingQuestionnaire={isUploadingQuestionnaire}
          fullPosition={selectedFullPosition}
          onBackToCard={() => setMobilePane("card")}
          onBackToList={() => setMobilePane("list")}
          onAddQuestionnaire={beginQuestionnaireReview}
          onOpenQuestionnaire={() => void openQuestionnairePreview()}
          onExportQuestionnaire={() => void exportCurrentQuestionnaire()}
          onRevealQuestionnaire={() => void revealCurrentQuestionnaireInFinder()}
          onDeleteQuestionnaire={() => void deleteSelectedQuestionnaire()}
          onOpenDocument={onOpenDocuments}
        />
      </section>

      <PersonnelPageOverlays
        personName={selectedSummary.name}
        selectedPhoto={selectedPhoto}
        isPhotoLightboxOpen={isPhotoLightboxOpen}
        onClosePhoto={() => setIsPhotoLightboxOpen(false)}
        photoCropFile={photoCropFile}
        isPhotoCropOpen={isPhotoCropOpen}
        isDiskFloatingCrop={isDiskFloatingCrop}
        onClosePhotoCrop={() => {
          setIsPhotoCropOpen(false);
          setIsDiskFloatingCrop(false);
        }}
        onSavePhoto={savePersonPhoto}
        onMessage={setMessage}
        isPhotoExtractOpen={isPhotoExtractOpen}
        photoExtractPeople={questionnairePhotoExtractTargets}
        onClosePhotoExtract={() => setIsPhotoExtractOpen(false)}
        onPhotoExtractSaved={(externalId, photoData) => {
          setPhotoByExternalId((current) => ({
            ...current,
            [externalId]: photoData,
          }));
          notifyPersonnelAttachmentChanged(externalId, "photo");
          if (Date.now() < holdStatusUntilRef.current) return;
          setMessage(`Фото з анкети збережено: ${externalId}.`);
        }}
        isDiskSearchOpen={isDiskSearchOpen}
        diskSearchPeople={missingDiskSearchPeople}
        onCloseDiskSearch={() => setIsDiskSearchOpen(false)}
        onDiskSearchConfirmed={(externalId) => {
          markQuestionnaireInDb(externalId);
          focusPersonByExternalId(externalId);
          notifyPersonnelAttachmentChanged(externalId, "questionnaire");
          void api
            .getPersonQuestionnaire(externalId)
            .then((next) => {
              setQuestionnaire(next);
            })
            .catch(() => undefined);
          setMessage(`Анкету підтверджено та збережено для ID ${externalId}.`);
        }}
        onDiskSearchPhotoSaved={(externalId, photoData) => {
          setPhotoByExternalId((current) => ({
            ...current,
            [externalId]: photoData,
          }));
          if (Date.now() < holdStatusUntilRef.current) return;
          setMessage(
            `Фото автоматично знайдено в PDF і додано до preview для ID ${externalId}.`,
          );
        }}
        onDiskSearchPreview={(file, title, externalId) => {
          focusPersonByExternalId(externalId);
          openDiskQuestionnairePreview(file, title);
        }}
        onDiskSearchCrop={openDiskPhotoCrop}
        isQuestionnairePreviewOpen={isQuestionnairePreviewOpen}
        questionnairePreviewTitle={questionnairePreviewTitle}
        questionnairePreviewUrl={questionnairePreviewUrl}
        pendingQuestionnaireFile={Boolean(pendingQuestionnaireFile)}
        isUploadingQuestionnaire={isUploadingQuestionnaire}
        isLoadingQuestionnairePreview={isLoadingQuestionnairePreview}
        isDiskFloatingPreview={isDiskFloatingPreview}
        questionnaireCropFile={questionnaireCropFileSource}
        onCloseQuestionnairePreview={closeQuestionnairePreview}
        onOpenQuestionnaireTab={() => void openQuestionnaireInNewTab()}
        onDownloadQuestionnaire={() => void exportCurrentQuestionnaire()}
        onSaveQuestionnaire={() => void confirmPendingQuestionnaire()}
        questionnaireExportFileName={questionnaireExportFileName}
        questionnaireShareSource={currentQuestionnaireShareSource}
      />
    </main>
  );
}
