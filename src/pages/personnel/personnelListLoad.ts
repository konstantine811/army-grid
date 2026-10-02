import { bootstrapPersonnelAppData } from "../../data/personnelBootstrap";
import {
  personnelDatasetToPreview,
  type PersonnelDataset,
} from "../../data/personnelDataset";
import type { DbPreviewState, EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  clearPersonnelFocusTarget,
  findPersonnelRowByFocusTarget,
  readPersonnelFocusTarget,
  type PersonnelFocusTarget,
} from "./personnelFocus";
import { buildPersonnelListIndex, type PersonnelListIndex } from "./personnelListIndex";
import {
  filterPersonnelScopeRows,
  withClosedScopeCounts,
} from "./personnelListScope";
import { buildQuestionnairePresencePeople } from "./personAttachments";
import { getPersonDisplayName, isLikelyPersonnelRow } from "./personnelUtils";
import { personnelListStatusMessage } from "./personnelQuestionnaireOpen";
import {
  loadStaffScopePreview,
  staffQuestionnaireRows,
} from "./loadStaffScopePreview";

export const personnelDatasetLoadingMessage =
  "Завантажую актуальну Штатку та готую список особового складу…";

export const personnelDatasetLoadError = (error: unknown) =>
  error instanceof Error
    ? error.message
    : "Не вдалося завантажити особовий склад.";

export const personnelPreviewError = (error: unknown) =>
  error instanceof Error
    ? error.message
    : "Не вдалося підготувати список особового складу.";

export const staffSheetLoadingMessage = "Завантажую Штатку…";

export const emptyStaffSheetMessage = "У штатці немає рядків.";

export const staffSheetLoadError = (error: unknown) =>
  error instanceof Error ? error.message : "Не вдалося завантажити штатку.";

export const nextPersonnelRowId = (
  current: string,
  focusedId: string | undefined,
  rows: EjournalPreviewRow[],
) => {
  if (focusedId) return focusedId;
  if (current && rows.some((row) => row.__dbRowId === current)) return current;
  return rows[0]?.__dbRowId ?? "";
};

export const publishedPersonnelScope = (
  source: EjournalPreviewRow[],
  scope: { all: boolean; archive: boolean },
) => {
  const visibleRows = filterPersonnelScopeRows(source, scope);
  return {
    visibleRows,
    index: withClosedScopeCounts(buildPersonnelListIndex(visibleRows), scope),
  };
};

const scheduleAttachmentHeal = (task: () => void) => {
  if ("requestIdleCallback" in window) {
    window.requestIdleCallback(task, { timeout: 10_000 });
    return;
  }
  globalThis.setTimeout(task, 3_000);
};

type PreviewRefs = {
  focusLockRef: { current: PersonnelFocusTarget | null };
  fingerprintRef: { current: string };
  sourceRowsRef: { current: EjournalPreviewRow[] };
  scopeRef: { current: { all: boolean; archive: boolean } };
  allRowsRef: { current: EjournalPreviewRow[] };
  listIndexRef: { current: PersonnelListIndex };
  presencePeopleRef: {
    current: ReturnType<typeof buildQuestionnairePresencePeople>;
  };
};

export const applyPersonnelListPreview = async ({
  preview,
  options,
  refs,
  resetPhotoRequests,
  setSelectedRowId,
  setMobilePane,
  setPersonnelDataEpoch,
  setMessage,
}: {
  preview: DbPreviewState;
  options?: {
    fromCache?: boolean;
    isCancelled?: () => boolean;
    datasetFingerprint?: string;
  };
  refs: PreviewRefs;
  resetPhotoRequests: () => void;
  setSelectedRowId: (value: string | ((current: string) => string)) => void;
  setMobilePane: (value: "list" | "card" | "side") => void;
  setPersonnelDataEpoch: (update: (value: number) => number) => void;
  setMessage: (value: string) => void;
}) => {
  const safePreview = Array.isArray(preview.rows)
    ? preview
    : { ...preview, rows: [] };
  if (options?.isCancelled?.()) return safePreview;

  try {
    const fingerprint = options?.datasetFingerprint?.trim() ?? "";
    const rows = safePreview.rows.filter(isLikelyPersonnelRow);
    const sheetName = safePreview.sheet?.name ?? "ООС";
    const storedFocus = readPersonnelFocusTarget();
    if (storedFocus.rowId || storedFocus.externalId || storedFocus.search) {
      refs.focusLockRef.current = storedFocus;
    }
    const focusTarget = refs.focusLockRef.current ?? {
      rowId: "",
      externalId: "",
      search: "",
    };

    const sameDataset =
      Boolean(fingerprint) &&
      fingerprint === refs.fingerprintRef.current &&
      refs.listIndexRef.current.records.length > 0;

    if (sameDataset) {
      const indexedRows = refs.listIndexRef.current.records.map(
        (record) => record.row,
      );
      const focusedRow = findPersonnelRowByFocusTarget(indexedRows, focusTarget);
      if (focusedRow?.__dbRowId) {
        setSelectedRowId(focusedRow.__dbRowId);
        setMobilePane("card");
        refs.focusLockRef.current = null;
        clearPersonnelFocusTarget();
      }
      return safePreview;
    }

    if (fingerprint && fingerprint !== refs.fingerprintRef.current) {
      resetPhotoRequests();
      refs.fingerprintRef.current = fingerprint;
    }
    const focusedRow = findPersonnelRowByFocusTarget(rows, focusTarget);
    refs.sourceRowsRef.current = safePreview.rows;
    const visibleRows = filterPersonnelScopeRows(
      safePreview.rows,
      refs.scopeRef.current,
    );
    const index = withClosedScopeCounts(
      buildPersonnelListIndex(visibleRows),
      refs.scopeRef.current,
    );
    refs.allRowsRef.current = visibleRows;
    refs.listIndexRef.current = index;
    refs.presencePeopleRef.current = buildQuestionnairePresencePeople(
      visibleRows.filter(isLikelyPersonnelRow),
    );
    setPersonnelDataEpoch((value) => value + 1);
    setSelectedRowId((current) =>
      nextPersonnelRowId(current, focusedRow?.__dbRowId, rows),
    );
    if (focusedRow?.__dbRowId) setMobilePane("card");
    if (focusedRow) {
      refs.focusLockRef.current = null;
      clearPersonnelFocusTarget();
    }

    setMessage(
      personnelListStatusMessage({
        focusedName: focusedRow
          ? getPersonDisplayName(focusedRow) || "особу"
          : "",
        fromCache: options?.fromCache,
        rowCount: rows.length,
        sheetName,
      }),
    );

    return safePreview;
  } catch (error) {
    console.error("[PersonnelPage] applyPersonnelPreview failed", error);
    setMessage(personnelPreviewError(error));
    return safePreview;
  }
};

export const loadFullPersonnelList = async ({
  signal,
  force,
  generationRef,
  allRowsRef,
  setIsLoading,
  setMessage,
  setRosterLabels,
  applyPreview,
  loadQuestionnaireIds,
  heal,
}: {
  signal?: AbortSignal;
  force?: boolean;
  generationRef: { current: number };
  allRowsRef: { current: EjournalPreviewRow[] };
  setIsLoading: (value: boolean) => void;
  setMessage: (value: string) => void;
  setRosterLabels: (labels: Record<string, string>) => void;
  applyPreview: (
    preview: DbPreviewState,
    options?: {
      fromCache?: boolean;
      isCancelled?: () => boolean;
      datasetFingerprint?: string;
    },
  ) => Promise<unknown>;
  loadQuestionnaireIds: (
    rows: EjournalPreviewRow[],
    signal: AbortSignal | undefined,
    prefetched: Promise<
      Array<{ personExternalId: string; fileName?: string | null }>
    >,
    options: { force?: boolean },
  ) => Promise<unknown>;
  heal: (
    rows: EjournalPreviewRow[],
    isCancelled: () => boolean,
    photos: Promise<Array<{ personExternalId: string; photoData: string }>>,
    questionnaires: Promise<Array<{ personExternalId: string }>>,
  ) => void;
}) => {
  const loadGeneration = ++generationRef.current;
  const isLoadCancelled = () =>
    loadGeneration !== generationRef.current || Boolean(signal?.aborted);
  const hasPaintedRows = allRowsRef.current.length > 0;
  if (!hasPaintedRows) {
    setIsLoading(true);
    setMessage(personnelDatasetLoadingMessage);
  }
  let cachedFingerprint: string | undefined;
  let paintedFromCache = false;
  let questionnaireItemsPromise: Promise<
    Array<{ personExternalId: string; fileName?: string | null }>
  > = Promise.resolve([]);
  const paintDataset = async (dataset: PersonnelDataset, fromCache: boolean) => {
    const preview = personnelDatasetToPreview(dataset);
    if (!preview || isLoadCancelled()) return;
    setRosterLabels(dataset.rosterLabels);
    await applyPreview(preview, {
      fromCache,
      isCancelled: isLoadCancelled,
      datasetFingerprint: dataset.fingerprint,
    });
  };
  try {
    const bootstrap = await bootstrapPersonnelAppData({
      force,
      signal,
      photoIndex: true,
      questionnaires: true,
      onCached: async (cached) => {
        paintedFromCache = true;
        cachedFingerprint = cached.fingerprint;
        await paintDataset(cached, true);
      },
    });
    const dataset = bootstrap.dataset;
    questionnaireItemsPromise =
      bootstrap.questionnairesPromise ??
      Promise.resolve(bootstrap.questionnaires ?? []);
    if (isLoadCancelled()) return;
    const needsFreshPaint =
      !paintedFromCache || force || cachedFingerprint !== dataset.fingerprint;
    if (needsFreshPaint) await paintDataset(dataset, false);
    if (isLoadCancelled()) return;

    const startAttachments = () => {
      if (isLoadCancelled()) return;
      const questionnairesPromise = loadQuestionnaireIds(
        dataset.rows,
        signal,
        questionnaireItemsPromise,
        { force },
      );
      const healAttachments = () => {
        if (isLoadCancelled()) return;
        heal(
          dataset.rows,
          isLoadCancelled,
          Promise.resolve([]),
          questionnairesPromise as Promise<Array<{ personExternalId: string }>>,
        );
      };
      scheduleAttachmentHeal(healAttachments);
    };
    startAttachments();
  } catch (error) {
    if (!isLoadCancelled()) setMessage(personnelDatasetLoadError(error));
  } finally {
    if (!isLoadCancelled()) setIsLoading(false);
  }
};

export const loadStaffPersonnelList = async ({
  signal,
  force,
  hasPaintedRows,
  setIsLoading,
  setMessage,
  applyPreview,
  setPhotoIndexReady,
  loadQuestionnaireIds,
}: {
  signal?: AbortSignal;
  force?: boolean;
  hasPaintedRows: boolean;
  setIsLoading: (value: boolean) => void;
  setMessage: (value: string) => void;
  applyPreview: (preview: DbPreviewState) => Promise<unknown>;
  setPhotoIndexReady: (update: (value: number) => number) => void;
  loadQuestionnaireIds: (
    rows: EjournalPreviewRow[],
    signal?: AbortSignal,
  ) => Promise<unknown>;
}) => {
  if (!hasPaintedRows) {
    setIsLoading(true);
    setMessage(staffSheetLoadingMessage);
  }
  try {
    const loaded = await loadStaffScopePreview({ force, signal });
    if (loaded.aborted || signal?.aborted) return;
    if (!loaded.preview?.rows.length) {
      setMessage(emptyStaffSheetMessage);
      return;
    }
    await applyPreview(loaded.preview);
    setPhotoIndexReady((value) => (value > 0 ? value : 1));
    void loadQuestionnaireIds(staffQuestionnaireRows(loaded.preview.rows), signal);
  } catch (error) {
    if (!signal?.aborted) setMessage(staffSheetLoadError(error));
  } finally {
    if (!signal?.aborted) setIsLoading(false);
  }
};
