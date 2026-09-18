import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { api, type BackendPersonQuestionnaireMeta } from "../../../api";
import {
  ANKETA_COLUMNS,
  isAnketaColumnReadonly,
  type AnketaColumnKey,
  type AnketaRow,
  type AnketaSheetSnapshot,
} from "../anketaSheet";
import {
  bulkWriteAnketaCellEdits,
  upsertAnketaCellEdit,
  countAnketaEdits,
} from "../anketaEdits";
import {
  ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
  anketaCellA1,
  anketaGapSkipKey,
  anketaRowHasQuestionnairePdf,
  buildAnketaRowsWithQuestionnairePdfSet,
  applyAbsentQuestionnaireClearsToRows,
  applyAbsentQuestionnaireFillsToRows,
  buildAbsentQuestionnaireAnketaRows,
  collectAbsentQuestionnaireCellClears,
  collectAbsentQuestionnaireCellFills,
  countAnketaEmptyCells,
  filterAnketaRowsForGapSearch,
  gateAnketaGapSearchRowsInStaffFirst,
  orderAnketaGapSearchRows,
  findNextAnketaEmptyCell,
  findNextAnketaPersonEmptyCell,
  isAnketaGapCellTrulyEmpty,
  isAnketaGapPersonWithoutTrulyEmptyCells,
  listAnketaPersonGapSkipKeys,
  removePresentQuestionnairesFromMissingNameKeys,
  type AnketaGapSearchOptions,
  pushAnketaCellToGoogle,
  summarizeAnketaGaps,
  updateAnketaRowCell,
  type AnketaEmptyCell,
} from "../anketaGaps";
import { loadAnketaMissingNames } from "../anketaMissingList";
import {
  buildAnketaInStaffRowIdSet,
  expandAnketaNameKeySet,
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnel,
  type AnketaPersonnelIndex,
} from "../anketaPersonMatch";
import {
  loadPersonPhotoForRow,
  loadPersonQuestionnaireForRow,
} from "../../personnel/personAttachments";
import {
  buildAnketaGapReviewEntry,
  countActiveAnketaGapReviews,
  filterSkippedAnketaGapRows,
  loadAnketaGapReviewProgress,
  mergeAnketaGapReviewEntry,
  saveAnketaGapReviewProgress,
  type AnketaGapReviewMap,
} from "../anketaGapProgress";

type UseAnketaGapSearchOptions = {
  rows: AnketaRow[];
  gapColumnKeys: AnketaColumnKey[];
  missingQuestionnaireNames: Set<string>;
  setMissingQuestionnaireNames: (keys: Set<string>) => void;
  appsScriptUrl: string;
  canEdit?: boolean;
  persistSnapshot: (
    updater: (current: AnketaSheetSnapshot) => AnketaSheetSnapshot,
    note?: string,
  ) => void;
  setMessage: (message: string) => void;
  setIsSyncing: (value: boolean) => void;
  setEditsCount: (value: number) => void;
  setGapColumnsOpen: (value: boolean) => void;
  gapColumnsOpen: boolean;
};

const prefetchNextGapPerson = (
  rows: AnketaRow[],
  current: AnketaEmptyCell,
  columnKeys: AnketaColumnKey[],
  options: AnketaGapSearchOptions,
) => {
  const next =
    findNextAnketaPersonEmptyCell(rows, current, columnKeys, options) ??
    findNextAnketaEmptyCell(rows, current, columnKeys, options);
  if (!next || next.rowId === current.rowId) return;
  const row = rows.find((item) => item.__rowId === next.rowId);
  if (!row) return;
  void loadPersonnelIndexForAnketa()
    .then((index) => {
      const match = matchAnketaRowToPersonnel(row, index);
      if (!match) return;
      const hints = {
        anketaExternalId: row.externalId,
        anketaFullName: row.fullName,
        anketaBirthDate: row.birthDate,
      };
      return Promise.all([
        loadPersonQuestionnaireForRow(match.row, hints),
        loadPersonPhotoForRow(match.row, hints),
      ]);
    })
    .catch(() => {
      /* prefetch must not block search */
    });
};

export function useAnketaGapSearch({
  rows,
  gapColumnKeys,
  missingQuestionnaireNames,
  setMissingQuestionnaireNames,
  appsScriptUrl,
  canEdit = true,
  persistSnapshot,
  setMessage,
  setIsSyncing,
  setEditsCount,
  setGapColumnsOpen,
  gapColumnsOpen,
}: UseAnketaGapSearchOptions) {
  const [focusedEmpty, setFocusedEmpty] = useState<AnketaEmptyCell | null>(
    null,
  );
  const [focusEpoch, setFocusEpoch] = useState(0);
  const [deferredGapKeys, setDeferredGapKeys] = useState<string[]>([]);
  const [personPanelOpen, setPersonPanelOpen] = useState(false);
  const [emptySearchActive, setEmptySearchActive] = useState(false);
  const [isFillingAbsent, setIsFillingAbsent] = useState(false);
  const [personnelIndex, setPersonnelIndex] = useState<AnketaPersonnelIndex | null>(
    null,
  );
  const [questionnaireMeta, setQuestionnaireMeta] = useState<
    BackendPersonQuestionnaireMeta[] | null
  >(null);
  const [gapReviewProgress, setGapReviewProgress] = useState<AnketaGapReviewMap>(
    {},
  );
  const gapReviewProgressRef = useRef<AnketaGapReviewMap>({});
  const gapReviewProgressLoadedRef = useRef(false);
  const suppressCellBlurSaveRef = useRef(false);

  useEffect(() => {
    let cancelled = false;
    void loadPersonnelIndexForAnketa()
      .then((index) => {
        if (!cancelled) setPersonnelIndex(index);
      })
      .catch(() => {
        if (!cancelled) setPersonnelIndex(null);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    void api
      .listPersonQuestionnaires()
      .then((items) => {
        if (!cancelled) setQuestionnaireMeta(items);
      })
      .catch(() => {
        if (!cancelled) setQuestionnaireMeta(null);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    let cancelled = false;
    void loadAnketaGapReviewProgress()
      .then((progress) => {
        if (cancelled) return;
        gapReviewProgressLoadedRef.current = true;
        gapReviewProgressRef.current = progress;
        setGapReviewProgress(progress);
      })
      .catch(() => {
        if (cancelled) return;
        gapReviewProgressLoadedRef.current = true;
        gapReviewProgressRef.current = {};
        setGapReviewProgress({});
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    gapReviewProgressRef.current = gapReviewProgress;
  }, [gapReviewProgress]);

  const markGapPersonReviewed = useCallback(
    (row: AnketaRow | null | undefined) => {
      if (!row || !gapColumnKeys.length) return;
      const entry = buildAnketaGapReviewEntry(
        row,
        gapColumnKeys,
        questionnaireMeta,
      );
      if (!entry) return;
      setGapReviewProgress((current) => {
        const next = mergeAnketaGapReviewEntry(current, entry);
        void saveAnketaGapReviewProgress(next);
        return next;
      });
    },
    [gapColumnKeys, questionnaireMeta],
  );

  const ensureGapReviewProgress = useCallback(async () => {
    if (gapReviewProgressLoadedRef.current) {
      return gapReviewProgressRef.current;
    }
    try {
      const progress = await loadAnketaGapReviewProgress();
      gapReviewProgressLoadedRef.current = true;
      gapReviewProgressRef.current = progress;
      setGapReviewProgress(progress);
      return progress;
    } catch {
      gapReviewProgressLoadedRef.current = true;
      gapReviewProgressRef.current = {};
      setGapReviewProgress({});
      return {};
    }
  }, []);

  const buildOrderedGapSearchRows = useCallback(
    (
      sourceRows: AnketaRow[],
      indexOverride?: AnketaPersonnelIndex | null,
      progressOverride?: AnketaGapReviewMap,
    ) => {
      const index = indexOverride ?? personnelIndex;
      const progress = progressOverride ?? gapReviewProgress;
      const inStaffRowIds = buildAnketaInStaffRowIdSet(sourceRows, index);
      const withPdfRowIds = buildAnketaRowsWithQuestionnairePdfSet(
        sourceRows,
        questionnaireMeta,
      );
      const ordered = orderAnketaGapSearchRows(
        sourceRows,
        inStaffRowIds,
        gapColumnKeys,
        withPdfRowIds,
      );
      const eligible = filterAnketaRowsForGapSearch(ordered);
      const afterReview = filterSkippedAnketaGapRows(
        eligible,
        gapColumnKeys,
        questionnaireMeta,
        progress,
      );
      return gateAnketaGapSearchRowsInStaffFirst(
        afterReview,
        inStaffRowIds,
        gapColumnKeys,
      );
    },
    [
      personnelIndex,
      questionnaireMeta,
      gapColumnKeys,
      gapReviewProgress,
    ],
  );

  const ensurePersonnelIndex = useCallback(async () => {
    if (personnelIndex) return personnelIndex;
    setMessage("Завантажую Штатку для пошуку порожніх…");
    const index = await loadPersonnelIndexForAnketa({ force: true });
    setPersonnelIndex(index);
    return index;
  }, [personnelIndex, setMessage]);

  const gapSearchRows = useMemo(
    () => buildOrderedGapSearchRows(rows),
    [rows, buildOrderedGapSearchRows],
  );

  const deferredGapKeySet = useMemo(
    () => new Set(deferredGapKeys),
    [deferredGapKeys],
  );
  const gapSearchOptions = useMemo(
    () => ({
      skipKeys: deferredGapKeySet,
    }),
    [deferredGapKeySet],
  );
  const emptyCount = useMemo(
    () => countAnketaEmptyCells(gapSearchRows, gapColumnKeys, gapSearchOptions),
    [gapSearchRows, gapColumnKeys, gapSearchOptions],
  );
  const gapStats = useMemo(
    () => summarizeAnketaGaps(rows, gapColumnKeys),
    [rows, gapColumnKeys],
  );
  const gapReviewedCount = useMemo(() => {
    const ordered = orderAnketaGapSearchRows(
      rows,
      buildAnketaInStaffRowIdSet(rows, personnelIndex),
      gapColumnKeys,
      buildAnketaRowsWithQuestionnairePdfSet(rows, questionnaireMeta),
    );
    const eligible = filterAnketaRowsForGapSearch(ordered);
    return countActiveAnketaGapReviews(
      eligible,
      gapColumnKeys,
      questionnaireMeta,
      gapReviewProgress,
    );
  }, [
    rows,
    personnelIndex,
    questionnaireMeta,
    gapColumnKeys,
    gapReviewProgress,
  ]);
  const focusedAnketaRow = useMemo(() => {
    if (!focusedEmpty) return null;
    return rows.find((row) => row.__rowId === focusedEmpty.rowId) ?? null;
  }, [focusedEmpty, rows]);

  useEffect(() => {
    if (focusedEmpty) setPersonPanelOpen(true);
  }, [focusedEmpty]);

  const clearGapFocus = () => {
    setEmptySearchActive(false);
    setFocusedEmpty(null);
    setPersonPanelOpen(false);
    setDeferredGapKeys([]);
  };

  const stopEmptySearch = () => {
    suppressCellBlurSaveRef.current = true;
    setEmptySearchActive(false);
    setFocusedEmpty(null);
    setPersonPanelOpen(false);
    setGapColumnsOpen(false);
    setMessage("Пошук порожніх зупинено · Esc.");
  };

  const activateEmptyCell = (cell: AnketaEmptyCell | null) => {
    if (cell && isAnketaColumnReadonly(cell.columnId)) return;
    queueMicrotask(() => {
      setFocusedEmpty(cell);
      if (cell) setFocusEpoch((epoch) => epoch + 1);
    });
  };

  const cancelCellEdit = () => {
    suppressCellBlurSaveRef.current = true;
    setFocusedEmpty(null);
  };

  useEffect(() => {
    if (!focusedEmpty) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (gapColumnsOpen) return;
      event.preventDefault();
      if (emptySearchActive) stopEmptySearch();
      else cancelCellEdit();
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [focusedEmpty, gapColumnsOpen, emptySearchActive]);

  const patchCell = async (
    rowId: string,
    columnId: AnketaColumnKey,
    value: string,
    options?: { advance?: boolean },
  ) => {
    if (!canEdit) {
      setMessage("Немає права редагувати анкетні дані.");
      return;
    }
    if (isAnketaColumnReadonly(columnId)) return;
    const columnIndex = ANKETA_COLUMNS.findIndex(
      (column) => column.key === columnId,
    );
    const currentRow = rows.find((item) => item.__rowId === rowId);
    const rowNumber = currentRow?.__rowNumber ?? focusedEmpty?.rowNumber ?? 0;

    let nextRows: AnketaRow[] = rows;
    persistSnapshot((current) => {
      nextRows = updateAnketaRowCell(current.rows, rowId, columnId, value);
      return { ...current, rows: nextRows };
    });

    let serverSynced = true;
    let serverError: string | undefined;
    try {
      const result = await upsertAnketaCellEdit({
        rowNumber,
        columnId,
        value,
        externalId: currentRow?.externalId,
        fullName: currentRow?.fullName,
      });
      setEditsCount(countAnketaEdits(result.edits));
      serverSynced = result.serverSynced;
      serverError = result.serverError;
    } catch {
      setMessage("Не вдалося записати правку в локальну БД (IndexedDB).");
      return;
    }

    const editedGapKey = `${rowId}:${columnId}`;
    const skipKeysAfterSave = deferredGapKeys.filter(
      (key) => key !== editedGapKey,
    );
    setDeferredGapKeys(skipKeysAfterSave);

    const row = nextRows.find((item) => item.__rowId === rowId);
    const cellA1 =
      row && columnIndex >= 0
        ? anketaCellA1(row.__rowNumber, columnIndex)
        : focusedEmpty?.a1;

    const saveNote = serverSynced
      ? "Збережено в серверну БД"
      : `Локально: ок. Сервер: ${serverError ?? "недоступний"}`;

    if (appsScriptUrl && cellA1) {
      setIsSyncing(true);
      try {
        await pushAnketaCellToGoogle({ a1: cellA1, value });
        setMessage(`${saveNote} і Google · ${cellA1}.`);
      } catch (error) {
        setMessage(
          error instanceof Error
            ? `${saveNote}. Google sync: ${error.message}`
            : `${saveNote}. Google sync не вдався.`,
        );
      } finally {
        setIsSyncing(false);
      }
    } else {
      setMessage(
        cellA1
          ? `${saveNote} · ${cellA1}. Для Google — Apps Script URL.`
          : saveNote,
      );
    }

    const savedRow = nextRows.find((item) => item.__rowId === rowId);
    if (
      savedRow &&
      isAnketaGapPersonWithoutTrulyEmptyCells(savedRow, gapColumnKeys)
    ) {
      markGapPersonReviewed(savedRow);
    }

    const shouldAdvance = Boolean(options?.advance);
    if (shouldAdvance) {
      const currentGap: AnketaEmptyCell = {
        rowId,
        columnId,
        rowNumber: row?.__rowNumber ?? rowNumber,
        columnIndex: Math.max(columnIndex, 0),
        header: ANKETA_COLUMNS[columnIndex]?.header ?? columnId,
        a1: cellA1 ?? "",
      };
      const orderedNextRows = buildOrderedGapSearchRows(nextRows);
      const next = findNextAnketaEmptyCell(
        orderedNextRows,
        currentGap,
        gapColumnKeys,
        {
          skipKeys: skipKeysAfterSave,
        },
      );
      setFocusedEmpty(next);
      if (next) {
        setFocusEpoch((epoch) => epoch + 1);
        setMessage(
          `Збережено. Наступна порожня: ${next.a1} · ${next.header}. Залишилось: ${countAnketaEmptyCells(orderedNextRows, gapColumnKeys, {
            skipKeys: skipKeysAfterSave,
          })}.`,
        );
        prefetchNextGapPerson(orderedNextRows, next, gapColumnKeys, {
          skipKeys: skipKeysAfterSave,
        });
      } else if (value.trim()) {
        setMessage("Усі порожні комірки (у вибраних колонках) заповнені.");
      }
    } else {
      setFocusedEmpty(null);
    }
  };

  const patchCells = async (
    rowId: string,
    items: Array<{ columnId: AnketaColumnKey; value: string }>,
    options?: { advance?: boolean },
  ) => {
    if (!canEdit) {
      setMessage("Немає права редагувати анкетні дані.");
      return;
    }
    const writes = items
      .map((item) => ({
        columnId: item.columnId,
        value: String(item.value ?? "").trim(),
      }))
      .filter((item) => item.value && !isAnketaColumnReadonly(item.columnId));
    if (!writes.length) {
      setMessage("Немає полів для запису.");
      return;
    }

    const currentRow = rows.find((item) => item.__rowId === rowId);
    const rowNumber = currentRow?.__rowNumber ?? focusedEmpty?.rowNumber ?? 0;

    let nextRows: AnketaRow[] = rows;
    persistSnapshot((current) => {
      let updatedRows = current.rows;
      for (const item of writes) {
        updatedRows = updateAnketaRowCell(
          updatedRows,
          rowId,
          item.columnId,
          item.value,
        );
      }
      nextRows = updatedRows;
      return { ...current, rows: updatedRows };
    });

    let serverSynced = true;
    let serverError: string | undefined;
    try {
      const result = await bulkWriteAnketaCellEdits(
        writes.map((item) => ({
          rowNumber,
          columnId: item.columnId,
          value: item.value,
          externalId: currentRow?.externalId,
          fullName: currentRow?.fullName,
        })),
      );
      setEditsCount(countAnketaEdits(result.edits));
      serverSynced = result.serverSynced;
      serverError = result.serverError;
    } catch {
      setMessage("Не вдалося записати правки в локальну БД (IndexedDB).");
      return;
    }

    const editedGapKeys = new Set(
      writes.map((item) => `${rowId}:${item.columnId}`),
    );
    const skipKeysAfterSave = deferredGapKeys.filter(
      (key) => !editedGapKeys.has(key),
    );
    setDeferredGapKeys(skipKeysAfterSave);

    const saveNote = serverSynced
      ? "Збережено в серверну БД"
      : `Локально: ок. Сервер: ${serverError ?? "недоступний"}`;

    if (appsScriptUrl) {
      setIsSyncing(true);
      try {
        const row = nextRows.find((item) => item.__rowId === rowId);
        let synced = 0;
        for (const item of writes) {
          const columnIndex = ANKETA_COLUMNS.findIndex(
            (column) => column.key === item.columnId,
          );
          if (!row || columnIndex < 0) continue;
          const a1 = anketaCellA1(row.__rowNumber, columnIndex);
          await pushAnketaCellToGoogle({ a1, value: item.value });
          synced += 1;
        }
        setMessage(`${saveNote} і Google · ${synced} комірок.`);
      } catch (error) {
        setMessage(
          error instanceof Error
            ? `${saveNote}. Google sync: ${error.message}`
            : `${saveNote}. Google sync не вдався.`,
        );
      } finally {
        setIsSyncing(false);
      }
    } else {
      setMessage(`${saveNote} · ${writes.length} комірок.`);
    }

    const savedRow = nextRows.find((item) => item.__rowId === rowId);
    if (
      savedRow &&
      isAnketaGapPersonWithoutTrulyEmptyCells(savedRow, gapColumnKeys)
    ) {
      markGapPersonReviewed(savedRow);
    }

    if (!options?.advance) return;

    const lastWrite = writes[writes.length - 1]!;
    const lastColumnIndex = ANKETA_COLUMNS.findIndex(
      (column) => column.key === lastWrite.columnId,
    );
    const row = nextRows.find((item) => item.__rowId === rowId);
    const currentGap: AnketaEmptyCell = {
      rowId,
      columnId: lastWrite.columnId,
      rowNumber: row?.__rowNumber ?? rowNumber,
      columnIndex: Math.max(lastColumnIndex, 0),
      header: ANKETA_COLUMNS[lastColumnIndex]?.header ?? lastWrite.columnId,
      a1:
        row && lastColumnIndex >= 0
          ? anketaCellA1(row.__rowNumber, lastColumnIndex)
          : focusedEmpty?.a1 ?? "",
    };
    const orderedNextRows = buildOrderedGapSearchRows(nextRows);
    const next = findNextAnketaEmptyCell(
      orderedNextRows,
      currentGap,
      gapColumnKeys,
      {
        skipKeys: skipKeysAfterSave,
      },
    );
    setFocusedEmpty(next);
    if (next) {
      setFocusEpoch((epoch) => epoch + 1);
      setMessage(
        `AI збережено. Наступна порожня: ${next.a1} · ${next.header}. Залишилось: ${countAnketaEmptyCells(orderedNextRows, gapColumnKeys, {
          skipKeys: skipKeysAfterSave,
        })}.`,
      );
      prefetchNextGapPerson(orderedNextRows, next, gapColumnKeys, {
        skipKeys: skipKeysAfterSave,
      });
    } else {
      setMessage("Усі порожні комірки (у вибраних колонках) заповнені.");
    }
  };

  const goToEmptyCell = async (direction: "first" | "next" | "nextPerson") => {
    if (!rows.length) {
      setMessage("Спочатку завантажте анкетні дані.");
      return;
    }
    if (!gapColumnKeys.length) {
      setMessage("Оберіть хоча б одну колонку для пошуку пропусків.");
      setGapColumnsOpen(true);
      return;
    }

    setMessage("Шукаю порожні комірки…");
    if (!personnelIndex) {
      void ensurePersonnelIndex().catch(() => {
        /* штатка лише для порядку, пошук без неї теж працює */
      });
    }

    try {
      const progressForSearch = await ensureGapReviewProgress();
      const searchRows = buildOrderedGapSearchRows(
        rows,
        personnelIndex,
        progressForSearch,
      );
      const inStaffCount = buildAnketaInStaffRowIdSet(
        rows,
        personnelIndex,
      ).size;

      let skipKeys = deferredGapKeys;
      if (direction === "first") {
        skipKeys = [];
        setDeferredGapKeys([]);
      } else if (direction === "nextPerson" && focusedEmpty) {
        markGapPersonReviewed(
          rows.find((row) => row.__rowId === focusedEmpty.rowId),
        );
        const personKeys = listAnketaPersonGapSkipKeys(
          rows,
          focusedEmpty.rowId,
          gapColumnKeys,
          gapSearchOptions,
        );
        const merged = new Set([...deferredGapKeys, ...personKeys]);
        skipKeys = [...merged];
        setDeferredGapKeys(skipKeys);
      } else if (focusedEmpty) {
        const focusedRow = rows.find((row) => row.__rowId === focusedEmpty.rowId);
        const stillEmpty = Boolean(
          focusedRow &&
            isAnketaGapCellTrulyEmpty(focusedRow, focusedEmpty.columnId),
        );
        if (stillEmpty) {
          const key = anketaGapSkipKey(focusedEmpty);
          if (!deferredGapKeySet.has(key)) {
            skipKeys = [...deferredGapKeys, key];
            setDeferredGapKeys(skipKeys);
          }
        }
      }

      const nextSearchOptions = {
        skipKeys,
      };
      const next =
        direction === "nextPerson"
          ? findNextAnketaPersonEmptyCell(
              searchRows,
              focusedEmpty,
              gapColumnKeys,
              nextSearchOptions,
            )
          : findNextAnketaEmptyCell(
              searchRows,
              direction === "first" ? null : focusedEmpty,
              gapColumnKeys,
              nextSearchOptions,
            );
      if (!next) {
        setFocusedEmpty(null);
        const inStaffNote = inStaffCount
          ? " (спочатку Штатка, потім решта)"
          : "";
        setMessage(
          direction === "nextPerson"
            ? skipKeys.length
              ? `Немає іншого службовця з пропусками${inStaffNote} (відкладено: ${skipKeys.length}).`
              : `Немає іншого службовця з порожніми комірками${inStaffNote}.`
            : skipKeys.length
              ? `Немає інших порожніх${inStaffNote} (відкладено: ${skipKeys.length}). «Перша порожня» скине відкладені.`
              : `Порожніх комірок у вибраних колонках немає${inStaffNote}.`,
        );
        return;
      }
      activateEmptyCell(next);
      setEmptySearchActive(true);
      const personName =
        rows.find((row) => row.__rowId === next.rowId)?.fullName?.trim() || "";
      const remaining = countAnketaEmptyCells(
        searchRows,
        gapColumnKeys,
        nextSearchOptions,
      );
      const inStaffPrefix = buildAnketaInStaffRowIdSet(
        rows,
        personnelIndex,
      ).has(next.rowId)
        ? "Штатка · "
        : inStaffCount
          ? "Решта · "
          : "";
      setMessage(
        direction === "nextPerson"
          ? `${inStaffPrefix}Наступний службовець · ${personName || `рядок ${next.rowNumber}`} · ${next.a1} · ${next.header}. Залишилось: ${remaining}.`
          : `${inStaffPrefix}Порожня комірка ${next.a1} · ${next.header} · ${personName || `рядок ${next.rowNumber}`}. Залишилось: ${remaining}${
              skipKeys.length ? ` · відкладено: ${skipKeys.length}` : ""
            }.`,
      );
      prefetchNextGapPerson(searchRows, next, gapColumnKeys, nextSearchOptions);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося знайти порожню комірку: ${error.message}`
          : "Не вдалося знайти порожню комірку.",
      );
    }
  };

  const fillAbsentQuestionnaireCells = async () => {
    if (!canEdit) {
      setMessage("Немає права редагувати анкетні дані.");
      return;
    }
    if (!rows.length) {
      setMessage("Спочатку завантажте анкетні дані.");
      return;
    }
    if (!gapColumnKeys.length) {
      setMessage("Оберіть хоча б одну колонку — запис лише у вибрані.");
      setGapColumnsOpen(true);
      return;
    }

    setIsFillingAbsent(true);
    setMessage("Шукаю осіб без анкет…");
    try {
      const [missingList, questionnaires] = await Promise.all([
        loadAnketaMissingNames().catch(() => ({ names: [] as string[] })),
        api.listPersonQuestionnaires(),
      ]);
      const missingKeys = expandAnketaNameKeySet([
        ...missingQuestionnaireNames,
        ...missingList.names,
      ]);

      const hasPdf = (row: AnketaRow) =>
        anketaRowHasQuestionnairePdf(row, questionnaires);
      const stillMissingKeys =
        removePresentQuestionnairesFromMissingNameKeys(
          rows,
          missingKeys,
          hasPdf,
        );
      setMissingQuestionnaireNames(stillMissingKeys);
      const fills = collectAbsentQuestionnaireCellFills(
        rows,
        gapColumnKeys,
        stillMissingKeys,
        hasPdf,
      );
      const clears = collectAbsentQuestionnaireCellClears(
        rows,
        gapColumnKeys,
        hasPdf,
      );
      const newRows = buildAbsentQuestionnaireAnketaRows(
        rows,
        missingList.names,
        gapColumnKeys,
      ).filter((row) => !hasPdf(row));
      if (!fills.length && !newRows.length && !clears.length) {
        setMessage(
          "Немає порожніх вибраних комірок у осіб без PDF-анкети, і «дані відсутні» знімати теж ніде.",
        );
        return;
      }

      setMessage(
        `Пишу «${ANKETA_ABSENT_QUESTIONNAIRE_VALUE}»… ${fills.length} комірок${
          clears.length ? ` · знімаю: ${clears.length}` : ""
        }${newRows.length ? ` · нових рядків: ${newRows.length}` : ""}`,
      );

      persistSnapshot((current) => ({
        ...current,
        rows: [
          ...applyAbsentQuestionnaireClearsToRows(
            applyAbsentQuestionnaireFillsToRows(current.rows, fills),
            clears,
          ),
          ...newRows,
        ],
      }));

      const newRowEdits = newRows.flatMap((row) => {
        const base = {
          rowNumber: row.__rowNumber,
          externalId: row.externalId,
          fullName: row.fullName,
        };
        return [
          { ...base, columnId: "fullName" as const, value: row.fullName },
          ...gapColumnKeys
            .filter((key) => !isAnketaColumnReadonly(key))
            .map((columnId) => ({
              ...base,
              columnId,
              value: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
            })),
        ];
      });

      const result = await bulkWriteAnketaCellEdits([
        ...fills.map((fill) => ({
          rowNumber: fill.rowNumber,
          columnId: fill.columnId,
          value: ANKETA_ABSENT_QUESTIONNAIRE_VALUE,
          externalId: fill.externalId,
          fullName: fill.fullName,
        })),
        ...clears.map((clear) => ({
          rowNumber: clear.rowNumber,
          columnId: clear.columnId,
          value: "",
          externalId: clear.externalId,
          fullName: clear.fullName,
        })),
        ...newRowEdits,
      ]);
      setEditsCount(countAnketaEdits(result.edits));
      const persons =
        new Set([
          ...fills.map((fill) => fill.rowId),
          ...clears.map((clear) => clear.rowId),
        ]).size + newRows.length;
      const syncNote = result.serverSynced
        ? "збережено в БД"
        : `локально: ок${result.serverError ? ` · сервер: ${result.serverError}` : ""}`;
      setMessage(
        `Без анкет · «${ANKETA_ABSENT_QUESTIONNAIRE_VALUE}»: ${fills.length} комірок · знято (є анкета): ${clears.length} · додано рядків: ${newRows.length} · осіб: ${persons} · ${syncNote}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося записати «дані відсутні».",
      );
    } finally {
      setIsFillingAbsent(false);
    }
  };

  const clearGapReviewProgress = useCallback(() => {
    gapReviewProgressRef.current = {};
    gapReviewProgressLoadedRef.current = true;
    setGapReviewProgress({});
    void saveAnketaGapReviewProgress({});
    setMessage("Список пройдених службовців скинуто.");
  }, [setMessage]);

  const markCurrentGapPersonReviewed = useCallback(() => {
    if (!focusedAnketaRow) {
      setMessage("Немає поточного службовця в пошуку порожніх.");
      return;
    }
    markGapPersonReviewed(focusedAnketaRow);
    const name = focusedAnketaRow.fullName?.trim() || "службовець";
    setMessage(`Запам’ятано як пройденого: ${name}.`);
  }, [focusedAnketaRow, markGapPersonReviewed, setMessage]);

  return {
    focusedEmpty,
    focusEpoch,
    emptySearchActive,
    isFillingAbsent,
    deferredGapKeys,
    gapReviewedCount,
    personPanelOpen,
    setPersonPanelOpen,
    emptyCount,
    gapStats,
    focusedAnketaRow,
    clearGapFocus,
    clearGapReviewProgress,
    markCurrentGapPersonReviewed,
    stopEmptySearch,
    activateEmptyCell,
    cancelCellEdit,
    patchCell,
    patchCells,
    goToEmptyCell,
    fillAbsentQuestionnaireCells,
  };
}
