import { useEffect, useMemo, useRef, useState } from "react";
import { api } from "../../../api";
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
  applyAbsentQuestionnaireClearsToRows,
  applyAbsentQuestionnaireFillsToRows,
  buildAbsentQuestionnaireAnketaRows,
  collectAbsentQuestionnaireCellClears,
  collectAbsentQuestionnaireCellFills,
  countAnketaEmptyCells,
  findNextAnketaEmptyCell,
  findNextAnketaPersonEmptyCell,
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
  expandAnketaNameKeySet,
  loadPersonnelIndexForAnketa,
  matchAnketaRowToPersonnel,
} from "../anketaPersonMatch";
import {
  loadPersonPhotoForRow,
  loadPersonQuestionnaireForRow,
} from "../../personnel/personAttachments";

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
  const suppressCellBlurSaveRef = useRef(false);

  const deferredGapKeySet = useMemo(
    () => new Set(deferredGapKeys),
    [deferredGapKeys],
  );
  const gapSearchOptions = useMemo(
    () => ({
      skipKeys: deferredGapKeySet,
      excludeNameKeys: missingQuestionnaireNames,
    }),
    [deferredGapKeySet, missingQuestionnaireNames],
  );
  const emptyCount = useMemo(
    () => countAnketaEmptyCells(rows, gapColumnKeys, gapSearchOptions),
    [rows, gapColumnKeys, gapSearchOptions],
  );
  const gapStats = useMemo(
    () =>
      summarizeAnketaGaps(rows, gapColumnKeys, {
        excludeNameKeys: missingQuestionnaireNames,
      }),
    [rows, gapColumnKeys, missingQuestionnaireNames],
  );
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
    setFocusedEmpty(cell);
    if (cell) setFocusEpoch((epoch) => epoch + 1);
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
      const next = findNextAnketaEmptyCell(
        nextRows,
        currentGap,
        gapColumnKeys,
        {
          skipKeys: skipKeysAfterSave,
          excludeNameKeys: missingQuestionnaireNames,
        },
      );
      setFocusedEmpty(next);
      if (next) {
        setFocusEpoch((epoch) => epoch + 1);
        setMessage(
          `Збережено. Наступна порожня: ${next.a1} · ${next.header}. Залишилось: ${countAnketaEmptyCells(nextRows, gapColumnKeys, {
            skipKeys: skipKeysAfterSave,
            excludeNameKeys: missingQuestionnaireNames,
          })}.`,
        );
        prefetchNextGapPerson(nextRows, next, gapColumnKeys, {
          skipKeys: skipKeysAfterSave,
          excludeNameKeys: missingQuestionnaireNames,
        });
      } else if (value.trim()) {
        setMessage("Усі порожні комірки (у вибраних колонках) заповнені.");
      }
    } else {
      setFocusedEmpty(null);
    }
  };

  const goToEmptyCell = (direction: "first" | "next" | "nextPerson") => {
    if (!rows.length) {
      setMessage("Спочатку завантажте анкетні дані.");
      return;
    }
    if (!gapColumnKeys.length) {
      setMessage("Оберіть хоча б одну колонку для пошуку пропусків.");
      setGapColumnsOpen(true);
      return;
    }

    let skipKeys = deferredGapKeys;
    if (direction === "first") {
      skipKeys = [];
      setDeferredGapKeys([]);
    } else if (direction === "nextPerson" && focusedEmpty) {
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
      const stillEmpty = !String(
        rows.find((row) => row.__rowId === focusedEmpty.rowId)?.[
          focusedEmpty.columnId
        ] ?? "",
      ).trim();
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
      excludeNameKeys: missingQuestionnaireNames,
    };
    const next =
      direction === "first"
        ? findNextAnketaEmptyCell(rows, null, gapColumnKeys, nextSearchOptions)
        : direction === "nextPerson"
          ? findNextAnketaPersonEmptyCell(
              rows,
              focusedEmpty,
              gapColumnKeys,
              nextSearchOptions,
            )
          : findNextAnketaEmptyCell(rows, focusedEmpty, gapColumnKeys, nextSearchOptions);
    if (!next) {
      setFocusedEmpty(null);
      setMessage(
        direction === "nextPerson"
          ? skipKeys.length
            ? `Немає іншого службовця з пропусками (відкладено: ${skipKeys.length}).`
            : "Немає іншого службовця з порожніми комірками."
          : skipKeys.length
            ? `Немає інших порожніх (відкладено: ${skipKeys.length}). «Перша порожня» скине відкладені.`
            : "Порожніх комірок у вибраних колонках немає.",
      );
      return;
    }
    activateEmptyCell(next);
    setEmptySearchActive(true);
    const personName =
      rows.find((row) => row.__rowId === next.rowId)?.fullName?.trim() || "";
    const remaining = countAnketaEmptyCells(rows, gapColumnKeys, nextSearchOptions);
    setMessage(
      direction === "nextPerson"
        ? `Наступний службовець · ${personName || `рядок ${next.rowNumber}`} · ${next.a1} · ${next.header}. Залишилось: ${remaining}.`
        : `Порожня комірка ${next.a1} · ${next.header} · рядок ${next.rowNumber}. Залишилось: ${remaining}${
            skipKeys.length ? ` · відкладено: ${skipKeys.length}` : ""
          }.`,
    );
    prefetchNextGapPerson(rows, next, gapColumnKeys, nextSearchOptions);
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
      const allWritableColumnKeys = ANKETA_COLUMNS.map(
        (column) => column.key,
      ).filter((key) => !isAnketaColumnReadonly(key));
      const clears = collectAbsentQuestionnaireCellClears(
        rows,
        allWritableColumnKeys,
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

  return {
    focusedEmpty,
    focusEpoch,
    emptySearchActive,
    isFillingAbsent,
    deferredGapKeys,
    personPanelOpen,
    setPersonPanelOpen,
    emptyCount,
    gapStats,
    focusedAnketaRow,
    clearGapFocus,
    stopEmptySearch,
    activateEmptyCell,
    cancelCellEdit,
    patchCell,
    goToEmptyCell,
    fillAbsentQuestionnaireCells,
  };
}
