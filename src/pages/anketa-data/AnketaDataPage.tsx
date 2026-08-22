import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  Alert,
  Box,
  Button,
  LinearProgress,
  Stack,
  TextField,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArrowRightOutlinedIcon,
  FileDownloadOutlinedIcon,
  FileUploadOutlinedIcon,
  SkipNextOutlinedIcon,
  SyncAltOutlinedIcon,
} from "@/components/sci/icons";
import type {
  MRT_ColumnDef,
  SciDataTableCellRef,
  SciDataTableExportContext,
} from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import {
  ANKETA_COLUMNS,
  anketaSheetEditUrl,
  isAnketaColumnReadonly,
  loadAnketaSheetFromFile,
  loadAnketaSheetPreferCache,
  type AnketaColumnKey,
  type AnketaRow,
  type AnketaSheetSnapshot,
} from "./anketaSheet";
import {
  applyAnketaEditsToSnapshot,
  countAnketaEdits,
  loadAnketaEdits,
  upsertAnketaCellEdit,
  type AnketaEditsMap,
} from "./anketaEdits";
import {
  ANKETA_APPS_SCRIPT_TEMPLATE,
  anketaCellA1,
  anketaGapSkipKey,
  anketaGoogleCellUrl,
  findNextAnketaEmptyCell,
  findNextAnketaPersonEmptyCell,
  listAnketaEmptyCells,
  listAnketaPersonGapSkipKeys,
  pushAnketaCellToGoogle,
  readAnketaAppsScriptUrl,
  readAnketaGapColumns,
  summarizeAnketaGaps,
  updateAnketaRowCell,
  writeAnketaAppsScriptUrl,
  writeAnketaGapColumns,
  type AnketaEmptyCell,
} from "./anketaGaps";
import { AnketaCellEditor } from "./AnketaCellEditor";
import { AnketaPersonSidePanel } from "./AnketaPersonSidePanel";
import { exportAnketaSheetExcel } from "./anketaExcelExport";

export function AnketaDataPage() {
  const [snapshot, setSnapshot] = useState<AnketaSheetSnapshot | null>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [message, setMessage] = useState(
    "Завантажте анкетні дані з Google Sheets.",
  );
  const [focusedEmpty, setFocusedEmpty] = useState<AnketaEmptyCell | null>(
    null,
  );
  const [focusEpoch, setFocusEpoch] = useState(0);
  const [dirtyCount, setDirtyCount] = useState(0);
  const [editsCount, setEditsCount] = useState(0);
  const [deferredGapKeys, setDeferredGapKeys] = useState<string[]>([]);
  const [appsScriptUrl, setAppsScriptUrl] = useState(() =>
    readAnketaAppsScriptUrl(),
  );
  const [showSyncHelp, setShowSyncHelp] = useState(false);
  const [gapColumnKeys, setGapColumnKeys] = useState<AnketaColumnKey[]>(() =>
    readAnketaGapColumns(),
  );
  const [gapColumnsOpen, setGapColumnsOpen] = useState(false);
  const [gapMenuPos, setGapMenuPos] = useState({ top: 0, left: 0 });
  const [personPanelOpen, setPersonPanelOpen] = useState(false);
  const gapTriggerRef = useRef<HTMLDivElement | null>(null);
  const gapMenuRef = useRef<HTMLDivElement | null>(null);
  const suppressCellBlurSaveRef = useRef(false);
  const [emptySearchActive, setEmptySearchActive] = useState(false);

  const rows = snapshot?.rows ?? [];
  const gapKeySet = useMemo(() => new Set(gapColumnKeys), [gapColumnKeys]);
  const deferredGapKeySet = useMemo(
    () => new Set(deferredGapKeys),
    [deferredGapKeys],
  );
  const emptyCells = useMemo(
    () =>
      listAnketaEmptyCells(rows, gapColumnKeys, {
        skipKeys: deferredGapKeySet,
      }),
    [rows, gapColumnKeys, deferredGapKeySet],
  );
  const emptyCount = emptyCells.length;
  const gapStats = useMemo(
    () => summarizeAnketaGaps(rows, gapColumnKeys),
    [rows, gapColumnKeys],
  );
  const focusedAnketaRow = useMemo(() => {
    if (!focusedEmpty) return null;
    return rows.find((row) => row.__rowId === focusedEmpty.rowId) ?? null;
  }, [focusedEmpty, rows]);

  useEffect(() => {
    if (focusedEmpty) setPersonPanelOpen(true);
  }, [focusedEmpty]);

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

  const mergeWithEdits = async (base: AnketaSheetSnapshot) => {
    const edits = await loadAnketaEdits();
    setEditsCount(countAnketaEdits(edits));
    return applyAnketaEditsToSnapshot(base, edits);
  };

  useEffect(() => {
    writeAnketaGapColumns(gapColumnKeys);
  }, [gapColumnKeys]);

  const updateGapMenuPosition = () => {
    const rect = gapTriggerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const width = Math.min(360, window.innerWidth - 24);
    const left = Math.min(
      Math.max(12, rect.left),
      Math.max(12, window.innerWidth - width - 12),
    );
    setGapMenuPos({ top: rect.bottom + 6, left });
  };

  useLayoutEffect(() => {
    if (!gapColumnsOpen) return;
    updateGapMenuPosition();
  }, [gapColumnsOpen]);

  useEffect(() => {
    if (!gapColumnsOpen) return;
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (gapMenuRef.current?.contains(target)) return;
      if (gapTriggerRef.current?.contains(target)) return;
      setGapColumnsOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setGapColumnsOpen(false);
    };
    const onReposition = () => updateGapMenuPosition();
    window.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("resize", onReposition);
    window.addEventListener("scroll", onReposition, true);
    return () => {
      window.removeEventListener("pointerdown", onPointerDown);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("resize", onReposition);
      window.removeEventListener("scroll", onReposition, true);
    };
  }, [gapColumnsOpen]);

  const toggleGapColumn = (key: AnketaColumnKey) => {
    if (isAnketaColumnReadonly(key)) return;
    setGapColumnKeys((current) => {
      const next = current.includes(key)
        ? current.filter((item) => item !== key)
        : ANKETA_COLUMNS.map((column) => column.key).filter(
            (item) => item === key || current.includes(item),
          );
      writeAnketaGapColumns(next);
      return next;
    });
  };

  const selectAllGapColumns = () => {
    const next = ANKETA_COLUMNS.filter((column) => !column.readonly).map(
      (column) => column.key,
    );
    writeAnketaGapColumns(next);
    setGapColumnKeys(next);
  };

  const clearGapColumns = () => {
    writeAnketaGapColumns([]);
    setGapColumnKeys([]);
  };

  const persistSnapshot = (
    updater: (current: AnketaSheetSnapshot) => AnketaSheetSnapshot,
    note?: string,
  ) => {
    setSnapshot((current) => {
      if (!current) return current;
      return updater(current);
    });
    setDirtyCount((value) => value + 1);
    if (note) setMessage(note);
  };

  const patchCell = async (
    rowId: string,
    columnId: AnketaColumnKey,
    value: string,
    options?: { advance?: boolean },
  ) => {
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
        { skipKeys: skipKeysAfterSave },
      );
      setFocusedEmpty(next);
      if (next) {
        setFocusEpoch((epoch) => epoch + 1);
        setMessage(
          `Збережено. Наступна порожня: ${next.a1} · ${next.header}. Залишилось: ${listAnketaEmptyCells(nextRows, gapColumnKeys, { skipKeys: skipKeysAfterSave }).length}.`,
        );
      } else if (value.trim()) {
        setMessage("Усі порожні комірки (у вибраних колонках) заповнені.");
      }
    } else {
      setFocusedEmpty(null);
    }
  };

  const loadFromGoogle = async () => {
    setIsLoading(true);
    try {
      const next = await loadAnketaSheetPreferCache({
        mergeEdits: mergeWithEdits,
        onCached: (cached) => {
          setSnapshot(cached);
          setIsLoading(false);
          setMessage(
            `Кеш + локальні правки · ${cached.rows.length} анкет · оновлюю з Google…`,
          );
        },
      });
      setSnapshot(next);
      setEmptySearchActive(false);
      setFocusedEmpty(null);
      setDirtyCount(0);
      setMessage(
        next.source === "cache"
          ? `Показано кеш + локальні правки (${next.rows.length}). Мережа недоступна.`
          : `Оновлено з Google + накладено локальні правки · ${new Date(next.fetchedAt).toLocaleString("uk-UA")}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? `Не вдалося завантажити таблицю: ${error.message}`
          : "Не вдалося завантажити Google Sheets.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const loadFromCsvFile = async (file: File | undefined) => {
    if (!file) return;
    setIsLoading(true);
    try {
      const base = await loadAnketaSheetFromFile(file);
      const next = await mergeWithEdits(base);
      setSnapshot(next);
      setEmptySearchActive(false);
      setFocusedEmpty(null);
      setDirtyCount(0);
      setMessage(`Імпортовано файл ${file.name}: ${next.rows.length} анкет.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати CSV файл.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    void loadFromGoogle();
    void loadAnketaEdits().then((edits: AnketaEditsMap) => {
      setEditsCount(countAnketaEdits(edits));
    });
  }, []);

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
        { skipKeys: deferredGapKeys },
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

    const next =
      direction === "first"
        ? findNextAnketaEmptyCell(rows, null, gapColumnKeys, { skipKeys })
        : direction === "nextPerson"
          ? findNextAnketaPersonEmptyCell(
              rows,
              focusedEmpty,
              gapColumnKeys,
              { skipKeys },
            )
          : findNextAnketaEmptyCell(rows, focusedEmpty, gapColumnKeys, {
              skipKeys,
            });
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
    const remaining = listAnketaEmptyCells(rows, gapColumnKeys, {
      skipKeys,
    }).length;
    setMessage(
      direction === "nextPerson"
        ? `Наступний службовець · ${personName || `рядок ${next.rowNumber}`} · ${next.a1} · ${next.header}. Залишилось: ${remaining}.`
        : `Порожня комірка ${next.a1} · ${next.header} · рядок ${next.rowNumber}. Залишилось: ${remaining}${
            skipKeys.length ? ` · відкладено: ${skipKeys.length}` : ""
          }.`,
    );
  };

  const exportTable = async (
    context: SciDataTableExportContext<AnketaRow>,
  ) => {
    await exportAnketaSheetExcel({
      columns: context.columns,
      rows: context.rows,
    });
    setMessage(
      `Експортовано ${context.rows.length} рядків у Excel (стиль Google Sheets).`,
    );
  };

  const focusedCell: SciDataTableCellRef | null = focusedEmpty
    ? {
        rowId: focusedEmpty.rowId,
        columnId: focusedEmpty.columnId,
        focusEpoch,
      }
    : null;

  const columns = useMemo<Array<MRT_ColumnDef<AnketaRow>>>(
    () => [
      {
        id: "__rowNumber",
        accessorFn: (row) => row.__rowNumber,
        header: "№",
        size: 64,
        pin: "left",
        enableColumnFilter: false,
        enableGlobalFilter: false,
        enableHiding: false,
        Cell: ({ row }) => (
          <span className="anketa-row-number">{row.original.__rowNumber}</span>
        ),
        exportValue: (row) => String(row.__rowNumber),
      },
      ...ANKETA_COLUMNS.map((column) => ({
        id: column.key,
        accessorKey: column.key,
        header: column.header,
        size: column.size,
        pin: column.pin,
        filterVariant: column.filterVariant,
        enableColumnFilter: true,
        enableGlobalFilter: true,
        Cell: ({ row }) => {
          const rowId = row.original.__rowId;
          const value = String(row.original[column.key] ?? "");
          const isFocused =
            focusedEmpty?.rowId === rowId &&
            focusedEmpty?.columnId === column.key;
          const isEmpty = !value.trim();
          const isReadonly = Boolean(column.readonly);

          if (isReadonly) {
            return (
              <span className="anketa-cell-readonly" title="ПІБ з Google Sheet — не редагується">
                {value || "—"}
              </span>
            );
          }

          if (isFocused) {
            return (
              <AnketaCellEditor
                key={`${rowId}:${column.key}:${focusEpoch}`}
                columnKey={column.key}
                columnHeader={column.header}
                rowNumber={row.original.__rowNumber}
                value={value}
                isEmpty={isEmpty}
                advanceOnSave={emptySearchActive}
                onCancel={cancelCellEdit}
                onSave={(next, advance) => {
                  void patchCell(rowId, column.key, next, {
                    advance: advance || emptySearchActive,
                  });
                }}
              />
            );
          }

          return (
            <button
              type="button"
              className={
                isEmpty ? "anketa-cell-button is-empty" : "anketa-cell-button"
              }
              onClick={() => {
                const columnIndex = ANKETA_COLUMNS.findIndex(
                  (item) => item.key === column.key,
                );
                activateEmptyCell({
                  rowId,
                  columnId: column.key,
                  rowNumber: row.original.__rowNumber,
                  columnIndex: Math.max(columnIndex, 0),
                  header: column.header,
                  a1: anketaCellA1(
                    row.original.__rowNumber,
                    Math.max(columnIndex, 0),
                  ),
                });
              }}
            >
              {value || "—"}
            </button>
          );
        },
      })),
    ],
    [focusEpoch, focusedEmpty, rows, appsScriptUrl, emptySearchActive],
  );

  const table = useMaterialReactTable({
    columns,
    data: rows,
    emptyMessage: isLoading
      ? "Завантаження…"
      : "Немає анкетних даних. Оновіть з Google або імпортуйте CSV.",
    exportLabel: "Експорт Excel",
    globalFilterPlaceholder: "Пошук: ПІБ, ID, звання, РНОКПП…",
    getRowId: (row) => row.__rowId,
    onExport: (context) => void exportTable(context),
    enableColumnFilters: true,
    enableColumnVisibility: true,
    enableGlobalFilter: true,
    enableRowVirtualization: true,
    estimatedRowHeight: 38,
    focusedCell,
    getTdProps: ({ rowId, columnId, row }) => {
      if (columnId === "__rowNumber") {
        return { className: "anketa-td-row-number" };
      }
      const key = columnId as AnketaColumnKey;
      const empty = !String(row[key] ?? "").trim();
      const inGapScope = gapKeySet.has(key);
      const focused =
        focusedEmpty?.rowId === rowId && focusedEmpty?.columnId === columnId;
      return {
        className: [
          empty && inGapScope ? "anketa-td-empty" : "",
          focused ? "is-focused-empty-cell" : "",
        ]
          .filter(Boolean)
          .join(" "),
      };
    },
    initialState: {
      density: "compact",
      pagination: { pageIndex: 0, pageSize: 200 },
      columnPinning: { left: ["__rowNumber", "fullName"] },
    },
  });

  return (
    <main className="main-panel anketa-data-page">
      <header className="topbar">
        <Box>
          <Typography component="h1" variant="h4">
            Анкетні дані
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Заповнення пропусків · локальне редагування · експорт / Google sync
          </Typography>
        </Box>
        <Stack direction="row" spacing={1} sx={{ flexWrap: "wrap", alignItems: "center" }}>
          <div className="anketa-gap-columns" ref={gapTriggerRef}>
            <Button
              variant="outlined"
              aria-expanded={gapColumnsOpen}
              aria-haspopup="listbox"
              onClick={() => setGapColumnsOpen((value) => !value)}
            >
              Колонки пропусків · {gapColumnKeys.length}
            </Button>
            {gapColumnsOpen
              ? createPortal(
                  <div
                    ref={gapMenuRef}
                    className="anketa-gap-columns-menu"
                    role="listbox"
                    style={{ top: gapMenuPos.top, left: gapMenuPos.left }}
                    onPointerDown={(event) => event.stopPropagation()}
                  >
                    <div className="anketa-gap-columns-actions">
                      <button type="button" onClick={selectAllGapColumns}>
                        Усі
                      </button>
                      <button type="button" onClick={clearGapColumns}>
                        Жодної
                      </button>
                    </div>
                    <div className="anketa-gap-columns-list">
                      {ANKETA_COLUMNS.map((column) => {
                        const checked = gapKeySet.has(column.key);
                        const isReadonly = Boolean(column.readonly);
                        return (
                          <button
                            key={column.key}
                            type="button"
                            role="option"
                            aria-selected={checked}
                            disabled={isReadonly}
                            className={
                              isReadonly
                                ? "anketa-gap-columns-option is-readonly"
                                : checked
                                  ? "anketa-gap-columns-option is-active"
                                  : "anketa-gap-columns-option"
                            }
                            onClick={() => toggleGapColumn(column.key)}
                          >
                            <span
                              className="sci-data-table-check"
                              aria-hidden="true"
                            >
                              {checked ? "■" : ""}
                            </span>
                            <span>
                              {column.header}
                              {isReadonly ? " · лише перегляд" : ""}
                            </span>
                          </button>
                        );
                      })}
                    </div>
                  </div>,
                  document.body,
                )
              : null}
          </div>
          <Button
            variant="outlined"
            startIcon={<SkipNextOutlinedIcon />}
            disabled={isLoading || !rows.length}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => goToEmptyCell("first")}
          >
            Перша порожня
          </Button>
          <Button
            variant="outlined"
            startIcon={<ArrowRightOutlinedIcon />}
            disabled={isLoading || !rows.length}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => goToEmptyCell("next")}
          >
            Наступна порожня
          </Button>
          <Button
            variant="outlined"
            startIcon={<SkipNextOutlinedIcon />}
            disabled={isLoading || !rows.length}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => goToEmptyCell("nextPerson")}
            title="Пропустити поточного службовця і перейти до наступного з пропусками"
          >
            Наступний службовець
          </Button>
          <Button
            variant="outlined"
            disabled={!emptySearchActive}
            onClick={stopEmptySearch}
            title="Esc"
          >
            Стоп · Esc
          </Button>
          <Button
            variant="outlined"
            startIcon={<SyncAltOutlinedIcon />}
            disabled={isLoading}
            onClick={() => void loadFromGoogle()}
          >
            Оновити з Google
          </Button>
          <Button
            component="label"
            variant="outlined"
            startIcon={<FileUploadOutlinedIcon />}
            disabled={isLoading}
          >
            CSV файл
            <input
              hidden
              type="file"
              accept=".csv,text/csv"
              onChange={(event) => {
                void loadFromCsvFile(event.target.files?.[0]);
                event.currentTarget.value = "";
              }}
            />
          </Button>
          <Button
            variant="outlined"
            startIcon={<FileDownloadOutlinedIcon />}
            component="a"
            href={
              focusedEmpty
                ? anketaGoogleCellUrl(
                    focusedEmpty.rowNumber,
                    focusedEmpty.columnIndex,
                  )
                : anketaSheetEditUrl()
            }
            target="_blank"
            rel="noreferrer"
          >
            {focusedEmpty ? `Google · ${focusedEmpty.a1}` : "Відкрити таблицю"}
          </Button>
        </Stack>
      </header>

      {isLoading || isSyncing ? <LinearProgress sx={{ mb: 1 }} /> : null}

      <Alert severity="info" className="personnel-page-alert" sx={{ mb: 1.5 }}>
        {message}
        {snapshot
          ? ` · вибрані колонки: ${gapStats.columns} · порожніх ячійок: ${gapStats.emptyCells} · осіб з пропусками: ${gapStats.personsWithGaps} · усього рядків/осіб: ${gapStats.totalRows} · правок у БД: ${editsCount}${
              emptySearchActive
                ? ` · у пошуку лишилось: ${emptyCount}`
                : ""
            }${
              deferredGapKeys.length
                ? ` · відкладено: ${deferredGapKeys.length}`
                : ""
            }${dirtyCount ? ` · сесія: ${dirtyCount}` : ""}`
          : ""}
      </Alert>

      <section className="analytics-panel anketa-sync-panel">
        <div className="panel-heading">
          Синхронізація з Google
          <Button
            size="small"
            variant="text"
            onClick={() => setShowSyncHelp((value) => !value)}
          >
            {showSyncHelp ? "Сховати інструкцію" : "Як підключити запис"}
          </Button>
        </div>
        <div className="anketa-sync-body">
          <TextField
            size="small"
            fullWidth
            label="URL Google Apps Script Web App"
            placeholder="https://script.google.com/macros/s/.../exec"
            value={appsScriptUrl}
            onChange={(event) => setAppsScriptUrl(event.target.value)}
            onBlur={() => writeAnketaAppsScriptUrl(appsScriptUrl)}
          />
          <Typography variant="caption" color="text.secondary">
            Правки завжди пишуться в локальну БД (IndexedDB) і не зникають після
            «Оновити з Google». З URL Apps Script — додатково в Google Sheet.
          </Typography>
          {showSyncHelp ? (
            <pre className="anketa-apps-script-template">
              {ANKETA_APPS_SCRIPT_TEMPLATE}
            </pre>
          ) : null}
        </div>
      </section>

      <div
        className={[
          "anketa-workspace",
          personPanelOpen && focusedAnketaRow ? "has-person-panel" : "",
        ]
          .filter(Boolean)
          .join(" ")}
      >
        <section className="analytics-panel anketa-data-table-panel">
          <div className="panel-heading">
            Анкети
            {snapshot ? (
              <span className="personnel-list-questionnaire-count">
                · {snapshot.rows.length}
                {focusedEmpty ? ` · фокус ${focusedEmpty.a1}` : ""}
              </span>
            ) : null}
          </div>
          <div className="anketa-data-table-wrap">
            <MaterialReactTable table={table} />
          </div>
        </section>

        {personPanelOpen && focusedAnketaRow ? (
          <AnketaPersonSidePanel
            anketaRow={focusedAnketaRow}
            focusedEmpty={focusedEmpty}
            gapColumnKeys={gapColumnKeys}
            onClose={() => setPersonPanelOpen(false)}
            onFillMissing={
              focusedEmpty && emptySearchActive
                ? (value) => {
                    void patchCell(
                      focusedEmpty.rowId,
                      focusedEmpty.columnId,
                      value,
                      { advance: true },
                    );
                  }
                : undefined
            }
          />
        ) : null}
      </div>
    </main>
  );
}
