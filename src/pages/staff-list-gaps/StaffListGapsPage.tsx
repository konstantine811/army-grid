import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  Alert,
  Box,
  Button,
  LinearProgress,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArrowRightOutlinedIcon,
  FileDownloadOutlinedIcon,
  FileUploadOutlinedIcon,
  SkipNextOutlinedIcon,
} from "@/components/sci/icons";
import {
  MaterialReactTable,
  useMaterialReactTable,
  type MRT_ColumnDef,
} from "@/components/sci/SciDataTable";
import { useAuth } from "../../auth/AuthProvider";
import { loadAnketaSheetPreferCache } from "../anketa-data/anketaSheet";
import { StaffListAnketaDock } from "./StaffListAnketaDock";
import { StaffListCellEditor } from "./StaffListCellEditor";
import {
  buildAnketaRowLookup,
  matchAnketaRowForStaffList,
  resolveStaffListAnketaRow,
} from "./staffListGapsAnketaMatch";
import {
  parseStaffListExcelFile,
  type StaffListRow,
  type StaffListSnapshot,
} from "./staffListGapsParse";
import { formatApiDateTime } from "../../shared/format";
import { exportStaffListExcel } from "./staffListGapsExport";
import {
  buildStaffListGapsSession,
  loadStaffListGapsSession,
  saveStaffListGapsSession,
} from "./staffListGapsStorage";
import {
  buildStaffListFocusedCell,
  countStaffListEmptyCells,
  findNextStaffListEmptyCell,
  isStaffListCellEmpty,
  readStaffListGapColumnIds,
  resolveStaffListGapColumnIds,
  staffListGapSkipKey,
  summarizeStaffListGaps,
  writeStaffListGapColumnIds,
  type StaffListEmptyCell,
} from "./staffListGapsSearch";

export function StaffListGapsPage() {
  const { canEditArea } = useAuth();
  const canEdit = canEditArea("anketaData");
  const [snapshot, setSnapshot] = useState<StaffListSnapshot | null>(null);
  const [rows, setRows] = useState<StaffListRow[]>([]);
  const [message, setMessage] = useState(
    "Завантажте Excel-список (наприклад, номери телефонів).",
  );
  const [isLoading, setIsLoading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [fileData, setFileData] = useState<ArrayBuffer | null>(null);
  const [gapColumnIds, setGapColumnIds] = useState<string[]>(() =>
    readStaffListGapColumnIds(),
  );
  const [gapColumnsOpen, setGapColumnsOpen] = useState(false);
  const [focusedEmpty, setFocusedEmpty] = useState<StaffListEmptyCell | null>(
    null,
  );
  const [focusEpoch, setFocusEpoch] = useState(0);
  const [emptySearchActive, setEmptySearchActive] = useState(false);
  const [deferredGapKeys, setDeferredGapKeys] = useState<string[]>([]);
  const [personPanelOpen, setPersonPanelOpen] = useState(false);
  const [anketaRows, setAnketaRows] = useState<
    import("../anketa-data/anketaSheet").AnketaRow[]
  >([]);
  const gapMenuRef = useRef<HTMLDivElement | null>(null);
  const gapTriggerRef = useRef<HTMLDivElement | null>(null);
  const skipPersistRef = useRef(true);

  const ensureGapColumns = useCallback((nextSnapshot: StaffListSnapshot) => {
    return resolveStaffListGapColumnIds(
      nextSnapshot,
      readStaffListGapColumnIds(),
    );
  }, []);

  const persistSession = useCallback(
    async (
      nextSnapshot: StaffListSnapshot,
      nextRows: StaffListRow[],
      nextGapColumnIds: string[],
      nextFileData: ArrayBuffer | null,
    ) => {
      if (!nextFileData) return;
      await saveStaffListGapsSession(
        buildStaffListGapsSession(
          nextSnapshot,
          nextRows,
          nextGapColumnIds,
          nextFileData,
        ),
      );
    },
    [],
  );

  useEffect(() => {
    let cancelled = false;
    void loadStaffListGapsSession()
      .then((session) => {
        if (cancelled) return;
        if (!session?.snapshot.rows.length) {
          skipPersistRef.current = false;
          return;
        }
        skipPersistRef.current = true;
        setSnapshot(session.snapshot);
        setRows(session.snapshot.rows);
        setFileData(session.fileData ?? null);
        setGapColumnIds(
          resolveStaffListGapColumnIds(
            session.snapshot,
            session.gapColumnIds.length
              ? session.gapColumnIds
              : readStaffListGapColumnIds(),
          ),
        );
        setMessage(
          `Відновлено ${session.snapshot.fileName} · ${session.snapshot.rows.length} рядків · збережено ${formatApiDateTime(session.savedAt)}.`,
        );
        window.setTimeout(() => {
          skipPersistRef.current = false;
        }, 0);
      })
      .catch(() => {
        skipPersistRef.current = false;
      });
    return () => {
      cancelled = true;
    };
  }, [ensureGapColumns]);

  useEffect(() => {
    if (!snapshot || !rows.length || !fileData || skipPersistRef.current) return;
    const timer = window.setTimeout(() => {
      void persistSession(snapshot, rows, gapColumnIds, fileData).catch(() => {
        setMessage("Не вдалося зберегти сесію в локальну БД (IndexedDB).");
      });
    }, 400);
    return () => window.clearTimeout(timer);
  }, [snapshot, rows, gapColumnIds, fileData, persistSession]);

  useEffect(() => {
    let cancelled = false;
    void loadAnketaSheetPreferCache()
      .then((sheet) => {
        if (!cancelled) setAnketaRows(sheet.rows);
      })
      .catch(() => {
        if (!cancelled) setAnketaRows([]);
      });
    return () => {
      cancelled = true;
    };
  }, []);

  useEffect(() => {
    if (!snapshot) return;
    setGapColumnIds((current) => resolveStaffListGapColumnIds(snapshot, current));
  }, [snapshot]);

  useEffect(() => {
    writeStaffListGapColumnIds(gapColumnIds);
  }, [gapColumnIds]);

  useEffect(() => {
    if (!gapColumnsOpen) return;
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (gapMenuRef.current?.contains(target)) return;
      if (gapTriggerRef.current?.contains(target)) return;
      setGapColumnsOpen(false);
    };
    window.addEventListener("pointerdown", onPointerDown);
    return () => window.removeEventListener("pointerdown", onPointerDown);
  }, [gapColumnsOpen]);

  useEffect(() => {
    if (focusedEmpty) setPersonPanelOpen(true);
  }, [focusedEmpty]);

  const anketaLookup = useMemo(
    () => buildAnketaRowLookup(anketaRows),
    [anketaRows],
  );
  const gapKeySet = useMemo(() => new Set(gapColumnIds), [gapColumnIds]);
  const deferredGapKeySet = useMemo(
    () => new Set(deferredGapKeys),
    [deferredGapKeys],
  );
  const gapStats = useMemo(
    () => summarizeStaffListGaps(rows, gapColumnIds),
    [rows, gapColumnIds],
  );
  const emptyCount = useMemo(
    () =>
      countStaffListEmptyCells(rows, gapColumnIds, {
        skipKeys: deferredGapKeySet,
      }),
    [rows, gapColumnIds, deferredGapKeySet],
  );

  const focusedStaffRow = useMemo(
    () => rows.find((row) => row.__rowId === focusedEmpty?.rowId) ?? null,
    [focusedEmpty, rows],
  );

  const focusedAnketaRow = useMemo(() => {
    if (!focusedStaffRow) return null;
    return resolveStaffListAnketaRow(anketaLookup, focusedStaffRow);
  }, [anketaLookup, focusedStaffRow]);

  const importFile = async (file?: File | null) => {
    if (!file) return;
    setIsLoading(true);
    setMessage(`Читаю ${file.name}…`);
    try {
      const parsed = await parseStaffListExcelFile(file);
      const next = parsed.snapshot;
      const nextGapColumnIds = ensureGapColumns(next);
      skipPersistRef.current = true;
      setSnapshot(next);
      setRows(next.rows);
      setFileData(parsed.fileData);
      setGapColumnIds(nextGapColumnIds);
      setFocusedEmpty(null);
      setEmptySearchActive(false);
      setDeferredGapKeys([]);
      setPersonPanelOpen(false);
      await persistSession(next, next.rows, nextGapColumnIds, parsed.fileData);
      skipPersistRef.current = false;
      setMessage(
        `Завантажено ${next.rows.length} рядків · ${next.sheetName} · ${next.fileName}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося прочитати Excel.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const stopEmptySearch = () => {
    setEmptySearchActive(false);
    setFocusedEmpty(null);
    setPersonPanelOpen(false);
    setGapColumnsOpen(false);
    setMessage("Пошук порожніх зупинено · Esc.");
  };

  const activateEmptyCell = (cell: StaffListEmptyCell | null) => {
    setFocusedEmpty(cell);
    if (cell) setFocusEpoch((epoch) => epoch + 1);
  };

  const focusCell = useCallback(
    (cell: StaffListEmptyCell) => {
      activateEmptyCell(cell);
      setPersonPanelOpen(true);
      const staffRow = rows.find((row) => row.__rowId === cell.rowId);
      const anketa = matchAnketaRowForStaffList(
        anketaLookup,
        cell.pib,
        staffRow?.birthDate ?? "",
      );
      const anketaNote = anketa
        ? ` · анкетні дані: ${anketa.fullName}`
        : " · шукаю в особовому складі";
      setMessage(`Комірка · ${cell.pib} · ${cell.header}${anketaNote}.`);
    },
    [anketaLookup, rows],
  );

  const goToEmptyCell = async (direction: "first" | "next") => {
    if (!snapshot || !rows.length) {
      setMessage("Спочатку завантажте Excel-список.");
      return;
    }
    if (!gapColumnIds.length) {
      setMessage("Оберіть хоча б одну колонку для пошуку пропусків.");
      setGapColumnsOpen(true);
      return;
    }

    let skipKeys = deferredGapKeys;
    if (direction === "first") {
      skipKeys = [];
      setDeferredGapKeys([]);
    } else if (focusedEmpty) {
      skipKeys = [...deferredGapKeys, staffListGapSkipKey(focusedEmpty)];
      setDeferredGapKeys(skipKeys);
    }

    const next = findNextStaffListEmptyCell(
      rows,
      snapshot.columns,
      gapColumnIds,
      direction === "first" ? null : focusedEmpty,
      { skipKeys: new Set(skipKeys) },
    );

    if (!next) {
      setFocusedEmpty(null);
      setMessage(
        skipKeys.length
          ? `Немає інших порожніх (відкладено: ${skipKeys.length}).`
          : "Порожніх комірок у вибраних колонках немає.",
      );
      return;
    }

    activateEmptyCell(next);
    setEmptySearchActive(true);

    const anketa = matchAnketaRowForStaffList(
      anketaLookup,
      next.pib,
      rows.find((row) => row.__rowId === next.rowId)?.birthDate ?? "",
    );
    const anketaNote = anketa
      ? ` · анкетні дані: ${anketa.fullName}`
      : " · шукаю в особовому складі";
    setMessage(
      `Порожня комірка · ${next.pib} · ${next.header}${anketaNote}. Залишилось: ${countStaffListEmptyCells(rows, gapColumnIds, {
        skipKeys: new Set(skipKeys),
      })}.`,
    );
  };

  const patchCell = (
    rowId: string,
    columnId: string,
    value: string,
    options?: { advance?: boolean },
  ) => {
    if (!canEdit) {
      setMessage("Немає права редагувати.");
      return;
    }
    setRows((current) => {
      const nextRows = current.map((row) =>
        row.__rowId === rowId
          ? { ...row, values: { ...row.values, [columnId]: value } }
          : row,
      );
      if (snapshot && fileData) {
        void persistSession(snapshot, nextRows, gapColumnIds, fileData).catch(
          () => {
            setMessage("Не вдалося зберегти в локальну БД (IndexedDB).");
          },
        );
      }
      return nextRows;
    });
    setMessage(`Збережено · ${value || "порожньо"}.`);
    if (options?.advance && emptySearchActive) {
      window.setTimeout(() => {
        void goToEmptyCell("next");
      }, 0);
    }
  };

  const exportExcel = async () => {
    if (!snapshot || !rows.length) {
      setMessage("Спочатку завантажте Excel-список.");
      return;
    }
    if (!fileData) {
      setMessage("Немає оригінального файлу — завантажте Excel знову для експорту.");
      return;
    }
    setIsExporting(true);
    try {
      await exportStaffListExcel(snapshot, rows, fileData);
      setMessage(`Експортовано ${rows.length} рядків · стилі збережено.`);
    } catch (error) {
      setMessage(
        error instanceof Error ? error.message : "Не вдалося експортувати Excel.",
      );
    } finally {
      setIsExporting(false);
    }
  };

  useEffect(() => {
    if (!focusedEmpty) return;
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      if (gapColumnsOpen) return;
      event.preventDefault();
      if (emptySearchActive) stopEmptySearch();
      else setFocusedEmpty(null);
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [focusedEmpty, gapColumnsOpen, emptySearchActive]);

  const columns = useMemo<MRT_ColumnDef<StaffListRow>[]>(() => {
    if (!snapshot) return [];
    return snapshot.columns.map((column) => ({
      accessorKey: column.id,
      header: column.label,
      accessorFn: (row) => row.values[column.id] ?? "",
      size: column.id === snapshot.pibColumnId ? 280 : 140,
      Cell: ({ row, cell }) => {
        const rowId = row.original.__rowId;
        const isFocused =
          focusedEmpty?.rowId === rowId &&
          focusedEmpty.columnId === column.id;
        const value = String(cell.getValue() ?? "");
        const isEmpty = isStaffListCellEmpty(row.original, column.id);
        const inGapScope = gapKeySet.has(column.id);

        if (isFocused && focusedEmpty) {
          return (
            <StaffListCellEditor
              key={`${rowId}:${column.id}:${focusEpoch}`}
              columnHeader={focusedEmpty.header}
              pib={focusedEmpty.pib}
              value={isEmpty ? "" : value}
              disabled={!canEdit}
              advanceOnSave={emptySearchActive}
              onCancel={() => {
                if (emptySearchActive) stopEmptySearch();
                else setFocusedEmpty(null);
              }}
              onSave={(next, advance) => {
                patchCell(rowId, column.id, next, { advance });
              }}
            />
          );
        }

        return (
          <button
            type="button"
            className={
              isEmpty && inGapScope
                ? "anketa-cell-button is-empty staff-list-cell-button"
                : "anketa-cell-button staff-list-cell-button"
            }
            onClick={() => {
              focusCell(buildStaffListFocusedCell(row.original, column));
            }}
          >
            {isEmpty ? "—" : value}
          </button>
        );
      },
    }));
  }, [
    snapshot,
    focusedEmpty,
    focusEpoch,
    emptySearchActive,
    canEdit,
    gapKeySet,
    focusCell,
  ]);

  const getTdProps = useCallback(
    ({
      rowId,
      columnId,
      row,
    }: {
      rowId: string;
      columnId: string;
      row: StaffListRow;
    }) => {
      const empty = isStaffListCellEmpty(row, columnId);
      const inGapScope = gapKeySet.has(columnId);
      return {
        className: empty && inGapScope ? "anketa-td-empty" : "",
      };
    },
    [gapKeySet],
  );

  const focusedCell = focusedEmpty
    ? {
        rowId: focusedEmpty.rowId,
        columnId: focusedEmpty.columnId,
        focusEpoch,
      }
    : null;

  const table = useMaterialReactTable({
    columns,
    data: rows,
    emptyMessage: isLoading ? "Завантаження…" : "Завантажте Excel-список.",
    globalFilterPlaceholder: "Пошук за ПІБ, посадою, підрозділом…",
    getRowId: (row) => row.__rowId,
    enableColumnFilters: true,
    enableGlobalFilter: true,
    enableRowVirtualization: true,
    estimatedRowHeight: 38,
    focusedCell,
    getTdProps,
    initialState: {
      density: "compact",
      pagination: { pageIndex: 0, pageSize: 200 },
      columnPinning: snapshot?.pibColumnId
        ? { left: [snapshot.pibColumnId] }
        : undefined,
    },
  });

  const personPanelVisible = Boolean(
    personPanelOpen && focusedAnketaRow && focusedEmpty,
  );

  const toggleGapColumn = (columnId: string) => {
    setGapColumnIds((current) =>
      current.includes(columnId)
        ? current.filter((item) => item !== columnId)
        : [...current, columnId],
    );
  };

  return (
    <main className="main-panel anketa-data-page staff-list-gaps-page">
      <header className="topbar anketa-topbar">
        <Box className="anketa-topbar-copy">
          <Typography component="h1" variant="h4">
            Списки Excel
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Завантажте список, знайдіть порожні колонки та відкрийте анкету за
            ПІБ.
          </Typography>
        </Box>

        <div className="anketa-toolbar">
          <div className="anketa-toolbar-group" aria-label="Файл">
            <Button
              component="label"
              variant="outlined"
              disabled={isLoading}
              startIcon={<FileUploadOutlinedIcon />}
            >
              Завантажити Excel
              <input
                hidden
                type="file"
                accept=".xlsx,.xlsm,.xls"
                onChange={(event) => {
                  void importFile(event.target.files?.[0]);
                  event.target.value = "";
                }}
              />
            </Button>
            <Button
              variant="outlined"
              disabled={isLoading || isExporting || !snapshot || !fileData}
              startIcon={<FileDownloadOutlinedIcon />}
              onClick={() => void exportExcel()}
            >
              {isExporting ? "Експорт…" : "Експорт Excel"}
            </Button>
          </div>

          <div className="anketa-toolbar-group" aria-label="Пропуски">
            <div className="anketa-gap-columns" ref={gapTriggerRef}>
              <Button
                variant="outlined"
                aria-expanded={gapColumnsOpen}
                disabled={!snapshot}
                onClick={() => setGapColumnsOpen((value) => !value)}
              >
                Колонки пропусків · {gapColumnIds.length}
              </Button>
              {gapColumnsOpen && snapshot ? (
                <div ref={gapMenuRef} className="anketa-gap-columns-menu">
                  <div className="anketa-gap-columns-actions">
                    <button
                      type="button"
                      onClick={() =>
                        setGapColumnIds(snapshot.columns.map((column) => column.id))
                      }
                    >
                      Усі
                    </button>
                    <button type="button" onClick={() => setGapColumnIds([])}>
                      Жодної
                    </button>
                  </div>
                  <div className="anketa-gap-columns-list">
                    {snapshot.columns.map((column) => {
                      const checked = gapKeySet.has(column.id);
                      const isPib = column.id === snapshot.pibColumnId;
                      return (
                        <button
                          key={column.id}
                          type="button"
                          className={
                            checked
                              ? "anketa-gap-columns-option is-active"
                              : "anketa-gap-columns-option"
                          }
                          disabled={isPib}
                          onClick={() => toggleGapColumn(column.id)}
                        >
                          <span className="sci-data-table-check" aria-hidden="true">
                            {checked ? "■" : ""}
                          </span>
                          <span>
                            {column.label}
                            {isPib ? " · ПІБ" : ""}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              ) : null}
            </div>

            <Button
              variant="outlined"
              startIcon={<SkipNextOutlinedIcon />}
              disabled={isLoading || !rows.length}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => void goToEmptyCell("first")}
            >
              Перша порожня
            </Button>
            <Button
              variant="outlined"
              startIcon={<ArrowRightOutlinedIcon />}
              disabled={isLoading || !rows.length}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => void goToEmptyCell("next")}
            >
              Наступна порожня
            </Button>
            <Button
              variant="outlined"
              disabled={!emptySearchActive}
              onClick={stopEmptySearch}
            >
              Стоп
            </Button>
          </div>
        </div>
      </header>

      {isLoading ? <LinearProgress sx={{ mb: 1 }} /> : null}

      <Alert severity="info" className="personnel-page-alert anketa-status-alert">
        <div className="anketa-status-bar">
          <p className="anketa-status-message">{message}</p>
          {snapshot ? (
            <div className="anketa-status-metrics" aria-label="Статистика">
              <span>кол. {gapStats.columns}</span>
              <span>∅ {gapStats.emptyCells}</span>
              <span>осіб {gapStats.personsWithGaps}</span>
              <span>рядків {gapStats.totalRows}</span>
              {emptySearchActive ? <span>пошук {emptyCount}</span> : null}
              {deferredGapKeys.length ? (
                <span>відкл. {deferredGapKeys.length}</span>
              ) : null}
            </div>
          ) : null}
        </div>
      </Alert>

      <div
        className={[
          "anketa-workspace",
          "staff-list-gaps-workspace",
          personPanelVisible ? "has-questionnaire-dock" : "",
        ]
          .filter(Boolean)
          .join(" ")}
      >
        <section className="analytics-panel anketa-data-table-panel">
          {focusedEmpty ? (
            <div className="panel-heading anketa-table-heading">
              <span className="personnel-list-questionnaire-count">
                фокус · {focusedEmpty.pib} · {focusedEmpty.header}
              </span>
            </div>
          ) : null}
          <div className="anketa-data-table-wrap">
            <MaterialReactTable table={table} />
          </div>
        </section>

        {personPanelVisible ? (
          <>
            <button
              type="button"
              className="anketa-person-panel-backdrop"
              aria-label="Закрити анкету"
              onClick={() => setPersonPanelOpen(false)}
            />
            <StaffListAnketaDock
              key={focusedAnketaRow?.__rowId || "staff-list-anketa-dock"}
              anketaRow={focusedAnketaRow}
              personLabel={focusedEmpty?.pib}
              onClose={() => setPersonPanelOpen(false)}
              onMessage={setMessage}
            />
          </>
        ) : null}
      </div>
    </main>
  );
}
