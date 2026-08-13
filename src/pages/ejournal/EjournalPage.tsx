import { useEffect, useMemo, useState, type FocusEvent } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  Divider,
  LinearProgress,
  MenuItem,
  Select,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { UploadFileOutlinedIcon } from "@/components/sci/icons";
import type { MRT_ColumnDef } from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import { api, type BackendEjournalImport } from "../../api";
import {
  type ExcelWorkbookSnapshot,
  createWorkbookDebugPayload,
  hasRowData,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import { cellValueToJson } from "../../shared/format";
import type {
  DbPreviewState,
  EjournalColumn,
  EjournalPreviewRow,
} from "./ejournalTypes";
import {
  buildImportColumns,
  localRowsToPreviewRows,
  parseDbColumns,
  previewValueToDisplay,
} from "./ejournalUtils";
import {
  analyzeEjournalWorkflow,
  exportEjournalTemplateWithProtocol,
  resetProcessedMovementRegistry,
} from "./ejournalWorkflow";
import {
  buildPersonSummary,
  createDefaultActionForm,
  personActions,
  type PersonAction,
  type PersonActionForm,
} from "../personnel/personnelUtils";

export function EjournalPage() {
  const [snapshot, setSnapshot] = useState<ExcelWorkbookSnapshot | null>(null);
  const [templateSnapshot, setTemplateSnapshot] =
    useState<ExcelWorkbookSnapshot | null>(null);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [imports, setImports] = useState<BackendEjournalImport[]>([]);
  const [selectedImportId, setSelectedImportId] = useState("");
  const [selectedDbSheetId, setSelectedDbSheetId] = useState("");
  const [dbPreview, setDbPreview] = useState<DbPreviewState | null>(null);
  const [selectedDbRowId, setSelectedDbRowId] = useState("");
  const [activePersonAction, setActivePersonAction] =
    useState<PersonAction | null>(null);
  const [personActionForm, setPersonActionForm] = useState<PersonActionForm>(
    () => createDefaultActionForm(),
  );
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const [isLoading, setIsLoading] = useState(false);
  const [movementRegistryVersion, setMovementRegistryVersion] = useState(0);

  const activeSheet =
    snapshot?.sheets.find((sheet) => sheet.sheetIndex === activeSheetIndex) ??
    snapshot?.sheets[0];
  const activeRows = useMemo(
    () => activeSheet?.rows.filter((row) => hasRowData(row.values)) ?? [],
    [activeSheet],
  );
  const activeColumns = useMemo<EjournalColumn[]>(
    () => (activeSheet ? buildImportColumns(activeSheet) : []),
    [activeSheet],
  );
  const selectedImport = useMemo(
    () => imports.find((item) => item.id === selectedImportId) ?? imports[0],
    [imports, selectedImportId],
  );
  const selectedDbSheet = useMemo(
    () =>
      selectedImport?.sheets.find((sheet) => sheet.id === selectedDbSheetId) ??
      selectedImport?.sheets[0],
    [selectedDbSheetId, selectedImport],
  );
  const workbookPayload = useMemo(() => {
    if (!snapshot) return null;

    return {
      name: snapshot.fileName.replace(/\.xlsx$/i, ""),
      sourceFileName: snapshot.fileName,
      notes:
        "Первинний staging-імпорт ЕЖООС. Нормалізація буде виконана після затвердження правил розбору.",
      sheets: snapshot.sheets.map((sheet) => {
        const columns = buildImportColumns(sheet);
        const rows = sheet.rows
          .filter((row) => hasRowData(row.values))
          .map((row) => ({
            excelRowNumber: row.excelRowNumber,
            values: Object.fromEntries(
              columns.map((column, index) => [
                column.key,
                cellValueToJson(row.values[index]),
              ]),
            ),
          }));

        return {
          name: sheet.sheetName,
          sheetIndex: sheet.sheetIndex,
          columns,
          rows,
        };
      }),
    };
  }, [snapshot]);
  const previewColumns = dbPreview?.columns ?? activeColumns;
  const previewRows = useMemo(
    () => dbPreview?.rows ?? localRowsToPreviewRows(activeRows, activeColumns),
    [activeColumns, activeRows, dbPreview],
  );
  const selectedPersonRow = useMemo(
    () => previewRows.find((row) => row.__dbRowId === selectedDbRowId) ?? null,
    [previewRows, selectedDbRowId],
  );
  const selectedPerson = useMemo(
    () => buildPersonSummary(selectedPersonRow),
    [selectedPersonRow],
  );
  const workflowAnalysis = useMemo(
    () => {
      void movementRegistryVersion;
      return snapshot
        ? analyzeEjournalWorkflow(snapshot, templateSnapshot)
        : null;
    },
    [snapshot, templateSnapshot, movementRegistryVersion],
  );

  async function saveDbCell(rowId: string, key: string, value: string) {
    try {
      const updatedRow = await api.updateEjournalRowValues(rowId, {
        [key]: value,
      });
      setDbPreview((currentPreview) => {
        if (!currentPreview) return currentPreview;

        return {
          ...currentPreview,
          rows: currentPreview.rows.map((row) =>
            row.__dbRowId === rowId
              ? {
                  ...row,
                  ...updatedRow.values,
                  __dbRowId: rowId,
                }
              : row,
          ),
        };
      });
      setMessage(`Збережено зміну в БД: ${key}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося зберегти зміну в БД.",
      );
    }
  }

  async function submitPersonAction() {
    if (!selectedPersonRow?.__dbRowId || !activePersonAction) return;

    setIsLoading(true);
    try {
      const payload = {
        actionType: activePersonAction.type,
        validFrom: personActionForm.validFrom || undefined,
        validTo: personActionForm.validTo || undefined,
        reason: personActionForm.reason || undefined,
        place: personActionForm.place || undefined,
        note: personActionForm.note || undefined,
        positionIndex: personActionForm.positionIndex || undefined,
        positionTitle: personActionForm.positionTitle || undefined,
        rank: personActionForm.rank || undefined,
      };

      await api.createEjournalRowAction(selectedPersonRow.__dbRowId, payload);
      setMessage(
        `Дію збережено: ${activePersonAction.label} · ${selectedPerson.name}.`,
      );
      setActivePersonAction(null);
      setPersonActionForm(createDefaultActionForm());
      await loadImports();
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося виконати дію для особи.",
      );
    } finally {
      setIsLoading(false);
    }
  }

  const tableColumns = useMemo<MRT_ColumnDef<EjournalPreviewRow>[]>(
    () => [
      {
        id: "excel-row-number",
        accessorKey: "__rowNumber",
        header: "#",
        Cell: ({ cell }) => previewValueToDisplay(cell.getValue()),
        enableEditing: false,
        size: 80,
      },
      ...previewColumns.map((column) => ({
        id: column.key,
        accessorKey: column.key,
        header: column.label || column.letter || "",
        Cell: ({ cell }: { cell: { getValue: () => unknown } }) =>
          previewValueToDisplay(cell.getValue()),
        muiEditTextFieldProps: ({
          row,
        }: {
          row: { original: EjournalPreviewRow };
        }) => ({
          multiline: true,
          minRows: 1,
          maxRows: 5,
          onBlur: (
            event: FocusEvent<HTMLInputElement | HTMLTextAreaElement>,
          ) => {
            const rowId = row.original.__dbRowId;
            const nextValue = event.target.value;
            const previousValue = row.original[column.key];

            if (!rowId || String(previousValue ?? "") === nextValue) return;
            void saveDbCell(rowId, column.key, nextValue);
          },
        }),
        size: (column.label || column.letter || "").length > 28 ? 260 : 180,
      })),
    ],
    [previewColumns],
  );
  const table = useMaterialReactTable({
    columns: tableColumns,
    data: previewRows,
    editDisplayMode: "cell",
    enableColumnResizing: true,
    enableColumnVirtualization: true,
    enableEditing: Boolean(dbPreview),
    enableRowVirtualization: true,
    enableStickyHeader: true,
    initialState: {
      density: "compact",
      pagination: { pageIndex: 0, pageSize: 25 },
    },
    muiTablePaperProps: {
      elevation: 0,
      sx: { backgroundColor: "transparent" },
    },
    muiTableContainerProps: {
      sx: { maxHeight: 560, backgroundColor: "transparent" },
    },
    muiTableHeadCellProps: {
      sx: {
        backgroundColor: "#131311",
        color: "#d9d49d",
        borderColor: "rgba(230,224,190,0.12)",
      },
    },
    muiTableBodyCellProps: {
      sx: { color: "#f2eee1", borderColor: "rgba(230,224,190,0.1)" },
    },
    muiTableBodyRowProps: ({ row }: any) => ({
      onClick: () => {
        if (row.original.__dbRowId) setSelectedDbRowId(row.original.__dbRowId);
      },
      sx: {
        cursor: row.original.__dbRowId ? "pointer" : "default",
        "& td": {
          backgroundColor:
            row.original.__dbRowId === selectedDbRowId
              ? "rgba(214, 215, 133, 0.16)"
              : "#11110f",
        },
        "&:hover td": {
          backgroundColor:
            row.original.__dbRowId === selectedDbRowId
              ? "rgba(214, 215, 133, 0.2)"
              : "rgba(214, 215, 133, 0.08)",
        },
      },
    }),
    muiTopToolbarProps: { sx: { backgroundColor: "rgba(17,17,15,0.92)" } },
    muiBottomToolbarProps: { sx: { backgroundColor: "rgba(17,17,15,0.92)" } },
  });

  useEffect(() => {
    if (!selectedImportId && imports[0]) {
      setSelectedImportId(imports[0].id);
    }
  }, [imports, selectedImportId]);

  useEffect(() => {
    if (!selectedDbSheetId && selectedImport?.sheets[0]) {
      setSelectedDbSheetId(selectedImport.sheets[0].id);
    }
  }, [selectedDbSheetId, selectedImport]);

  const loadImports = async () => {
    setIsLoading(true);
    try {
      const nextImports = await api.listEjournalImports();
      setImports(nextImports);
      setSelectedImportId(
        (currentImportId) => currentImportId || nextImports[0]?.id || "",
      );
      setSelectedDbSheetId(
        (currentSheetId) =>
          currentSheetId || nextImports[0]?.sheets[0]?.id || "",
      );
      setMessage(`У БД знайдено імпортів ЕЖООС: ${nextImports.length}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати імпорти ЕЖООС.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    void loadImports();
  }, []);

  const loadFile = async (file: File | undefined) => {
    if (!file) return;

    setIsLoading(true);
    try {
      const nextSnapshot = await readWorkbookSnapshot(file);
      setSnapshot(nextSnapshot);
      setDbPreview(null);
      setActiveSheetIndex(nextSnapshot.sheets[0]?.sheetIndex ?? 0);
      setMessage(
        `Розпарсено ${nextSnapshot.fileName}: ${nextSnapshot.sheets.length} вкладок.`,
      );
      console.groupCollapsed(
        `[army-grid] EЖООС parsed workbook: ${nextSnapshot.fileName}`,
      );
      console.log(createWorkbookDebugPayload(nextSnapshot));
      console.groupEnd();
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати ЕЖООС Excel.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const loadTemplateFile = async (file: File | undefined) => {
    if (!file) return;

    setIsLoading(true);
    try {
      const nextTemplate = await readWorkbookSnapshot(file);
      setTemplateSnapshot(nextTemplate);
      setMessage(
        `Завантажено шаблон ЕЖООС: ${nextTemplate.fileName}. Вкладок: ${nextTemplate.sheets.length}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати Excel-шаблон ЕЖООС.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const saveImport = async () => {
    if (!workbookPayload) return;

    setIsLoading(true);
    try {
      console.groupCollapsed(
        `[army-grid] EЖООС import payload: ${workbookPayload.sourceFileName}`,
      );
      console.log(workbookPayload);
      console.groupEnd();
      const createdImport = await api.importEjournalWorkbook(workbookPayload);
      setMessage(
        `Імпорт записано в БД: ${createdImport.name}. Вкладок: ${createdImport.sheetCount}, рядків: ${createdImport.totalRows}.`,
      );
      await loadImports();
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося записати імпорт ЕЖООС у БД.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const loadDbSheetPreview = async () => {
    if (!selectedDbSheet) return;

    setIsLoading(true);
    try {
      const response = await api.listEjournalSheetRows(selectedDbSheet.id, {
        limit: 500,
        offset: 0,
      });
      const columns = parseDbColumns(response.columns);
      const rows = response.items.map((row) => ({
        __dbRowId: row.id,
        __rowNumber: row.excelRowNumber,
        ...row.values,
      }));
      setSelectedDbRowId(rows[0]?.__dbRowId ?? "");
      setDbPreview({
        sheet: selectedDbSheet,
        columns,
        rows,
        total: response.total,
        offset: response.offset,
        limit: response.limit,
      });
      setMessage(
        `Показано з БД: ${selectedDbSheet.name}. ${rows.length} з ${response.total} рядків.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося завантажити рядки вкладки з БД.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const exportPreparedTemplate = async () => {
    if (!snapshot || !templateSnapshot || !workflowAnalysis) return;

    setIsLoading(true);
    try {
      await exportEjournalTemplateWithProtocol(
        templateSnapshot,
        snapshot,
        workflowAnalysis,
      );
      setMovementRegistryVersion((version) => version + 1);
      setMessage(
        `Експортовано шаблон ЕЖООС. Нових рухів позначено виконаними: ${workflowAnalysis.pending.length}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати підготовлений шаблон ЕЖООС.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const resetMovementRegistry = () => {
    resetProcessedMovementRegistry();
    setMovementRegistryVersion((version) => version + 1);
    setMessage("Реєстр виконаних рухів ЕЖООС очищено для цього браузера.");
  };

  return (
    <main className="main-panel">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            ЕЖООС
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Електронний журнал обліку особового складу · staging-імпорт перед
            нормалізацією
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <Button
            component="label"
            variant="outlined"
            startIcon={<UploadFileOutlinedIcon />}
          >
            Завантажити Рух
            <input
              hidden
              type="file"
              accept=".xlsx,.xls"
              onChange={(event) => void loadFile(event.target.files?.[0])}
            />
          </Button>
          <Button
            component="label"
            variant="outlined"
            startIcon={<UploadFileOutlinedIcon />}
          >
            Завантажити шаблон
            <input
              hidden
              type="file"
              accept=".xlsx,.xls"
              onChange={(event) =>
                void loadTemplateFile(event.target.files?.[0])
              }
            />
          </Button>
          <Button variant="outlined" onClick={() => void loadImports()}>
            Оновити імпорти
          </Button>
          <Button
            variant="outlined"
            startIcon={<FileDownloadOutlinedIcon />}
            disabled={!snapshot || !templateSnapshot || !workflowAnalysis}
            onClick={() => void exportPreparedTemplate()}
          >
            Експорт шаблону
          </Button>
          <Button
            variant="contained"
            disabled={!workbookPayload}
            onClick={() => void saveImport()}
            sx={{ color: "#1a1a14" }}
          >
            Записати в БД
          </Button>
        </Stack>
      </header>
      {isLoading && <LinearProgress color="primary" />}
      <Alert
        severity={snapshot ? "success" : "info"}
        variant="outlined"
        sx={{ mb: 2 }}
      >
        {message}
      </Alert>

      <section className="ejournal-layout">
        <div className="analytics-panel ejournal-workflow-panel">
          <div className="panel-heading">Рух → шаблон</div>
          <div className="structure-grid">
            <div>
              <strong>{workflowAnalysis?.pending.length ?? 0}</strong>
              <span>нових рухів</span>
            </div>
            <div>
              <strong>{workflowAnalysis?.conflicts.length ?? 0}</strong>
              <span>конфліктів</span>
            </div>
            <div>
              <strong>{workflowAnalysis?.archivePeriods.length ?? 0}</strong>
              <span>archive</span>
            </div>
          </div>
          <div className="ejournal-workflow-files">
            <span>Рух: {snapshot?.fileName ?? "не завантажено"}</span>
            <span>Шаблон: {templateSnapshot?.fileName ?? "не завантажено"}</span>
          </div>
          <div className="ejournal-workflow-actions">
            <Button
              size="small"
              variant="outlined"
              disabled={!workflowAnalysis}
              onClick={resetMovementRegistry}
            >
              Скинути реєстр рухів
            </Button>
          </div>
        </div>
        <div className="analytics-panel">
          <div className="panel-heading">Вкладки файлу</div>
          <div className="ejournal-sheet-list">
            {snapshot ? (
              snapshot.sheets.map((sheet) => (
                <button
                  className={
                    sheet.sheetIndex === activeSheetIndex ? "active" : ""
                  }
                  key={`${sheet.sheetIndex}-${sheet.sheetName}`}
                  type="button"
                  onClick={() => setActiveSheetIndex(sheet.sheetIndex)}
                >
                  <strong>{sheet.sheetName}</strong>
                  <span>
                    {sheet.rows.filter((row) => hasRowData(row.values)).length}{" "}
                    рядків · {sheet.columnCount} колонок
                  </span>
                </button>
              ))
            ) : (
              <Typography variant="body2" color="text.secondary">
                Завантажте файл ЕЖООС, щоб побачити 16 вкладок і структуру
                колонок.
              </Typography>
            )}
          </div>
        </div>
        <div className="analytics-panel">
          <div className="panel-heading">Що формує експорт</div>
          <div className="ejournal-plan">
            {[
              ["1", "Оновлюємо поточні 1. ШПО і 6. Табель із фінального sh."],
              [
                "2",
                "Додаємо кадрову історію з Рух у 2. ООС без повторів по № руху.",
              ],
              [
                "3",
                "Переносимо ПЕРЕВ і ЗВІЛЬН у накопичувальний 3. Виключені.",
              ],
              ["4", "Закриваємо/додаємо 5. Тимчасово відсутні тільки з archive."],
            ].map(([step, text]) => (
              <div key={step}>
                <span>{step}</span>
                <p>{text}</p>
              </div>
            ))}
          </div>
        </div>
        <div className="analytics-panel">
          <div className="panel-heading">БД</div>
          <div className="structure-grid">
            <div>
              <strong>{imports.length}</strong>
              <span>імпортів</span>
            </div>
            <div>
              <strong>
                {imports.reduce((sum, item) => sum + item.sheetCount, 0)}
              </strong>
              <span>вкладок</span>
            </div>
            <div>
              <strong>
                {imports.reduce((sum, item) => sum + item.totalRows, 0)}
              </strong>
              <span>рядків</span>
            </div>
          </div>
          <div className="ejournal-db-controls">
            <label>
              <span>Імпорт</span>
              <Select
                value={selectedImport?.id ?? ""}
                onChange={(event) => {
                  const nextImport = imports.find(
                    (item) => item.id === event.target.value,
                  );
                  setSelectedImportId(event.target.value);
                  setSelectedDbSheetId(nextImport?.sheets[0]?.id ?? "");
                  setDbPreview(null);
                }}
              >
                {imports.map((item) => (
                  <MenuItem key={item.id} value={item.id}>
                    {item.sourceFileName ?? item.name} ·{" "}
                    {new Date(item.createdAt).toLocaleString("uk-UA")}
                  </MenuItem>
                ))}
              </Select>
            </label>
            <label>
              <span>Вкладка</span>
              <Select
                value={selectedDbSheet?.id ?? ""}
                onChange={(event) => {
                  setSelectedDbSheetId(event.target.value);
                  setDbPreview(null);
                }}
              >
                {(selectedImport?.sheets ?? []).map((sheet) => (
                  <MenuItem key={sheet.id} value={sheet.id}>
                    {sheet.name} · {sheet.rowCount} рядків
                  </MenuItem>
                ))}
              </Select>
            </label>
            <Button
              disabled={!selectedDbSheet}
              variant="outlined"
              onClick={() => void loadDbSheetPreview()}
            >
              Показати з БД
            </Button>
          </div>
          <pre className="debug-json">
            {JSON.stringify(imports.slice(0, 3), null, 2)}
          </pre>
        </div>
      </section>

      {workflowAnalysis && (
        <section className="ejournal-workflow-results">
          <div className="analytics-panel">
            <div className="panel-heading">Нові рухи до експорту</div>
            <div className="ejournal-type-grid">
              {Object.entries(workflowAnalysis.statsByType).map(
                ([type, count]) => (
                  <div key={type}>
                    <strong>{count}</strong>
                    <span>{type}</span>
                  </div>
                ),
              )}
              {workflowAnalysis.pending.length === 0 && (
                <Typography variant="body2" color="text.secondary">
                  Нових рухів немає.
                </Typography>
              )}
            </div>
            <div className="ejournal-event-list">
              {workflowAnalysis.pending.slice(0, 12).map((event) => (
                <div key={event.movementNumber}>
                  <strong>№ {event.movementNumber}</strong>
                  <span>
                    {event.type} · {event.name} · {event.orderDate || "без дати"}
                  </span>
                </div>
              ))}
            </div>
          </div>
          <div className="analytics-panel">
            <div className="panel-heading">Конфлікти</div>
            <div className="ejournal-event-list conflict">
              {workflowAnalysis.conflicts.length === 0 ? (
                <Typography variant="body2" color="text.secondary">
                  Конфліктів не знайдено.
                </Typography>
              ) : (
                workflowAnalysis.conflicts.slice(0, 12).map((event) => (
                  <div key={event.movementNumber}>
                    <strong>№ {event.movementNumber}</strong>
                    <span>
                      {event.type} · {event.name} · {event.conflict}
                    </span>
                  </div>
                ))
              )}
            </div>
          </div>
        </section>
      )}

      {dbPreview && (
        <section className="person-action-panel">
          <div className="person-action-main">
            <div>
              <div className="panel-heading">Панель особи</div>
              <Typography component="h2" variant="h5">
                {selectedPerson.name}
              </Typography>
              <div className="person-action-tags">
                {selectedPerson.rank && (
                  <Chip label={selectedPerson.rank} size="small" />
                )}
                {selectedPerson.externalId && (
                  <Chip
                    label={`ID: ${selectedPerson.externalId}`}
                    size="small"
                  />
                )}
                {selectedPerson.positionIndex && (
                  <Chip
                    label={`Посада: ${selectedPerson.positionIndex}`}
                    size="small"
                  />
                )}
                {selectedPerson.serviceType && (
                  <Chip label={selectedPerson.serviceType} size="small" />
                )}
              </div>
            </div>
            <div className="person-action-fields">
              <span>
                <strong>Дата народження</strong>
                {selectedPerson.birthDate || "—"}
              </span>
              <span>
                <strong>РНОКПП</strong>
                {selectedPerson.rnokpp || "—"}
              </span>
              <span>
                <strong>Дислокація</strong>
                {selectedPerson.location || "—"}
              </span>
            </div>
          </div>
          <Divider />
          <div className="person-action-buttons">
            {personActions.map((action) => (
              <Button
                key={action.type}
                disabled={!selectedPersonRow}
                variant={
                  activePersonAction?.type === action.type
                    ? "contained"
                    : "outlined"
                }
                size="small"
                onClick={() => {
                  setActivePersonAction(action);
                  setPersonActionForm(createDefaultActionForm());
                }}
                sx={
                  activePersonAction?.type === action.type
                    ? { color: "#1a1a14" }
                    : undefined
                }
              >
                {action.label}
              </Button>
            ))}
          </div>
          {activePersonAction && (
            <div className="person-action-form">
              <div className="panel-heading">{activePersonAction.label}</div>
              <label>
                <span>Дата від</span>
                <input
                  type="date"
                  value={personActionForm.validFrom}
                  onChange={(event) =>
                    setPersonActionForm((form) => ({
                      ...form,
                      validFrom: event.target.value,
                    }))
                  }
                />
              </label>
              <label>
                <span>Дата до</span>
                <input
                  type="date"
                  value={personActionForm.validTo}
                  onChange={(event) =>
                    setPersonActionForm((form) => ({
                      ...form,
                      validTo: event.target.value,
                    }))
                  }
                />
              </label>
              <label>
                <span>Причина / статус</span>
                <input
                  value={personActionForm.reason}
                  onChange={(event) =>
                    setPersonActionForm((form) => ({
                      ...form,
                      reason: event.target.value,
                    }))
                  }
                  placeholder={activePersonAction.label}
                />
              </label>
              <label>
                <span>Місце</span>
                <input
                  value={personActionForm.place}
                  onChange={(event) =>
                    setPersonActionForm((form) => ({
                      ...form,
                      place: event.target.value,
                    }))
                  }
                  placeholder="Місце / підрозділ / заклад"
                />
              </label>
              {activePersonAction.type === "POSITION_CHANGE" && (
                <>
                  <label>
                    <span>Індекс посади</span>
                    <input
                      value={personActionForm.positionIndex}
                      onChange={(event) =>
                        setPersonActionForm((form) => ({
                          ...form,
                          positionIndex: event.target.value,
                        }))
                      }
                      placeholder={selectedPerson.positionIndex || "0600000"}
                    />
                  </label>
                  <label>
                    <span>Назва посади</span>
                    <input
                      value={personActionForm.positionTitle}
                      onChange={(event) =>
                        setPersonActionForm((form) => ({
                          ...form,
                          positionTitle: event.target.value,
                        }))
                      }
                      placeholder="Нова посада"
                    />
                  </label>
                </>
              )}
              {activePersonAction.type === "RANK_CHANGE" && (
                <label>
                  <span>Нове звання</span>
                  <input
                    value={personActionForm.rank}
                    onChange={(event) =>
                      setPersonActionForm((form) => ({
                        ...form,
                        rank: event.target.value,
                      }))
                    }
                    placeholder={selectedPerson.rank || "солдат"}
                  />
                </label>
              )}
              <label className="person-action-note">
                <span>Примітка</span>
                <textarea
                  value={personActionForm.note}
                  onChange={(event) =>
                    setPersonActionForm((form) => ({
                      ...form,
                      note: event.target.value,
                    }))
                  }
                  placeholder="Наказ, уточнення або коментар"
                />
              </label>
              <div className="person-action-submit">
                <Button
                  variant="outlined"
                  onClick={() => setActivePersonAction(null)}
                >
                  Скасувати
                </Button>
                <Button
                  variant="contained"
                  onClick={() => void submitPersonAction()}
                  sx={{ color: "#1a1a14" }}
                >
                  Зберегти дію
                </Button>
              </div>
            </div>
          )}
        </section>
      )}

      <section className="panel table-panel">
        <div className="panel-heading">
          {dbPreview
            ? `${dbPreview.sheet.name} · preview з БД · ${dbPreview.rows.length}/${dbPreview.total}`
            : activeSheet
              ? `${activeSheet.sheetName} · preview з Excel`
              : "Preview ЕЖООС"}
        </div>
        <div className="panel-body">
          {dbPreview || activeSheet ? (
            <MaterialReactTable table={table} />
          ) : (
            <div className="drop-zone">
              <Box>
                <CloudUploadOutlinedIcon color="disabled" />
                <Typography variant="body2">Додайте `.xlsx` ЕЖООС</Typography>
              </Box>
            </div>
          )}
        </div>
      </section>
    </main>
  );
}
