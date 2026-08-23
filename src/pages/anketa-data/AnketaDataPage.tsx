import { useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  LinearProgress,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArrowRightOutlinedIcon,
  FileDownloadOutlinedIcon,
  FileUploadOutlinedIcon,
  SkipNextOutlinedIcon,
  SyncAltOutlinedIcon,
} from "@/components/sci/icons";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import { anketaSheetEditUrl, type AnketaRow } from "./anketaSheet";
import { anketaGoogleCellUrl } from "./anketaGaps";
import { AnketaPersonSidePanel } from "./AnketaPersonSidePanel";
import { exportAnketaSheetExcel } from "./anketaExcelExport";
import {
  buildAnketaFocusedCell,
  buildAnketaTableColumns,
  buildAnketaTdProps,
} from "./buildAnketaTableColumns";
import { AnketaGapColumnsMenu } from "./components/AnketaGapColumnsMenu";
import { AnketaSyncPanel } from "./components/AnketaSyncPanel";
import { useAnketaGapColumnsMenu } from "./hooks/useAnketaGapColumnsMenu";
import { useAnketaGapSearch } from "./hooks/useAnketaGapSearch";
import { useAnketaSheetLoader } from "./hooks/useAnketaSheetLoader";

export function AnketaDataPage() {
  const [showSyncHelp, setShowSyncHelp] = useState(false);

  const sheet = useAnketaSheetLoader();
  const gapColumns = useAnketaGapColumnsMenu();
  const gap = useAnketaGapSearch({
    rows: sheet.rows,
    gapColumnKeys: gapColumns.gapColumnKeys,
    missingQuestionnaireNames: sheet.missingQuestionnaireNames,
    appsScriptUrl: sheet.appsScriptUrl,
    persistSnapshot: sheet.persistSnapshot,
    setMessage: sheet.setMessage,
    setIsSyncing: sheet.setIsSyncing,
    setEditsCount: sheet.setEditsCount,
    setGapColumnsOpen: gapColumns.setGapColumnsOpen,
    gapColumnsOpen: gapColumns.gapColumnsOpen,
  });

  const handleLoadFromGoogle = async () => {
    const result = await sheet.loadFromGoogle();
    if (result?.clearSearch) gap.clearGapFocus();
  };

  const handleLoadFromCsv = async (file: File | undefined) => {
    const result = await sheet.loadFromCsvFile(file);
    if (result?.clearSearch) gap.clearGapFocus();
  };

  const exportTable = async (
    context: SciDataTableExportContext<AnketaRow>,
  ) => {
    await exportAnketaSheetExcel({
      columns: context.columns,
      rows: context.rows,
    });
    sheet.setMessage(
      `Експортовано ${context.rows.length} рядків у Excel (стиль Google Sheets).`,
    );
  };

  const columns = useMemo(
    () =>
      buildAnketaTableColumns({
        focusedEmpty: gap.focusedEmpty,
        focusEpoch: gap.focusEpoch,
        emptySearchActive: gap.emptySearchActive,
        onActivateCell: gap.activateEmptyCell,
        onCancelEdit: gap.cancelCellEdit,
        onPatchCell: (rowId, columnId, value, options) => {
          void gap.patchCell(rowId, columnId, value, options);
        },
      }),
    [
      gap.focusEpoch,
      gap.focusedEmpty,
      gap.emptySearchActive,
      gap.activateEmptyCell,
      gap.cancelCellEdit,
      gap.patchCell,
    ],
  );

  const focusedCell = buildAnketaFocusedCell(gap.focusedEmpty, gap.focusEpoch);
  const getTdProps = useMemo(
    () => buildAnketaTdProps(gapColumns.gapKeySet, gap.focusedEmpty),
    [gapColumns.gapKeySet, gap.focusedEmpty],
  );

  const table = useMaterialReactTable({
    columns,
    data: sheet.rows,
    emptyMessage: sheet.isLoading
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
    getTdProps,
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
          <AnketaGapColumnsMenu
            gapColumnKeys={gapColumns.gapColumnKeys}
            gapKeySet={gapColumns.gapKeySet}
            gapColumnsOpen={gapColumns.gapColumnsOpen}
            setGapColumnsOpen={gapColumns.setGapColumnsOpen}
            gapMenuPos={gapColumns.gapMenuPos}
            gapTriggerRef={gapColumns.gapTriggerRef}
            gapMenuRef={gapColumns.gapMenuRef}
            onToggleColumn={gapColumns.toggleGapColumn}
            onSelectAll={gapColumns.selectAllGapColumns}
            onClear={gapColumns.clearGapColumns}
          />
          <Button
            variant="outlined"
            startIcon={<SkipNextOutlinedIcon />}
            disabled={sheet.isLoading || !sheet.rows.length}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => gap.goToEmptyCell("first")}
          >
            Перша порожня
          </Button>
          <Button
            variant="outlined"
            startIcon={<ArrowRightOutlinedIcon />}
            disabled={sheet.isLoading || !sheet.rows.length}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => gap.goToEmptyCell("next")}
          >
            Наступна порожня
          </Button>
          <Button
            variant="outlined"
            startIcon={<SkipNextOutlinedIcon />}
            disabled={sheet.isLoading || !sheet.rows.length}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => gap.goToEmptyCell("nextPerson")}
            title="Пропустити поточного службовця і перейти до наступного з пропусками"
          >
            Наступний службовець
          </Button>
          <Button
            variant="outlined"
            disabled={!gap.emptySearchActive}
            onClick={gap.stopEmptySearch}
            title="Esc"
          >
            Стоп · Esc
          </Button>
          <Button
            variant="outlined"
            startIcon={<SyncAltOutlinedIcon />}
            disabled={sheet.isLoading || sheet.isMergingPersonnel || !sheet.rows.length}
            onClick={() => void sheet.mergeToPersonnel()}
          >
            {sheet.isMergingPersonnel ? "Злиття…" : "Оновити особовий склад"}
          </Button>
          <Button
            variant="outlined"
            startIcon={<SyncAltOutlinedIcon />}
            disabled={sheet.isLoading}
            onClick={() => void handleLoadFromGoogle()}
          >
            Оновити з Google
          </Button>
          <Button
            component="label"
            variant="outlined"
            startIcon={<FileUploadOutlinedIcon />}
            disabled={sheet.isLoading}
          >
            CSV файл
            <input
              hidden
              type="file"
              accept=".csv,text/csv"
              onChange={(event) => {
                void handleLoadFromCsv(event.target.files?.[0]);
                event.currentTarget.value = "";
              }}
            />
          </Button>
          <Button
            variant="outlined"
            startIcon={<FileDownloadOutlinedIcon />}
            component="a"
            href={
              gap.focusedEmpty
                ? anketaGoogleCellUrl(
                    gap.focusedEmpty.rowNumber,
                    gap.focusedEmpty.columnIndex,
                  )
                : anketaSheetEditUrl()
            }
            target="_blank"
            rel="noreferrer"
          >
            {gap.focusedEmpty ? `Google · ${gap.focusedEmpty.a1}` : "Відкрити таблицю"}
          </Button>
        </Stack>
      </header>

      {sheet.isLoading || sheet.isSyncing || sheet.isMergingPersonnel ? (
        <LinearProgress sx={{ mb: 1 }} />
      ) : null}

      <Alert severity="info" className="personnel-page-alert anketa-status-alert">
        {sheet.message}
        {sheet.snapshot
          ? ` · вибрані колонки: ${gap.gapStats.columns} · порожніх ячійок: ${gap.gapStats.emptyCells} · осіб з пропусками: ${gap.gapStats.personsWithGaps} · усього рядків/осіб: ${gap.gapStats.totalRows} · правок у БД: ${sheet.editsCount}${
              gap.emptySearchActive
                ? ` · у пошуку лишилось: ${gap.emptyCount}`
                : ""
            }${
              gap.deferredGapKeys.length
                ? ` · відкладено: ${gap.deferredGapKeys.length}`
                : ""
            }${sheet.dirtyCount ? ` · сесія: ${sheet.dirtyCount}` : ""}${
              sheet.missingQuestionnaireNames.size
                ? ` · без анкет (пропущено): ${sheet.missingQuestionnaireNames.size}`
                : ""
            }`
          : ""}
      </Alert>

      <AnketaSyncPanel
        appsScriptUrl={sheet.appsScriptUrl}
        onAppsScriptUrlChange={sheet.setAppsScriptUrl}
        showSyncHelp={showSyncHelp}
        onToggleSyncHelp={() => setShowSyncHelp((value) => !value)}
      />

      <div
        className={[
          "anketa-workspace",
          gap.personPanelOpen && gap.focusedAnketaRow ? "has-person-panel" : "",
        ]
          .filter(Boolean)
          .join(" ")}
      >
        <section className="analytics-panel anketa-data-table-panel">
          <div className="panel-heading">
            Анкети
            {sheet.snapshot ? (
              <span className="personnel-list-questionnaire-count">
                · {sheet.snapshot.rows.length}
                {gap.focusedEmpty ? ` · фокус ${gap.focusedEmpty.a1}` : ""}
              </span>
            ) : null}
          </div>
          <div className="anketa-data-table-wrap">
            <MaterialReactTable table={table} />
          </div>
        </section>

        {gap.personPanelOpen && gap.focusedAnketaRow ? (
          <AnketaPersonSidePanel
            anketaRow={gap.focusedAnketaRow}
            focusedEmpty={gap.focusedEmpty}
            gapColumnKeys={gapColumns.gapColumnKeys}
            onClose={() => gap.setPersonPanelOpen(false)}
            onFillMissing={
              gap.focusedEmpty && gap.emptySearchActive
                ? (value) => {
                    void gap.patchCell(
                      gap.focusedEmpty!.rowId,
                      gap.focusedEmpty!.columnId,
                      value,
                      { advance: true },
                    );
                  }
                : undefined
            }
            onMessage={sheet.setMessage}
          />
        ) : null}
      </div>
    </main>
  );
}
