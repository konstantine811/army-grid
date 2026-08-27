import { useMemo, useState } from "react";
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
import { useAuth } from "../../auth/AuthProvider";

export function AnketaDataPage() {
  const { canEditArea, isAdmin } = useAuth();
  const canEdit = canEditArea("anketaData");
  const [showSyncHelp, setShowSyncHelp] = useState(false);
  const [syncExpanded, setSyncExpanded] = useState(false);

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
    canEdit,
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

  const personPanelVisible = Boolean(
    gap.personPanelOpen && gap.focusedAnketaRow,
  );

  return (
    <main className="main-panel anketa-data-page">
      <header className="topbar anketa-topbar">
        <Box className="anketa-topbar-copy">
          <Typography component="h1" variant="h4">
            Анкетні дані
          </Typography>
          <Typography
            variant="body2"
            color="text.secondary"
            className="anketa-topbar-subtitle"
          >
            Заповнення пропусків · локальне редагування · експорт / Google sync
          </Typography>
        </Box>

        <div className="anketa-toolbar">
          <div className="anketa-toolbar-group" aria-label="Пропуски">
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
              title="Перша порожня"
            >
              <span className="anketa-label-full">Перша порожня</span>
              <span className="anketa-label-short" aria-hidden="true">
                1-ша ∅
              </span>
            </Button>
            <Button
              variant="outlined"
              startIcon={<ArrowRightOutlinedIcon />}
              disabled={sheet.isLoading || !sheet.rows.length}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => gap.goToEmptyCell("next")}
              title="Наступна порожня"
            >
              <span className="anketa-label-full">Наступна порожня</span>
              <span className="anketa-label-short" aria-hidden="true">
                Далі ∅
              </span>
            </Button>
            <Button
              variant="outlined"
              startIcon={<SkipNextOutlinedIcon />}
              disabled={sheet.isLoading || !sheet.rows.length}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => gap.goToEmptyCell("nextPerson")}
              title="Пропустити поточного службовця і перейти до наступного з пропусками"
            >
              <span className="anketa-label-full">Наступний службовець</span>
              <span className="anketa-label-short" aria-hidden="true">
                Наст. особа
              </span>
            </Button>
            <Button
              variant="outlined"
              disabled={!gap.emptySearchActive}
              onClick={gap.stopEmptySearch}
              title="Стоп · Esc"
            >
              Стоп
            </Button>
          </div>

          <div className="anketa-toolbar-group" aria-label="Дані">
            <Button
              variant="outlined"
              startIcon={<SyncAltOutlinedIcon />}
              disabled={
                sheet.isLoading ||
                sheet.isMergingFromPersonnel ||
                sheet.isMergingPersonnel ||
                sheet.isPushingStaffSheet ||
                !sheet.rows.length ||
                !canEdit
              }
              onClick={() => void sheet.mergeFromPersonnel()}
              title="Доповнити порожні поля анкет з БД особового складу (РНОКПП, дата народження тощо)"
            >
              <span className="anketa-label-full">
                {sheet.isMergingFromPersonnel ? "З ООС…" : "З особового складу"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {sheet.isMergingFromPersonnel ? "З ООС…" : "← ООС"}
              </span>
            </Button>
            <Button
              variant="outlined"
              startIcon={<SyncAltOutlinedIcon />}
              disabled={
                sheet.isLoading ||
                sheet.isMergingPersonnel ||
                sheet.isMergingFromPersonnel ||
                sheet.isPushingStaffSheet ||
                !sheet.rows.length ||
                !canEdit
              }
              onClick={() => void sheet.mergeToPersonnel()}
              title="Оновити особовий склад"
            >
              <span className="anketa-label-full">
                {sheet.isMergingPersonnel
                  ? "Злиття…"
                  : "Оновити особовий склад"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {sheet.isMergingPersonnel ? "Злиття…" : "→ ООС"}
              </span>
            </Button>
            <Button
              component="label"
              variant="contained"
              startIcon={<FileUploadOutlinedIcon />}
              disabled={
                sheet.isLoading ||
                sheet.isImportingStaffSheet ||
                sheet.isPushingStaffSheet ||
                sheet.isDownloadingStaffSheet ||
                !canEdit
              }
              title="Єдиний імпорт «Штатки» (.xlsx) — для анкет, ранкового звіту та БД персоналу"
              sx={{ color: "#1a1a14" }}
            >
              <span className="anketa-label-full">
                {sheet.isImportingStaffSheet ? "Імпорт…" : "Імпорт Штатки"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {sheet.isImportingStaffSheet ? "…" : "↑ Штатка"}
              </span>
              <input
                hidden
                type="file"
                accept=".xlsx,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12"
                disabled={!canEdit}
                onChange={(event) => {
                  void sheet.importStaffSheetFile(event.target.files?.[0]);
                  event.currentTarget.value = "";
                }}
              />
            </Button>
            <Button
              variant="outlined"
              startIcon={<FileDownloadOutlinedIcon />}
              disabled={
                sheet.isLoading ||
                sheet.isMergingPersonnel ||
                sheet.isMergingFromPersonnel ||
                sheet.isPushingStaffSheet ||
                sheet.isDownloadingStaffSheet ||
                sheet.isImportingStaffSheet ||
                !sheet.staffSheetImportName ||
                !canEdit
              }
              onClick={() => void sheet.downloadStaffSheetExcel()}
              title="Скачати «Штатку» з колонками: анкета (так), ІПН, дата, рік, повних років"
            >
              <span className="anketa-label-full">
                {sheet.isDownloadingStaffSheet ? "Штатка…" : "Скачати Штатку"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {sheet.isDownloadingStaffSheet ? "…" : "↓ Штатка"}
              </span>
            </Button>
            <Button
              variant="outlined"
              startIcon={<SyncAltOutlinedIcon />}
              disabled={
                sheet.isLoading ||
                sheet.isMergingPersonnel ||
                sheet.isMergingFromPersonnel ||
                sheet.isPushingStaffSheet ||
                sheet.isDownloadingStaffSheet ||
                sheet.isImportingStaffSheet ||
                !sheet.staffSheetImportName ||
                !canEdit
              }
              onClick={() => void sheet.pushStaffSheetEnrichment()}
              title="Записати в Google Sheet «Штатка»: анкета (так або порожньо), ІПН, дата/рік/вік"
            >
              <span className="anketa-label-full">
                {sheet.isPushingStaffSheet ? "Штатка…" : "→ Google Штатка"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {sheet.isPushingStaffSheet ? "Штатка…" : "→ Штатка"}
              </span>
            </Button>
            <Button
              variant="outlined"
              startIcon={<SyncAltOutlinedIcon />}
              disabled={sheet.isLoading || sheet.isImportingStaffSheet}
              onClick={() => void handleLoadFromGoogle()}
              title="Оновити таблицю анкет з Google (не Штатку)"
            >
              <span className="anketa-label-full">Оновити анкети</span>
              <span className="anketa-label-short" aria-hidden="true">
                Анкети
              </span>
            </Button>
            <Button
              component="label"
              variant="outlined"
              startIcon={<FileUploadOutlinedIcon />}
              disabled={sheet.isLoading || !canEdit}
              title="Імпорт CSV"
            >
              <span className="anketa-label-full">CSV файл</span>
              <span className="anketa-label-short" aria-hidden="true">
                CSV
              </span>
              <input
                hidden
                type="file"
                accept=".csv,text/csv"
                disabled={!canEdit}
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
              title={
                gap.focusedEmpty
                  ? `Google · ${gap.focusedEmpty.a1}`
                  : "Відкрити таблицю"
              }
            >
              <span className="anketa-label-full">
                {gap.focusedEmpty
                  ? `Google · ${gap.focusedEmpty.a1}`
                  : "Відкрити таблицю"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {gap.focusedEmpty ? gap.focusedEmpty.a1 : "Таблиця"}
              </span>
            </Button>
          </div>
        </div>
      </header>

      {sheet.isLoading ||
      sheet.isSyncing ||
      sheet.isMergingPersonnel ||
      sheet.isMergingFromPersonnel ||
      sheet.isPushingStaffSheet ||
      sheet.isDownloadingStaffSheet ||
      sheet.isImportingStaffSheet ? (
        <LinearProgress sx={{ mb: 1 }} />
      ) : null}

      <Alert severity="info" className="personnel-page-alert anketa-status-alert">
        <div className="anketa-status-bar">
          <p className="anketa-status-message">{sheet.message}</p>
          {sheet.snapshot ? (
            <div className="anketa-status-metrics" aria-label="Статистика">
              <span title="Вибрані колонки пропусків">
                кол. {gap.gapStats.columns}
              </span>
              <span title="Порожніх ячійок">
                ∅ {gap.gapStats.emptyCells}
              </span>
              <span title="Осіб з пропусками">
                осіб {gap.gapStats.personsWithGaps}
              </span>
              <span title="Усього рядків">
                рядків {gap.gapStats.totalRows}
              </span>
              <span title="Правок у БД">правок {sheet.editsCount}</span>
              {gap.emptySearchActive ? (
                <span title="Залишилось у пошуку">
                  пошук {gap.emptyCount}
                </span>
              ) : null}
              {gap.deferredGapKeys.length ? (
                <span title="Відкладено">
                  відкл. {gap.deferredGapKeys.length}
                </span>
              ) : null}
              {sheet.dirtyCount ? (
                <span title="Несинхронізовані в сесії">
                  сесія {sheet.dirtyCount}
                </span>
              ) : null}
              {sheet.staffSheetImportName ? (
                <span
                  title={
                    sheet.staffSheetImportedAt
                      ? `${sheet.staffSheetImportName} · ${sheet.staffSheetSource || "?"} · ${new Date(sheet.staffSheetImportedAt).toLocaleString("uk-UA")}${sheet.staffSheetPersonCount ? ` · ${sheet.staffSheetPersonCount} осіб` : ""}`
                      : sheet.staffSheetImportName
                  }
                >
                  штатка ✓{sheet.staffSheetPersonCount ? ` ${sheet.staffSheetPersonCount}` : ""}
                </span>
              ) : null}
              {sheet.missingQuestionnaireNames.size ? (
                <span title="Без анкет (пропущено)">
                  без анкет {sheet.missingQuestionnaireNames.size}
                </span>
              ) : null}
            </div>
          ) : null}
        </div>
      </Alert>

      {isAdmin ? (
        <AnketaSyncPanel
          appsScriptUrl={sheet.appsScriptUrl}
          onAppsScriptUrlChange={sheet.setAppsScriptUrl}
          expanded={syncExpanded}
          onToggleExpanded={() => setSyncExpanded((value) => !value)}
          showSyncHelp={showSyncHelp}
          onToggleSyncHelp={() => setShowSyncHelp((value) => !value)}
        />
      ) : null}

      <div
        className={[
          "anketa-workspace",
          personPanelVisible ? "has-person-panel" : "",
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

        {personPanelVisible ? (
          <>
            <button
              type="button"
              className="anketa-person-panel-backdrop"
              aria-label="Закрити картку службовця"
              onClick={() => gap.setPersonPanelOpen(false)}
            />
            <AnketaPersonSidePanel
              key={gap.focusedAnketaRow?.__rowId || "anketa-person-panel"}
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
          </>
        ) : null}
      </div>
    </main>
  );
}
