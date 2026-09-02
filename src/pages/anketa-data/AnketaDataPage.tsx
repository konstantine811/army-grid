import { useMemo } from "react";
import {
  Alert,
  Box,
  Button,
  LinearProgress,
  Typography,
} from "@/components/sci/SciPrimitives";
import {
  ArrowRightOutlinedIcon,
  SkipNextOutlinedIcon,
} from "@/components/sci/icons";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import { type AnketaRow } from "./anketaSheet";
import { AnketaPersonSidePanel } from "./AnketaPersonSidePanel";
import { exportAnketaSheetExcel } from "./anketaExcelExport";
import {
  buildAnketaFocusedCell,
  buildAnketaTableColumns,
  buildAnketaTdProps,
} from "./buildAnketaTableColumns";
import { AnketaGapColumnsMenu } from "./components/AnketaGapColumnsMenu";
import { useAnketaGapColumnsMenu } from "./hooks/useAnketaGapColumnsMenu";
import { useAnketaGapSearch } from "./hooks/useAnketaGapSearch";
import { useAnketaSheetLoader } from "./hooks/useAnketaSheetLoader";
import { useAuth } from "../../auth/AuthProvider";

export function AnketaDataPage() {
  const { canEditArea } = useAuth();
  const canEdit = canEditArea("anketaData");

  const sheet = useAnketaSheetLoader();
  const gapColumns = useAnketaGapColumnsMenu();
  const gap = useAnketaGapSearch({
    rows: sheet.rows,
    gapColumnKeys: gapColumns.gapColumnKeys,
    missingQuestionnaireNames: sheet.missingQuestionnaireNames,
    setMissingQuestionnaireNames: sheet.setMissingQuestionnaireNames,
    appsScriptUrl: sheet.appsScriptUrl,
    persistSnapshot: sheet.persistSnapshot,
    setMessage: sheet.setMessage,
    setIsSyncing: sheet.setIsSyncing,
    setEditsCount: sheet.setEditsCount,
    canEdit,
    setGapColumnsOpen: gapColumns.setGapColumnsOpen,
    gapColumnsOpen: gapColumns.gapColumnsOpen,
  });

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
            <Button
              variant="outlined"
              disabled={
                !canEdit ||
                sheet.isLoading ||
                gap.isFillingAbsent ||
                !sheet.rows.length
              }
              onClick={() => void gap.fillAbsentQuestionnaireCells()}
              title="Для осіб без анкет записати «дані відсутні» у порожні вибрані колонки. Якщо анкета вже з’явилась — зняти «дані відсутні», щоб пошук міг заповнити."
            >
              <span className="anketa-label-full">
                {gap.isFillingAbsent
                  ? "Пишу…"
                  : "Без анкет → дані відсутні"}
              </span>
              <span className="anketa-label-short" aria-hidden="true">
                {gap.isFillingAbsent ? "…" : "∅ → н/д"}
              </span>
            </Button>
          </div>
        </div>
      </header>

      {sheet.isLoading ||
      sheet.isSyncing ||
      sheet.isMergingPersonnel ||
      sheet.isAddingFromEjoos ||
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
              <span title="Осіб з пропусками (пошук порожніх, без «дані відсутні»)">
                осіб {gap.gapStats.personsWithGaps}
              </span>
              <span title="Службовці, у яких хоча б одне вибране поле ще зовсім порожнє">
                ще пусті {gap.gapStats.personsWithBlankFields}
              </span>
              <span title="Службовці, у яких усі вибрані поля ще зовсім порожні">
                усі пусті {gap.gapStats.personsFullyBlank}
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
