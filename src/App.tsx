import {
  useEffect,
  useState,
  type Dispatch,
  type SetStateAction,
} from "react";
import { Box, Button, Stack, Typography } from "@/components/sci/SciPrimitives";
import { DescriptionOutlinedIcon } from "@/components/sci/icons";
import { HelpOutlineOutlinedIcon } from "@/components/sci/icons";
import {
  type ExcelRow,
  type ExcelWorkbookSnapshot,
  createWorkbookDebugPayload,
  exportStyledWorkbook,
  hasRowData,
  mergeRows,
  readWorkbookSnapshot,
} from "./excelRoundTrip";
import "./App.css";
import {
  buildDocumentRoute,
  buildPersonnelRoute,
  getCurrentRouteKey,
  getPageFromPath,
  navigateToPage,
  pagePaths,
  pushAppRoute,
  type AppPage,
} from "./app/navigation";
import { Sidebar } from "./app/layout/Sidebar";
import { formatDateTime, formatFileSize } from "./shared/format";
import { DataSources } from "./pages/import/DataSources";
import { MappingPanel } from "./pages/import/MappingPanel";
import { ValidationPanel } from "./pages/import/ValidationPanel";
import { ExcelRoundTripPanel } from "./pages/import/ExcelRoundTripPanel";
import type { DataSourceFile } from "./pages/import/types";
import { AnalyticsPage } from "./pages/analytics/AnalyticsPage";
import { PersonnelPage } from "./pages/personnel/PersonnelPage";
import { DocumentsPage } from "./pages/documents/DocumentsPage";
import { EjournalPage } from "./pages/ejournal/EjournalPage";
import type { EjournalPreviewRow } from "./pages/ejournal/ejournalTypes";
import { buildPersonSummary } from "./pages/personnel/personnelUtils";
import { BchsLabPage } from "./pages/bchs/BchsLabPage";
import { BchsPage } from "./pages/bchs/BchsPage";
import { OverviewPage } from "./pages/overview/OverviewPage";
import { ExcelFillPage } from "./pages/excel-fill/ExcelFillPage";
import { QuestionnaireParserPage } from "./pages/questionnaire-parser/QuestionnaireParserPage";
import { SocPassportPage } from "./pages/soc-passport/SocPassportPage";
import { DocumentSignatoriesSettingsPage } from "./pages/document-settings/DocumentSignatoriesSettingsPage";
import { SciScrollbars } from "./components/sci/SciScrollbars";
import { SciLiveFeedback } from "./components/sci/SciLiveFeedback";

function App() {
  const [activePage, setActivePage] = useState<AppPage>(() =>
    getPageFromPath(window.location.pathname),
  );
  const [routeKey, setRouteKey] = useState(getCurrentRouteKey);
  const [snapshot, setSnapshot] = useState<ExcelWorkbookSnapshot | null>(null);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [sheetRowsByIndex, setSheetRowsByIndex] = useState<
    Record<number, ExcelRow[]>
  >({});
  const [sources, setSources] = useState<DataSourceFile[]>([]);
  const [isBusy, setIsBusy] = useState(false);
  const [message, setMessage] = useState(
    "Завантажте основний Excel у блоці “Джерела даних”.",
  );
  const [mergeSummary, setMergeSummary] = useState("");
  const activeSourceSheet =
    snapshot?.sheets.find((sheet) => sheet.sheetIndex === activeSheetIndex) ??
    snapshot?.sheets[0];
  const activeSnapshot =
    snapshot && activeSourceSheet
      ? {
          ...snapshot,
          sheetName: activeSourceSheet.sheetName,
          headerRows: activeSourceSheet.headerRows,
          rows:
            sheetRowsByIndex[activeSourceSheet.sheetIndex] ??
            activeSourceSheet.rows,
          columnCount: activeSourceSheet.columnCount,
          columnIndexes: activeSourceSheet.columnIndexes,
          dataStartRow: activeSourceSheet.dataStartRow,
        }
      : snapshot;
  const rows = activeSourceSheet
    ? (sheetRowsByIndex[activeSourceSheet.sheetIndex] ?? activeSourceSheet.rows)
    : [];
  const setRows: Dispatch<SetStateAction<ExcelRow[]>> = (nextRows) => {
    const sheetIndex = activeSourceSheet?.sheetIndex ?? activeSheetIndex;

    setSheetRowsByIndex((currentRowsByIndex) => {
      const currentRows =
        currentRowsByIndex[sheetIndex] ?? activeSourceSheet?.rows ?? [];
      const resolvedRows =
        typeof nextRows === "function" ? nextRows(currentRows) : nextRows;

      return {
        ...currentRowsByIndex,
        [sheetIndex]: resolvedRows,
      };
    });
  };

  const changeActiveSheet = (sheetIndex: number) => {
    setActiveSheetIndex(sheetIndex);
    const sheet = snapshot?.sheets.find(
      (item) => item.sheetIndex === sheetIndex,
    );
    if (sheet) {
      setMessage(
        `Відкрито sheet: ${sheet.sheetName}. Рядків із даними: ${sheet.rows.filter((row) => hasRowData(row.values)).length}.`,
      );
    }
  };

  const loadPrimaryFile = async (file: File | undefined) => {
    if (!file) return;

    setIsBusy(true);
    setMergeSummary("");
    try {
      const nextSnapshot = await readWorkbookSnapshot(file);
      console.groupCollapsed(
        `[army-grid] Parsed Excel: ${nextSnapshot.fileName}`,
      );
      console.log(createWorkbookDebugPayload(nextSnapshot));
      console.groupEnd();
      setSnapshot(nextSnapshot);
      setActiveSheetIndex(nextSnapshot.sheets[0]?.sheetIndex ?? 0);
      setSheetRowsByIndex(
        Object.fromEntries(
          nextSnapshot.sheets.map((sheet) => [sheet.sheetIndex, sheet.rows]),
        ),
      );
      setSources([
        {
          id: `primary-${file.name}-${file.lastModified}`,
          name: file.name,
          size: formatFileSize(file.size),
          rows: nextSnapshot.rows.filter((row) => hasRowData(row.values))
            .length,
          columns: nextSnapshot.columnCount,
          uploadedAt: formatDateTime(),
          role: "Основний файл",
        },
      ]);
      setMessage(
        `Завантажено ${nextSnapshot.fileName}: ${nextSnapshot.sheetName}. Дані відкрито в робочій таблиці нижче.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати Excel-файл.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const loadMergeFile = async (file: File | undefined) => {
    if (!file || !snapshot) return;

    setIsBusy(true);
    try {
      const mergeSnapshot = await readWorkbookSnapshot(file);
      console.groupCollapsed(
        `[army-grid] Parsed merge Excel: ${mergeSnapshot.fileName}`,
      );
      console.log(createWorkbookDebugPayload(mergeSnapshot));
      console.groupEnd();
      const mergeSheet =
        mergeSnapshot.sheets.find(
          (sheet) => sheet.sheetIndex === activeSheetIndex,
        ) ?? mergeSnapshot.sheets[0];
      const result = mergeRows(rows, mergeSheet?.rows ?? mergeSnapshot.rows);
      setRows(result.rows);
      setSources(
        (currentSources) =>
          [
            currentSources[0],
            {
              id: `merge-${file.name}-${file.lastModified}`,
              name: file.name,
              size: formatFileSize(file.size),
              rows: (mergeSheet?.rows ?? mergeSnapshot.rows).filter((row) =>
                hasRowData(row.values),
              ).length,
              columns: mergeSheet?.columnCount ?? mergeSnapshot.columnCount,
              uploadedAt: formatDateTime(),
              role: "Файл для merge",
            },
          ].filter(Boolean) as DataSourceFile[],
      );
      setMergeSummary(
        `Merge: оновлено ${result.stats.updated}, створено ${result.stats.created}, пропущено ${result.stats.skipped}.`,
      );
    } catch (error) {
      setMergeSummary(
        error instanceof Error ? error.message : "Не вдалося виконати merge.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  const exportWorkbook = async () => {
    if (!snapshot) return;

    setIsBusy(true);
    try {
      await exportStyledWorkbook(snapshot, rows, activeSheetIndex);
      setMessage(
        `Експортовано XLSX зі змінами на sheet: ${activeSourceSheet?.sheetName ?? snapshot.sheetName}.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати Excel-файл.",
      );
    } finally {
      setIsBusy(false);
    }
  };

  useEffect(() => {
    const handlePopState = () => {
      setActivePage(getPageFromPath(window.location.pathname));
      setRouteKey(getCurrentRouteKey());
    };

    window.addEventListener("popstate", handlePopState);
    return () => window.removeEventListener("popstate", handlePopState);
  }, []);

  useEffect(() => {
    if (window.location.pathname === "/") {
      window.history.replaceState(
        { page: activePage },
        "",
        pagePaths[activePage],
      );
    }
  }, [activePage]);

  const applyRoute = (route: { page: AppPage; routeKey: string }) => {
    setActivePage(route.page);
    setRouteKey(route.routeKey);
  };

  const changePage = (page: AppPage) => {
    applyRoute(navigateToPage(page));
  };
  const openPersonnelForPerson = (target: {
    rowId: string;
    externalId: string;
  }) => {
    window.localStorage.setItem(
      "army-grid:focus-personnel",
      JSON.stringify(target),
    );
    applyRoute(
      pushAppRoute(
        buildPersonnelRoute(target),
        "personnel",
      ),
    );
  };
  const openDocumentsForPerson = (
    row: EjournalPreviewRow,
    mode:
      | "default"
      | "salaryPowerAttorney"
      | "ubdReport"
      | "ubdRestoreReport"
      | "form6Report"
      | "form12Report"
      | "temporaryMilitaryId" = "default",
  ) => {
    const externalId = buildPersonSummary(row).externalId;
    window.localStorage.setItem(
      "army-grid:selected-person",
      JSON.stringify(row),
    );
    window.localStorage.setItem("army-grid:selected-document-mode", mode);
    applyRoute(
      pushAppRoute(
        buildDocumentRoute({
          personExternalId: externalId,
          rowId: row.__dbRowId,
          type: mode,
        }),
        "documents",
      ),
    );
  };

  return (
    <div className="app-shell">
      <SciLiveFeedback />
      <SciScrollbars />
      <Sidebar activePage={activePage} onPageChange={changePage} />

      {activePage === "overview" ? (
        <OverviewPage
          onOpenImport={() => changePage("import")}
          onOpenPersonnel={openPersonnelForPerson}
        />
      ) : activePage === "analytics" ? (
        <AnalyticsPage />
      ) : activePage === "bchs" ? (
        <BchsPage />
      ) : activePage === "bchsLab" ? (
        <BchsLabPage />
      ) : activePage === "excelFill" ? (
        <ExcelFillPage />
      ) : activePage === "questionnaireParser" ? (
        <QuestionnaireParserPage />
      ) : activePage === "socPassport" ? (
        <SocPassportPage />
      ) : activePage === "ejournal" ? (
        <EjournalPage />
      ) : activePage === "personnel" ? (
        <PersonnelPage onOpenDocuments={openDocumentsForPerson} />
      ) : activePage === "documents" ? (
        <DocumentsPage
          key={routeKey}
          onNavigate={(path) => {
            applyRoute(pushAppRoute(path, getPageFromPath(path)));
          }}
        />
      ) : activePage === "documentSettings" ? (
        <DocumentSignatoriesSettingsPage />
      ) : (
        <main className="main-panel">
          <header className="topbar">
            <Box>
              <Typography component="h1" variant="h5">
                Імпорт та об’єднання даних
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Excel/CSV → перевірка → єдина база → журнали статусів і
                документи
              </Typography>
            </Box>
            <Stack direction="row" spacing={1}>
              <Button
                variant="outlined"
                startIcon={<DescriptionOutlinedIcon />}
              >
                Журнал імпортів
              </Button>
              <Button variant="outlined" aria-label="Довідка">
                <HelpOutlineOutlinedIcon />
              </Button>
            </Stack>
          </header>

          <section className="stepper" aria-label="Етапи імпорту">
            {["Файл", "Поля", "Перевірка", "Імпорт"].map((step, index) => (
              <div className={`step${index === 0 ? " active" : ""}`} key={step}>
                <span className="step-number">
                  {String(index + 1).padStart(2, "0")}
                </span>
                <Typography variant="body2" sx={{ textTransform: "uppercase" }}>
                  {step}
                </Typography>
              </div>
            ))}
          </section>

          <div className="dashboard-grid">
            <DataSources
              sources={sources}
              canMerge={Boolean(snapshot)}
              onPrimaryFile={(file) => void loadPrimaryFile(file)}
              onMergeFile={(file) => void loadMergeFile(file)}
            />
            <MappingPanel />
            <ValidationPanel />
          </div>

          <ExcelRoundTripPanel
            snapshot={activeSnapshot}
            rows={rows}
            setRows={setRows}
            activeSheetIndex={activeSheetIndex}
            onActiveSheetChange={changeActiveSheet}
            isBusy={isBusy}
            message={message}
            mergeSummary={mergeSummary}
            onExport={() => void exportWorkbook()}
          />
        </main>
      )}
    </div>
  );
}

export default App;
