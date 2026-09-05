import { useCallback, useEffect, useMemo, useState, Suspense } from "react";
import { Canvas } from "@react-three/fiber";
import {
  Alert,
  Box,
  IconButton,
  LinearProgress,
  MenuItem,
  Select,
  Stack,
  Typography,
} from "@/components/sci/SciPrimitives";
import { BuildOutlinedIcon } from "@/components/sci/icons";
import { CheckCircleOutlineIcon } from "@/components/sci/icons";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { DeleteOutlineOutlinedIcon } from "@/components/sci/icons";
import { ErrorOutlineOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { FormatListBulletedOutlinedIcon } from "@/components/sci/icons";
import { FullscreenExitOutlinedIcon } from "@/components/sci/icons";
import { FullscreenOutlinedIcon } from "@/components/sci/icons";
import { GroupsOutlinedIcon } from "@/components/sci/icons";
import { HelpOutlineOutlinedIcon } from "@/components/sci/icons";
import { LogoutOutlinedIcon } from "@/components/sci/icons";
import { MilitaryTechOutlinedIcon } from "@/components/sci/icons";
import { PushPinOutlinedIcon } from "@/components/sci/icons";
import { SearchOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import { SyncAltOutlinedIcon } from "@/components/sci/icons";
import { WarningAmberOutlinedIcon } from "@/components/sci/icons";
import type { MRT_ColumnDef } from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import {
  api,
  type BackendEjournalImport,
  type BackendPersonnelRosterLatest,
} from "../../api";
import { useAuth } from "../../auth/AuthProvider";
import { CacheKeys, peekDataCache, subscribeDataCache } from "../../data/idbDataCache";
import { loadSharedRosterLatest } from "../../data/sharedAppData";
import { runHeavyJob } from "../../workers/runHeavyJob";
import { Button as SciButton } from "../../components/ui/button/button";
import {
  type ExcelWorkbookSnapshot,
  exportBlankWorkbookSheetWithMutations,
  exportTemplateWorkbookWithMutations,
  exportWorkbookWithMutations,
  hasRowData,
} from "../../excelRoundTrip";
import {
  getInitialBchsAnalyticsTab,
  type BchsAnalyticsTab,
} from "../../app/navigation";
import { BarList, BchsAbsenceDonut } from "../analytics/charts";
import { CombatStructureScene, type CombatBlock } from "../analytics/combatScene";
import { CombatModelsPreloader } from "../analytics/combatModelsPreloader";
import { MAX_VISUAL_SOLDIERS } from "../analytics/combatSoldiers";
import type {
  DbPreviewState,
  EjournalPreviewRow,
} from "../ejournal/ejournalTypes";
import {
  buildImportColumns,
  localRowsToPreviewRows,
  parseDbColumns,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";
import {
  applyBchsPersonnelDerivedColumns,
  buildBchsAnalytics,
  buildBchsAnalyticsFromWorkbook,
  buildBchsAppendixSupplement,
  buildBchsPersonnelBzvpSupplement,
  buildBchsSupplementComparisonRow,
  computeBchsUnitAttachedStats,
  computeBchsUnitAbsenceCategoryStats,
  computeBchsUnitAwayStats,
  computeBchsUnitBusinessTripCount,
  computeBchsUnitRankListedAvailableStats,
  computeBchsUnitUnassignedNewcomers,
  computeBchsUnitSearchInProgressCount,
  computeBchsUnitCombatComponentStats,
  sumBchsComparisonCombatComponent,
  applyBchsRankListedAvailableToComparisonRow,
  sumBchsComparisonRankListedAvailable,
  createBchsComparisonRow,
  enrichBchsAnalyticsForExport,
  extractBchsAwayPeopleFromSheet,
  extractBchsNovaPeopleFromSheet,
  filterBchsNovaPeople,
  formatRatioPercent,
  summarizeBchsPersonnelStay,
  isBchsAppendixSheet,
  isBchsDetachedStatus,
  isBchsPersonnelGeneralListSheet,
  isBchsPersonnelBzvpSheet,
  isBchsTotalUnit,
  isLegacyBchsAnalyticsSheet,
  parseBchsDestinationText,
  parseBchsImportAnalytics,
  resolveBchsFullRosterPeople,
  resolveBchsRosterPeople,
  summarizeBchsCommanderReserve,
  toPercent,
} from "./bchsCalc";
import {
  BCHS_EXPORT_DATA_END_ROW,
  BCHS_EXPORT_DATA_START_ROW,
  buildBchsMainExportRows,
  clearBchsSheetPaddingRows,
  fitBchsTextExportColumns,
  fitSheetColumnToText,
  materializeBchsSheetFormulas,
  styleBchsMainAttachedHeaderColumns,
  styleBchsMainAbsenceHeaderColumns,
  styleBchsMainDataRows,
  updateBchsSheetHeaderDate,
  writeBchsAppendixRow,
  writeBchsLegacyRow,
  writeBchsPersonnelBzvp1PbTemplate,
  writeBchsPersonnelBzvpRow,
  writeGeneratedBchsWorkbook,
} from "./bchsExport";
import type {
  BchsAnalyticsSnapshot,
  BchsDataIssue,
  BchsPersonnelAwayPerson,
  BchsSupplementSnapshot,
} from "./bchsTypes";

const getTodayBchsExportDate = () =>
  new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(new Date());

const buildMainBchsExportFileName = (reportDate: string) =>
  `1ПБ БЧС ${reportDate}.xlsx`;

const keepOnlyWorkbookSheet = (workbook: any, sheetToKeep: any) => {
  [...workbook.sheets()].reverse().forEach((sheet: any) => {
    if (sheet !== sheetToKeep) workbook.deleteSheet(sheet);
  });
};

const waitForNextBchsDownload = () =>
  new Promise((resolve) => window.setTimeout(resolve, 450));

const forceWorkbookFullCalculation = (workbook: any) => {
  const workbookNode = workbook?._node;
  const children = Array.isArray(workbookNode?.children)
    ? workbookNode.children
    : null;
  if (!children) return;

  let calcPr = children.find((child: any) => child?.name === "calcPr");
  if (!calcPr) {
    calcPr = { name: "calcPr", attributes: {}, children: [] };
    children.push(calcPr);
  }

  calcPr.attributes = {
    ...(calcPr.attributes ?? {}),
    calcMode: "auto",
    fullCalcOnLoad: 1,
    forceFullCalc: 1,
  };
};

const writeFormulaDrivenBchsCalculationWorkbook = (
  workbook: any,
  exportDate: string,
  analytics: BchsAnalyticsSnapshot,
  people: BchsPersonnelAwayPerson[],
) => {
  const calculationSheet = workbook.sheet("Аркуш1") ?? workbook.sheet(0);
  writeMainBchsCalculationSheet(
    calculationSheet,
    exportDate,
    analytics,
    people,
  );
  keepOnlyWorkbookSheet(workbook, calculationSheet);
  forceWorkbookFullCalculation(workbook);
};

const writeMainBchsCalculationSheet = (
  sheet: any,
  exportDate: string,
  analytics: ReturnType<typeof buildBchsAnalyticsFromWorkbook>,
  people: BchsPersonnelAwayPerson[] = [],
) => {
  updateBchsSheetHeaderDate(sheet, exportDate);
  materializeBchsSheetFormulas(sheet);
  const rosterPeople = filterBchsNovaPeople(people);
  const enriched = enrichBchsAnalyticsForExport(analytics, people);
  const exportRows = buildBchsMainExportRows(enriched).map((row) => {
    if (rosterPeople.length === 0 || row.rowNumber === 11 || isBchsTotalUnit(row.unit)) {
      return row;
    }
    // Always recompute away from roster at write time (нова + Відком. за межі).
    const away = computeBchsUnitAwayStats(rosterPeople, row.unit);
    const attached = computeBchsUnitAttachedStats(people, row.unit);
    const rankStats = computeBchsUnitRankListedAvailableStats(
      rosterPeople,
      row.unit,
    );
    const absenceStats = computeBchsUnitAbsenceCategoryStats(
      rosterPeople,
      row.unit,
    );
    const businessTrip = computeBchsUnitBusinessTripCount(
      rosterPeople,
      row.unit,
    );
    const unassignedNewcomers = computeBchsUnitUnassignedNewcomers(
      people,
      row.unit,
    );
    const searchInProgress = computeBchsUnitSearchInProgressCount(
      rosterPeople,
      row.unit,
    );
    const combat = computeBchsUnitCombatComponentStats(people, row.unit);
    return applyBchsRankListedAvailableToComparisonRow(
      createBchsComparisonRow({
        ...row,
        awayOfficers: away.officers,
        awaySergeants: away.sergeants,
        awaySoldiers: away.soldiers,
        awayInOtherUnits: away.total,
        awayDestinationsText: away.destinationText,
        attachedOfficers: attached.officers,
        attachedSergeants: attached.sergeants,
        attachedSoldiers: attached.soldiers,
        attachedFromOtherUnits: attached.total,
        attachedSourcesText: attached.sourcesText,
        businessTrip,
        unassignedNewcomers,
        searchInProgress,
        ...combat,
        training: absenceStats.training,
        hospitalWounded: absenceStats.hospitalWounded,
        hospitalIllness: absenceStats.hospitalIllness,
        vacation: absenceStats.vacation,
        awol: absenceStats.awol,
        missing: absenceStats.missing,
        killed: absenceStats.killed,
        medWounded: absenceStats.medWounded,
        medIllness: absenceStats.medIllness,
      }),
      rankStats,
    );
  });

  // Recompute total away/attached from detail rows (Excel: no AO text on Усього).
  const detailRows = exportRows.filter(
    (row) => row.rowNumber !== 11 && !isBchsTotalUnit(row.unit),
  );
  const totalAwayOfficers = detailRows.reduce((sum, row) => sum + row.awayOfficers, 0);
  const totalAwaySergeants = detailRows.reduce((sum, row) => sum + row.awaySergeants, 0);
  const totalAwaySoldiers = detailRows.reduce((sum, row) => sum + row.awaySoldiers, 0);
  const totalAttachedOfficers = detailRows.reduce((sum, row) => sum + row.attachedOfficers, 0);
  const totalAttachedSergeants = detailRows.reduce((sum, row) => sum + row.attachedSergeants, 0);
  const totalAttachedSoldiers = detailRows.reduce((sum, row) => sum + row.attachedSoldiers, 0);
  const totalRankStats = sumBchsComparisonRankListedAvailable(detailRows);
  const totalBusinessTrip = detailRows.reduce(
    (sum, row) => sum + row.businessTrip,
    0,
  );
  const totalUnassignedNewcomers = detailRows.reduce(
    (sum, row) => sum + row.unassignedNewcomers,
    0,
  );
  const totalSearchInProgress = detailRows.reduce(
    (sum, row) => sum + row.searchInProgress,
    0,
  );
  const totalCombat = sumBchsComparisonCombatComponent(detailRows);
  const exportRowsWithTotal = exportRows.map((row) => {
    if (row.rowNumber !== 11 && !isBchsTotalUnit(row.unit)) return row;
    return applyBchsRankListedAvailableToComparisonRow(
      createBchsComparisonRow({
        ...row,
        awayOfficers: totalAwayOfficers,
        awaySergeants: totalAwaySergeants,
        awaySoldiers: totalAwaySoldiers,
        awayInOtherUnits: totalAwayOfficers + totalAwaySergeants + totalAwaySoldiers,
        awayDestinationsText: "",
        attachedOfficers: totalAttachedOfficers,
        attachedSergeants: totalAttachedSergeants,
        attachedSoldiers: totalAttachedSoldiers,
        attachedFromOtherUnits:
          totalAttachedOfficers + totalAttachedSergeants + totalAttachedSoldiers,
        attachedSourcesText: "",
        businessTrip: totalBusinessTrip,
        unassignedNewcomers: totalUnassignedNewcomers,
        searchInProgress: totalSearchInProgress,
        ...totalCombat,
        training: detailRows.reduce((sum, item) => sum + item.training, 0),
        hospitalWounded: detailRows.reduce(
          (sum, item) => sum + item.hospitalWounded,
          0,
        ),
        hospitalIllness: detailRows.reduce(
          (sum, item) => sum + item.hospitalIllness,
          0,
        ),
        vacation: detailRows.reduce((sum, item) => sum + item.vacation, 0),
        awol: detailRows.reduce((sum, item) => sum + item.awol, 0),
        missing: detailRows.reduce((sum, item) => sum + item.missing, 0),
        killed: detailRows.reduce((sum, item) => sum + item.killed, 0),
        medWounded: detailRows.reduce((sum, item) => sum + item.medWounded, 0),
        medIllness: detailRows.reduce((sum, item) => sum + item.medIllness, 0),
      }),
      totalRankStats,
    );
  });

  for (
    let rowNumber = BCHS_EXPORT_DATA_START_ROW;
    rowNumber <= BCHS_EXPORT_DATA_END_ROW;
    rowNumber += 1
  ) {
    writeBchsLegacyRow(sheet, createBchsComparisonRow({ rowNumber, unit: "" }));
  }
  exportRowsWithTotal.forEach((row) => {
    writeBchsLegacyRow(sheet, row);
  });
  clearBchsSheetPaddingRows(sheet);
  styleBchsMainAbsenceHeaderColumns(sheet);
  styleBchsMainAttachedHeaderColumns(sheet);
  styleBchsMainDataRows(
    sheet,
    exportRowsWithTotal
      .map((row) => row.rowNumber)
      .filter((rowNumber): rowNumber is number => Boolean(rowNumber)),
  );
  fitBchsTextExportColumns(
    sheet,
    exportRowsWithTotal
      .map((row) => row.rowNumber)
      .filter((rowNumber): rowNumber is number => Boolean(rowNumber)),
  );
};

export function BchsPage({ active = true }: { active?: boolean }) {
  const { canEditArea } = useAuth();
  const canEdit = canEditArea("bchs");
  const [snapshot] = useState<ExcelWorkbookSnapshot | null>(null);
  const [imports, setImports] = useState<BackendEjournalImport[]>([]);
  const [dbPreview, setDbPreview] = useState<DbPreviewState | null>(null);
  const [selectedImportId, setSelectedImportId] = useState("");
  const [selectedComparisonRow, setSelectedComparisonRow] = useState<
    number | null
  >(null);
  const [comparisonQuery, setComparisonQuery] = useState("");
  const [comparisonSort, setComparisonSort] = useState("critical");
  const [activeBchsAnalyticsTab, setActiveBchsAnalyticsTab] =
    useState<BchsAnalyticsTab>(getInitialBchsAnalyticsTab);
  const [selectedCombatUnit, setSelectedCombatUnit] = useState("3 рота");
  const [combatFullscreen, setCombatFullscreen] = useState(false);
  const [combatCameraResetKey, setCombatCameraResetKey] = useState(0);
  const [combatSceneReady, setCombatSceneReady] = useState(false);
  const [message, setMessage] = useState(`API: ${api.baseUrl}`);
  const [isLoading, setIsLoading] = useState(false);
  const [personnelAwayPeople, setPersonnelAwayPeople] = useState<
    BchsPersonnelAwayPerson[]
  >([]);

  const analyticsSheet =
    snapshot?.sheets.find(
      (sheet) => sheet.sheetName.trim().toLowerCase() === "аркуш1",
    ) ?? snapshot?.sheets[0];
  const dataSheet =
    snapshot?.sheets.find(isBchsPersonnelGeneralListSheet) ??
    snapshot?.sheets.find(
      (sheet) => sheet.sheetName.trim().toLowerCase() === "аркуш2",
    );
  const selectedImport = useMemo(
    () => imports.find((item) => item.id === selectedImportId) ?? imports[0],
    [imports, selectedImportId],
  );
  const analyticsFromDb = useMemo(
    () => parseBchsImportAnalytics(selectedImport),
    [selectedImport],
  );
  const analytics = useMemo(() => {
    const base = snapshot
      ? buildBchsAnalyticsFromWorkbook(snapshot, analyticsSheet)
      : (analyticsFromDb ?? buildBchsAnalytics(undefined));

    const people = personnelAwayPeople.length
      ? personnelAwayPeople
      : resolveBchsFullRosterPeople(
          dataSheet ? extractBchsAwayPeopleFromSheet(dataSheet) : undefined,
        );

    return applyBchsPersonnelDerivedColumns(base, people);
  }, [analyticsFromDb, analyticsSheet, dataSheet, personnelAwayPeople, snapshot]);
  const personnelStayStats = useMemo(() => {
    const people = personnelAwayPeople.length
      ? personnelAwayPeople
      : resolveBchsFullRosterPeople(
          dataSheet ? extractBchsAwayPeopleFromSheet(dataSheet) : undefined,
        );
    return summarizeBchsPersonnelStay(people);
  }, [dataSheet, personnelAwayPeople]);
  const dataIssues = (analytics.dataIssues ?? []) as BchsDataIssue[];
  const detachedDestinationIssues = dataIssues.filter(
    (issue) =>
      issue.kind === "destination" ||
      issue.reason === "невідоме місце відкомандирування",
  );
  const rankDataIssues = dataIssues.filter(
    (issue) =>
      issue.kind === "rank" ||
      /званн|капітан|капитан|латиниц/i.test(issue.reason),
  );
  const anomalyDataIssues = dataIssues.filter(
    (issue) => issue.kind === "anomaly" || issue.kind === "brez",
  );
  const statusDataIssues = dataIssues.filter(
    (issue) =>
      !detachedDestinationIssues.includes(issue) &&
      !rankDataIssues.includes(issue) &&
      !anomalyDataIssues.includes(issue),
  );
  const personnelSupplementForExport = useMemo(() => {
    if (analytics.supplement?.kind === "personnel-bzvp")
      return analytics.supplement;

    return imports
      .map(parseBchsImportAnalytics)
      .map((item) => item?.supplement)
      .find(
        (supplement): supplement is BchsSupplementSnapshot =>
          supplement?.kind === "personnel-bzvp" && supplement.rows.length > 0,
      );
  }, [analytics.supplement, imports]);
  const comparisonRows = useMemo(() => {
    const query = comparisonQuery.trim().toLowerCase();
    const sourceRows = analytics.comparisonRows.filter(
      (row) => row.rowNumber > 11 && row.unit.trim(),
    );
    const filteredRows = query
      ? sourceRows.filter((row) => row.unit.toLowerCase().includes(query))
      : sourceRows;
    const sortedRows = [...filteredRows];

    if (comparisonSort === "critical") {
      sortedRows.sort(
        (left, right) =>
          left.levelPercent - right.levelPercent || right.absent - left.absent,
      );
    } else if (comparisonSort === "absent") {
      sortedRows.sort(
        (left, right) =>
          right.absent - left.absent || left.levelPercent - right.levelPercent,
      );
    } else if (comparisonSort === "staff") {
      sortedRows.sort(
        (left, right) =>
          right.staff - left.staff || left.unit.localeCompare(right.unit, "uk"),
      );
    } else {
      sortedRows.sort((left, right) => left.rowNumber - right.rowNumber);
    }

    return sortedRows;
  }, [analytics.comparisonRows, comparisonQuery, comparisonSort]);
  const selectedComparison = useMemo(
    () =>
      comparisonRows.find((row) => row.rowNumber === selectedComparisonRow) ??
      comparisonRows[0] ??
      null,
    [comparisonRows, selectedComparisonRow],
  );
  const comparisonMax = useMemo(
    () =>
      Math.max(
        1,
        ...comparisonRows.flatMap((row) => [
          row.staff,
          row.listed,
          row.inRanksActually,
        ]),
      ),
    [comparisonRows],
  );
  const totalComparison = useMemo(
    () =>
      analytics.comparisonRows.find((row) => row.rowNumber === 11) ??
      createBchsComparisonRow(analytics.total),
    [analytics.comparisonRows, analytics.total],
  );
  const overviewUnits = useMemo(() => {
    const namedRows = analytics.comparisonRows.filter(
      (row) =>
        row.rowNumber > 11 &&
        (/піхотна рота/i.test(row.unit) || /рбпак/i.test(row.unit)),
    );

    return namedRows.length > 0
      ? namedRows.slice(0, 4)
      : analytics.comparisonRows
          .filter((row) => row.rowNumber > 11)
          .slice(0, 4);
  }, [analytics.comparisonRows]);
  const keySignals = useMemo(
    () =>
      analytics.comparisonRows
        .filter((row) => row.rowNumber > 11 && row.staff > 0)
        .sort(
          (left, right) =>
            left.levelPercent - right.levelPercent ||
            right.absent - left.absent,
        )
        .slice(0, 3),
    [analytics.comparisonRows],
  );
  const combatComposition = useMemo(
    () => [
      {
        label: "Піхота",
        value: analytics.total.assaultTotal,
        tone: "infantry",
      },
      {
        label: "Екіпажі БПЛА",
        value: analytics.total.droneCrew,
        tone: "drone",
      },
      {
        label: "Екіпажі техніки",
        value: analytics.total.vehicleCrew,
        tone: "vehicle",
      },
      {
        label: "Колективне озброєння",
        value: analytics.total.crewServedWeapons,
        tone: "weapon",
      },
      {
        label: "Управління",
        value: analytics.total.commandCombat,
        tone: "command",
      },
      {
        label: "Забезпечення",
        value: analytics.total.supportCombat,
        tone: "support",
      },
    ],
    [analytics.total],
  );
  const combatReadiness = useMemo(
    () => [
      { label: "Готові", value: analytics.total.assaultReady, tone: "ready" },
      {
        label: "На відновленні",
        value: analytics.total.assaultRecovery,
        tone: "recovery",
      },
      {
        label: "На виконанні",
        value: analytics.total.assaultExecution,
        tone: "execution",
      },
      { label: "Без БЗВП", value: analytics.total.noBzvp, tone: "no-bzvp" },
    ],
    [analytics.total],
  );
  const combatBlocks = useMemo<CombatBlock[]>(() => {
    const findUnit = (pattern: RegExp, label: string, x: number, z: number) => {
      const row = analytics.comparisonRows.find((item) =>
        pattern.test(item.unit),
      );

      return {
        label,
        staff: row?.staff ?? 1,
        actual: row?.inRanksActually ?? 0,
        x,
        z,
      };
    };

    return [
      findUnit(/1\s*піхотна рота/i, "1 рота", -1.9, -0.95),
      findUnit(/2\s*піхотна рота/i, "2 рота", 0, -0.95),
      findUnit(/3\s*піхотна рота/i, "3 рота", 1.9, -0.95),
      findUnit(/рота безпілот|рбпак/i, "РБпАК", -1.8, 1.05),
      {
        label: "Підтримка",
        staff: Math.max(1, analytics.total.supportCombat),
        actual: analytics.total.supportCombat,
        x: 0,
        z: 1.05,
      },
      {
        label: "Управління",
        staff: Math.max(1, analytics.total.commandCombat),
        actual: analytics.total.commandCombat,
        x: 1.8,
        z: 1.05,
      },
    ];
  }, [analytics.comparisonRows, analytics.total]);
  const selectedCombatBlock = useMemo(
    () =>
      combatBlocks.find((block) => block.label === selectedCombatUnit) ??
      combatBlocks[0],
    [combatBlocks, selectedCombatUnit],
  );
  const resetCombatCamera = () => {
    setCombatCameraResetKey((key) => key + 1);
  };
  const selectCombatUnit = (label: string) => {
    setSelectedCombatUnit(label);
    resetCombatCamera();
  };
  const setCombatFullscreenMode = (next: boolean) => {
    setCombatFullscreen(next);
    resetCombatCamera();
  };
  const markCombatSceneReady = useCallback(() => {
    setCombatSceneReady(true);
  }, []);
  const analyticsTable = useMemo(
    () => ({
      columns: analytics.table?.columns ?? [],
      rows: analytics.table?.rows ?? [],
    }),
    [analytics.table],
  );
  const hasBchsAnalyticsTable =
    analyticsTable.columns.length > 0 &&
    analyticsTable.rows.some((row) =>
      analyticsTable.columns.some(
        (column) => String(row.values[column.key] ?? "").trim() !== "",
      ),
    );
  const isMissingBchsAnalyticsSnapshot = Boolean(
    selectedImport && !snapshot && !analyticsFromDb,
  );
  const selectedDbSheet = selectedImport?.sheets[0];
  const localColumns = useMemo(
    () => (dataSheet ? buildImportColumns(dataSheet) : []),
    [dataSheet],
  );
  const previewColumns = useMemo(
    () => dbPreview?.columns ?? localColumns,
    [dbPreview, localColumns],
  );
  const previewRows = useMemo(
    () =>
      dbPreview?.rows ??
      (dataSheet
        ? localRowsToPreviewRows(
            dataSheet.rows.filter((row) => hasRowData(row.values)),
            localColumns,
          )
        : []),
    [dataSheet, dbPreview, localColumns],
  );
  const tableColumns = useMemo<MRT_ColumnDef<EjournalPreviewRow>[]>(
    () => [
      {
        id: "excel-row-number",
        accessorKey: "__rowNumber",
        header: "#",
        Cell: ({ cell }) => previewValueToDisplay(cell.getValue()),
        size: 72,
      },
      ...previewColumns.map((column) => ({
        id: column.key,
        accessorKey: column.key,
        header: column.label || column.letter || "",
        Cell: ({ cell }: { cell: { getValue: () => unknown } }) =>
          previewValueToDisplay(cell.getValue()),
        size: (column.label || column.letter || "").length > 24 ? 240 : 160,
      })),
    ],
    [previewColumns],
  );
  const table = useMaterialReactTable({
    columns: tableColumns,
    data: previewRows,
    enableColumnResizing: true,
    enableColumnVirtualization: true,
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
      sx: { maxHeight: 520, backgroundColor: "transparent" },
    },
    muiTableHeadCellProps: {
      sx: {
        backgroundColor: "#131311",
        color: "#d9d49d",
        borderColor: "rgba(230,224,190,0.12)",
      },
    },
    muiTableBodyCellProps: {
      sx: {
        backgroundColor: "#11110f",
        color: "#f2eee1",
        borderColor: "rgba(230,224,190,0.1)",
      },
    },
    muiTopToolbarProps: { sx: { backgroundColor: "rgba(17,17,15,0.92)" } },
    muiBottomToolbarProps: { sx: { backgroundColor: "rgba(17,17,15,0.92)" } },
  });
  const loadImports = async () => {
    setIsLoading(true);
    try {
      const nextImports = await api.listBchsImports();
      setImports(nextImports);
      setSelectedImportId((currentId) => currentId || nextImports[0]?.id || "");
      setMessage(`У БД знайдено імпортів БЧС: ${nextImports.length}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося прочитати імпорти БЧС.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    void loadImports();
  }, []);

  useEffect(() => {
    if (active) return;
    setCombatFullscreen(false);
  }, [active]);

  useEffect(() => {
    if (!combatFullscreen) return;

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "Escape") return;
      setCombatFullscreen(false);
      setCombatCameraResetKey((key) => key + 1);
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [combatFullscreen]);

  useEffect(() => {
    document.body.classList.toggle("combat-fs-open", combatFullscreen);
    setCombatSceneReady(false);
    const frame = window.requestAnimationFrame(() => {
      window.dispatchEvent(new Event("resize"));
    });
    return () => {
      window.cancelAnimationFrame(frame);
      document.body.classList.remove("combat-fs-open");
    };
  }, [combatFullscreen]);

  useEffect(() => {
    let cancelled = false;

    const applyRoster = async (latest: BackendPersonnelRosterLatest | null) => {
      if (!latest?.sheet) {
        setPersonnelAwayPeople([]);
        return;
      }

      const rows = latest.rows.map((row) =>
        row.values &&
        typeof row.values === "object" &&
        !Array.isArray(row.values)
          ? row.values
          : {},
      );
      const { people, novaCount } = await runHeavyJob({
        type: "bchsExtractPeople",
        rows,
        columns: parseDbColumns(latest.sheet.columns),
      });
      if (cancelled) return;
      const label =
        latest.sourceFileName?.trim() || latest.importName?.trim() || "Штатка";
      setPersonnelAwayPeople(people);
      setMessage(
        `БЧС зі Штатки «${label}»: ${people.length} осіб (нова: ${novaCount}). Аркуш1 — шаблон з БД.`,
      );
      console.info(
        `[BCHS] Люди зі Штатки: ${people.length} (нова: ${novaCount}) · ${label}`,
      );
    };

    const loadStaffRosterPeople = async () => {
      try {
        const latest = await loadSharedRosterLatest();
        if (cancelled) return;
        await applyRoster(latest);
      } catch (error) {
        console.warn("[BCHS] Не вдалося підвантажити Штатку", error);
        if (!cancelled) setPersonnelAwayPeople([]);
      }
    };

    void loadStaffRosterPeople();
    const unsubscribe = subscribeDataCache(CacheKeys.rosterLatest, () => {
      if (cancelled) return;
      void applyRoster(peekDataCache<BackendPersonnelRosterLatest | null>(CacheKeys.rosterLatest));
    });

    return () => {
      cancelled = true;
      unsubscribe();
    };
  }, []);

  const deleteSelectedImport = async () => {
    if (!selectedImport) return;

    const importName = selectedImport.sourceFileName ?? selectedImport.name;
    const confirmed = window.confirm(
      `Видалити імпорт БЧС "${importName}" з бази даних?`,
    );
    if (!confirmed) return;

    setIsLoading(true);
    try {
      await api.deleteBchsImport(selectedImport.id);
      setDbPreview(null);
      setSelectedImportId("");
      await loadImports();
      setMessage(`Імпорт БЧС видалено з БД: ${importName}.`);
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося видалити імпорт БЧС.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const loadDbPreview = async () => {
    if (!selectedDbSheet) return;

    setIsLoading(true);
    try {
      const response = await api.listBchsSheetRows(selectedDbSheet.id, {
        limit: 500,
        offset: 0,
      });
      const columns = parseDbColumns(response.columns);
      setDbPreview({
        sheet: selectedDbSheet,
        columns,
        rows: response.items.map((row) => ({
          __dbRowId: row.id,
          __rowNumber: row.excelRowNumber,
          ...row.values,
        })),
        total: response.total,
        offset: response.offset,
        limit: response.limit,
      });
      setMessage(
        `Показано БЧС з БД: ${response.items.length} з ${response.total} рядків.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося показати БЧС з БД.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  const exportComparison = () => {
    const headers = [
      "Підрозділ",
      "Штат",
      "Список",
      "Фактично",
      "Рівень",
      "Відсутні",
      "В інших підрозділах",
      "Прикомандировані",
    ];
    const rows = comparisonRows.map((row) => [
      row.unit,
      row.staff,
      row.listed,
      row.inRanksActually,
      formatRatioPercent(row.inRanksActually, row.staff),
      row.absent,
      row.awayInOtherUnits,
      row.attachedFromOtherUnits,
    ]);
    const csv = [headers, ...rows]
      .map((row) =>
        row.map((cell) => `"${String(cell).replace(/"/g, '""')}"`).join(";"),
      )
      .join("\n");
    const blob = new Blob([`\uFEFF${csv}`], {
      type: "text/csv;charset=utf-8;",
    });
    const link = document.createElement("a");

    link.href = URL.createObjectURL(blob);
    link.download = "bchs-comparison.csv";
    link.click();
    URL.revokeObjectURL(link.href);
  };

  const exportBchsExcel = async () => {
    if (!snapshot && !analyticsFromDb) {
      setMessage("Завантажте Excel-шаблон БЧС або виберіть імпорт з БД.");
      return;
    }

    setIsLoading(true);
    try {
      const exportDate = getTodayBchsExportDate();
      const hasLegacyCalculationSheet = Boolean(
        snapshot?.sheets.some(isLegacyBchsAnalyticsSheet),
      );

      if (dataSheet) {
        const peopleForExport = personnelAwayPeople.length
          ? personnelAwayPeople
          : resolveBchsFullRosterPeople(
              dataSheet ? extractBchsAwayPeopleFromSheet(dataSheet) : undefined,
            );
        console.info(
          "[BCHS] Резерв командира полку",
          summarizeBchsCommanderReserve(peopleForExport),
        );
        const novaDetached = filterBchsNovaPeople(peopleForExport).filter((person) =>
          isBchsDetachedStatus(person.status),
        ).length;
        await exportTemplateWorkbookWithMutations(
          "/templates/bchs-calculation-template.xlsx",
          (workbook) =>
            writeFormulaDrivenBchsCalculationWorkbook(
              workbook,
              exportDate,
              analytics,
              peopleForExport,
            ),
          buildMainBchsExportFileName(exportDate),
        );
        if (novaDetached === 0) {
          console.warn(
            "[BCHS export] Немає відкомандированих з battalion=нова — AK–AO будуть порожні.",
          );
        } else {
          console.info(
            `[BCHS export] Відкомандировані (нова): ${novaDetached} осіб → пишемо AK–AO.`,
          );
        }
      } else if (snapshot && hasLegacyCalculationSheet) {
        const legacyPeople = personnelAwayPeople.length
          ? personnelAwayPeople
          : resolveBchsFullRosterPeople(
              dataSheet
                ? extractBchsAwayPeopleFromSheet(dataSheet)
                : snapshot?.sheets.find(isBchsPersonnelGeneralListSheet)
                  ? extractBchsAwayPeopleFromSheet(
                      snapshot.sheets.find(isBchsPersonnelGeneralListSheet),
                    )
                  : undefined,
            );
        await exportWorkbookWithMutations(
          snapshot,
          (workbook) => {
            let calculationSheet: any = null;

            snapshot.sheets.forEach((sheetSnapshot) => {
              const sheet = workbook.sheet(sheetSnapshot.sheetIndex);

              if (isLegacyBchsAnalyticsSheet(sheetSnapshot)) {
                calculationSheet ??= sheet;
                writeMainBchsCalculationSheet(
                  sheet,
                  exportDate,
                  analytics,
                  legacyPeople,
                );
                return;
              }

              updateBchsSheetHeaderDate(sheet, exportDate);

              if (isBchsAppendixSheet(sheetSnapshot)) {
                const supplement =
                  analytics.supplement?.kind === "appendix"
                    ? analytics.supplement
                    : buildBchsAppendixSupplement(sheetSnapshot);
                const sourceRows =
                  supplement?.rows.map(buildBchsSupplementComparisonRow) ?? [
                    totalComparison,
                  ];

                sourceRows.forEach((row) => {
                  if (row.rowNumber)
                    writeBchsAppendixRow(sheet, row.rowNumber, row);
                });
                fitSheetColumnToText(
                  sheet,
                  "B",
                  sourceRows
                    .map((row) => row.rowNumber)
                    .filter((rowNumber): rowNumber is number => Boolean(rowNumber)),
                  {
                    minWidth: 18,
                    maxWidth: 42,
                    headerTexts: ["Батальйон (підрозділ)"],
                  },
                );
                return;
              }

              if (isBchsPersonnelBzvpSheet(sheetSnapshot)) {
                const supplement =
                  analytics.supplement?.kind === "personnel-bzvp"
                    ? analytics.supplement
                    : buildBchsPersonnelBzvpSupplement(sheetSnapshot);

                supplement?.rows.forEach((row) =>
                  writeBchsPersonnelBzvpRow(sheet, row),
                );
                fitSheetColumnToText(
                  sheet,
                  "B",
                  (supplement?.rows ?? [])
                    .map((row) => row.rowNumber)
                    .filter((rowNumber): rowNumber is number => Boolean(rowNumber)),
                  {
                    minWidth: 18,
                    maxWidth: 48,
                    headerTexts: ["ПІДРОЗДІЛ"],
                  },
                );
              }
            });

            if (calculationSheet) keepOnlyWorkbookSheet(workbook, calculationSheet);
          },
          buildMainBchsExportFileName(exportDate),
        );
      } else {
        await exportTemplateWorkbookWithMutations(
          "/templates/bchs-calculation-template.xlsx",
          (workbook) => {
            const sheet = workbook.sheet("Аркуш1") ?? workbook.sheet(0);
            writeMainBchsCalculationSheet(
              sheet,
              exportDate,
              analytics,
              personnelAwayPeople,
            );
            keepOnlyWorkbookSheet(workbook, sheet);
          },
          buildMainBchsExportFileName(exportDate),
        );
      }
      await waitForNextBchsDownload();
      const peopleForPersonnelExport = resolveBchsRosterPeople(
        personnelAwayPeople,
        extractBchsNovaPeopleFromSheet(dataSheet),
      );
      await exportTemplateWorkbookWithMutations(
        "/templates/bchs-personnel-bzvp-template.xlsx",
        (workbook) => {
          const sheet = workbook.sheet(0);
          updateBchsSheetHeaderDate(sheet, exportDate);
          writeBchsPersonnelBzvp1PbTemplate(
            sheet,
            analytics,
            peopleForPersonnelExport,
          );
        },
        `1ПБ Особовий склад підрозділів + БЗВП_${exportDate}.xlsx`,
      );
      await waitForNextBchsDownload();
      await exportBlankWorkbookSheetWithMutations(
        (workbook) =>
          writeGeneratedBchsWorkbook(
            workbook,
            analytics,
            personnelSupplementForExport,
            exportDate,
            "Аркуш1",
            peopleForPersonnelExport,
          ),
        "БЧС додаток",
        `1ПБ БЧС додаток ${exportDate}.xlsx`,
      );
      setMessage(
        `Excel експортовано 3 файла: ${buildMainBchsExportFileName(exportDate)}, 1ПБ Особовий склад підрозділів + БЗВП_${exportDate}.xlsx, 1ПБ БЧС додаток ${exportDate}.xlsx.`,
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Не вдалося експортувати БЧС Excel.",
      );
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <main className="main-panel bchs-page">
      <header className="topbar analytics-topbar">
        <Box>
          <Typography component="h1" variant="h4">
            БЧС
          </Typography>
          <Typography variant="body2" color="text.secondary">
            Бойовий чисельний склад · люди зі Штатки (Особовий склад) · Аркуш1 як
            розрахунковий шаблон
          </Typography>
        </Box>
        <Stack direction="row" spacing={1}>
          <SciButton
            disabled={!snapshot && !analyticsFromDb}
            variant="OUTLINE"
            onClick={() => void exportBchsExcel()}
          >
            <FileDownloadOutlinedIcon fontSize="small" />
            Експорт Excel
          </SciButton>
        </Stack>
      </header>
      {isLoading && <LinearProgress color="primary" />}
      <Alert
        severity={snapshot ? "success" : "info"}
        variant="outlined"
        sx={{ mb: 2 }}
      >
        {isMissingBchsAnalyticsSnapshot
          ? `${message} Для вибраного імпорту немає аналітики Аркуш1, бо він був збережений старою версією.`
          : message}
      </Alert>

      <section className="bchs-kpi-grid bchs-shared-metrics">
        {[
          {
            label: "За штатом",
            value: analytics.total.staff,
            suffix: "посад",
            icon: <GroupsOutlinedIcon />,
          },
          {
            label: "За списком",
            value: analytics.total.listed,
            suffix: `осіб · ${toPercent(analytics.total.staffedPercent)}`,
            icon: <FormatListBulletedOutlinedIcon />,
          },
          {
            label: "В строю за штатом",
            value: analytics.total.available,
            icon: <MilitaryTechOutlinedIcon />,
          },
          {
            label: "Фактично в строю",
            value: analytics.total.inRanksActually,
            suffix: formatRatioPercent(
              analytics.total.inRanksActually,
              analytics.total.staff,
            ),
            icon: <ShieldOutlinedIcon />,
          },
          {
            label: "Некомплект",
            value: analytics.total.shortage,
            tone: "warn",
            icon: <WarningAmberOutlinedIcon />,
          },
          {
            label: "Відсутні",
            value: analytics.total.absent,
            tone: "danger",
            icon: <ErrorOutlineOutlinedIcon />,
          },
        ].map((item) => (
          <div
            className={`bchs-kpi-card ${
              item.tone === "warn" || item.tone === "danger" ? item.tone : ""
            }`}
            key={item.label}
          >
            <span className="bchs-kpi-icon">{item.icon}</span>
            <span>{item.label}</span>
            <div className="bchs-kpi-value">
              <strong>{item.value}</strong>
              {item.suffix && <em>· {item.suffix}</em>}
            </div>
          </div>
        ))}
      </section>

      <div className="bchs-primary-tabs" aria-label="Розділи аналітики БЧС">
        {[
          ["overview", "Аналітика БЧС"],
          ["comparison", "Порівняння підрозділів"],
          ["combat", "Бойова складова"],
          ["supplement", "Додаткова"],
        ].map(([tab, label]) => (
          <button
            className={activeBchsAnalyticsTab === tab ? "active" : ""}
            key={tab}
            type="button"
            onClick={() => setActiveBchsAnalyticsTab(tab as BchsAnalyticsTab)}
          >
            {label}
          </button>
        ))}
      </div>

      {activeBchsAnalyticsTab === "overview" && (
        <section className="bchs-overview">
          <section className="bchs-stay-stats-grid" aria-label="Перебування зі Штатки">
            <div className="bchs-stay-panel bchs-stay-panel--places">
              <div className="panel-heading">
                <span>
                  <PushPinOutlinedIcon fontSize="small" />
                  Місце перебування
                </span>
                <em>{personnelStayStats.total}</em>
              </div>
              {personnelStayStats.stayPlaces.length > 0 ? (
                <BarList items={personnelStayStats.stayPlaces} />
              ) : (
                <Typography variant="body2" color="text.secondary">
                  Немає людей зі Штатки, щоб порахувати місця.
                </Typography>
              )}
            </div>

            <div className="bchs-stay-stack">
              <div className="bchs-stay-panel bchs-stay-panel--exit">
                <div className="panel-heading">
                  <span>
                    <LogoutOutlinedIcon fontSize="small" />
                    На виході
                  </span>
                  <em>{personnelStayStats.onExit.length}</em>
                </div>
                {personnelStayStats.onExit.length > 0 ? (
                  <ul className="bchs-exit-list">
                    {personnelStayStats.onExit.map((person, index) => (
                      <li
                        key={`${person.fullName}-${person.rosterUnit}-${index}`}
                      >
                        <strong>{person.fullName || "Без ПІБ"}</strong>
                        <span>
                          {[
                            person.rosterUnit,
                            person.medicalPlace,
                            person.fighterExitDate
                              ? `вихід ${person.fighterExitDate}`
                              : "",
                          ]
                            .filter(Boolean)
                            .join(" · ")}
                        </span>
                      </li>
                    ))}
                  </ul>
                ) : (
                  <Typography variant="body2" color="text.secondary">
                    Нікого на виході за Штаткою.
                  </Typography>
                )}
              </div>

              <div className="bchs-stay-panel bchs-stay-panel--bg">
                <div className="panel-heading">
                  <span>
                    <ShieldOutlinedIcon fontSize="small" />
                    Статус БГ
                  </span>
                  <em>{personnelStayStats.total}</em>
                </div>
                <div className="bchs-bg-split">
                  <div className="bchs-bg-item ready">
                    <strong>{personnelStayStats.battleReady}</strong>
                    <span>БГ</span>
                    <em>
                      {formatRatioPercent(
                        personnelStayStats.battleReady,
                        personnelStayStats.total,
                      )}
                    </em>
                  </div>
                  <div className="bchs-bg-item">
                    <strong>{personnelStayStats.notBattleReady}</strong>
                    <span>не БГ</span>
                    <em>
                      {formatRatioPercent(
                        personnelStayStats.notBattleReady,
                        personnelStayStats.total,
                      )}
                    </em>
                  </div>
                </div>
                <div
                  className="bchs-bg-track"
                  aria-hidden={personnelStayStats.total === 0}
                >
                  <i
                    style={{
                      width: formatRatioPercent(
                        personnelStayStats.battleReady,
                        personnelStayStats.total,
                      ),
                    }}
                  />
                </div>
              </div>
            </div>
          </section>

          <div className="bchs-overview-title">
            <Typography component="h2" variant="h3">
              Аналітика БЧС
            </Typography>
            <Typography variant="body1">
              Станом на {analytics.reportDate || "дату з Аркуш1"}
            </Typography>
          </div>

          <div className="bchs-overview-grid">
            <div className="bchs-balance-panel">
              <div className="panel-heading">Баланс фактичного строю</div>
              <div className="bchs-balance-flow">
                {[
                  ["за списком", totalComparison.listed, null],
                  ["відсутні", totalComparison.absent, "-"],
                  [
                    "в інших підрозділах",
                    totalComparison.awayInOtherUnits,
                    "-",
                  ],
                  [
                    "прикомандировані",
                    totalComparison.attachedFromOtherUnits,
                    "+",
                  ],
                ].map(([label, value, operator]) => (
                  <div className="bchs-balance-item" key={label}>
                    {operator && (
                      <span className="bchs-balance-op">{operator}</span>
                    )}
                    <strong>{value}</strong>
                    <span>{label}</span>
                  </div>
                ))}
                <div className="bchs-balance-item result">
                  <span className="bchs-balance-op">=</span>
                  <strong>{totalComparison.balanceActual}</strong>
                  <span>фактично</span>
                </div>
              </div>
            </div>

            <div className="bchs-signal-panel">
              <div className="panel-heading">Ключові сигнали</div>
              <div className="bchs-signal-list">
                {keySignals.map((row, index) => (
                  <div className="bchs-signal-row" key={row.rowNumber}>
                    <span className="bchs-signal-icon">
                      {index === 0 ? (
                        <GroupsOutlinedIcon />
                      ) : index === 1 ? (
                        <ShieldOutlinedIcon />
                      ) : (
                        <SyncAltOutlinedIcon />
                      )}
                    </span>
                    <span>
                      <strong>{row.unit}</strong>
                      <em>
                        {formatRatioPercent(row.inRanksActually, row.staff)}{" "}
                        фактично
                      </em>
                    </span>
                  </div>
                ))}
                <div className="bchs-signal-row warn">
                  <span className="bchs-signal-icon">
                    <WarningAmberOutlinedIcon />
                  </span>
                  <span>
                    <strong>{analytics.total.shortage}</strong>
                    <em>вакантних посад</em>
                  </span>
                </div>
              </div>
            </div>

            <div className="bchs-fill-panel">
              <div className="panel-heading">
                Укомплектованість підрозділів
                <span>
                  <i /> За списком <i className="actual" /> Фактично
                </span>
              </div>
              <div className="bchs-fill-bars">
                {overviewUnits.map((row) => {
                  const listedPercent =
                    row.staff > 0 ? row.listed / row.staff : 0;
                  const actualPercent =
                    row.staff > 0 ? row.inRanksActually / row.staff : 0;

                  return (
                    <div className="bchs-fill-row" key={row.rowNumber}>
                      <span>{row.unit}</span>
                      <div>
                        <i
                          style={{
                            width: `${Math.min(100, listedPercent * 100)}%`,
                          }}
                        />
                        <strong>
                          {formatRatioPercent(row.listed, row.staff)}
                        </strong>
                        <i
                          className="actual"
                          style={{
                            width: `${Math.min(100, actualPercent * 100)}%`,
                          }}
                        />
                        <strong>
                          {formatRatioPercent(row.inRanksActually, row.staff)}
                        </strong>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            <div className="bchs-absence-panel">
              <div className="panel-heading">
                Причини відсутності · {analytics.total.absent}
              </div>
              <div className="bchs-absence-content">
                <BchsAbsenceDonut
                  items={analytics.absenceReasons}
                  total={analytics.total.absent}
                />
                <div className="bchs-absence-legend">
                  {analytics.absenceReasons.map((item, index) => (
                    <div key={item.label}>
                      <span
                        style={{
                          background: [
                            "#7f9f43",
                            "#ead15d",
                            "#e3a128",
                            "#b56549",
                            "#8aa0a0",
                            "#9a9a91",
                          ][index % 6],
                        }}
                      />
                      <em>{item.label}</em>
                      <strong>{item.value}</strong>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </section>
      )}

      {activeBchsAnalyticsTab === "combat" && (
        <section className="bchs-combat">
          <div className="bchs-combat-header">
            <div>
              <Typography component="h2" variant="h3">
                Бойова складова та готовність
              </Typography>
              <Typography variant="body1">
                Фактично в строю · {analytics.total.inRanksActually}
              </Typography>
            </div>
            <SciButton
              variant="OUTLINE"
              onClick={() => setActiveBchsAnalyticsTab("comparison")}
            >
              Перейти до осіб
            </SciButton>
          </div>

          <div className="bchs-combat-layout">
            <div className="bchs-combat-left">
              <div className="bchs-combat-panel">
                <div className="panel-heading">
                  Склад фактичного строю{" "}
                  <strong>{analytics.total.inRanksActually}</strong>
                </div>
                <div className="combat-composition-list">
                  {combatComposition.map((item) => (
                    <div
                      className={`combat-composition-row ${item.tone}`}
                      key={item.label}
                    >
                      <span>{item.label}</span>
                      <strong>{item.value}</strong>
                    </div>
                  ))}
                </div>
                <div className="combat-panel-footer">
                  <span>Усього фактично в строю</span>
                  <strong>{analytics.total.inRanksActually}</strong>
                </div>
              </div>

              <div className="bchs-combat-panel">
                <div className="panel-heading">
                  Піхота · {analytics.total.assaultTotal}
                </div>
                <div className="combat-readiness-grid">
                  {combatReadiness.map((item) => (
                    <div
                      className={`combat-readiness-card ${item.tone}`}
                      key={item.label}
                    >
                      <span>{item.label}</span>
                      <strong>{item.value}</strong>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div
              className={`bchs-combat-visual${combatFullscreen ? " is-fullscreen" : ""}`}
            >
              <div className="panel-heading">
                <div className="combat-visual-heading">
                  <div>
                    3D структура підрозділів
                    <span>
                      На сцені — пропорційно строю (до {MAX_VISUAL_SOLDIERS}{" "}
                      фігур) · Idle / Walk
                      {combatFullscreen ? " · OrbitControls" : ""}
                    </span>
                  </div>
                  <IconButton
                    aria-label={
                      combatFullscreen
                        ? "Вийти з повного екрана"
                        : "На весь екран"
                    }
                    className="combat-fullscreen-btn"
                    size="small"
                    title={
                      combatFullscreen
                        ? "Вийти з повного екрана (Esc)"
                        : "На весь екран"
                    }
                    onClick={() => setCombatFullscreenMode(!combatFullscreen)}
                  >
                    {combatFullscreen ? (
                      <FullscreenExitOutlinedIcon fontSize="small" />
                    ) : (
                      <FullscreenOutlinedIcon fontSize="small" />
                    )}
                  </IconButton>
                </div>
              </div>
              <div className="combat-canvas-wrap">
                <CombatModelsPreloader sceneReady={combatSceneReady} />
                <Canvas
                  key={combatFullscreen ? "combat-fs" : "combat-normal"}
                  camera={{
                    position: [1.8, 3.8, 8.4],
                    fov: combatFullscreen ? 42 : 38,
                  }}
                  dpr={[1, 1.25]}
                  style={{
                    pointerEvents: combatFullscreen ? "auto" : "none",
                  }}
                >
                  <Suspense fallback={null}>
                    <CombatStructureScene
                      blocks={combatBlocks}
                      cameraResetKey={combatCameraResetKey}
                      orbitEnabled={combatFullscreen}
                      selectedLabel={selectedCombatBlock?.label}
                      onReady={markCombatSceneReady}
                    />
                  </Suspense>
                </Canvas>
                <div className="combat-block-labels">
                  {combatBlocks.map((block) => (
                    <button
                      className={
                        selectedCombatBlock?.label === block.label
                          ? "active"
                          : ""
                      }
                      key={block.label}
                      type="button"
                      onClick={() => selectCombatUnit(block.label)}
                    >
                      <span>{block.label}</span>
                      <strong>
                        {block.actual} / {block.staff}
                      </strong>
                    </button>
                  ))}
                </div>
              </div>
            </div>

            <aside className="bchs-combat-attention">
              <div className="panel-heading danger-text">Потребує уваги</div>
              {[
                {
                  label: "без БЗВП",
                  value: analytics.total.noBzvp,
                  icon: <GroupsOutlinedIcon />,
                },
                {
                  label: "на відновленні",
                  value: analytics.total.assaultRecovery,
                  icon: <SyncAltOutlinedIcon />,
                },
                {
                  label: "відсутні",
                  value: analytics.total.absent,
                  icon: <ErrorOutlineOutlinedIcon />,
                },
                {
                  label: "некомплект",
                  value: analytics.total.shortage,
                  icon: <WarningAmberOutlinedIcon />,
                },
              ].map((item) => (
                <div className="combat-attention-row" key={item.label}>
                  <span>{item.icon}</span>
                  <strong>{item.value}</strong>
                  <em>{item.label}</em>
                </div>
              ))}
              {anomalyDataIssues.length > 0 && (
                <div className="bchs-data-issues bchs-anomaly-issues">
                  <div className="combat-attention-row">
                    <span>
                      <WarningAmberOutlinedIcon />
                    </span>
                    <strong>{anomalyDataIssues.length}</strong>
                    <em>БРЕЗ / аномалії підрахунку</em>
                  </div>
                  <div className="bchs-data-issues-list">
                    {anomalyDataIssues.slice(0, 10).map((issue, index) => (
                      <div
                        className="bchs-data-issue"
                        key={`${issue.fullName}-anomaly-${index}`}
                      >
                        <strong>{issue.fullName}</strong>
                        <span>
                          {issue.rosterUnit}
                          {issue.status ? ` · ${issue.status}` : ""}
                        </span>
                        <em>{issue.reason}</em>
                      </div>
                    ))}
                    {anomalyDataIssues.length > 10 && (
                      <p>ще {anomalyDataIssues.length - 10} у списку</p>
                    )}
                  </div>
                </div>
              )}
              {statusDataIssues.length > 0 && (
                <div className="bchs-data-issues">
                  <div className="combat-attention-row">
                    <span>
                      <HelpOutlineOutlinedIcon />
                    </span>
                    <strong>{statusDataIssues.length}</strong>
                    <em>невідомий статус</em>
                  </div>
                  <div className="bchs-data-issues-list">
                    {statusDataIssues.slice(0, 8).map((issue, index) => (
                      <div
                        className="bchs-data-issue"
                        key={`${issue.fullName}-${issue.status}-${index}`}
                      >
                        <strong>{issue.fullName}</strong>
                        <span>
                          {issue.rosterUnit} · {issue.status}
                        </span>
                        <em>{issue.reason}</em>
                      </div>
                    ))}
                    {statusDataIssues.length > 8 && (
                      <p>ще {statusDataIssues.length - 8} у списку аналітики</p>
                    )}
                  </div>
                </div>
              )}
              {rankDataIssues.length > 0 && (
                <div className="bchs-data-issues bchs-rank-issues">
                  <div className="combat-attention-row">
                    <span>
                      <WarningAmberOutlinedIcon />
                    </span>
                    <strong>{rankDataIssues.length}</strong>
                    <em>помилки звань</em>
                  </div>
                  <div className="bchs-data-issues-list">
                    {rankDataIssues.slice(0, 12).map((issue, index) => (
                      <div
                        className="bchs-data-issue"
                        key={`${issue.fullName}-rank-${index}`}
                      >
                        <strong>{issue.fullName}</strong>
                        <span>
                          {issue.rosterUnit}
                          {issue.rankTitle ? ` · ${issue.rankTitle}` : ""}
                          {issue.rankCategory ? ` · ${issue.rankCategory}` : ""}
                        </span>
                        <em>{issue.reason}</em>
                      </div>
                    ))}
                    {rankDataIssues.length > 12 && (
                      <p>ще {rankDataIssues.length - 12} у списку</p>
                    )}
                  </div>
                </div>
              )}
            </aside>
          </div>

          <div className="bchs-capability-flow">
            <div className="capability-flow-canvas">
              <h3>Від списку до спроможності</h3>
              <svg
                aria-hidden="true"
                className="capability-lines"
                preserveAspectRatio="none"
                viewBox="0 0 1280 330"
              >
                <defs>
                  <marker
                    id="cap-arrow-neutral"
                    markerHeight="5"
                    markerUnits="userSpaceOnUse"
                    markerWidth="5"
                    orient="auto"
                    refX="4.6"
                    refY="2.5"
                    viewBox="0 0 5 5"
                  >
                    <path d="M0,0 L5,2.5 L0,5 Z" />
                  </marker>
                  <marker
                    id="cap-arrow-red"
                    markerHeight="5"
                    markerUnits="userSpaceOnUse"
                    markerWidth="5"
                    orient="auto"
                    refX="4.6"
                    refY="2.5"
                    viewBox="0 0 5 5"
                  >
                    <path d="M0,0 L5,2.5 L0,5 Z" />
                  </marker>
                  <marker
                    id="cap-arrow-orange"
                    markerHeight="5"
                    markerUnits="userSpaceOnUse"
                    markerWidth="5"
                    orient="auto"
                    refX="4.6"
                    refY="2.5"
                    viewBox="0 0 5 5"
                  >
                    <path d="M0,0 L5,2.5 L0,5 Z" />
                  </marker>
                  <marker
                    id="cap-arrow-cyan"
                    markerHeight="5"
                    markerUnits="userSpaceOnUse"
                    markerWidth="5"
                    orient="auto"
                    refX="4.6"
                    refY="2.5"
                    viewBox="0 0 5 5"
                  >
                    <path d="M0,0 L5,2.5 L0,5 Z" />
                  </marker>
                  <marker
                    id="cap-arrow-green"
                    markerHeight="5"
                    markerUnits="userSpaceOnUse"
                    markerWidth="5"
                    orient="auto"
                    refX="4.6"
                    refY="2.5"
                    viewBox="0 0 5 5"
                  >
                    <path d="M0,0 L5,2.5 L0,5 Z" />
                  </marker>
                </defs>
                <path
                  className="capability-line neutral"
                  d="M248 164 C300 164 300 74 348 74"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-neutral)"
                />
                <path
                  className="capability-line neutral"
                  d="M248 164 L348 164"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-neutral)"
                />
                <path
                  className="capability-line neutral"
                  d="M248 164 C300 164 300 254 348 254"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-neutral)"
                />
                <path
                  className="capability-line red"
                  d="M660 74 C720 74 710 168 768 168"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-red)"
                />
                <path
                  className="capability-line orange"
                  d="M660 166 L768 166"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-orange)"
                />
                <path
                  className="capability-line cyan"
                  d="M660 254 C720 254 710 168 768 168"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-cyan)"
                />
                <path
                  className="capability-line green"
                  d="M970 168 C992 168 992 66 1016 66"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-green)"
                />
                <path
                  className="capability-line cyan"
                  d="M970 168 C992 168 992 141 1016 141"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-cyan)"
                />
                <path
                  className="capability-line neutral"
                  d="M970 168 C992 168 992 216 1016 216"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-neutral)"
                />
                <path
                  className="capability-line yellow"
                  d="M970 168 C992 168 992 291 1016 291"
                  vectorEffect="non-scaling-stroke"
                  markerEnd="url(#cap-arrow-orange)"
                />
              </svg>

              <div className="capability-node capability-start">
                <GroupsOutlinedIcon />
                <strong>{analytics.total.listed}</strong>
                <span>за списком</span>
              </div>
              <div className="capability-node capability-mid capability-bad">
                <ErrorOutlineOutlinedIcon />
                <strong>{analytics.total.absent}</strong>
                <span>відсутні</span>
              </div>
              <div className="capability-node capability-mid capability-warn">
                <SyncAltOutlinedIcon />
                <strong>{analytics.total.awayInOtherUnits}</strong>
                <span>в інших підрозділах</span>
              </div>
              <div className="capability-node capability-mid capability-add">
                <CheckCircleOutlineIcon />
                <strong>+{analytics.total.attachedFromOtherUnits}</strong>
                <span>прикомандировані</span>
              </div>
              <div className="capability-node capability-result">
                <GroupsOutlinedIcon />
                <strong>{analytics.total.inRanksActually}</strong>
                <span>фактично</span>
              </div>
              {[
                {
                  className: "capability-green",
                  icon: <GroupsOutlinedIcon />,
                  label: "піхота",
                  value: analytics.total.assaultTotal,
                },
                {
                  className: "capability-cyan",
                  icon: <MilitaryTechOutlinedIcon />,
                  label: "БПЛА",
                  value: analytics.total.droneCrew,
                },
                {
                  className: "capability-neutral",
                  icon: <BuildOutlinedIcon />,
                  label: "забезпечення",
                  value: analytics.total.supportCombat,
                },
                {
                  className: "capability-yellow",
                  icon: <WarningAmberOutlinedIcon />,
                  label: "інше",
                  value: Math.max(
                    0,
                    analytics.total.inRanksActually -
                      analytics.total.assaultTotal -
                      analytics.total.droneCrew -
                      analytics.total.supportCombat,
                  ),
                },
              ].map((item, index) => (
                <div
                  className={`capability-node capability-output ${item.className}`}
                  key={item.label}
                  style={{ top: `${34 + index * 75}px` }}
                >
                  {item.icon}
                  <strong>{item.value}</strong>
                  <span>{item.label}</span>
                </div>
              ))}
            </div>
          </div>
        </section>
      )}

      {activeBchsAnalyticsTab === "comparison" && (
        <section className="bchs-comparison">
          <div className="bchs-comparison-header">
            <div>
              <Typography component="h2" variant="h4">
                Порівняння підрозділів
              </Typography>
              <Typography variant="body2" color="text.secondary">
                Штат · список · наявність · фактичний стрій
              </Typography>
            </div>
            <div
              className="bchs-comparison-tabs"
              aria-label="Вкладки порівняння БЧС"
            >
              <button className="active" type="button">
                Укомплектованість
              </button>
              <button type="button" disabled>
                Фактична присутність
              </button>
              <button type="button" disabled>
                Бойова складова
              </button>
            </div>
            <label className="bchs-search">
              <SearchOutlinedIcon fontSize="small" />
              <input
                value={comparisonQuery}
                onChange={(event) => setComparisonQuery(event.target.value)}
                placeholder="Знайти підрозділ"
              />
            </label>
            <Select
              size="small"
              value={comparisonSort}
              onChange={(event) =>
                setComparisonSort(String(event.target.value))
              }
              sx={{
                minWidth: 190,
                color: "#d9d49d",
                backgroundColor: "rgba(10, 10, 9, 0.78)",
                fontFamily: "inherit",
              }}
              MenuProps={{
                slotProps: {
                  paper: {
                    sx: {
                      color: "#f2eee1",
                      backgroundColor: "#171713",
                      border: "1px solid rgba(230, 224, 190, 0.18)",
                    },
                  },
                },
              }}
            >
              <MenuItem value="critical">Спочатку критичні</MenuItem>
              <MenuItem value="absent">Більше відсутніх</MenuItem>
              <MenuItem value="staff">Більший штат</MenuItem>
              <MenuItem value="excel">Як в Excel</MenuItem>
            </Select>
          <SciButton
            variant="OUTLINE"
            onClick={exportComparison}
          >
              <FileDownloadOutlinedIcon fontSize="small" />
              Експорт
          </SciButton>
          </div>

          <div className="bchs-comparison-layout">
            <div className="bchs-comparison-table">
              <div className="bchs-comparison-row head">
                <span>Підрозділ</span>
                <span>Штат</span>
                <span>Список</span>
                <span>Фактично</span>
                <span>Рівень</span>
                <span>Відсутні</span>
              </div>
              {comparisonRows.map((row, index) => (
                <button
                  className={`bchs-comparison-row ${selectedComparison?.rowNumber === row.rowNumber ? "selected" : ""}`}
                  key={row.rowNumber}
                  type="button"
                  onClick={() => setSelectedComparisonRow(row.rowNumber)}
                >
                  <span className="unit-cell">
                    <em>{index + 1}</em>
                    {row.unit}
                  </span>
                  <span>
                    <strong>{row.staff}</strong>
                    <i
                      style={{
                        width: `${Math.min(100, (row.staff / comparisonMax) * 100)}%`,
                      }}
                    />
                  </span>
                  <span>
                    <strong>{row.listed}</strong>
                    <i
                      className="green"
                      style={{
                        width: `${Math.min(100, (row.listed / comparisonMax) * 100)}%`,
                      }}
                    />
                  </span>
                  <span>
                    <strong>{row.inRanksActually}</strong>
                    <i
                      className="amber"
                      style={{
                        width: `${Math.min(100, (row.inRanksActually / comparisonMax) * 100)}%`,
                      }}
                    />
                  </span>
                  <span
                    className={
                      row.levelPercent < 0.5
                        ? "danger-text"
                        : row.levelPercent < 0.7
                          ? "warn-text"
                          : "good-text"
                    }
                  >
                    {formatRatioPercent(row.inRanksActually, row.staff)}
                  </span>
                  <span className={row.absent > 0 ? "danger-text" : ""}>
                    {row.absent}
                  </span>
                </button>
              ))}
            </div>

            <aside className="bchs-comparison-detail">
              {selectedComparison ? (
                <>
                  <div className="panel-heading">
                    Обрано: {selectedComparison.unit}
                  </div>
                  <div className="bchs-presence-formula">
                    {[
                      ["за списком", selectedComparison.listed, "good"],
                      ["відсутні", selectedComparison.absent, "bad"],
                      [
                        "в інших підрозділах",
                        selectedComparison.awayInOtherUnits,
                        "info",
                      ],
                      [
                        "прикомандировані",
                        selectedComparison.attachedFromOtherUnits,
                        "info",
                      ],
                    ].map(([label, value, tone], index) => (
                      <div className={`formula-box ${tone}`} key={label}>
                        <strong>{value}</strong>
                        <span>{label}</span>
                        {index < 3 && <em>{index === 2 ? "+" : "-"}</em>}
                      </div>
                    ))}
                    <div className="formula-result">
                      <strong>{selectedComparison.balanceActual}</strong>
                      <span>фактично</span>
                    </div>
                  </div>

                  <div className="panel-heading">Склад за категоріями</div>
                  <BarList
                    maxValue={Math.max(
                      1,
                      selectedComparison.actualOfficers,
                      selectedComparison.actualSergeants,
                      selectedComparison.actualSoldiers,
                    )}
                    items={[
                      {
                        label: "Офіцери",
                        value: selectedComparison.actualOfficers,
                      },
                      {
                        label: "Сержанти",
                        value: selectedComparison.actualSergeants,
                      },
                      {
                        label: "Солдати",
                        value: selectedComparison.actualSoldiers,
                      },
                    ]}
                  />
                  {selectedComparison.attachedSourcesText ? (
                    <>
                      <div className="panel-heading">
                        Звідки прикомандировані ·{" "}
                        {selectedComparison.attachedFromOtherUnits}
                      </div>
                      <BarList
                        items={Array.from(
                          parseBchsDestinationText(
                            selectedComparison.attachedSourcesText,
                          ).entries(),
                        ).map(([label, value]) => ({ label, value }))}
                      />
                    </>
                  ) : null}

                  <div className="bchs-warning-line">
                    <WarningAmberOutlinedIcon fontSize="small" />
                    <span>Без БЗВП</span>
                    <strong>{selectedComparison.noBzvp}</strong>
                  </div>
                </>
              ) : (
                <Typography variant="body2" color="text.secondary">
                  Натисніть на рядок, щоб переглянути деталі підрозділу.
                </Typography>
              )}
            </aside>
          </div>

          <div className="bchs-comparison-note">
            <HelpOutlineOutlinedIcon fontSize="small" />
            <strong>Як читати</strong>
            <span>
              Понад 100% фактичної присутності означає підсилення
              прикомандированими, це не штатна укомплектованість.
            </span>
          </div>
        </section>
      )}

      {activeBchsAnalyticsTab === "supplement" && (
        <section className="bchs-supplement">
          <div className="bchs-overview-title">
            <Typography component="h2" variant="h4">
              Додаткова аналітика БЧС
            </Typography>
            <Typography variant="body2" color="text.secondary">
              {analytics.supplement
                ? `${analytics.supplement.title} · формули з колонок Excel`
                : "Завантажте файл `1ПБ Особовий склад підрозділів + БЗВП` або `1ПБ БЧС додаток`."}
            </Typography>
          </div>

          {analytics.supplement ? (
            <>
              <div className="bchs-supplement-kpis">
                {[
                  ["Штат", analytics.supplement.total.staff],
                  ["За списком", analytics.supplement.total.listed],
                  [
                    analytics.supplement.kind === "appendix"
                      ? "В строю"
                      : "Наявність",
                    analytics.supplement.total.inRanks ||
                      analytics.supplement.total.available,
                  ],
                  ["Відсутні", analytics.supplement.total.absent],
                  [
                    analytics.supplement.kind === "appendix"
                      ? "Бойова складова"
                      : "На виконанні БЗ",
                    analytics.supplement.kind === "appendix"
                      ? analytics.supplement.total.assaultTotal +
                          analytics.supplement.total.vehicleCrew +
                          analytics.supplement.total.droneCrew +
                          analytics.supplement.total.crewServedWeapons +
                          analytics.supplement.total.commandCombat +
                          analytics.supplement.total.supportCombat
                      : analytics.supplement.total.combatTask,
                  ],
                  [
                    analytics.supplement.kind === "appendix"
                      ? "Без БЗВП"
                      : "БЗВП разом",
                    analytics.supplement.kind === "appendix"
                      ? analytics.supplement.total.noBzvp
                      : analytics.supplement.total.totalBzvp,
                  ],
                ].map(([label, value]) => (
                  <div className="bchs-supplement-kpi" key={label}>
                    <span>{label}</span>
                    <strong>{value}</strong>
                  </div>
                ))}
              </div>

              <div className="bchs-supplement-grid">
                <div className="analytics-panel bchs-supplement-table-panel">
                  <div className="panel-heading">
                    {analytics.supplement.kind === "appendix"
                      ? "Рядки додатка"
                      : "Підсумки підрозділів"}
                  </div>
                  <div className="bchs-supplement-table-wrap">
                    <table className="bchs-supplement-table">
                      <thead>
                        <tr>
                          <th>Підрозділ</th>
                          <th>Штат</th>
                          <th>Список</th>
                          <th>
                            {analytics.supplement.kind === "appendix"
                              ? "В строю"
                              : "Наявність"}
                          </th>
                          <th>%</th>
                          <th>
                            {analytics.supplement.kind === "appendix"
                              ? "Відсутні"
                              : "БЗ"}
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {(analytics.supplement.kind === "appendix"
                          ? analytics.supplement.rows
                          : analytics.supplement.totals
                        ).map((row) => (
                          <tr key={`${row.rowNumber}-${row.unit}`}>
                            <td>
                              <strong>{row.unit}</strong>
                              {row.battalion && <span>{row.battalion}</span>}
                            </td>
                            <td>{row.staff}</td>
                            <td>{row.listed}</td>
                            <td>{row.inRanks || row.available}</td>
                            <td>{toPercent(row.staffedPercent)}</td>
                            <td>
                              {analytics.supplement?.kind === "appendix"
                                ? row.absent
                                : row.combatTask}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <div className="analytics-panel">
                  <div className="panel-heading">
                    {analytics.supplement.kind === "appendix"
                      ? "Причини відсутності"
                      : "Резерви"}
                  </div>
                  <BarList
                    items={
                      analytics.supplement.kind === "appendix"
                        ? analytics.supplement.absenceReasons
                        : analytics.supplement.reserveMetrics
                    }
                  />
                </div>

                <div className="analytics-panel">
                  <div className="panel-heading">
                    {analytics.supplement.kind === "appendix"
                      ? "Бойові категорії"
                      : "БЗВП по періодах"}
                  </div>
                  <BarList
                    items={
                      analytics.supplement.kind === "appendix"
                        ? analytics.supplement.combatCategories
                        : analytics.supplement.bzvpBuckets
                    }
                  />
                </div>
              </div>
            </>
          ) : (
            <div className="drop-zone">
              <Box>
                <CloudUploadOutlinedIcon color="disabled" />
                <Typography variant="body2">
                  Додаткова БЧС-аналітика для цього формату не знайдена.
                </Typography>
              </Box>
            </div>
          )}
        </section>
      )}

      <section className="bchs-grid">
        <div className="analytics-panel">
          <div className="panel-heading">Укомплектованість</div>
          <div className="bchs-big-percent">
            {toPercent(analytics.total.staffedPercent)}
          </div>
          <BarList
            items={[
              { label: "Наявність", value: analytics.total.available },
              { label: "Відсутні", value: analytics.total.absent },
              { label: "Некомплект", value: analytics.total.shortage },
              {
                label: "Бойова складова",
                value: analytics.total.combatComponent,
              },
            ]}
          />
        </div>
        <div className="analytics-panel">
          <div className="panel-heading">Причини відсутності</div>
          <BarList
            items={[
              { label: "Відрядження", value: analytics.total.businessTrip },
              { label: "Навчання", value: analytics.total.training },
              {
                label: "Шпиталь поранення",
                value: analytics.total.hospitalWounded,
              },
              {
                label: "Шпиталь хвороба",
                value: analytics.total.hospitalIllness,
              },
              { label: "Відпустка", value: analytics.total.vacation },
              { label: "СЗЧ", value: analytics.total.awol },
              { label: "Зниклі безвісти", value: analytics.total.missing },
              { label: "Загиблі", value: analytics.total.killed },
            ]}
          />
        </div>
        <div className="analytics-panel">
          <div className="panel-heading">
            Куди відкомандировані ·{" "}
            {analytics.total.awayInOtherUnits}
          </div>
          {analytics.detachedDestinations.length > 0 ? (
            <BarList items={analytics.detachedDestinations} />
          ) : (
            <Typography variant="body2" color="text.secondary">
              Немає розпізнаних записів у форматі `Назва-кількість`.
            </Typography>
          )}
        </div>
        <div className="analytics-panel">
          <div className="panel-heading">
            Звідки прикомандировані ·{" "}
            {analytics.total.attachedFromOtherUnits}
          </div>
          {analytics.attachedSources.length > 0 ? (
            <BarList items={analytics.attachedSources} />
          ) : (
            <Typography variant="body2" color="text.secondary">
              Немає розпізнаних записів у форматі `Назва-кількість`.
            </Typography>
          )}
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
                {imports.reduce((sum, item) => sum + item.totalRows, 0)}
              </strong>
              <span>рядків</span>
            </div>
          </div>
          <div className="ejournal-db-controls">
            <label>
              <span>Імпорт</span>
              <div className="import-select-row">
                <Select
                  displayEmpty
                  size="small"
                  value={selectedImport?.id ?? ""}
                  onChange={(event) => {
                    setSelectedImportId(String(event.target.value));
                    setDbPreview(null);
                  }}
                  renderValue={(value) => {
                    const item = imports.find(
                      (importItem) => importItem.id === value,
                    );

                    if (!item) return "Імпорт не вибрано";

                    return `${item.sourceFileName ?? item.name} · ${new Date(item.createdAt).toLocaleString("uk-UA")}`;
                  }}
                  MenuProps={{
                    slotProps: {
                      paper: {
                        sx: {
                          color: "#f2eee1",
                          backgroundColor: "#171713",
                          border: "1px solid rgba(230, 224, 190, 0.18)",
                          borderRadius: "6px",
                        },
                      },
                    },
                  }}
                  sx={{
                    minWidth: 0,
                    color: "#f2eee1",
                    backgroundColor: "rgba(10, 10, 9, 0.78)",
                    fontFamily: "inherit",
                  }}
                >
                  {imports.length === 0 && (
                    <MenuItem value="">Імпортів немає</MenuItem>
                  )}
                  {imports.map((item) => (
                    <MenuItem key={item.id} value={item.id}>
                      <Box
                        component="span"
                        sx={{
                          overflow: "hidden",
                          textOverflow: "ellipsis",
                          whiteSpace: "nowrap",
                        }}
                      >
                        {item.sourceFileName ?? item.name} ·{" "}
                        {new Date(item.createdAt).toLocaleString("uk-UA")}
                      </Box>
                    </MenuItem>
                  ))}
                </Select>
                <IconButton
                  aria-label="Видалити імпорт БЧС"
                  disabled={!selectedImport || isLoading || !canEdit}
                  onClick={() => void deleteSelectedImport()}
                  sx={{
                    width: 40,
                    height: 40,
                    color: "#e66b5b",
                    backgroundColor: "rgba(10, 10, 9, 0.78)",
                    border: "1px solid rgba(230, 224, 190, 0.18)",
                    "&:hover": {
                      color: "#ff7565",
                      backgroundColor: "rgba(130, 45, 35, 0.22)",
                      borderColor: "rgba(230, 84, 65, 0.5)",
                    },
                  }}
                >
                  <DeleteOutlineOutlinedIcon fontSize="small" />
                </IconButton>
              </div>
            </label>
            <SciButton
              disabled={!selectedDbSheet}
              variant="OUTLINE"
              onClick={() => void loadDbPreview()}
            >
              Показати з БД
            </SciButton>
          </div>
        </div>
      </section>

      {detachedDestinationIssues.length > 0 && (
        <section className="analytics-panel bchs-unknown-destination-panel">
          <div className="panel-heading">
            Напрямок `невідомо` · {detachedDestinationIssues.length}
          </div>
          <div className="bchs-data-issues-list">
            {detachedDestinationIssues.map((issue, index) => (
              <div
                className="bchs-data-issue"
                key={`${issue.fullName}-${issue.destination}-${index}`}
              >
                <strong>{issue.fullName}</strong>
                <span>
                  {issue.rosterUnit} · {issue.status}
                </span>
                <em>
                  {issue.reason} · колонка `В якому підрозділі`:{" "}
                  {issue.destination}
                </em>
              </div>
            ))}
          </div>
        </section>
      )}

      {rankDataIssues.length > 0 && (
        <section className="analytics-panel bchs-rank-issues-panel">
          <div className="panel-heading">
            Помилки / розбіжності звань · {rankDataIssues.length}
          </div>
          <p className="bchs-rank-issues-hint">
            Показуємо лише підозріле написання звання (змішана латиниця/кирилиця).
            Українське «капітан» вважається правильним і в підрахунках враховується.
          </p>
          <div className="bchs-data-issues-list">
            {rankDataIssues.map((issue, index) => (
              <div
                className="bchs-data-issue"
                key={`${issue.fullName}-rank-panel-${index}`}
              >
                <strong>{issue.fullName}</strong>
                <span>
                  {issue.rosterUnit}
                  {issue.rankTitle ? ` · звання: ${issue.rankTitle}` : ""}
                  {issue.rankCategory
                    ? ` · категорія: ${issue.rankCategory}`
                    : ""}
                </span>
                <em>{issue.reason}</em>
              </div>
            ))}
          </div>
        </section>
      )}

      <section className="panel table-panel">
        <div className="panel-heading">Повна аналітика Аркуш1</div>
        {hasBchsAnalyticsTable ? (
          <div className="bchs-analytics-table-wrap">
            <table className="bchs-analytics-table">
              <thead>
                <tr>
                  {analyticsTable.columns.map((column) => (
                    <th
                      className={
                        ["B", "P", "AO", "AT"].includes(column.letter)
                          ? "bchs-text-col"
                          : undefined
                      }
                      key={column.key}
                      title={column.label}
                    >
                      <span>{column.letter}</span>
                      {column.label}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {analyticsTable.rows.map((row) => (
                  <tr key={row.rowNumber}>
                    {analyticsTable.columns.map((column) => (
                      <td
                        className={[
                          column.isPercent ? "percent" : "",
                          ["B", "P", "AO", "AT"].includes(column.letter)
                            ? "bchs-text-col"
                            : "",
                        ]
                          .filter(Boolean)
                          .join(" ") || undefined}
                        key={`${row.rowNumber}-${column.key}`}
                      >
                        {row.values[column.key]}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="drop-zone">
            <Box>
              <CloudUploadOutlinedIcon color="disabled" />
              <Typography variant="body2">
                Для цього імпорту немає snapshot `Аркуш1`.
              </Typography>
              <Typography variant="caption" className="muted">
                Завантажте `Рассчет.xlsx` і натисніть `Записати Аркуш2`, щоб
                зберегти повну аналітику разом з імпортом.
              </Typography>
            </Box>
          </div>
        )}
      </section>

      <section className="panel table-panel">
        <div className="panel-heading">
          {dbPreview
            ? `Аркуш2 з БД · ${dbPreview.rows.length}/${dbPreview.total}`
            : "Аркуш2 · preview"}
        </div>
        <div className="panel-body">
          {previewRows.length > 0 ? (
            <MaterialReactTable table={table} />
          ) : (
            <div className="drop-zone">
              <Box>
                <CloudUploadOutlinedIcon color="disabled" />
                <Typography variant="body2">
                  Додайте `Рассчет.xlsx` або відкрийте імпорт з БД
                </Typography>
              </Box>
            </div>
          )}
        </div>
      </section>
    </main>
  );
}
