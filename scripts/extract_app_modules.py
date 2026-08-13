#!/usr/bin/env python3
"""Mechanical extraction of App.tsx into modules. Logic preserved; exports/imports added."""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path("/Users/konstantinabramkin/Desktop/projects/army_projects/army-grid/src")
APP = Path("/tmp/App.tsx.original")
lines = APP.read_text().splitlines()


def slice_lines(start: int, end: int) -> str:
    """1-indexed inclusive."""
    return "\n".join(lines[start - 1 : end])


def export_top_level(code: str) -> str:
    out = []
    for line in code.splitlines():
        if re.match(r"^(type|function|const) ", line):
            out.append("export " + line)
        else:
            out.append(line)
    return "\n".join(out)


def write(rel: str, content: str) -> None:
    path = ROOT / rel
    path.parent.mkdir(parents=True, exist_ok=True)
    if not content.endswith("\n"):
        content += "\n"
    path.write_text(content)
    print(f"wrote {rel} ({len(content.splitlines())} lines)")


# ---------------------------------------------------------------------------
# shared/format.ts
# ---------------------------------------------------------------------------
write(
    "shared/format.ts",
    '''export const formatFileSize = (size: number) => `${(size / 1024).toFixed(1)} KB`;

export const formatDateTime = () =>
  new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(new Date());

export const normalizeDatasetKey = (value: string) =>
  value
    .trim()
    .replace(/\\s+/g, "_")
    .replace(/[^\\p{L}\\p{N}_-]/gu, "")
    .toLowerCase();

export const cellValueToJson = (value: unknown) => {
  if (value === undefined || value === null) return null;
  if (value instanceof Date) return value.toISOString();
  return value;
};
''',
)

# ---------------------------------------------------------------------------
# app/navigation.ts
# ---------------------------------------------------------------------------
write(
    "app/navigation.ts",
    '''export type AppPage =
  | "import"
  | "analytics"
  | "ejournal"
  | "bchs"
  | "personnel"
  | "documents";
export type BchsAnalyticsTab = "overview" | "comparison" | "combat" | "supplement";

export const pagePaths: Record<AppPage, string> = {
  import: "/import",
  analytics: "/analytics",
  ejournal: "/ejournal",
  bchs: "/bchs",
  personnel: "/personnel",
  documents: "/documents",
};

const pathPages = Object.fromEntries(
  Object.entries(pagePaths).map(([page, path]) => [path, page]),
) as Record<string, AppPage>;

export const getPageFromPath = (path: string): AppPage =>
  pathPages[path] ?? (path === "/" ? "import" : "import");

export const getInitialBchsAnalyticsTab = (): BchsAnalyticsTab => {
  const tab = new URLSearchParams(window.location.search).get("bchsTab");

  return tab === "comparison" ||
    tab === "combat" ||
    tab === "overview" ||
    tab === "supplement"
    ? tab
    : "overview";
};

export const navigateToPage = (page: AppPage) => {
  const nextPath = pagePaths[page];

  if (window.location.pathname !== nextPath) {
    window.history.pushState({ page }, "", nextPath);
  }
};
''',
)

# ---------------------------------------------------------------------------
# app/layout/Sidebar.tsx
# ---------------------------------------------------------------------------
write(
    "app/layout/Sidebar.tsx",
    '''import type { ReactNode } from "react";
import { Box, Stack, Typography } from "@mui/material";
import AnalyticsOutlinedIcon from "@mui/icons-material/AnalyticsOutlined";
import ArticleOutlinedIcon from "@mui/icons-material/ArticleOutlined";
import DashboardOutlinedIcon from "@mui/icons-material/DashboardOutlined";
import FileDownloadOutlinedIcon from "@mui/icons-material/FileDownloadOutlined";
import GridViewOutlinedIcon from "@mui/icons-material/GridViewOutlined";
import HelpOutlineOutlinedIcon from "@mui/icons-material/HelpOutlineOutlined";
import MenuOpenOutlinedIcon from "@mui/icons-material/MenuOpenOutlined";
import PersonSearchOutlinedIcon from "@mui/icons-material/PersonSearchOutlined";
import SettingsOutlinedIcon from "@mui/icons-material/SettingsOutlined";
import ShieldOutlinedIcon from "@mui/icons-material/ShieldOutlined";
import TableChartOutlinedIcon from "@mui/icons-material/TableChartOutlined";
import type { AppPage } from "../navigation";

const navItems: Array<{ label: string; page?: AppPage; icon: ReactNode }> = [
  { label: "Огляд", icon: <DashboardOutlinedIcon /> },
  { label: "Особовий склад", page: "personnel", icon: <PersonSearchOutlinedIcon /> },
  { label: "Статуси", icon: <GridViewOutlinedIcon /> },
  { label: "ЕЖООС", page: "ejournal", icon: <TableChartOutlinedIcon /> },
  { label: "БЧС", page: "bchs", icon: <ShieldOutlinedIcon /> },
  { label: "Аналітика", page: "analytics", icon: <AnalyticsOutlinedIcon /> },
  { label: "Імпорт", page: "import", icon: <FileDownloadOutlinedIcon /> },
  { label: "Документи", page: "documents", icon: <ArticleOutlinedIcon /> },
];

'''
    + export_top_level(slice_lines(198, 263))
    + "\n",
)

# ---------------------------------------------------------------------------
# pages/import/*
# ---------------------------------------------------------------------------
write(
    "pages/import/types.ts",
    '''import type { Dispatch, SetStateAction } from "react";
import type { ExcelRow, ExcelWorkbookSnapshot } from "../../excelRoundTrip";

export type DataSourceFile = {
  id: string;
  name: string;
  size: string;
  rows: number;
  columns: number;
  uploadedAt: string;
  role: "Основний файл" | "Файл для merge";
};

export type ExcelRoundTripPanelProps = {
  snapshot: ExcelWorkbookSnapshot | null;
  rows: ExcelRow[];
  setRows: Dispatch<SetStateAction<ExcelRow[]>>;
  activeSheetIndex: number;
  onActiveSheetChange: (sheetIndex: number) => void;
  isBusy: boolean;
  message: string;
  mergeSummary: string;
  onExport: () => void;
};

export type ExcelCellEditorProps = {
  value: string;
  wrapText: boolean;
  onCommit: (value: string) => void;
};
''',
)

write(
    "pages/import/DataSources.tsx",
    '''import { Box, Button, Chip, Stack, Typography } from "@mui/material";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutlineOutlined";
import CloudUploadOutlinedIcon from "@mui/icons-material/CloudUploadOutlined";
import MoreVertOutlinedIcon from "@mui/icons-material/MoreVertOutlined";
import SyncAltOutlinedIcon from "@mui/icons-material/SyncAltOutlined";
import TableChartOutlinedIcon from "@mui/icons-material/TableChartOutlined";
import UploadFileOutlinedIcon from "@mui/icons-material/UploadFileOutlined";
import type { DataSourceFile } from "./types";

'''
    + export_top_level(slice_lines(265, 394))
    + "\n",
)

write(
    "pages/import/MappingPanel.tsx",
    '''import { Button, Typography } from "@mui/material";
import SettingsOutlinedIcon from "@mui/icons-material/SettingsOutlined";
import SyncAltOutlinedIcon from "@mui/icons-material/SyncAltOutlined";

const mappingRows = [
  ["ПІБ", "Особа", "Точний збіг", "136 зіставлено"],
  ["Дата від", "Початок статусу", "Найближча дата", "140 зіставлено"],
  ["Дата до", "Кінець статусу", "Найближча дата", "138 зіставлено"],
  ["Статус", "Поточний статус", "За словником", "136 зіставлено"],
  ["Підрозділ", "Підрозділ", "Точний збіг", "134 зіставлено"],
  ["Звання", "Звання", "Точний збіг", "132 зіставлено"],
  ["Посада", "Посада", "Точний збіг", "130 зіставлено"],
  ["Коментар", "Примітка", "Додати/оновити", "46 зіставлено"],
];

'''
    + export_top_level(slice_lines(396, 469))
    + "\n",
)

write(
    "pages/import/ValidationPanel.tsx",
    '''import { Button, Typography } from "@mui/material";
import CheckCircleOutlineIcon from "@mui/icons-material/CheckCircleOutlineOutlined";
import ErrorOutlineOutlinedIcon from "@mui/icons-material/ErrorOutlineOutlined";
import WarningAmberOutlinedIcon from "@mui/icons-material/WarningAmberOutlined";

'''
    + export_top_level(slice_lines(471, 567))
    + "\n",
)

write(
    "pages/import/ExcelCellEditor.tsx",
    '''import { useLayoutEffect, useRef } from "react";
import type { ExcelCellEditorProps } from "./types";

'''
    + export_top_level(slice_lines(587, 624))
    + "\n",
)

write(
    "pages/import/ExcelRoundTripPanel.tsx",
    '''import { useMemo } from "react";
import { Button, MenuItem, Select, Stack, Typography } from "@mui/material";
import CloudUploadOutlinedIcon from "@mui/icons-material/CloudUploadOutlined";
import FileDownloadOutlinedIcon from "@mui/icons-material/FileDownloadOutlined";
import ShieldOutlinedIcon from "@mui/icons-material/ShieldOutlined";
import type { MRT_ColumnDef } from "material-react-table";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "material-react-table";
import {
  type ExcelRow,
  getColumnHeader,
  updateCell,
  valueToDisplay,
} from "../../excelRoundTrip";
import { ExcelCellEditor } from "./ExcelCellEditor";
import type { ExcelRoundTripPanelProps } from "./types";

'''
    + export_top_level(slice_lines(626, 948))
    + "\n",
)

# ---------------------------------------------------------------------------
# pages/analytics/*
# ---------------------------------------------------------------------------
analytics_data = slice_lines(950, 2130)
analytics_imports = [
    'import type { ExcelRow, ExcelWorkbookSnapshot } from "../../excelRoundTrip";',
]
if re.search(r"\bvalueToDisplay\b", analytics_data):
    analytics_imports = [
        "import {",
        "  type ExcelRow,",
        "  type ExcelWorkbookSnapshot,",
        "  valueToDisplay,",
        '} from "../../excelRoundTrip";',
    ]
write(
    "pages/analytics/analyticsData.ts",
    "\n".join(analytics_imports) + "\n\n" + export_top_level(analytics_data) + "\n",
)

write(
    "pages/analytics/charts.tsx",
    '''import * as d3 from "d3";
import type { AnalyticsMetric } from "./analyticsData";

'''
    + export_top_level(slice_lines(2132, 2247))
    + "\n\n"
    + export_top_level(slice_lines(2353, 2563))
    + "\n",
)

write(
    "pages/analytics/combatScene.tsx",
    '''import { useEffect, useRef } from "react";
import { useFrame, useThree } from "@react-three/fiber";
import type { Mesh } from "three";

'''
    + export_top_level(slice_lines(2249, 2351))
    + "\n",
)

write(
    "pages/analytics/reports.tsx",
    '''import { Button, Typography } from "@mui/material";
import ContentCopyOutlinedIcon from "@mui/icons-material/ContentCopyOutlined";
import type { AnalyticsData, AnalyticsMetric } from "./analyticsData";

'''
    + export_top_level(slice_lines(2565, 2847))
    + "\n",
)

# AnalyticsPage icons - scan
analytics_page = slice_lines(2849, 3273)
write(
    "pages/analytics/AnalyticsPage.tsx",
    '''import { useEffect, useMemo, useState } from "react";
import { Button, Typography } from "@mui/material";
import CloudUploadOutlinedIcon from "@mui/icons-material/CloudUploadOutlined";
import type { MRT_ColumnDef } from "material-react-table";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "material-react-table";
import { api } from "../../api";
import {
  type ExcelWorkbookSnapshot,
  readWorkbookSnapshot,
  valueToDisplay,
} from "../../excelRoundTrip";
import {
  makeAnalyticsData,
  type AnalyticsData,
  type AnalyticsRecord,
} from "./analyticsData";
import {
  BarList,
  DonutChart,
  HorizontalSvgChart,
  VerticalSvgChart,
} from "./charts";
import { HintAnalytics, ShortAnalytics, TextReports } from "./reports";

'''
    + export_top_level(analytics_page)
    + "\n",
)

# ---------------------------------------------------------------------------
# ejournal types + utils
# ---------------------------------------------------------------------------
write(
    "pages/ejournal/ejournalTypes.ts",
    '''import type { BackendEjournalImportSheet } from "../../api";

export type EjournalColumn = {
  key: string;
  label: string;
  order: number;
  originalIndex?: number;
  letter?: string;
};

export type EjournalPreviewRow = {
  __dbRowId?: string;
  __rowNumber?: number | null;
  [key: string]: unknown;
};

export type DbPreviewState = {
  sheet: BackendEjournalImportSheet;
  columns: EjournalColumn[];
  rows: EjournalPreviewRow[];
  total: number;
  offset: number;
  limit: number;
};
''',
)

# buildImportColumns 3275-3316; isEjournalColumn..getRowKeyByKeyPart 3380-3453
write(
    "pages/ejournal/ejournalUtils.ts",
    '''import type { ExcelRow, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { getColumnHeader, getColumnLabel } from "../../excelRoundTrip";
import { normalizeDatasetKey } from "../../shared/format";
import type { EjournalColumn, EjournalPreviewRow } from "./ejournalTypes";

'''
    + export_top_level(slice_lines(3275, 3316))
    + "\n\n"
    + export_top_level(slice_lines(3380, 3453))
    + "\n",
)

# ---------------------------------------------------------------------------
# personnel utils + components
# ---------------------------------------------------------------------------
# PersonAction types/actions 3341-3378; getPersonExternalId..loadAll 3455-3575
write(
    "pages/personnel/personnelUtils.ts",
    '''import type {
  BackendEjournalImport,
  EjournalRowActionType,
} from "../../api";
import { api } from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  getRowKeyByKeyPart,
  getRowValueByKeyPart,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";

export {
  getRowKeyByKeyPart,
  getRowValueByKeyPart,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";

'''
    + export_top_level(slice_lines(3341, 3378))
    + "\n\n"
    + export_top_level(slice_lines(3455, 3575))
    + "\n",
)

write(
    "pages/personnel/PhotoCropDialog.tsx",
    '''import { useEffect, useRef, useState, type PointerEvent } from "react";
import {
  Button,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  LinearProgress,
  Stack,
  Typography,
} from "@mui/material";
import * as pdfjs from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

'''
    + export_top_level(slice_lines(3577, 3647))
    + "\n\n"
    + export_top_level(slice_lines(3649, 3890))
    + "\n",
)

write(
    "pages/personnel/PersonnelVirtualList.tsx",
    '''import { useEffect, useRef } from "react";
import { useVirtualizer } from "@tanstack/react-virtual";
import type { PersonnelRecord } from "./personnelUtils";

'''
    + export_top_level(slice_lines(3892, 4010))
    + "\n",
)

write(
    "pages/personnel/PersonnelPage.tsx",
    '''import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  Divider,
  IconButton,
  Stack,
  Typography,
} from "@mui/material";
import AddPhotoAlternateOutlinedIcon from "@mui/icons-material/AddPhotoAlternateOutlined";
import DeleteOutlineOutlinedIcon from "@mui/icons-material/DeleteOutlineOutlined";
import DescriptionOutlinedIcon from "@mui/icons-material/DescriptionOutlined";
import GroupsOutlinedIcon from "@mui/icons-material/GroupsOutlined";
import MilitaryTechOutlinedIcon from "@mui/icons-material/MilitaryTechOutlined";
import PictureAsPdfOutlinedIcon from "@mui/icons-material/PictureAsPdfOutlined";
import SearchOutlinedIcon from "@mui/icons-material/SearchOutlined";
import {
  api,
  type BackendEjournalImport,
  type BackendPersonQuestionnaire,
} from "../../api";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { PhotoCropDialog } from "./PhotoCropDialog";
import { PersonnelVirtualList } from "./PersonnelVirtualList";
import {
  buildPersonSummary,
  createDefaultActionForm,
  dataUrlToObjectUrl,
  findEjournalPersonnelSheet,
  getRowKeyByKeyPart,
  isLikelyPersonnelRow,
  loadAllEjournalSheetRows,
  personActions,
  previewValueToDisplay,
  readFileAsDataUrl,
  type PersonAction,
  type PersonActionForm,
  type PersonnelRecord,
} from "./personnelUtils";

'''
    + export_top_level(slice_lines(4012, 4898))
    + "\n",
)

write(
    "pages/documents/DocumentsPage.tsx",
    '''import { useEffect, useMemo, useState } from "react";
import { Button, Typography } from "@mui/material";
import DescriptionOutlinedIcon from "@mui/icons-material/DescriptionOutlined";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { buildPersonSummary } from "../personnel/personnelUtils";

'''
    + export_top_level(slice_lines(4900, 5016))
    + "\n",
)

write(
    "pages/ejournal/EjournalPage.tsx",
    '''import { useEffect, useMemo, useState } from "react";
import {
  Alert,
  Button,
  MenuItem,
  Select,
  Stack,
  Typography,
} from "@mui/material";
import CloudUploadOutlinedIcon from "@mui/icons-material/CloudUploadOutlined";
import type { MRT_ColumnDef } from "material-react-table";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "material-react-table";
import { api, type BackendEjournalImport } from "../../api";
import {
  type ExcelRow,
  type ExcelWorkbookSnapshot,
  createWorkbookDebugPayload,
  hasRowData,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import { cellValueToJson } from "../../shared/format";
import type { DbPreviewState, EjournalPreviewRow } from "./ejournalTypes";
import {
  buildImportColumns,
  localRowsToPreviewRows,
  parseDbColumns,
  previewValueToDisplay,
} from "./ejournalUtils";
import {
  buildPersonSummary,
  createDefaultActionForm,
  personActions,
  type PersonAction,
  type PersonActionForm,
} from "../personnel/personnelUtils";

'''
    + export_top_level(slice_lines(5018, 5789))
    + "\n",
)

# ---------------------------------------------------------------------------
# BCHS: types / calc / export
# ---------------------------------------------------------------------------
helpers = slice_lines(5791, 8465)
type_names = [
    "BchsAnalyticsRow",
    "BchsComparisonRow",
    "BchsAnalyticsTableColumn",
    "BchsAnalyticsTableRow",
    "BchsSupplementRow",
    "BchsSupplementSnapshot",
    "BchsAnalyticsSnapshot",
    "BchsPersonnelAwayPerson",
    "BchsUnitAwayStats",
    "BchsUnitAttachedStats",
    "GeneratedWorkbook",
    "GeneratedSheet",
]

remaining = helpers
type_blocks: list[str] = []
for name in type_names:
    # Match type Name = { ... };  (brace-balanced)
    m = re.search(rf"type {name} = ", remaining)
    if not m:
        print("WARN: type not found", name)
        continue
    start = m.start()
    # Find '=' then parse (supports `any`, `{...}`, and `Foo & {...}`)
    i = remaining.find("=", start) + 1
    while i < len(remaining) and remaining[i] in " \t\n":
        i += 1
    if remaining.startswith("any", i):
        end = remaining.find(";", i) + 1
        block = remaining[start:end]
    else:
        brace = remaining.find("{", i)
        if brace < 0:
            end = remaining.find(";", i) + 1
            block = remaining[start:end]
        else:
            depth = 0
            j = brace
            while j < len(remaining):
                ch = remaining[j]
                if ch == "{":
                    depth += 1
                elif ch == "}":
                    depth -= 1
                    if depth == 0:
                        j += 1
                        break
                j += 1
            while j < len(remaining) and remaining[j] in " \t":
                j += 1
            if j < len(remaining) and remaining[j] == ";":
                j += 1
            block = remaining[start:j]
    type_blocks.append("export " + block.strip())
    remaining = remaining[:start] + "\n" + remaining[start + len(block) :]

write(
    "pages/bchs/bchsTypes.ts",
    '''import type { CellValue } from "../../excelRoundTrip";
import type { AnalyticsMetric } from "../analytics/analyticsData";

'''
    + "\n\n".join(type_blocks)
    + "\n",
)

set_idx = remaining.index("const setWorkbookCellValue")
parse_idx = remaining.index("const parseBchsImportAnalytics")
calc_part = remaining[:set_idx]
export_part = remaining[set_idx:parse_idx]
parse_part = remaining[parse_idx:]

# Clean excessive blank lines in calc
calc_part = re.sub(r"\n{3,}", "\n\n", calc_part).strip() + "\n"
export_part = re.sub(r"\n{3,}", "\n\n", export_part).strip() + "\n"
parse_part = re.sub(r"\n{3,}", "\n\n", parse_part).strip() + "\n"

write(
    "pages/bchs/bchsCalc.ts",
    '''import type { BackendEjournalImport } from "../../api";
import type {
  CellValue,
  ExcelRow,
  ExcelWorkbookSnapshot,
} from "../../excelRoundTrip";
import { getColumnLabel, valueToDisplay } from "../../excelRoundTrip";
import type { AnalyticsMetric } from "../analytics/analyticsData";
import type {
  BchsAnalyticsRow,
  BchsAnalyticsSnapshot,
  BchsAnalyticsTableColumn,
  BchsAnalyticsTableRow,
  BchsComparisonRow,
  BchsPersonnelAwayPerson,
  BchsSupplementRow,
  BchsSupplementSnapshot,
  BchsUnitAttachedStats,
  BchsUnitAwayStats,
} from "./bchsTypes";

'''
    + export_top_level(calc_part)
    + "\n"
    + export_top_level(parse_part)
    + "\n",
)

write(
    "pages/bchs/bchsExport.ts",
    '''import type { CellValue, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { getColumnLabel, valueToDisplay } from "../../excelRoundTrip";
import {
  BCHS_PERCENT_COLUMNS,
  bchsToNumber,
  columnLetterToIndex,
  emptyBchsSupplementRow,
} from "./bchsCalc";
import type {
  BchsAnalyticsSnapshot,
  BchsAnalyticsTableRow,
  BchsComparisonRow,
  BchsSupplementRow,
  GeneratedSheet,
  GeneratedWorkbook,
} from "./bchsTypes";

'''
    + export_top_level(export_part)
    + "\n",
)

# BchsPage - scan icons in body
write(
    "pages/bchs/BchsPage.tsx",
    '''import { useEffect, useMemo, useState } from "react";
import { Canvas } from "@react-three/fiber";
import {
  Alert,
  Box,
  Button,
  Chip,
  MenuItem,
  Select,
  Stack,
  Switch,
  Typography,
} from "@mui/material";
import CloudUploadOutlinedIcon from "@mui/icons-material/CloudUploadOutlined";
import FileDownloadOutlinedIcon from "@mui/icons-material/FileDownloadOutlined";
import type { MRT_ColumnDef } from "material-react-table";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "material-react-table";
import { api, type BackendEjournalImport } from "../../api";
import {
  type ExcelRow,
  type ExcelWorkbookSnapshot,
  createWorkbookDebugPayload,
  exportBlankWorkbookWithMutations,
  exportWorkbookWithMutations,
  hasRowData,
  readWorkbookSnapshot,
} from "../../excelRoundTrip";
import {
  getInitialBchsAnalyticsTab,
  type BchsAnalyticsTab,
} from "../../app/navigation";
import { cellValueToJson } from "../../shared/format";
import { BarList, BchsAbsenceDonut } from "../analytics/charts";
import { CombatStructureScene } from "../analytics/combatScene";
import {
  buildImportColumns,
  previewValueToDisplay,
} from "../ejournal/ejournalUtils";
import {
  buildBchsAnalytics,
  buildBchsAnalyticsFromWorkbook,
  createBchsComparisonRow,
  parseBchsImportAnalytics,
} from "./bchsCalc";
import {
  materializeBchsAnalyticsTableRow,
  materializeBchsSheetFormulas,
  writeBchsAppendixRow,
  writeBchsLegacyRow,
  writeBchsPersonnelBzvpRow,
  writeGeneratedBchsWorkbook,
} from "./bchsExport";
import type { BchsComparisonRow } from "./bchsTypes";

'''
    + export_top_level(slice_lines(8466, 10488))
    + "\n",
)

# ---------------------------------------------------------------------------
# shared/types.ts
# ---------------------------------------------------------------------------
write(
    "shared/types.ts",
    '''/** Cross-page type re-exports. */
export type { AppPage, BchsAnalyticsTab } from "../app/navigation";
export type { DataSourceFile } from "../pages/import/types";
export type { AnalyticsMetric } from "../pages/analytics/analyticsData";
export type {
  DbPreviewState,
  EjournalColumn,
  EjournalPreviewRow,
} from "../pages/ejournal/ejournalTypes";
export type {
  PersonAction,
  PersonActionForm,
  PersonnelRecord,
} from "../pages/personnel/personnelUtils";
''',
)

# ---------------------------------------------------------------------------
# thin App.tsx
# ---------------------------------------------------------------------------
app_fn = export_top_level(slice_lines(10490, 10780))
# Keep default export; App itself should not be named export only
app_fn = app_fn.replace("export function App(", "function App(")

write(
    "App.tsx",
    '''import {
  useEffect,
  useState,
  type Dispatch,
  type SetStateAction,
} from "react";
import { Button, Typography } from "@mui/material";
import BuildOutlinedIcon from "@mui/icons-material/BuildOutlined";
import FormatListBulletedOutlinedIcon from "@mui/icons-material/FormatListBulletedOutlined";
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
import { getPageFromPath, navigateToPage, type AppPage } from "./app/navigation";
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
import { BchsPage } from "./pages/bchs/BchsPage";

'''
    + app_fn
    + "\n",
)

print("App.tsx lines:", len((ROOT / "App.tsx").read_text().splitlines()))
print("DONE")
