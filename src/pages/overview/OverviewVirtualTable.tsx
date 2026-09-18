import { useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import { Chip, IconButton } from "@/components/sci/SciPrimitives";
import { InfoOutlinedIcon, MoreHorizOutlinedIcon, PersonOutlinedIcon } from "@/components/sci/icons";
import {
  MaterialReactTable,
  type MRT_ColumnDef,
  type SciDataTableExportContext,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import type { BackendPersonnelOverviewRow } from "../../api";
import { openPersonnelFromOverview } from "../../app/navigation";
import { OVERVIEW_STAFF_COLUMN_HEADERS } from "./overviewStaffColumns";
import {
  buildOverviewStaffSheetColumnDefs,
  DEFAULT_OVERVIEW_STAFF_COLUMN_VISIBILITY,
} from "./overviewStaffSheetColumns";
import { resolveOverviewPhoto, overviewPhotoPreviewUrl } from "./overviewPhotos";
import { overviewPersonMatchKey } from "./overviewPersonnelAssets";
import { overviewStatusFilterLabel } from "./overviewRosterMerge";

export type OverviewPersonDocumentSummary = {
  count: number;
  labels: string[];
};

const overviewDocumentPresence = (
  row: BackendPersonnelOverviewRow,
  documentsByExternalId?: Record<string, OverviewPersonDocumentSummary>,
) => {
  const count =
    (row.externalId && documentsByExternalId?.[row.externalId]?.count) ||
    documentsByExternalId?.[overviewPersonMatchKey(row.name)]?.count ||
    0;
  return count > 0 ? "Є" : "Немає";
};

const overviewDocumentChipLabel = (
  row: BackendPersonnelOverviewRow,
  documentsByExternalId?: Record<string, OverviewPersonDocumentSummary>,
) => {
  const summary =
    (row.externalId && documentsByExternalId?.[row.externalId]) ||
    documentsByExternalId?.[overviewPersonMatchKey(row.name)] ||
    null;
  if (!summary?.count) return "Немає";
  return summary.count > 1 ? `Є · ${summary.count}` : "Є";
};

export const overviewStatusTone = (status: string) => {
  switch (status) {
    case "ON_DUTY":
      return "ok";
    case "BUSINESS_TRIP":
      return "trip";
    case "LEAVE":
      return "leave";
    case "MEDICAL":
      return "medical";
    case "AWOL":
    case "MISSING":
    case "CAPTIVITY":
      return "danger";
    default:
      return "other";
  }
};

export type OverviewPersonTarget = {
  rowId: string;
  externalId: string;
};

export type OverviewQuestionnaireTarget = {
  externalId: string;
  name: string;
  rowId: string;
  hasQuestionnaire: boolean;
};

const PHOTO_HOVER_SIZE = 220;

function OverviewPersonPhoto({
  row,
  photos,
  onNeedPhoto,
}: {
  row: BackendPersonnelOverviewRow;
  photos: Record<string, string>;
  onNeedPhoto?: (row: BackendPersonnelOverviewRow) => void;
}) {
  const photo = resolveOverviewPhoto(row, photos);
  const thumbRef = useRef<HTMLSpanElement>(null);
  const [preview, setPreview] = useState<{
    src: string;
    top: number;
    left: number;
  } | null>(null);

  useEffect(() => {
    if (!photo) onNeedPhoto?.(row);
  }, [onNeedPhoto, photo, row]);

  useEffect(() => {
    if (!preview) return;
    const hide = () => setPreview(null);
    window.addEventListener("scroll", hide, true);
    return () => window.removeEventListener("scroll", hide, true);
  }, [preview]);

  const showPreview = () => {
    if (!photo || !thumbRef.current) return;
    const rect = thumbRef.current.getBoundingClientRect();
    const left = Math.min(
      rect.right + 10,
      window.innerWidth - PHOTO_HOVER_SIZE - 12,
    );
    const top = Math.min(
      Math.max(12, rect.top + rect.height / 2 - PHOTO_HOVER_SIZE / 2),
      window.innerHeight - PHOTO_HOVER_SIZE - 12,
    );
    setPreview({ src: overviewPhotoPreviewUrl(photo), top, left });
  };

  return (
    <>
      <span
        ref={thumbRef}
        className="overview-avatar"
        aria-hidden
        onMouseEnter={showPreview}
        onMouseLeave={() => setPreview(null)}
      >
        {photo ? <img alt="" src={photo} /> : <PersonOutlinedIcon fontSize="small" />}
      </span>
      {preview
        ? createPortal(
            <div
              className="overview-photo-hover"
              style={{ top: preview.top, left: preview.left }}
            >
              <img alt="" src={preview.src} />
            </div>,
            document.body,
          )
        : null}
    </>
  );
}

export function OverviewVirtualTable({
  rows,
  photos,
  questionnaireByExternalId,
  questionnaireLoading = false,
  questionnairePresenceStatus = "ready",
  documentsByExternalId,
  onOpenQuestionnaire,
  onNeedPhoto,
  onVisibleRowsChange,
  onExport,
  onImportantExport,
  onRotaGudzExport,
  onRotaBchsMorningExport,
  onPpdLocationExport,
  onCommandersExport,
  copyTextBuilder,
  rotaCopyTextBuilder,
  locationCopyTextBuilder,
  emptyMessage = "Немає записів за поточними фільтрами.",
  onColumnVisibilityChange,
  onColumnFiltersChange,
}: {
  rows: BackendPersonnelOverviewRow[];
  photos: Record<string, string>;
  questionnaireByExternalId?: Record<string, true>;
  questionnaireLoading?: boolean;
  documentsByExternalId?: Record<string, OverviewPersonDocumentSummary>;
  onOpenQuestionnaire?: (target: OverviewQuestionnaireTarget) => void;
  onNeedPhoto?: (row: BackendPersonnelOverviewRow) => void;
  onVisibleRowsChange?: (rows: BackendPersonnelOverviewRow[]) => void;
  onExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  onImportantExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  onRotaGudzExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  onRotaBchsMorningExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  onPpdLocationExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  onCommandersExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  copyTextBuilder?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => string | Promise<string>;
  rotaCopyTextBuilder?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => string | Promise<string>;
  locationCopyTextBuilder?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => string | Promise<string>;
  emptyMessage?: string;
  questionnairePresenceStatus?: "idle" | "loading" | "ready";
  onColumnVisibilityChange?: (visibility: Record<string, boolean>) => void;
  onColumnFiltersChange?: (columnFilters: Record<string, string[]>) => void;
}) {
  const photosRef = useRef(photos);
  photosRef.current = photos;

  const questionnaireByExternalIdRef = useRef(questionnaireByExternalId);
  questionnaireByExternalIdRef.current = questionnaireByExternalId;
  const documentsByExternalIdRef = useRef(documentsByExternalId);
  documentsByExternalIdRef.current = documentsByExternalId;
  const questionnaireLoadingRef = useRef(questionnaireLoading);
  questionnaireLoadingRef.current = questionnaireLoading;
  const questionnairePresenceStatusRef = useRef(questionnairePresenceStatus);
  questionnairePresenceStatusRef.current = questionnairePresenceStatus;

  const rowHasQuestionnaire = (row: BackendPersonnelOverviewRow) => {
    const map = questionnaireByExternalIdRef.current;
    return Boolean(
      map?.[row.externalId] ||
        map?.[overviewPersonMatchKey(row.name)],
    );
  };

  const questionnaireCellLabel = (row: BackendPersonnelOverviewRow) =>
    rowHasQuestionnaire(row)
      ? "Є"
      : questionnaireLoadingRef.current &&
          questionnairePresenceStatusRef.current !== "idle"
        ? "Завантаження…"
        : questionnairePresenceStatusRef.current === "idle"
          ? "—"
          : "Немає";

  const openPerson = (row: BackendPersonnelOverviewRow) => {
    openPersonnelFromOverview({ externalId: row.externalId });
  };

  const openQuestionnaire = (row: BackendPersonnelOverviewRow) => {
    if (!row.externalId) return;
    const hasQuestionnaire = rowHasQuestionnaire(row);
    if (onOpenQuestionnaire) {
      onOpenQuestionnaire({
        externalId: row.externalId,
        name: row.name,
        rowId: row.id,
        hasQuestionnaire,
      });
      return;
    }
    if (hasQuestionnaire) return;
    openPersonnelFromOverview({ externalId: row.externalId });
  };

  const rowNumberById = useMemo(
    () => new Map(rows.map((row, index) => [row.id, index + 1])),
    [rows],
  );

  const columns = useMemo<Array<MRT_ColumnDef<BackendPersonnelOverviewRow>>>(
    () => [
      {
        id: "person",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.name,
        size: 360,
        pin: "left",
        accessorFn: (row) => row.name,
        exportValue: (row) => row.name,
        Cell: ({ row }) => (
          <button
            type="button"
            className="overview-person-cell"
            onClick={() => openPerson(row.original)}
          >
            <OverviewPersonPhoto
              row={row.original}
              photos={photosRef.current}
              onNeedPhoto={onNeedPhoto}
            />
            <span>
              <strong>{row.original.name}</strong>
            </span>
          </button>
        ),
      },
      {
        id: "rowNumber",
        header: "№",
        size: 80,
        enableColumnFilter: false,
        accessorFn: (row) => String(rowNumberById.get(row.id) ?? ""),
      },
      {
        accessorKey: "unit",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.unit,
        size: 240,
      },
      {
        accessorKey: "rank",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.rank,
        size: 180,
      },
      {
        accessorKey: "positionTitle",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.positionTitle,
        size: 280,
        accessorFn: (row) => row.positionTitle?.trim() || "—",
        exportValue: (row) => row.positionTitle?.trim() || "",
      },
      {
        id: "status",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.status,
        size: 220,
        accessorFn: overviewStatusFilterLabel,
        exportValue: (row) => row.statusLabel,
        Cell: ({ row }) => (
          <Chip
            className={`overview-status-chip tone-${overviewStatusTone(row.original.status)}`}
            label={row.original.statusLabel}
            size="small"
            variant="outlined"
          />
        ),
      },
      {
        id: "questionnaire",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.questionnaire,
        size: 130,
        accessorFn: (row) => questionnaireCellLabel(row),
        Cell: ({ row }) => {
          const hasQuestionnaire = rowHasQuestionnaire(row.original);
          const label = questionnaireCellLabel(row.original);
          const loading = questionnaireLoadingRef.current;
          return (
            <button
              type="button"
              className="overview-chip-button overview-questionnaire-hit"
              aria-label={
                hasQuestionnaire
                  ? `Відкрити анкету: ${row.original.name}`
                  : loading
                    ? `Перевіряю анкету: ${row.original.name}`
                    : `Анкети немає — перейти до картки: ${row.original.name}`
              }
              disabled={
                !row.original.externalId ||
                (!hasQuestionnaire &&
                  questionnaireLoadingRef.current &&
                  questionnairePresenceStatusRef.current !== "idle")
              }
              onClick={() => openQuestionnaire(row.original)}
            >
              <Chip
                className={`overview-status-chip ${
                  hasQuestionnaire ? "tone-ok" : "tone-other"
                }`}
                label={label}
                size="small"
                variant="outlined"
                component="span"
              />
            </button>
          );
        },
      },
      {
        id: "documents",
        header: "Документи",
        size: 150,
        accessorFn: (row) =>
          overviewDocumentPresence(row, documentsByExternalIdRef.current),
        exportValue: (row) =>
          overviewDocumentChipLabel(row, documentsByExternalIdRef.current),
        Cell: ({ row }) => {
          const documents = documentsByExternalIdRef.current;
          const summary =
            (row.original.externalId &&
              documents?.[row.original.externalId]) ||
            documents?.[overviewPersonMatchKey(row.original.name)] ||
            null;
          const hasDocuments = Boolean(summary?.count);
          return (
            <span title={summary?.labels.join(" · ") || "Немає створених документів"}>
              <Chip
                className={`overview-status-chip ${
                  hasDocuments ? "tone-ok" : "tone-other"
                }`}
                label={overviewDocumentChipLabel(row.original, documents)}
                size="small"
                variant="outlined"
              />
            </span>
          );
        },
      },
      {
        accessorKey: "fighterDirection",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.fighterDirection,
        size: 190,
      },
      {
        accessorKey: "fighterExitDate",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.fighterExitDate,
        size: 150,
      },
      {
        accessorKey: "fighterReturnDate",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.fighterReturnDate,
        size: 220,
      },
      {
        accessorKey: "fighterTotalDays",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.fighterTotalDays,
        size: 110,
        filterVariant: "number-range",
      },
      {
        accessorKey: "fighterStatus",
        header: OVERVIEW_STAFF_COLUMN_HEADERS.fighterStatus,
        size: 150,
      },
      ...buildOverviewStaffSheetColumnDefs(),
      {
        accessorKey: "updatedAt",
        header: "Оновлено",
        size: 190,
      },
      {
        id: "actions",
        header: "",
        columnMenuLabel: "Дії",
        size: 90,
        pin: "right",
        enableColumnFilter: false,
        enableGlobalFilter: false,
        enableHiding: false,
        Cell: ({ row }) => (
          <div className="overview-table-actions">
            <IconButton
              size="small"
              aria-label="Деталі"
              onClick={() =>
                openPersonnelFromOverview({
                  externalId: row.original.externalId,
                })
              }
            >
              <InfoOutlinedIcon fontSize="small" />
            </IconButton>
            <IconButton size="small" aria-label="Меню">
              <MoreHorizOutlinedIcon fontSize="small" />
            </IconButton>
          </div>
        ),
      },
    ],
    [onNeedPhoto, onOpenQuestionnaire, rowNumberById],
  );

  const table = useMaterialReactTable({
    columns,
    data: rows,
    emptyMessage,
    exportLabel: "Експорт",
    copyLabel: "Копіювати",
    enableCopyText: true,
    copyTextBuilder,
    secondaryCopyLabel: "Копіювати Роту",
    secondaryCopyTextBuilder: rotaCopyTextBuilder,
    tertiaryCopyLabel: "Копіювати місця",
    tertiaryCopyTextBuilder: locationCopyTextBuilder,
    enableGlobalFilter: false,
    enableRowVirtualization: true,
    estimatedRowHeight: 44,
    getRowId: (row) => row.id,
    getTdProps: ({ columnId }) =>
      columnId === "questionnaire"
        ? { className: "overview-questionnaire-cell" }
        : undefined,
    onExport,
    onSecondaryExport: onImportantExport,
    secondaryExportLabel: "Експорт важливих колонок",
    onTertiaryExport: onRotaGudzExport,
    tertiaryExportLabel: "Звіт роти (ГУД)",
    onQuaternaryExport: onRotaBchsMorningExport,
    quaternaryExportLabel: "БЧС (ранковий ПБ)",
    onQuinaryExport: onPpdLocationExport,
    quinaryExportLabel: "ППД / Полігон",
    onSenaryExport: onCommandersExport,
    senaryExportLabel: "Командири",
    onColumnVisibilityChange,
    onColumnFiltersChange,
    onVisibleRowsChange,
    initialState: {
      pagination: {
        pageSize: 1000,
      },
      columnVisibility: DEFAULT_OVERVIEW_STAFF_COLUMN_VISIBILITY,
      columnPinning: {
        left: ["person"],
        right: ["actions"],
      },
    },
  });

  return <MaterialReactTable table={table} />;
}
