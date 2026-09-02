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
import { openPersonnelInNewTab } from "../../app/navigation";
import { resolveOverviewPhoto } from "./overviewPhotos";
import { overviewPersonMatchKey } from "./overviewPersonnelAssets";

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
    setPreview({ src: photo, top, left });
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
  documentsByExternalId,
  onOpenQuestionnaire,
  onNeedPhoto,
  onExport,
  emptyMessage = "Немає записів за поточними фільтрами.",
}: {
  rows: BackendPersonnelOverviewRow[];
  photos: Record<string, string>;
  questionnaireByExternalId?: Record<string, true>;
  documentsByExternalId?: Record<string, OverviewPersonDocumentSummary>;
  onOpenQuestionnaire?: (target: OverviewQuestionnaireTarget) => void;
  onNeedPhoto?: (row: BackendPersonnelOverviewRow) => void;
  onExport?: (
    context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  ) => void | Promise<void>;
  emptyMessage?: string;
}) {
  const openPerson = (row: BackendPersonnelOverviewRow) => {
    openPersonnelInNewTab({
      rowId: row.id,
      externalId: row.externalId,
    });
  };

  const openQuestionnaire = (row: BackendPersonnelOverviewRow) => {
    if (!row.externalId) return;
    const hasQuestionnaire = Boolean(
      questionnaireByExternalId?.[row.externalId] ||
        questionnaireByExternalId?.[overviewPersonMatchKey(row.name)],
    );
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
    openPersonnelInNewTab({ rowId: row.id, externalId: row.externalId });
  };

  const columns = useMemo<Array<MRT_ColumnDef<BackendPersonnelOverviewRow>>>(
    () => [
      {
        id: "person",
        header: "ПІБ",
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
              photos={photos}
              onNeedPhoto={onNeedPhoto}
            />
            <span>
              <strong>{row.original.name}</strong>
            </span>
          </button>
        ),
      },
      {
        accessorKey: "unit",
        header: "Підрозділ",
        size: 240,
      },
      {
        accessorKey: "rank",
        header: "Звання",
        size: 180,
      },
      {
        accessorKey: "positionTitle",
        header: "Посада",
        size: 280,
        accessorFn: (row) => row.positionTitle?.trim() || "—",
        exportValue: (row) => row.positionTitle?.trim() || "",
      },
      {
        id: "status",
        header: "Поточний статус",
        size: 220,
        accessorFn: (row) => row.statusLabel,
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
        header: "Анкета",
        size: 130,
        accessorFn: (row) =>
          questionnaireByExternalId?.[row.externalId] ||
          questionnaireByExternalId?.[overviewPersonMatchKey(row.name)]
            ? "Є"
            : "Немає",
        Cell: ({ row }) => {
          const hasQuestionnaire = Boolean(
            questionnaireByExternalId?.[row.original.externalId] ||
              questionnaireByExternalId?.[overviewPersonMatchKey(row.original.name)],
          );
          const label = hasQuestionnaire ? "Є" : "Немає";
          return (
            <button
              type="button"
              className="overview-chip-button overview-questionnaire-hit"
              aria-label={
                hasQuestionnaire
                  ? `Відкрити анкету: ${row.original.name}`
                  : `Анкети немає — перейти до картки: ${row.original.name}`
              }
              disabled={!row.original.externalId}
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
          overviewDocumentPresence(row, documentsByExternalId),
        exportValue: (row) =>
          overviewDocumentChipLabel(row, documentsByExternalId),
        Cell: ({ row }) => {
          const summary =
            (row.original.externalId &&
              documentsByExternalId?.[row.original.externalId]) ||
            null;
          const hasDocuments = Boolean(summary?.count);
          return (
            <span title={summary?.labels.join(" · ") || "Немає створених документів"}>
              <Chip
                className={`overview-status-chip ${
                  hasDocuments ? "tone-ok" : "tone-other"
                }`}
                label={overviewDocumentChipLabel(
                  row.original,
                  documentsByExternalId,
                )}
                size="small"
                variant="outlined"
              />
            </span>
          );
        },
      },
      {
        accessorKey: "fighterDirection",
        header: "Напрямок",
        size: 190,
      },
      {
        accessorKey: "fighterExitDate",
        header: "Дата виходу",
        size: 150,
      },
      {
        accessorKey: "fighterReturnDate",
        header: "Дата повернення",
        size: 160,
      },
      {
        accessorKey: "fighterTotalDays",
        header: "Днів",
        size: 110,
        filterVariant: "number-range",
      },
      {
        accessorKey: "fighterStatus",
        header: "Статус бійців",
        size: 150,
      },
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
                openPersonnelInNewTab({
                  rowId: row.original.id,
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
    [
      documentsByExternalId,
      onNeedPhoto,
      onOpenQuestionnaire,
      photos,
      questionnaireByExternalId,
    ],
  );

  const table = useMaterialReactTable({
    columns,
    data: rows,
    emptyMessage,
    exportLabel: "Експорт",
    copyLabel: "Копіювати",
    enableCopyText: true,
    enableGlobalFilter: false,
    getRowId: (row) => row.id,
    getTdProps: ({ columnId }) =>
      columnId === "questionnaire"
        ? { className: "overview-questionnaire-cell" }
        : undefined,
    onExport,
    initialState: {
      pagination: {
        pageSize: 1000,
      },
      columnPinning: {
        left: ["person"],
        right: ["actions"],
      },
    },
  });

  return <MaterialReactTable table={table} />;
}
