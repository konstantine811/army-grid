import { useEffect, useRef } from "react";
import { Chip, IconButton, Typography } from "@/components/sci/SciPrimitives";
import { InfoOutlinedIcon } from "@/components/sci/icons";
import { MoreHorizOutlinedIcon } from "@/components/sci/icons";
import { PersonOutlinedIcon } from "@/components/sci/icons";
import { useVirtualizer } from "@tanstack/react-virtual";
import type { BackendPersonnelOverviewRow } from "../../api";

const ROW_HEIGHT = 64;

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

export function OverviewVirtualTable({
  rows,
  photos,
  onOpenPersonnel,
}: {
  rows: BackendPersonnelOverviewRow[];
  photos: Record<string, string>;
  onOpenPersonnel?: (target: OverviewPersonTarget) => void;
}) {
  const openPerson = (row: BackendPersonnelOverviewRow) => {
    onOpenPersonnel?.({
      rowId: row.id,
      externalId: row.externalId,
    });
  };
  const parentRef = useRef<HTMLDivElement | null>(null);
  const rowVirtualizer = useVirtualizer({
    count: rows.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => ROW_HEIGHT,
    overscan: 16,
  });

  useEffect(() => {
    parentRef.current?.scrollTo({ top: 0 });
  }, [rows]);

  return (
    <div className="overview-table-wrap">
      <div className="overview-table-header" role="row">
        <span>ПІБ</span>
        <span>Підрозділ</span>
        <span>Звання</span>
        <span>Поточний статус</span>
        <span>Від</span>
        <span>Днів</span>
        <span>Оновлено</span>
        <span />
      </div>

      <div className="overview-table-body" ref={parentRef}>
        {rows.length === 0 ? (
          <Typography variant="body2" color="text.secondary" sx={{ p: 2 }}>
            Немає записів за поточними фільтрами.
          </Typography>
        ) : (
          <div
            className="overview-table-virtual"
            style={{ height: rowVirtualizer.getTotalSize() }}
          >
            {rowVirtualizer.getVirtualItems().map((virtualRow) => {
              const row = rows[virtualRow.index];
              if (!row) return null;

              return (
                <div
                  className="overview-table-row"
                  key={row.id}
                  role="row"
                  style={{
                    height: virtualRow.size,
                    transform: `translateY(${virtualRow.start}px)`,
                  }}
                >
                  <div>
                    <button
                      type="button"
                      className="overview-person-cell"
                      onClick={() => openPerson(row)}
                    >
                      <span className="overview-avatar" aria-hidden>
                        {row.externalId && photos[row.externalId] ? (
                          <img alt="" src={photos[row.externalId]} />
                        ) : (
                          <PersonOutlinedIcon fontSize="small" />
                        )}
                      </span>
                      <span>
                        <strong>{row.name}</strong>
                        <small>{row.externalId || "без ID"}</small>
                      </span>
                    </button>
                  </div>
                  <div>{row.unit || "—"}</div>
                  <div>{row.rank || "—"}</div>
                  <div>
                    <Chip
                      className={`overview-status-chip tone-${overviewStatusTone(row.status)}`}
                      label={row.statusLabel}
                      size="small"
                      variant="outlined"
                    />
                  </div>
                  <div>{row.validFrom || "—"}</div>
                  <div>{row.days ?? "—"}</div>
                  <div>{row.updatedAt}</div>
                  <div className="overview-table-actions">
                    <IconButton
                      size="small"
                      aria-label="Деталі"
                      onClick={() => openPerson(row)}
                    >
                      <InfoOutlinedIcon fontSize="small" />
                    </IconButton>
                    <IconButton size="small" aria-label="Меню">
                      <MoreHorizOutlinedIcon fontSize="small" />
                    </IconButton>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}
