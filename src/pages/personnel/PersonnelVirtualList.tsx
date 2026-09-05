import { useEffect, useRef } from "react";
import { PersonSearchOutlinedIcon } from "@/components/sci/icons";
import { useVirtualizer } from "@tanstack/react-virtual";
import type { PersonnelRecord } from "./personnelUtils";

export function PersonnelVirtualList({
  items,
  selectedRowId,
  photoByExternalId,
  onNeedPhotos,
  onSelect,
  keyboardEnabled = true,
}: {
  items: PersonnelRecord[];
  selectedRowId: string;
  photoByExternalId: Record<string, string>;
  onNeedPhotos?: (externalIds: string[]) => void;
  onSelect: (rowId: string) => void;
  keyboardEnabled?: boolean;
}) {
  const parentRef = useRef<HTMLDivElement | null>(null);
  const lastScrolledIdRef = useRef("");
  const rowVirtualizer = useVirtualizer({
    count: items.length,
    getScrollElement: () => parentRef.current,
    estimateSize: () => 66,
    overscan: 12,
    gap: 8,
  });
  const virtualItems = rowVirtualizer.getVirtualItems();
  const visiblePhotoIds = virtualItems
    .map((item) => items[item.index]?.summary.externalId ?? "")
    .filter(Boolean)
    .join("\u0000");

  useEffect(() => {
    if (!visiblePhotoIds) return;
    onNeedPhotos?.(visiblePhotoIds.split("\u0000"));
  }, [onNeedPhotos, visiblePhotoIds]);

  useEffect(() => {
    if (!selectedRowId || items.length === 0) return;

    const index = items.findIndex(
      (item) => item.row.__dbRowId === selectedRowId,
    );
    if (index < 0) return;

    // Only auto-scroll when selection changes (e.g. from Overview), not on every re-render.
    if (lastScrolledIdRef.current === selectedRowId) return;
    lastScrolledIdRef.current = selectedRowId;

    let innerFrame = 0;
    const timeout = window.setTimeout(() => {
      innerFrame = window.requestAnimationFrame(() => {
        rowVirtualizer.scrollToIndex(index, { align: "center" });
      });
    }, 50);

    return () => {
      window.clearTimeout(timeout);
      window.cancelAnimationFrame(innerFrame);
    };
  }, [items, rowVirtualizer, selectedRowId]);

  useEffect(() => {
    if (!keyboardEnabled || items.length === 0) return;

    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key !== "ArrowDown" && event.key !== "ArrowUp") return;
      if (event.metaKey || event.ctrlKey || event.altKey) return;

      const target = event.target as HTMLElement | null;
      if (
        target &&
        (target.closest("input, textarea, select, [contenteditable='true']") ||
          target.closest('[role="dialog"]'))
      ) {
        return;
      }

      const currentIndex = items.findIndex(
        (item) => item.row.__dbRowId === selectedRowId,
      );
      const nextIndex =
        event.key === "ArrowDown"
          ? Math.min(
              items.length - 1,
              currentIndex < 0 ? 0 : currentIndex + 1,
            )
          : Math.max(0, currentIndex < 0 ? 0 : currentIndex - 1);

      const nextId = items[nextIndex]?.row.__dbRowId;
      if (!nextId || nextId === selectedRowId) return;

      event.preventDefault();
      onSelect(nextId);
      rowVirtualizer.scrollToIndex(nextIndex, { align: "auto" });
    };

    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [
    items,
    keyboardEnabled,
    onSelect,
    rowVirtualizer,
    selectedRowId,
  ]);

  return (
    <div className="personnel-list" ref={parentRef}>
      <div
        className="personnel-list-virtual"
        style={{ height: `${rowVirtualizer.getTotalSize()}px` }}
      >
        {virtualItems.map((virtualRow) => {
          const record = items[virtualRow.index];
          if (!record) return null;

          const photo =
            (record.summary.externalId &&
              photoByExternalId[record.summary.externalId]) ||
            "";

          return (
            <button
              className={
                record.row.__dbRowId === selectedRowId ? "active" : ""
              }
              data-index={virtualRow.index}
              key={record.row.__dbRowId ?? virtualRow.key}
              ref={rowVirtualizer.measureElement}
              type="button"
              style={{ transform: `translateY(${virtualRow.start}px)` }}
              onClick={() => onSelect(record.row.__dbRowId ?? "")}
            >
              <span className="personnel-list-thumb" aria-hidden>
                {photo ? (
                  <img alt="" src={photo} />
                ) : (
                  <PersonSearchOutlinedIcon fontSize="small" />
                )}
              </span>
              <span className="personnel-list-meta">
                <strong>{record.summary.name}</strong>
                <span>
                  {[
                    record.summary.rank || "звання не вказано",
                    record.summary.callSign,
                    record.summary.birthDate,
                  ]
                    .filter(Boolean)
                    .join(" · ")}
                </span>
              </span>
            </button>
          );
        })}
      </div>
    </div>
  );
}
