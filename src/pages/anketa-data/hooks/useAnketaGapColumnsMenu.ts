import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import {
  ANKETA_COLUMNS,
  type AnketaColumnKey,
} from "../anketaSheet";
import { readAnketaGapColumns, writeAnketaGapColumns } from "../anketaGaps";

export function useAnketaGapColumnsMenu() {
  const [gapColumnKeys, setGapColumnKeys] = useState<AnketaColumnKey[]>(() =>
    readAnketaGapColumns(),
  );
  const [gapColumnsOpen, setGapColumnsOpen] = useState(false);
  const [gapMenuPos, setGapMenuPos] = useState({ top: 0, left: 0 });
  const gapTriggerRef = useRef<HTMLDivElement | null>(null);
  const gapMenuRef = useRef<HTMLDivElement | null>(null);

  const gapKeySet = useMemo(() => new Set(gapColumnKeys), [gapColumnKeys]);

  useEffect(() => {
    writeAnketaGapColumns(gapColumnKeys);
  }, [gapColumnKeys]);

  const updateGapMenuPosition = () => {
    const rect = gapTriggerRef.current?.getBoundingClientRect();
    if (!rect) return;
    const width = Math.min(360, window.innerWidth - 24);
    const left = Math.min(
      Math.max(12, rect.left),
      Math.max(12, window.innerWidth - width - 12),
    );
    setGapMenuPos({ top: rect.bottom + 6, left });
  };

  useLayoutEffect(() => {
    if (!gapColumnsOpen) return;
    const frame = window.requestAnimationFrame(() => updateGapMenuPosition());
    return () => window.cancelAnimationFrame(frame);
  }, [gapColumnsOpen]);

  useEffect(() => {
    if (!gapColumnsOpen) return;
    let repositionTimer = 0;
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target as Node;
      if (gapMenuRef.current?.contains(target)) return;
      if (gapTriggerRef.current?.contains(target)) return;
      setGapColumnsOpen(false);
    };
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") setGapColumnsOpen(false);
    };
    const onReposition = () => {
      window.clearTimeout(repositionTimer);
      repositionTimer = window.setTimeout(() => updateGapMenuPosition(), 32);
    };
    window.addEventListener("pointerdown", onPointerDown);
    window.addEventListener("keydown", onKeyDown);
    window.addEventListener("resize", onReposition);
    window.addEventListener("scroll", onReposition, true);
    return () => {
      window.clearTimeout(repositionTimer);
      window.removeEventListener("pointerdown", onPointerDown);
      window.removeEventListener("keydown", onKeyDown);
      window.removeEventListener("resize", onReposition);
      window.removeEventListener("scroll", onReposition, true);
    };
  }, [gapColumnsOpen]);

  const toggleGapColumn = (key: AnketaColumnKey, isReadonly: boolean) => {
    if (isReadonly) return;
    setGapColumnKeys((current) => {
      const next = current.includes(key)
        ? current.filter((item) => item !== key)
        : ANKETA_COLUMNS.map((column) => column.key).filter(
            (item) => item === key || current.includes(item),
          );
      writeAnketaGapColumns(next);
      return next;
    });
  };

  const selectAllGapColumns = () => {
    const next = ANKETA_COLUMNS.filter((column) => !column.readonly).map(
      (column) => column.key,
    );
    writeAnketaGapColumns(next);
    setGapColumnKeys(next);
  };

  const clearGapColumns = () => {
    writeAnketaGapColumns([]);
    setGapColumnKeys([]);
  };

  return {
    gapColumnKeys,
    gapKeySet,
    gapColumnsOpen,
    setGapColumnsOpen,
    gapMenuPos,
    gapTriggerRef,
    gapMenuRef,
    toggleGapColumn,
    selectAllGapColumns,
    clearGapColumns,
  };
}
