import { createPortal } from "react-dom";
import { Button } from "@/components/sci/SciPrimitives";
import { ANKETA_COLUMNS, type AnketaColumnKey } from "../anketaSheet";

type AnketaGapColumnsMenuProps = {
  gapColumnKeys: AnketaColumnKey[];
  gapKeySet: Set<AnketaColumnKey>;
  gapColumnsOpen: boolean;
  setGapColumnsOpen: (value: boolean | ((current: boolean) => boolean)) => void;
  gapMenuPos: { top: number; left: number };
  gapTriggerRef: React.RefObject<HTMLDivElement | null>;
  gapMenuRef: React.RefObject<HTMLDivElement | null>;
  onToggleColumn: (key: AnketaColumnKey, isReadonly: boolean) => void;
  onSelectAll: () => void;
  onClear: () => void;
};

export function AnketaGapColumnsMenu({
  gapColumnKeys,
  gapKeySet,
  gapColumnsOpen,
  setGapColumnsOpen,
  gapMenuPos,
  gapTriggerRef,
  gapMenuRef,
  onToggleColumn,
  onSelectAll,
  onClear,
}: AnketaGapColumnsMenuProps) {
  return (
    <div className="anketa-gap-columns" ref={gapTriggerRef}>
      <Button
        variant="outlined"
        aria-expanded={gapColumnsOpen}
        aria-haspopup="listbox"
        onClick={() => setGapColumnsOpen((value) => !value)}
      >
        Колонки пропусків · {gapColumnKeys.length}
      </Button>
      {gapColumnsOpen
        ? createPortal(
            <div
              ref={gapMenuRef}
              className="anketa-gap-columns-menu"
              role="listbox"
              style={{ top: gapMenuPos.top, left: gapMenuPos.left }}
              onPointerDown={(event) => event.stopPropagation()}
            >
              <div className="anketa-gap-columns-actions">
                <button type="button" onClick={onSelectAll}>
                  Усі
                </button>
                <button type="button" onClick={onClear}>
                  Жодної
                </button>
              </div>
              <div className="anketa-gap-columns-list">
                {ANKETA_COLUMNS.map((column) => {
                  const checked = gapKeySet.has(column.key);
                  const isReadonly = Boolean(column.readonly);
                  return (
                    <button
                      key={column.key}
                      type="button"
                      role="option"
                      aria-selected={checked}
                      disabled={isReadonly}
                      className={
                        isReadonly
                          ? "anketa-gap-columns-option is-readonly"
                          : checked
                            ? "anketa-gap-columns-option is-active"
                            : "anketa-gap-columns-option"
                      }
                      onClick={() => onToggleColumn(column.key, isReadonly)}
                    >
                      <span className="sci-data-table-check" aria-hidden="true">
                        {checked ? "■" : ""}
                      </span>
                      <span>
                        {column.header}
                        {isReadonly ? " · лише перегляд" : ""}
                      </span>
                    </button>
                  );
                })}
              </div>
            </div>,
            document.body,
          )
        : null}
    </div>
  );
}
