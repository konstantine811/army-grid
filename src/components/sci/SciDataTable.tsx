import type { CSSProperties, ReactNode } from "react";
import { useDeferredValue, useEffect, useMemo, useRef, useState } from "react";
import { useVirtualizer } from "@tanstack/react-virtual";
import {
  PushPinOutlinedIcon,
  SortArrowDownIcon,
  SortArrowUpDownIcon,
  SortArrowUpIcon,
} from "@/components/sci/icons";
import {
  Select,
  SelectContent,
  SelectTrigger,
} from "@/components/ui/select/select";

export type SciDataTablePin = "left" | "right";

export type MRT_ColumnDef<TData> = {
  id?: string;
  accessorKey?: keyof TData | string;
  header?: ReactNode;
  Header?: () => ReactNode;
  accessorFn?: (row: TData) => ReactNode;
  exportValue?: (row: TData) => string;
  Cell?: (props: {
    row: { original: TData; index: number };
    cell: { getValue: () => unknown };
  }) => ReactNode;
  size?: number;
  minSize?: number;
  enableColumnFilter?: boolean;
  filterVariant?: "facet" | "number-range";
  enableHiding?: boolean;
  enableGlobalFilter?: boolean;
  pin?: SciDataTablePin;
  enableSorting?: boolean;
  muiEditTextFieldProps?: unknown;
};

export type SciDataTableExportContext<TData> = {
  rows: TData[];
  columns: Array<{
    id: string;
    label: string;
    value: (row: TData) => string;
  }>;
};

export type SciDataTableCellRef = {
  rowId: string;
  columnId: string;
  /** Bumps on each navigate so scroll/focus re-runs even for the same cell. */
  focusEpoch?: number;
};

type SciTableOptions<TData> = {
  columns: MRT_ColumnDef<TData>[];
  data: TData[];
  initialState?: {
    density?: string;
    pagination?: {
      pageIndex?: number;
      pageSize?: number;
    };
    columnVisibility?: Record<string, boolean>;
    columnPinning?: {
      left?: string[];
      right?: string[];
    };
  };
  enableColumnFilters?: boolean;
  enableColumnVisibility?: boolean;
  enableGlobalFilter?: boolean;
  globalFilterPlaceholder?: string;
  exportLabel?: string;
  onExport?: (
    context: SciDataTableExportContext<TData>,
  ) => void | Promise<void>;
  emptyMessage?: string;
  getRowId?: (row: TData, index: number) => string;
  /** Extra props/class for body cells (editing highlight, data attrs). */
  getTdProps?: (args: {
    row: TData;
    rowIndex: number;
    rowId: string;
    columnId: string;
  }) =>
    | {
        className?: string;
        style?: CSSProperties;
        title?: string;
      }
    | undefined;
  /** Scroll/focus target cell after render. */
  focusedCell?: SciDataTableCellRef | null;
  /** Virtualize body rows (default: on when data length > 80). */
  enableRowVirtualization?: boolean;
  /** Estimated row height for virtualizer. */
  estimatedRowHeight?: number;
  muiTableBodyCellProps?:
    | unknown
    | ((props: {
        cell: { column: { id: string } };
        row: { original: TData; index: number };
      }) => unknown);
  muiTableBodyRowProps?:
    | unknown
    | ((props: { row: { original: TData; index: number } }) => unknown);
  [key: string]: unknown;
};

type PreparedColumn<TData> = MRT_ColumnDef<TData> & {
  columnId: string;
  label: string;
  width: number;
  pin?: SciDataTablePin;
  sortable: boolean;
};

type ColumnRangeFilters = Record<string, { min: string; max: string }>;
type ColumnSortDirection = "asc" | "desc";
type ColumnSortState = {
  columnId: string;
  direction: ColumnSortDirection;
};

export function useMaterialReactTable<TData>(options: SciTableOptions<TData>) {
  return options;
}

export function MaterialReactTable<TData>({
  table,
}: {
  table: SciTableOptions<TData>;
}) {
  const preparedColumns = useMemo(
    () =>
      table.columns.map((column, index) => prepareColumn(column, index, table)),
    [table.columns, table.initialState?.columnPinning],
  );
  const [globalFilter, setGlobalFilter] = useState("");
  const deferredGlobalFilter = useDeferredValue(globalFilter);
  const [columnFilters, setColumnFilters] = useState<Record<string, string[]>>(
    {},
  );
  const [columnRangeFilters, setColumnRangeFilters] =
    useState<ColumnRangeFilters>({});
  const [columnVisibility, setColumnVisibility] = useState<
    Record<string, boolean>
  >(() => table.initialState?.columnVisibility ?? {});
  const [pinOverrides, setPinOverrides] = useState<
    Record<string, SciDataTablePin | "off">
  >({});
  const [openMenu, setOpenMenu] = useState<string | null>(null);
  const [sorting, setSorting] = useState<ColumnSortState | null>(null);
  const scrollParentRef = useRef<HTMLDivElement | null>(null);
  const pinnedColumns = useMemo(
    () =>
      preparedColumns.map((column) => ({
        ...column,
        pin: resolveColumnPin(column, pinOverrides),
      })),
    [pinOverrides, preparedColumns],
  );
  const visibleColumns = useMemo(() => {
    const visible = pinnedColumns.filter(
      (column) => columnVisibility[column.columnId] !== false,
    );
    const left = visible.filter((column) => column.pin === "left");
    const right = visible.filter((column) => column.pin === "right");
    const rest = visible.filter((column) => !column.pin);
    return [...left, ...rest, ...right];
  }, [columnVisibility, pinnedColumns]);
  const pinnedOffsets = useMemo(
    () => calculatePinnedOffsets(visibleColumns),
    [visibleColumns],
  );
  const filteredRows = useMemo(
    () =>
      table.data.filter((row) =>
        rowMatchesFilters(
          row,
          preparedColumns,
          deferredGlobalFilter,
          columnFilters,
          columnRangeFilters,
        ),
      ),
    [
      columnFilters,
      columnRangeFilters,
      deferredGlobalFilter,
      preparedColumns,
      table.data,
    ],
  );
  const sortedRows = useMemo(() => {
    try {
      return sortRows(filteredRows, preparedColumns, sorting);
    } catch {
      return filteredRows;
    }
  }, [filteredRows, preparedColumns, sorting]);
  const facetedFilterOptions = useMemo(() => {
    const rowCount = table.data.length;
    const facetMenuId = openMenu?.startsWith("facet:")
      ? openMenu.slice("facet:".length)
      : null;

    // На великих таблицях рахуємо опції лише для відкритого фільтра (O(rows), не O(rows×cols)).
    if (rowCount > 300) {
      if (!facetMenuId) return {};
      return {
        [facetMenuId]: getFacetedColumnOptions(
          table.data,
          preparedColumns,
          facetMenuId,
          deferredGlobalFilter,
          columnFilters,
          columnRangeFilters,
        ),
      };
    }

    return Object.fromEntries(
      preparedColumns.map((column) => [
        column.columnId,
        getFacetedColumnOptions(
          table.data,
          preparedColumns,
          column.columnId,
          deferredGlobalFilter,
          columnFilters,
          columnRangeFilters,
        ),
      ]),
    );
  }, [
    columnFilters,
    columnRangeFilters,
    deferredGlobalFilter,
    openMenu,
    preparedColumns,
    table.data,
  ]);
  const enableVirtualization =
    table.enableRowVirtualization ?? table.data.length > 80;
  const estimatedRowHeight = table.estimatedRowHeight ?? 40;
  const pageSize = table.initialState?.pagination?.pageSize ?? 200;
  const rows = enableVirtualization
    ? sortedRows
    : sortedRows.slice(0, pageSize);

  const rowVirtualizer = useVirtualizer({
    count: enableVirtualization ? rows.length : 0,
    getScrollElement: () => scrollParentRef.current,
    estimateSize: () => estimatedRowHeight,
    overscan: 14,
  });
  const rowVirtualizerRef = useRef(rowVirtualizer);
  rowVirtualizerRef.current = rowVirtualizer;

  const focusedRowId = table.focusedCell?.rowId ?? "";
  const focusedColumnId = table.focusedCell?.columnId ?? "";
  const focusedEpoch = table.focusedCell?.focusEpoch ?? 0;
  const filterClearAttemptRef = useRef<string>("");
  const hasActiveColumnFilters =
    Object.keys(columnFilters).length > 0 ||
    Object.keys(columnRangeFilters).length > 0;

  useEffect(() => {
    if (!focusedRowId || !focusedColumnId) {
      filterClearAttemptRef.current = "";
      return;
    }

    const focusKey = `${focusedRowId}::${focusedColumnId}::${focusedEpoch}`;
    const rowIndex = rows.findIndex(
      (row, index) =>
        String(table.getRowId?.(row, index) ?? index) === focusedRowId,
    );

    // Target hidden by filters. Clear them once, then wait for deferred filter flush.
    if (rowIndex < 0) {
      if (filterClearAttemptRef.current === focusKey) return;

      if (globalFilter.trim()) {
        filterClearAttemptRef.current = focusKey;
        queueMicrotask(() => setGlobalFilter(""));
        return;
      }
      if (deferredGlobalFilter.trim()) {
        return;
      }
      if (hasActiveColumnFilters) {
        filterClearAttemptRef.current = focusKey;
        queueMicrotask(() => {
          setColumnFilters({});
          setColumnRangeFilters({});
        });
        return;
      }
      return;
    }

    filterClearAttemptRef.current = "";

    let cancelled = false;
    let attempts = 0;
    let timer = 0;
    let raf = 0;

    const focusCellEditor = () => {
      if (cancelled) return;
      const node = document.querySelector<HTMLElement>(
        `[data-sci-cell="${CSS.escape(`${focusedRowId}::${focusedColumnId}`)}"]`,
      );
      if (node) {
        node.scrollIntoView({
          block: "nearest",
          inline: "center",
          behavior: "auto",
        });
        const input = node.querySelector<
          HTMLInputElement | HTMLTextAreaElement
        >(
          "input.anketa-cell-input, textarea.anketa-cell-input, input, textarea",
        );
        if (input) {
          if (input === document.activeElement) {
            return;
          }
          input.focus({ preventScroll: true });
          if ("select" in input && attempts === 0) input.select();
          return;
        }
      }

      if (attempts >= 12) return;
      attempts += 1;
      if (enableVirtualization) {
        rowVirtualizerRef.current.scrollToIndex(rowIndex, { align: "center" });
      }
      timer = window.setTimeout(focusCellEditor, 64);
    };

    raf = window.requestAnimationFrame(() => {
      if (enableVirtualization) {
        rowVirtualizerRef.current.scrollToIndex(rowIndex, { align: "center" });
      }
      timer = window.setTimeout(focusCellEditor, enableVirtualization ? 64 : 0);
    });

    return () => {
      cancelled = true;
      window.cancelAnimationFrame(raf);
      window.clearTimeout(timer);
    };
  }, [
    deferredGlobalFilter,
    enableVirtualization,
    focusedColumnId,
    focusedEpoch,
    focusedRowId,
    globalFilter,
    hasActiveColumnFilters,
    rows,
    table.getRowId,
  ]);

  const showToolbar =
    table.enableGlobalFilter !== false ||
    table.enableColumnVisibility !== false ||
    Boolean(table.onExport);
  const hasColumnFilters =
    table.enableColumnFilters !== false &&
    visibleColumns.some((column) => column.enableColumnFilter !== false);

  const toggleColumnPin = (column: PreparedColumn<TData>) => {
    const original = preparedColumns.find(
      (item) => item.columnId === column.columnId,
    );
    setPinOverrides((current) => {
      const resolved = resolveColumnPin(original ?? column, current);
      return {
        ...current,
        [column.columnId]: resolved
          ? "off"
          : original?.pin === "right" || column.columnId === "actions"
            ? "right"
            : "left",
      };
    });
  };

  const toggleColumnSort = (column: PreparedColumn<TData>) => {
    if (!column.sortable) return;
    setSorting((current) => {
      if (current?.columnId !== column.columnId) {
        return { columnId: column.columnId, direction: "asc" };
      }
      if (current.direction === "asc") {
        return { columnId: column.columnId, direction: "desc" };
      }
      return null;
    });
  };

  return (
    <div className="sci-data-table-shell">
      {showToolbar ? (
        <div className="sci-data-table-toolbar">
          {table.enableGlobalFilter !== false ? (
            <label className="sci-data-table-search">
              <span>Пошук</span>
              <input
                value={globalFilter}
                onChange={(event) => setGlobalFilter(event.target.value)}
                placeholder={
                  table.globalFilterPlaceholder ?? "Пошук по таблиці"
                }
              />
            </label>
          ) : null}
          {table.onExport ? (
            <button
              type="button"
              className="sci-data-table-export"
              onClick={() =>
                void table.onExport?.({
                  rows: filteredRows,
                  columns: preparedColumns
                    .filter(
                      (column) => columnVisibility[column.columnId] !== false,
                    )
                    .filter((column) => column.columnId !== "actions")
                    .map((column) => ({
                      id: column.columnId,
                      label: column.label,
                      value: (row) =>
                        column.exportValue?.(row) ??
                        getPlainCellValue(column, row),
                    })),
                })
              }
            >
              {table.exportLabel ?? "Експорт"}
            </button>
          ) : null}
          {table.enableColumnVisibility !== false ? (
            <div className="sci-data-table-columns">
              <TableSelectMenu
                id="columns"
                openId={openMenu}
                onOpenIdChange={setOpenMenu}
                trigger="Колонки"
                triggerClassName="sci-data-table-columns-trigger"
                contentClassName="sci-data-table-columns-menu"
              >
                {preparedColumns.map((column) => (
                  <label key={column.columnId}>
                    <input
                      type="checkbox"
                      checked={columnVisibility[column.columnId] !== false}
                      disabled={column.enableHiding === false}
                      onChange={(event) =>
                        setColumnVisibility((current) => ({
                          ...current,
                          [column.columnId]: event.target.checked,
                        }))
                      }
                    />
                    <span>{column.label}</span>
                  </label>
                ))}
              </TableSelectMenu>
            </div>
          ) : null}
        </div>
      ) : null}

      <div className="sci-data-table-wrap" ref={scrollParentRef}>
        <table className="sci-data-table">
          <thead>
            <tr>
              {visibleColumns.map((column) => (
                <th
                  key={column.columnId}
                  className={pinnedClassName(column.pin)}
                  style={pinnedStyle(column, pinnedOffsets)}
                >
                  <div className="sci-data-table-header-cell">
                    <span>
                      {column.Header ? column.Header() : column.header}
                    </span>
                    <div className="sci-data-table-header-actions">
                      {column.sortable ? (
                        <button
                          type="button"
                          className={
                            sorting?.columnId === column.columnId
                              ? "sci-data-table-sort is-active"
                              : "sci-data-table-sort"
                          }
                          aria-label={
                            sorting?.columnId === column.columnId
                              ? sorting.direction === "asc"
                                ? `Сортування ${column.label}: за зростанням`
                                : `Сортування ${column.label}: за спаданням`
                              : `Сортувати колонку ${column.label}`
                          }
                          aria-pressed={sorting?.columnId === column.columnId}
                          title={
                            sorting?.columnId === column.columnId
                              ? sorting.direction === "asc"
                                ? "За зростанням"
                                : "За спаданням"
                              : "Сортувати"
                          }
                          onClick={() => toggleColumnSort(column)}
                        >
                          {sorting?.columnId === column.columnId ? (
                            sorting.direction === "asc" ? (
                              <SortArrowUpIcon fontSize="small" />
                            ) : (
                              <SortArrowDownIcon fontSize="small" />
                            )
                          ) : (
                            <SortArrowUpDownIcon fontSize="small" />
                          )}
                        </button>
                      ) : null}
                      <button
                        type="button"
                        className={
                          column.pin
                            ? "sci-data-table-pin is-pinned"
                            : "sci-data-table-pin"
                        }
                        aria-label={
                          column.pin
                            ? `Відкріпити колонку ${column.label}`
                            : `Закріпити колонку ${column.label}`
                        }
                        aria-pressed={Boolean(column.pin)}
                        title={
                          column.pin
                            ? "Відкріпити колонку"
                            : "Закріпити колонку"
                        }
                        onClick={() => toggleColumnPin(column)}
                      >
                        <PushPinOutlinedIcon fontSize="small" />
                      </button>
                    </div>
                  </div>
                </th>
              ))}
            </tr>
            {hasColumnFilters ? (
              <tr className="sci-data-table-filter-row">
                {visibleColumns.map((column) => (
                  <th
                    key={`${column.columnId}:filter`}
                    className={pinnedClassName(column.pin)}
                    style={pinnedStyle(column, pinnedOffsets)}
                  >
                    {column.enableColumnFilter ===
                    false ? null : column.filterVariant === "number-range" ? (
                      <ColumnNumberRangeFilter
                        label={column.label}
                        openId={openMenu}
                        menuId={`range:${column.columnId}`}
                        onOpenIdChange={setOpenMenu}
                        value={
                          columnRangeFilters[column.columnId] ?? {
                            min: "",
                            max: "",
                          }
                        }
                        onChange={(range) =>
                          setColumnRangeFilters((current) => {
                            const next = { ...current };
                            if (range.min.trim() || range.max.trim()) {
                              next[column.columnId] = range;
                            } else {
                              delete next[column.columnId];
                            }
                            return next;
                          })
                        }
                      />
                    ) : (
                      <ColumnFacetFilter
                        label={column.label}
                        openId={openMenu}
                        menuId={`facet:${column.columnId}`}
                        onOpenIdChange={setOpenMenu}
                        options={facetedFilterOptions[column.columnId] ?? []}
                        selected={columnFilters[column.columnId] ?? []}
                        onChange={(selected) =>
                          setColumnFilters((current) => {
                            const next = { ...current };
                            if (selected.length)
                              next[column.columnId] = selected;
                            else delete next[column.columnId];
                            return next;
                          })
                        }
                      />
                    )}
                  </th>
                ))}
              </tr>
            ) : null}
          </thead>
          <tbody>
            {enableVirtualization && rows.length > 0 ? (
              (() => {
                const virtualItems = rowVirtualizer.getVirtualItems();
                const paddingTop = virtualItems[0]?.start ?? 0;
                const paddingBottom = Math.max(
                  0,
                  rowVirtualizer.getTotalSize() -
                    (virtualItems.at(-1)?.end ?? 0),
                );
                return (
                  <>
                    {paddingTop > 0 ? (
                      <tr
                        aria-hidden
                        className="sci-data-table-virtual-spacer"
                        style={{ height: paddingTop }}
                      >
                        <td colSpan={Math.max(visibleColumns.length, 1)} />
                      </tr>
                    ) : null}
                    {virtualItems.map((virtualRow) => {
                      const rowIndex = virtualRow.index;
                      const row = rows[rowIndex];
                      if (!row) return null;
                      const rowId = String(
                        table.getRowId?.(row, rowIndex) ?? rowIndex,
                      );
                      return (
                        <tr
                          key={rowId}
                          data-index={rowIndex}
                          ref={rowVirtualizer.measureElement}
                          style={{ height: virtualRow.size }}
                        >
                          {visibleColumns.map((column) => {
                            const tdProps = table.getTdProps?.({
                              row,
                              rowIndex,
                              rowId,
                              columnId: column.columnId,
                            });
                            const focused =
                              table.focusedCell?.rowId === rowId &&
                              table.focusedCell?.columnId === column.columnId;
                            return (
                              <td
                                key={column.columnId}
                                data-sci-cell={`${rowId}::${column.columnId}`}
                                className={[
                                  pinnedClassName(column.pin),
                                  tdProps?.className,
                                  focused ? "is-focused-empty-cell" : "",
                                ]
                                  .filter(Boolean)
                                  .join(" ")}
                                style={{
                                  ...pinnedStyle(column, pinnedOffsets),
                                  ...tdProps?.style,
                                }}
                                title={tdProps?.title}
                              >
                                {renderCell(column, row, rowIndex)}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                    {paddingBottom > 0 ? (
                      <tr
                        aria-hidden
                        className="sci-data-table-virtual-spacer"
                        style={{ height: paddingBottom }}
                      >
                        <td colSpan={Math.max(visibleColumns.length, 1)} />
                      </tr>
                    ) : null}
                  </>
                );
              })()
            ) : (
              <>
                {rows.map((row, rowIndex) => {
                  const rowId = String(
                    table.getRowId?.(row, rowIndex) ?? rowIndex,
                  );
                  return (
                    <tr key={rowId}>
                      {visibleColumns.map((column) => {
                        const tdProps = table.getTdProps?.({
                          row,
                          rowIndex,
                          rowId,
                          columnId: column.columnId,
                        });
                        const focused =
                          table.focusedCell?.rowId === rowId &&
                          table.focusedCell?.columnId === column.columnId;
                        return (
                          <td
                            key={column.columnId}
                            data-sci-cell={`${rowId}::${column.columnId}`}
                            className={[
                              pinnedClassName(column.pin),
                              tdProps?.className,
                              focused ? "is-focused-empty-cell" : "",
                            ]
                              .filter(Boolean)
                              .join(" ")}
                            style={{
                              ...pinnedStyle(column, pinnedOffsets),
                              ...tdProps?.style,
                            }}
                            title={tdProps?.title}
                          >
                            {renderCell(column, row, rowIndex)}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </>
            )}
            {rows.length === 0 && (
              <tr>
                <td
                  colSpan={Math.max(visibleColumns.length, 1)}
                  className="sci-data-table-empty"
                >
                  {table.emptyMessage ?? "Даних немає"}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {enableVirtualization ? (
        <div className="sci-data-table-footer">
          {globalFilter !== deferredGlobalFilter
            ? "Фільтрація…"
            : `Рядків: ${filteredRows.length}`}
        </div>
      ) : filteredRows.length > rows.length ? (
        <div className="sci-data-table-footer">
          Показано {rows.length} з {filteredRows.length}
        </div>
      ) : null}
    </div>
  );
}

function prepareColumn<TData>(
  column: MRT_ColumnDef<TData>,
  index: number,
  table: SciTableOptions<TData>,
): PreparedColumn<TData> {
  const columnId =
    column.id ??
    (column.accessorKey ? String(column.accessorKey) : `column_${index}`);
  const leftPinned =
    table.initialState?.columnPinning?.left?.includes(columnId);
  const rightPinned =
    table.initialState?.columnPinning?.right?.includes(columnId);

  return {
    ...column,
    columnId,
    label: labelText(column.header ?? columnId),
    width: column.size ?? column.minSize ?? 160,
    pin:
      column.pin ?? (leftPinned ? "left" : rightPinned ? "right" : undefined),
    sortable:
      column.enableSorting !== false &&
      columnId !== "actions" &&
      Boolean(column.header),
  };
}

function resolveColumnPin<TData>(
  column: PreparedColumn<TData>,
  overrides: Record<string, SciDataTablePin | "off">,
) {
  const override = overrides[column.columnId];
  if (override === "off") return undefined;
  if (override) return override;
  return column.pin;
}

function calculatePinnedOffsets<TData>(columns: Array<PreparedColumn<TData>>) {
  const left = new Map<string, number>();
  const right = new Map<string, number>();
  let leftOffset = 0;
  columns.forEach((column) => {
    if (column.pin !== "left") return;
    left.set(column.columnId, leftOffset);
    leftOffset += column.width;
  });

  let rightOffset = 0;
  [...columns].reverse().forEach((column) => {
    if (column.pin !== "right") return;
    right.set(column.columnId, rightOffset);
    rightOffset += column.width;
  });

  return { left, right };
}

function pinnedClassName(pin?: SciDataTablePin) {
  return pin ? `sci-data-table-pinned pinned-${pin}` : undefined;
}

function pinnedStyle<TData>(
  column: PreparedColumn<TData>,
  offsets: {
    left: Map<string, number>;
    right: Map<string, number>;
  },
): CSSProperties {
  const style: CSSProperties = {
    width: column.width,
    minWidth: column.width,
    maxWidth: column.width,
  };
  if (column.pin === "left")
    style.left = offsets.left.get(column.columnId) ?? 0;
  if (column.pin === "right")
    style.right = offsets.right.get(column.columnId) ?? 0;
  return style;
}

function TableSelectMenu({
  id,
  openId,
  onOpenIdChange,
  trigger,
  children,
  triggerClassName,
  contentClassName,
}: {
  id: string;
  openId: string | null;
  onOpenIdChange: (
    id: string | null | ((current: string | null) => string | null),
  ) => void;
  trigger: ReactNode;
  children: ReactNode;
  triggerClassName?: string;
  contentClassName?: string;
}) {
  return (
    <Select
      open={openId === id}
      onOpenChange={(nextOpen) =>
        onOpenIdChange((current) =>
          nextOpen ? id : current === id ? null : current,
        )
      }
    >
      <SelectTrigger
        className={triggerClassName ?? "sci-data-table-facet-trigger"}
      >
        {trigger}
      </SelectTrigger>
      <SelectContent
        className={
          contentClassName ??
          "sci-data-table-facet-menu sci-data-table-select-content"
        }
        position="popper"
        align="start"
        onCloseAutoFocus={(event) => event.preventDefault()}
      >
        {children}
      </SelectContent>
    </Select>
  );
}

function ColumnFacetFilter({
  label,
  options,
  selected,
  onChange,
  openId,
  menuId,
  onOpenIdChange,
}: {
  label: string;
  options: string[];
  selected: string[];
  onChange: (selected: string[]) => void;
  openId: string | null;
  menuId: string;
  onOpenIdChange: (
    id: string | null | ((current: string | null) => string | null),
  ) => void;
}) {
  const [optionQuery, setOptionQuery] = useState("");
  const selectedSet = new Set(selected);
  const caption = selected.length ? `Обрано: ${selected.length}` : "Фільтр";
  const monthGroups = useMemo(() => buildFacetMonthGroups(options), [options]);
  const query = normalizeFilter(optionQuery);
  const filteredOptions = options.filter((option) =>
    normalizeFilter(option).includes(query),
  );
  const filteredMonths = monthGroups.filter((group) => {
    if (!query) return true;
    return (
      normalizeFilter(group.label).includes(query) ||
      normalizeFilter(group.fromLabel).includes(query) ||
      group.dates.some((date) => normalizeFilter(date).includes(query))
    );
  });

  const toggleOption = (option: string) => {
    if (selectedSet.has(option)) {
      onChange(selected.filter((item) => item !== option));
      return;
    }
    onChange([...selected, option]);
  };

  const toggleMonth = (group: FacetMonthGroup) => {
    const allSelected = group.dates.every((date) => selectedSet.has(date));
    if (allSelected) {
      const remove = new Set(group.dates);
      onChange(selected.filter((item) => !remove.has(item)));
      return;
    }
    const next = new Set(selected);
    group.dates.forEach((date) => next.add(date));
    onChange([...next]);
  };

  const selectFromMonth = (group: FacetMonthGroup) => {
    const next = new Set(selected);
    monthGroups
      .filter((item) => item.key >= group.key)
      .forEach((item) => item.dates.forEach((date) => next.add(date)));
    onChange([...next]);
  };

  return (
    <TableSelectMenu
      id={menuId}
      openId={openId}
      onOpenIdChange={onOpenIdChange}
      trigger={<span>{caption}</span>}
    >
      <input
        className="sci-data-table-facet-search"
        value={optionQuery}
        onChange={(event) => setOptionQuery(event.target.value)}
        onKeyDown={(event) => event.stopPropagation()}
        onPointerDown={(event) => event.stopPropagation()}
        placeholder="Пошук значення"
      />
      <button
        type="button"
        className={!selected.length ? "is-active" : undefined}
        onClick={() => onChange([])}
      >
        <span className="sci-data-table-check" aria-hidden="true">
          {!selected.length ? "■" : ""}
        </span>
        Усі {label.toLocaleLowerCase("uk-UA")}
      </button>
      {filteredMonths.length ? (
        <>
          <div className="sci-data-table-facet-section">Місяці</div>
          {filteredMonths.map((group) => {
            const selectedCount = group.dates.filter((date) =>
              selectedSet.has(date),
            ).length;
            const allSelected = selectedCount === group.dates.length;
            return (
              <label key={group.key}>
                <input
                  type="checkbox"
                  checked={allSelected}
                  onChange={() => toggleMonth(group)}
                />
                <span>
                  {group.label}
                  <small> · {group.dates.length}</small>
                </span>
              </label>
            );
          })}
          <div className="sci-data-table-facet-section">Від місяця</div>
          {filteredMonths.map((group) => (
            <button
              type="button"
              className="sci-data-table-facet-from"
              key={`from-${group.key}`}
              onClick={() => selectFromMonth(group)}
            >
              <span className="sci-data-table-check" aria-hidden="true">
                ▸
              </span>
              {group.fromLabel}
            </button>
          ))}
          <div className="sci-data-table-facet-section">Дати</div>
        </>
      ) : null}
      {filteredOptions.map((option) => (
        <label key={option}>
          <input
            type="checkbox"
            checked={selectedSet.has(option)}
            onChange={() => toggleOption(option)}
          />
          <span>{option}</span>
        </label>
      ))}
      {!filteredOptions.length && !filteredMonths.length ? (
        <div className="sci-data-table-facet-empty">Немає значень</div>
      ) : null}
    </TableSelectMenu>
  );
}

function ColumnNumberRangeFilter({
  label,
  value,
  onChange,
  openId,
  menuId,
  onOpenIdChange,
}: {
  label: string;
  value: { min: string; max: string };
  onChange: (value: { min: string; max: string }) => void;
  openId: string | null;
  menuId: string;
  onOpenIdChange: (
    id: string | null | ((current: string | null) => string | null),
  ) => void;
}) {
  const hasRange = Boolean(value.min.trim() || value.max.trim());
  const caption = hasRange
    ? `${value.min.trim() || "0"}-${value.max.trim() || "∞"}`
    : "Фільтр";

  return (
    <TableSelectMenu
      id={menuId}
      openId={openId}
      onOpenIdChange={onOpenIdChange}
      trigger={<span>{caption}</span>}
      contentClassName="sci-data-table-facet-menu sci-data-table-select-content sci-data-table-range-menu"
    >
      <div className="sci-data-table-range-title">{label}</div>
      <label>
        <span>Від</span>
        <input
          className="sci-data-table-facet-search"
          inputMode="numeric"
          value={value.min}
          onChange={(event) => onChange({ ...value, min: event.target.value })}
          onKeyDown={(event) => event.stopPropagation()}
          onPointerDown={(event) => event.stopPropagation()}
          placeholder="0"
        />
      </label>
      <label>
        <span>До</span>
        <input
          className="sci-data-table-facet-search"
          inputMode="numeric"
          value={value.max}
          onChange={(event) => onChange({ ...value, max: event.target.value })}
          onKeyDown={(event) => event.stopPropagation()}
          onPointerDown={(event) => event.stopPropagation()}
          placeholder="без межі"
        />
      </label>
      <button
        type="button"
        className={!hasRange ? "is-active" : undefined}
        onClick={() => onChange({ min: "", max: "" })}
      >
        <span className="sci-data-table-check" aria-hidden="true">
          {!hasRange ? "■" : ""}
        </span>
        Усі {label.toLocaleLowerCase("uk-UA")}
      </button>
    </TableSelectMenu>
  );
}

function rowMatchesFilters<TData>(
  row: TData,
  columns: Array<PreparedColumn<TData>>,
  globalFilter: string,
  columnFilters: Record<string, string[]>,
  columnRangeFilters: ColumnRangeFilters,
  ignoredColumnId?: string,
) {
  const global = normalizeFilter(globalFilter);
  if (global) {
    const rowText = columns
      .filter((column) => column.enableGlobalFilter !== false)
      .map((column) => normalizeFilter(getPlainCellValue(column, row)))
      .join(" ");
    if (!rowText.includes(global)) return false;
  }

  const matchesFacets = Object.entries(columnFilters).every(
    ([columnId, filter]) => {
      if (columnId === ignoredColumnId) return true;
      const normalizedFilter = filter.map(normalizeFilter).filter(Boolean);
      if (!normalizedFilter.length) return true;
      const column = columns.find((item) => item.columnId === columnId);
      if (!column) return true;
      return normalizedFilter.includes(
        normalizeFilter(getPlainCellValue(column, row)),
      );
    },
  );
  if (!matchesFacets) return false;

  return Object.entries(columnRangeFilters).every(([columnId, range]) => {
    if (columnId === ignoredColumnId) return true;
    const min = parseFilterNumber(range.min);
    const max = parseFilterNumber(range.max);
    if (min == null && max == null) return true;
    const column = columns.find((item) => item.columnId === columnId);
    if (!column) return true;
    const value = parseFilterNumber(getPlainCellValue(column, row));
    if (value == null) return false;
    if (min != null && value < min) return false;
    if (max != null && value > max) return false;
    return true;
  });
}

function getFacetedColumnOptions<TData>(
  rows: TData[],
  columns: Array<PreparedColumn<TData>>,
  columnId: string,
  globalFilter: string,
  columnFilters: Record<string, string[]>,
  columnRangeFilters: ColumnRangeFilters,
) {
  const column = columns.find((item) => item.columnId === columnId);
  if (!column) return [];

  const options = new Set<string>();
  rows
    .filter((row) =>
      rowMatchesFilters(
        row,
        columns,
        globalFilter,
        columnFilters,
        columnRangeFilters,
        columnId,
      ),
    )
    .forEach((row) => {
      const value = getPlainCellValue(column, row).trim();
      if (value) options.add(value);
    });

  return [...options].sort((left, right) =>
    left.localeCompare(right, "uk", { numeric: true, sensitivity: "base" }),
  );
}

function renderCell<TData>(
  column: MRT_ColumnDef<TData>,
  row: TData,
  rowIndex: number,
) {
  const value = getCellValue(column, row);

  if (column.Cell) {
    return column.Cell({
      row: { original: row, index: rowIndex },
      cell: { getValue: () => value },
    });
  }

  return value;
}

function getPlainCellValue<TData>(column: MRT_ColumnDef<TData>, row: TData) {
  const value = getCellValue(column, row);
  if (typeof value === "string" || typeof value === "number")
    return String(value);
  if (value == null || typeof value === "boolean") return String(value ?? "");
  return "";
}

function getCellValue<TData>(column: MRT_ColumnDef<TData>, row: TData) {
  if (column.accessorFn) return column.accessorFn(row);
  if (column.accessorKey) {
    return (row as Record<string, unknown>)[
      String(column.accessorKey)
    ] as ReactNode;
  }
  return null;
}

function labelText(value: ReactNode) {
  if (typeof value === "string" || typeof value === "number")
    return String(value);
  return "Колонка";
}

function normalizeFilter(value: unknown) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");
}

function parseSortDate(value: string) {
  const text = value.trim();
  const dotted = text.match(
    /^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})(?:[,\s]+(\d{1,2}):(\d{2}))?/,
  );
  if (dotted) {
    let year = Number(dotted[3]);
    if (year < 100) year += year >= 50 ? 1900 : 2000;
    const stamp = Date.UTC(
      year,
      Number(dotted[2]) - 1,
      Number(dotted[1]),
      dotted[4] ? Number(dotted[4]) : 0,
      dotted[5] ? Number(dotted[5]) : 0,
    );
    return Number.isFinite(stamp) ? stamp : null;
  }

  const iso = Date.parse(text);
  return Number.isFinite(iso) ? iso : null;
}

function compareSortValues(left: string, right: string) {
  const leftDate = parseSortDate(left);
  const rightDate = parseSortDate(right);
  if (leftDate != null && rightDate != null && leftDate !== rightDate) {
    return leftDate - rightDate;
  }

  const leftNumber = parseFilterNumber(left);
  const rightNumber = parseFilterNumber(right);
  if (leftNumber != null && rightNumber != null && leftNumber !== rightNumber) {
    return leftNumber - rightNumber;
  }

  return left.localeCompare(right, "uk", {
    numeric: true,
    sensitivity: "base",
  });
}

function sortRows<TData>(
  rows: TData[],
  columns: Array<PreparedColumn<TData>>,
  sorting: ColumnSortState | null,
) {
  if (!sorting) return rows;
  const column = columns.find((item) => item.columnId === sorting.columnId);
  if (!column) return rows;

  const direction = sorting.direction === "asc" ? 1 : -1;
  return [...rows].sort((left, right) => {
    const leftValue = getPlainCellValue(column, left).trim();
    const rightValue = getPlainCellValue(column, right).trim();
    if (!leftValue && !rightValue) return 0;
    if (!leftValue) return 1;
    if (!rightValue) return -1;
    return compareSortValues(leftValue, rightValue) * direction;
  });
}

function parseFilterNumber(value: unknown) {
  const normalized = String(value ?? "")
    .replace(/\s+/g, "")
    .replace(",", ".")
    .match(/-?\d+(?:\.\d+)?/)?.[0];
  if (!normalized) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

const MONTH_NAMES_UK = [
  "Січень",
  "Лютий",
  "Березень",
  "Квітень",
  "Травень",
  "Червень",
  "Липень",
  "Серпень",
  "Вересень",
  "Жовтень",
  "Листопад",
  "Грудень",
] as const;

const MONTH_NAMES_UK_GENITIVE = [
  "січня",
  "лютого",
  "березня",
  "квітня",
  "травня",
  "червня",
  "липня",
  "серпня",
  "вересня",
  "жовтня",
  "листопада",
  "грудня",
] as const;

type FacetMonthGroup = {
  key: string;
  label: string;
  fromLabel: string;
  dates: string[];
};

function parseFacetDate(value: string) {
  const match = String(value ?? "")
    .trim()
    .match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  let year = Number(match[3]);
  if (year < 100) year += year >= 50 ? 1900 : 2000;
  if (day < 1 || day > 31 || month < 1 || month > 12 || year < 1900) {
    return null;
  }

  return {
    day,
    month,
    year,
    key: `${year}-${String(month).padStart(2, "0")}`,
  };
}

function buildFacetMonthGroups(options: string[]): FacetMonthGroup[] {
  const dated = options
    .map((option) => ({ option, parsed: parseFacetDate(option) }))
    .filter(
      (
        item,
      ): item is {
        option: string;
        parsed: NonNullable<ReturnType<typeof parseFacetDate>>;
      } => Boolean(item.parsed),
    );
  if (dated.length < 2 || dated.length < options.length / 2) return [];

  const groups = new Map<string, FacetMonthGroup>();
  dated.forEach(({ option, parsed }) => {
    const current = groups.get(parsed.key);
    if (current) {
      current.dates.push(option);
      return;
    }
    groups.set(parsed.key, {
      key: parsed.key,
      label: `${MONTH_NAMES_UK[parsed.month - 1]} ${parsed.year}`,
      fromLabel: `Від ${MONTH_NAMES_UK_GENITIVE[parsed.month - 1]} ${parsed.year}`,
      dates: [option],
    });
  });

  return [...groups.values()].sort((left, right) =>
    right.key.localeCompare(left.key),
  );
}
