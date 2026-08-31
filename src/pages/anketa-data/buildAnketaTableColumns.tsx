import type {
  MRT_ColumnDef,
  SciDataTableCellRef,
} from "@/components/sci/SciDataTable";
import {
  ANKETA_COLUMNS,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";
import { AnketaCellEditor } from "./AnketaCellEditor";
import { anketaCellA1, type AnketaEmptyCell } from "./anketaGaps";

type BuildAnketaTableColumnsOptions = {
  focusedEmpty: AnketaEmptyCell | null;
  focusEpoch: number;
  emptySearchActive: boolean;
  onActivateCell: (cell: AnketaEmptyCell) => void;
  onCancelEdit: () => void;
  onPatchCell: (
    rowId: string,
    columnId: AnketaColumnKey,
    value: string,
    options?: { advance?: boolean },
  ) => void;
};

export function buildAnketaTableColumns({
  focusedEmpty,
  focusEpoch,
  emptySearchActive,
  onActivateCell,
  onCancelEdit,
  onPatchCell,
}: BuildAnketaTableColumnsOptions): Array<MRT_ColumnDef<AnketaRow>> {
  return [
    {
      id: "__rowNumber",
      accessorFn: (row) => row.__rowNumber,
      header: "№",
      size: 64,
      pin: "left",
      enableColumnFilter: false,
      enableGlobalFilter: false,
      enableHiding: false,
      Cell: ({ row }) => (
        <span className="anketa-row-number">{row.original.__rowNumber}</span>
      ),
      exportValue: (row) => String(row.__rowNumber),
    },
    ...ANKETA_COLUMNS.map((column): MRT_ColumnDef<AnketaRow> => ({
      id: column.key,
      accessorKey: column.key,
      header: column.header,
      size: column.size,
      pin: column.pin,
      filterVariant: column.filterVariant,
      enableColumnFilter: true,
      enableGlobalFilter: true,
      Cell: ({ row }) => {
        const rowId = row.original.__rowId;
        const value = String(row.original[column.key] ?? "");
        const isFocused =
          focusedEmpty?.rowId === rowId &&
          focusedEmpty?.columnId === column.key;
        const isEmpty = !value.trim();
        const isReadonly = Boolean(column.readonly);

        if (isReadonly) {
          return (
            <span className="anketa-cell-readonly" title="ПІБ з Google Sheet — не редагується">
              {value || "—"}
            </span>
          );
        }

        if (isFocused) {
          return (
            <AnketaCellEditor
              key={`${rowId}:${column.key}:${focusEpoch}`}
              columnKey={column.key}
              columnHeader={column.header}
              rowNumber={row.original.__rowNumber}
              value={value}
              isEmpty={isEmpty}
              advanceOnSave={emptySearchActive}
              onCancel={onCancelEdit}
              onSave={(next, advance) => {
                onPatchCell(rowId, column.key, next, {
                  advance: advance || emptySearchActive,
                });
              }}
            />
          );
        }

        return (
          <button
            type="button"
            className={
              isEmpty ? "anketa-cell-button is-empty" : "anketa-cell-button"
            }
            onClick={() => {
              const columnIndex = ANKETA_COLUMNS.findIndex(
                (item) => item.key === column.key,
              );
              onActivateCell({
                rowId,
                columnId: column.key,
                rowNumber: row.original.__rowNumber,
                columnIndex: Math.max(columnIndex, 0),
                header: column.header,
                a1: anketaCellA1(
                  row.original.__rowNumber,
                  Math.max(columnIndex, 0),
                ),
              });
            }}
          >
            {value || "—"}
          </button>
        );
      },
    })),
  ];
}

export function buildAnketaFocusedCell(
  focusedEmpty: AnketaEmptyCell | null,
  focusEpoch: number,
): SciDataTableCellRef | null {
  if (!focusedEmpty) return null;
  return {
    rowId: focusedEmpty.rowId,
    columnId: focusedEmpty.columnId,
    focusEpoch,
  };
}

export function buildAnketaTdProps(
  gapKeySet: Set<AnketaColumnKey>,
  focusedEmpty: AnketaEmptyCell | null,
) {
  return ({
    rowId,
    columnId,
    row,
  }: {
    rowId: string;
    columnId: string;
    row: AnketaRow;
  }) => {
    if (columnId === "__rowNumber") {
      return { className: "anketa-td-row-number" };
    }
    const key = columnId as AnketaColumnKey;
    const empty = !String(row[key] ?? "").trim();
    const inGapScope = gapKeySet.has(key);
    const focused =
      focusedEmpty?.rowId === rowId && focusedEmpty?.columnId === columnId;
    return {
      className: [
        empty && inGapScope ? "anketa-td-empty" : "",
        focused ? "is-focused-empty-cell" : "",
      ]
        .filter(Boolean)
        .join(" "),
    };
  };
}
