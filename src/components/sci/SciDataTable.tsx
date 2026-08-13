import type { ReactNode } from "react";

export type MRT_ColumnDef<TData> = {
  id?: string;
  accessorKey?: keyof TData | string;
  header?: ReactNode;
  Header?: () => ReactNode;
  accessorFn?: (row: TData) => ReactNode;
  Cell?: (props: {
    row: { original: TData; index: number };
    cell: { getValue: () => unknown };
  }) => ReactNode;
  size?: number;
  muiEditTextFieldProps?: unknown;
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
  };
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

export function useMaterialReactTable<TData>(options: SciTableOptions<TData>) {
  return options;
}

export function MaterialReactTable<TData>({
  table,
}: {
  table: SciTableOptions<TData>;
}) {
  const pageSize = table.initialState?.pagination?.pageSize ?? 200;
  const rows = table.data.slice(0, pageSize);

  return (
    <div className="sci-data-table-wrap">
      <table className="sci-data-table">
        <thead>
          <tr>
            {table.columns.map((column, index) => (
              <th
                key={column.id ?? index}
                style={{ width: column.size, minWidth: column.size }}
              >
                {column.Header ? column.Header() : column.header}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, rowIndex) => (
            <tr key={rowIndex}>
              {table.columns.map((column, columnIndex) => (
                <td
                  key={column.id ?? columnIndex}
                  style={{ width: column.size, minWidth: column.size }}
                >
                  {renderCell(column, row, rowIndex)}
                </td>
              ))}
            </tr>
          ))}
          {rows.length === 0 && (
            <tr>
              <td colSpan={Math.max(table.columns.length, 1)} className="sci-data-table-empty">
                Даних немає
              </td>
            </tr>
          )}
        </tbody>
      </table>
      {table.data.length > rows.length && (
        <div className="sci-data-table-footer">
          Показано {rows.length} з {table.data.length}
        </div>
      )}
    </div>
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

function getCellValue<TData>(column: MRT_ColumnDef<TData>, row: TData) {
  if (column.accessorFn) return column.accessorFn(row);
  if (column.accessorKey) {
    return (row as Record<string, unknown>)[String(column.accessorKey)] as ReactNode;
  }
  return null;
}
