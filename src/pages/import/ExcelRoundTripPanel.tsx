import { useMemo, useState } from "react";
import {
  Alert,
  Box,
  Button,
  Chip,
  LinearProgress,
  MenuItem,
  Select,
  Stack,
  Switch,
  Typography,
} from "@/components/sci/SciPrimitives";
import { CloudUploadOutlinedIcon } from "@/components/sci/icons";
import { FileDownloadOutlinedIcon } from "@/components/sci/icons";
import { ShieldOutlinedIcon } from "@/components/sci/icons";
import type { MRT_ColumnDef } from "@/components/sci/SciDataTable";
import {
  MaterialReactTable,
  useMaterialReactTable,
} from "@/components/sci/SciDataTable";
import {
  type ExcelRow,
  getColumnHeader,
  hasRowData,
  updateCell,
  valueToDisplay,
} from "../../excelRoundTrip";
import { ExcelCellEditor } from "./ExcelCellEditor";
import type { ExcelRoundTripPanelProps } from "./types";

export function ExcelRoundTripPanel({
  snapshot,
  rows,
  setRows,
  activeSheetIndex,
  onActiveSheetChange,
  isBusy,
  message,
  mergeSummary,
  onExport,
}: ExcelRoundTripPanelProps) {
  const [wrappedColumns, setWrappedColumns] = useState<Set<number>>(
    () => new Set(),
  );
  const [pinFirstColumns, setPinFirstColumns] = useState(false);
  const visibleRows = useMemo(
    () => rows.filter((row) => hasRowData(row.values)),
    [rows],
  );
  const hasWrappedColumns = wrappedColumns.size > 0;

  const toggleColumnWrap = (columnIndex: number) => {
    setWrappedColumns((currentColumns) => {
      const nextColumns = new Set(currentColumns);
      if (nextColumns.has(columnIndex)) {
        nextColumns.delete(columnIndex);
      } else {
        nextColumns.add(columnIndex);
      }

      return nextColumns;
    });
  };

  const columns = useMemo<MRT_ColumnDef<ExcelRow>[]>(() => {
    if (!snapshot) return [];

    return Array.from({ length: snapshot.columnCount }, (_, columnIndex) => ({
      id: `column-${columnIndex}`,
      header: getColumnHeader(snapshot, columnIndex),
      Header: () => {
        const header = getColumnHeader(snapshot, columnIndex);
        const isWrapped = wrappedColumns.has(columnIndex);

        return (
          <div className="excel-header-cell" title={header}>
            <span className="excel-header-title">{header}</span>
            <Switch
              checked={isWrapped}
              color="primary"
              size="small"
              onChange={() => toggleColumnWrap(columnIndex)}
              onClick={(event) => event.stopPropagation()}
              onMouseDown={(event) => event.stopPropagation()}
              slotProps={{
                input: {
                  "aria-label": `Перенос тексту для колонки ${header}`,
                },
              }}
            />
          </div>
        );
      },
      accessorFn: (row) => valueToDisplay(row.values[columnIndex]),
      size: columnIndex === 1 ? 320 : 180,
      Cell: ({ row }) => {
        const cellValue = valueToDisplay(row.original.values[columnIndex]);
        const wrapText = wrappedColumns.has(columnIndex);

        return (
          <ExcelCellEditor
            key={`${row.original.id}-${columnIndex}-${cellValue}-${wrapText ? "wrap" : "compact"}`}
            value={cellValue}
            wrapText={wrapText}
            onCommit={(nextValue) => {
              setRows((currentRows) =>
                updateCell(
                  currentRows,
                  row.original.id,
                  columnIndex,
                  nextValue,
                ),
              );
            }}
          />
        );
      },
    }));
  }, [setRows, snapshot, wrappedColumns]);

  const table = useMaterialReactTable({
    columns,
    data: visibleRows,
    enableColumnActions: false,
    enableDensityToggle: false,
    enableFullScreenToggle: true,
    enableGlobalFilter: false,
    enableHiding: true,
    enablePagination: true,
    enableColumnPinning: true,
    enableColumnResizing: true,
    enableColumnVirtualization: !pinFirstColumns,
    enableRowVirtualization: !hasWrappedColumns,
    enableStickyHeader: true,
    columnResizeMode: "onChange",
    layoutMode: "grid",
    state: {
      columnPinning: pinFirstColumns
        ? { left: ["column-0", "column-1"] }
        : { left: [], right: [] },
    },
    initialState: {
      density: "compact",
      pagination: { pageIndex: 0, pageSize: 25 },
    },
  });

  return (
    <section className="panel table-panel">
      <div className="panel-heading">
        <span>Робоча таблиця даних</span>
        <Stack direction="row" spacing={1}>
          <Chip
            size="small"
            label={snapshot ? `${visibleRows.length} записів` : "очікує файл"}
            variant="outlined"
          />
          <Chip
            size="small"
            label="XLSX template"
            color="primary"
            variant="outlined"
          />
        </Stack>
      </div>
      {isBusy && <LinearProgress color="primary" />}
      <div className="panel-body">
        <Stack
          direction={{ xs: "column", md: "row" }}
          spacing={1.5}
          sx={{ mb: 2 }}
        >
          <Select
            className="db-input"
            disabled={!snapshot}
            value={String(activeSheetIndex)}
            onChange={(event) =>
              onActiveSheetChange(Number(event.target.value))
            }
          >
            {snapshot ? (
              snapshot.sheets.map((sheet) => (
                <MenuItem
                  value={String(sheet.sheetIndex)}
                  key={`${sheet.sheetIndex}-${sheet.sheetName}`}
                >
                  {sheet.sheetName} (
                  {sheet.rows.filter((row) => hasRowData(row.values)).length})
                </MenuItem>
              ))
            ) : (
              <MenuItem value="0">Sheet не вибрано</MenuItem>
            )}
          </Select>
          <Button
            variant="outlined"
            startIcon={<FileDownloadOutlinedIcon />}
            disabled={!snapshot}
            onClick={onExport}
          >
            Експорт у Excel зі стилями
          </Button>
          <Stack
            direction="row"
            spacing={1}
            sx={{
              alignItems: "center",
              minHeight: 38,
              px: 1.5,
              border: "1px solid rgba(214, 215, 133, 0.38)",
              color: "#d9d49d",
            }}
          >
            <Switch
              checked={pinFirstColumns}
              color="primary"
              disabled={!snapshot}
              size="small"
              onChange={(event) => setPinFirstColumns(event.target.checked)}
              slotProps={{
                input: {
                  "aria-label": "Закріпити перші колонки",
                },
              }}
            />
            <Typography variant="body2">Закріпити перші колонки</Typography>
          </Stack>
        </Stack>
        <Alert
          icon={<ShieldOutlinedIcon />}
          severity="info"
          variant="outlined"
          sx={{ mb: 2 }}
        >
          {message} {mergeSummary}
        </Alert>
        {snapshot ? (
          <MaterialReactTable table={table} />
        ) : (
          <div className="drop-zone">
            <Box>
              <CloudUploadOutlinedIcon color="disabled" />
              <Typography variant="body2">
                Завантажте `.xlsx`, щоб побачити реальні колонки файлу
              </Typography>
              <Typography variant="caption" className="muted">
                Перші 4 рядки лишаються як шапка Excel, дані редагуються з 5-го
                рядка.
              </Typography>
            </Box>
          </div>
        )}
      </div>
    </section>
  );
}
