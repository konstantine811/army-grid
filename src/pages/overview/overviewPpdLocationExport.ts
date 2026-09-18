import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import type { SheetData } from "write-excel-file/browser";
import writeXlsxFile from "write-excel-file/browser";
import type { BackendPersonnelOverviewRow } from "../../api";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import {
  overviewStatusFilterLabel,
  rosterRowMatchesOverviewUnit,
} from "./overviewRosterMerge";
import {
  resolveRotaGudzCallsign,
  resolveRotaGudzPersonName,
  resolveSelectedOverviewUnit,
} from "./overviewRotaGudzExport";

const EXPORT_COLUMNS = [
  { id: "index", label: "№", width: 6 },
  { id: "name", label: "ПІБ", width: 32 },
  { id: "callsign", label: "Позивний", width: 14 },
  { id: "status", label: "Статус", width: 16 },
  { id: "bg", label: "БГ", width: 12 },
  { id: "notes", label: "Примітки", width: 28 },
] as const;

type PpdLocationExportCell = (typeof EXPORT_COLUMNS)[number]["id"];
type LocationColumn = 31 | 35 | 40;

const PPD_VYSHNEVE = "ппд вишневе";
const PPD_DETAIL = "ппд";

const staffValue = (row: BackendPersonnelOverviewRow, columnNumber: number) =>
  row.staffSheetColumns?.[`staff_${columnNumber}`]?.trim() ?? "";

const normalizedLocationValue = (value: string) =>
  normalizeRosterMatchText(value);

export const isOverviewPpdVyshneveBase = (value: string) =>
  normalizedLocationValue(value) === PPD_VYSHNEVE;

export const isOverviewPpdDetailValue = (value: string) => {
  const normalized = normalizedLocationValue(value);
  return !normalized || normalized === PPD_DETAIL;
};

/** Кол. 35 «Місце перебування (уточнення)» = ПОЛІГОН. */
export const isOverviewPolygonDetailValue = (value: string) => {
  const normalized = normalizedLocationValue(value);
  if (!normalized) return false;
  return normalized.includes("полігон") || normalized.includes("полигон");
};

export type OverviewPpdLocationParts = {
  base: string;
  detail: string;
  baseColumn: 31 | null;
  detailColumn: 35 | 40 | null;
};

const readOverviewLocationColumn = (
  overviewRow: BackendPersonnelOverviewRow,
  rosterRow: EjournalPreviewRow | null | undefined,
  columnNumber: LocationColumn,
) => {
  if (rosterRow) {
    return readRosterColumnValue(rosterRow, columnNumber);
  }
  return staffValue(overviewRow, columnNumber);
};

const resolveOverviewPpdLocationDetail = (
  overviewRow: BackendPersonnelOverviewRow,
  rosterRow: EjournalPreviewRow | null | undefined,
  base: string,
) => {
  const detail35 = readOverviewLocationColumn(overviewRow, rosterRow, 35);
  if (detail35) {
    return { detail: detail35, detailColumn: 35 as const };
  }
  const detail40 = readOverviewLocationColumn(overviewRow, rosterRow, 40);
  if (
    detail40 &&
    normalizedLocationValue(detail40) !== normalizedLocationValue(base)
  ) {
    return { detail: detail40, detailColumn: 40 as const };
  }
  return { detail: "", detailColumn: 35 as const };
};

/** AE (кол.31) + AI (кол.35 «уточнення»). */
export const resolveOverviewPpdLocationParts = (
  row: BackendPersonnelOverviewRow,
  rosterRow?: EjournalPreviewRow | null,
): OverviewPpdLocationParts => {
  const base = readOverviewLocationColumn(row, rosterRow, 31);
  if (!isOverviewPpdVyshneveBase(base)) {
    return { base: "", detail: "", baseColumn: null, detailColumn: null };
  }
  const { detail, detailColumn } = resolveOverviewPpdLocationDetail(
    row,
    rosterRow,
    base,
  );
  return {
    base,
    detail,
    baseColumn: 31,
    detailColumn,
  };
};

/** Аркуш 1: кол.31 = ППД Вишневе, кол.35 порожня або «ППД» (не полігон). */
export const isOverviewPpdVyshneveLocation = (
  row: BackendPersonnelOverviewRow,
  rosterRow?: EjournalPreviewRow | null,
) => {
  const { base, detail } = resolveOverviewPpdLocationParts(row, rosterRow);
  if (!isOverviewPpdVyshneveBase(base)) return false;
  if (isOverviewPolygonDetailValue(detail)) return false;
  return isOverviewPpdDetailValue(detail);
};

/** Аркуш 2: кол.31 = ППД Вишневе, кол.35 = «ПОЛІГОН». */
export const isOverviewPolygonLocation = (
  row: BackendPersonnelOverviewRow,
  rosterRow?: EjournalPreviewRow | null,
) => {
  const { base, detail } = resolveOverviewPpdLocationParts(row, rosterRow);
  return isOverviewPpdVyshneveBase(base) && isOverviewPolygonDetailValue(detail);
};

const resolveExportPersonName = (row: BackendPersonnelOverviewRow) =>
  resolveRotaGudzPersonName(row) ||
  row.name?.trim() ||
  staffValue(row, 14);

const hasOverviewPersonName = (row: BackendPersonnelOverviewRow) =>
  Boolean(resolveExportPersonName(row));

export type OverviewPpdLocationRosterLookup = (
  row: BackendPersonnelOverviewRow,
) => EjournalPreviewRow | undefined;

export const buildOverviewPpdLocationRosterLookup = (
  rosterRows: EjournalPreviewRow[] | null | undefined,
): OverviewPpdLocationRosterLookup | undefined => {
  if (!rosterRows?.length) return undefined;

  const byKey = new Map<string, EjournalPreviewRow>();
  const byName = new Map<string, EjournalPreviewRow>();

  for (const rosterRow of rosterRows) {
    const dbRowId = String(rosterRow.__dbRowId ?? "").trim();
    if (dbRowId) {
      byKey.set(dbRowId, rosterRow);
      byKey.set(`roster:row:${dbRowId}`, rosterRow);
    }
    const rowNumber = String(rosterRow.__rowNumber ?? "").trim();
    if (rowNumber) {
      byKey.set(`google:${rowNumber}`, rosterRow);
      byKey.set(`roster:row:google:${rowNumber}`, rosterRow);
    }
    const name = normalizeRosterMatchText(readRosterColumnValue(rosterRow, 14));
    if (name) byName.set(name, rosterRow);
  }

  return (overviewRow) => {
    const candidates = [
      overviewRow.id,
      overviewRow.externalId,
      overviewRow.id.replace(/^roster:row:/, ""),
    ]
      .map((value) => String(value ?? "").trim())
      .filter(Boolean);
    for (const key of candidates) {
      const match = byKey.get(key);
      if (match) return match;
    }
    const nameKey = normalizeRosterMatchText(overviewRow.name);
    return nameKey ? byName.get(nameKey) : undefined;
  };
};

export const filterOverviewPpdVyshneveRows = (
  rows: BackendPersonnelOverviewRow[],
  rosterLookup?: OverviewPpdLocationRosterLookup,
) =>
  rows.filter((row) => {
    const rosterRow = rosterLookup?.(row);
    return (
      hasOverviewPersonName(row) &&
      isOverviewPpdVyshneveLocation(row, rosterRow)
    );
  });

export const filterOverviewPolygonRows = (
  rows: BackendPersonnelOverviewRow[],
  rosterLookup?: OverviewPpdLocationRosterLookup,
) =>
  rows.filter((row) => {
    const rosterRow = rosterLookup?.(row);
    return hasOverviewPersonName(row) && isOverviewPolygonLocation(row, rosterRow);
  });

const extractOverviewRotaNumber = (value: string) => {
  const match = normalizedLocationValue(value).match(/(\d+)\D*рот/u);
  return match?.[1] ?? "";
};

export const matchesOverviewRotaUnitFilter = (
  rowUnit: string,
  filterUnit: string,
) => {
  if (normalizedLocationValue(rowUnit) === normalizedLocationValue(filterUnit)) {
    return true;
  }
  if (rosterRowMatchesOverviewUnit(rowUnit, [filterUnit])) return true;
  const rowNumber = extractOverviewRotaNumber(rowUnit);
  const filterNumber = extractOverviewRotaNumber(filterUnit);
  return Boolean(rowNumber && filterNumber && rowNumber === filterNumber);
};

/** Усі рядки обраної роти з allRows (не лише видимі в таблиці). */
export const resolveOverviewPpdLocationExportRows = (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  preferredRows?: BackendPersonnelOverviewRow[],
) => {
  const unitValues =
    context.filters
      ?.find((filter) => filter.id === "unit")
      ?.values.filter(Boolean) ?? [];
  const sourceRows =
    preferredRows ?? context.allRows ?? context.rows ?? [];
  if (!unitValues.length) return sourceRows;

  return sourceRows.filter((row) =>
    unitValues.some((unit) => matchesOverviewRotaUnitFilter(row.unit, unit)),
  );
};

const resolveExportStatus = (row: BackendPersonnelOverviewRow) =>
  staffValue(row, 21) || overviewStatusFilterLabel(row) || "";

const DUTY_STATUS_PREFIX_RE = /^\s*в\s*строю\s*/i;

const resolveExportBg = (row: BackendPersonnelOverviewRow) => {
  const fromStatusBg = (value: string) =>
    value.replace(DUTY_STATUS_PREFIX_RE, "").replace(/\s+/g, " ").trim();
  return fromStatusBg(staffValue(row, 23)) || fromStatusBg(staffValue(row, 42));
};

const buildPpdLocationExportRow = (
  row: BackendPersonnelOverviewRow,
  index: number,
): Record<PpdLocationExportCell, string> => ({
  index: String(index),
  name: resolveExportPersonName(row),
  callsign: resolveRotaGudzCallsign(row),
  status: resolveExportStatus(row),
  bg: resolveExportBg(row),
  notes: staffValue(row, 32),
});

const buildPpdLocationTitleRow = (title: string): SheetData[number] => [
  {
    value: title,
    fontWeight: "bold" as const,
    fontSize: 16,
    align: "center" as const,
    alignVertical: "center" as const,
    height: 32,
    backgroundColor: "#EAF1EE",
    textColor: "#1F3D34",
    columnSpan: EXPORT_COLUMNS.length,
  },
  ...Array.from({ length: EXPORT_COLUMNS.length - 1 }, () => null),
];

const buildPpdLocationSheetData = (
  title: string,
  rows: BackendPersonnelOverviewRow[],
): SheetData => [
  buildPpdLocationTitleRow(title),
  EXPORT_COLUMNS.map((column) => ({
    value: column.label,
    fontWeight: "bold" as const,
    fontSize: 10,
    align: "center" as const,
    alignVertical: "center" as const,
    wrap: true,
    height: 30,
    backgroundColor: "#39735C",
    textColor: "#FFFFFF",
    borderColor: "#53605A",
    borderStyle: "thin" as const,
  })),
  ...rows.map((row, rowIndex) => {
    const values = buildPpdLocationExportRow(row, rowIndex + 1);
    return EXPORT_COLUMNS.map((column) => ({
      value: values[column.id],
      fontSize: 10,
      align:
        column.id === "name" || column.id === "notes"
          ? ("left" as const)
          : ("center" as const),
      alignVertical: "center" as const,
      wrap: true,
      height: 28,
      backgroundColor: rowIndex % 2 ? "#F7F9F8" : "#FFFFFF",
      borderColor: "#7D8983",
      borderStyle: "thin" as const,
    }));
  }),
];

export const buildOverviewPpdLocationExportSheets = (
  rows: BackendPersonnelOverviewRow[],
  rosterRows?: EjournalPreviewRow[] | null,
) => {
  const rosterLookup = buildOverviewPpdLocationRosterLookup(rosterRows);
  const sheetOptions = {
    columns: EXPORT_COLUMNS.map((column) => ({ width: column.width })),
    stickyRowsCount: 2,
    showGridLines: true,
    orientation: "landscape" as const,
  };

  return [
    {
      ...sheetOptions,
      sheet: "ППД Вишневе",
      data: buildPpdLocationSheetData(
        "ППД Вишневе",
        filterOverviewPpdVyshneveRows(rows, rosterLookup),
      ),
    },
    {
      ...sheetOptions,
      sheet: "Полігон",
      data: buildPpdLocationSheetData(
        "Полігон",
        filterOverviewPolygonRows(rows, rosterLookup),
      ),
    },
  ];
};

export const buildOverviewPpdLocationExportFileName = (
  unitLabel?: string,
  exportedAt = new Date().toISOString().slice(0, 10),
) => {
  const unit = unitLabel?.replace(/\s+/g, " ").trim();
  return unit
    ? `ППД Вишневе ${unit} ${exportedAt}.xlsx`
    : `ППД Вишневе ${exportedAt}.xlsx`;
};

export const exportOverviewPpdLocationReport = async (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  preferredRows?: BackendPersonnelOverviewRow[],
  rosterRows?: EjournalPreviewRow[] | null,
) => {
  const rows = resolveOverviewPpdLocationExportRows(context, preferredRows);
  const unitLabel = resolveSelectedOverviewUnit(context.filters);
  const rosterLookup = buildOverviewPpdLocationRosterLookup(rosterRows);

  await writeXlsxFile(buildOverviewPpdLocationExportSheets(rows, rosterRows), {
    fontFamily: "Arial",
    fontSize: 10,
  }).toFile(buildOverviewPpdLocationExportFileName(unitLabel ?? undefined));

  return {
    rows,
    ppdCount: filterOverviewPpdVyshneveRows(rows, rosterLookup).length,
    polygonCount: filterOverviewPolygonRows(rows, rosterLookup).length,
  };
};
