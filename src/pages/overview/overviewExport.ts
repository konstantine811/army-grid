import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import type { SheetData } from "write-excel-file/browser";
import type { BackendPersonnelOverviewRow } from "../../api";
import {
  extractBchsAwayPeopleFromDbRows,
  hasBchsFullName,
  isBchsAwolStatus,
  isBchsBusinessTripStatus,
  isBchsDetachedStatus,
  isBchsKilledStatus,
  isBchsMissingStatus,
  isBchsTrainingStatus,
  isBchsTreatmentStatus,
  isBchsVacationStatus,
  normalizeBchsText,
} from "../bchs/bchsCalc";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { canonicalName, dateMs } from "../ejournal/ejoosIdentity";
import type { PbArchivePeriod } from "../ejournal/ejoosParsers";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { overviewStatusFilterLabel } from "./overviewRosterMerge";
import { OVERVIEW_STAFF_COLUMN_HEADERS } from "./overviewStaffColumns";

export type OverviewExportFilters = {
  unit?: string;
  status?: string;
  activeFilters?: Array<{
    id: string;
    label: string;
    values: string[];
  }>;
};

const EXPORT_COLORS = {
  title: "#234F3E",
  header: "#39735C",
  stripe: "#F1F7F4",
  border: "#C8D8D1",
  text: "#17231E",
} as const;

const CENTERED_COLUMN_IDS = new Set([
  "status",
  "questionnaire",
  "documents",
  "fighterExitDate",
  "fighterReturnDate",
  "fighterTotalDays",
  "fighterStatus",
  "updatedAt",
]);

const overviewExportColumnWidth = (columnId: string) => {
  if (columnId === "person") return 30;
  if (columnId === "unit") return 22;
  if (columnId === "rank") return 16;
  if (columnId === "positionTitle") return 42;
  if (columnId === "status") return 16;
  if (columnId === "questionnaire") return 12;
  if (columnId === "documents") return 16;
  if (columnId === "fighterDirection") return 22;
  if (columnId === "fighterExitDate") return 15;
  if (columnId === "fighterReturnDate") return 17;
  if (columnId === "fighterTotalDays") return 10;
  if (columnId === "fighterStatus") return 16;
  if (columnId === "updatedAt") return 18;
  if (/^staff_\d+$/.test(columnId)) return 18;
  return 20;
};

const IMPORTANT_OVERVIEW_COLUMNS = [
  { id: "staff_5", label: "Посада", width: 28 },
  { id: "staff_8", label: "ШПК факт", width: 14 },
  { id: "person", label: "ПІБ", width: 30 },
  { id: "staff_15", label: "Позивний", width: 15 },
  { id: "staff_22", label: "Тип В\\С", width: 14 },
  { id: "staff_23", label: "Статус БГ", width: 15 },
  { id: "staff_27", label: "Відрядження (БРЕЗ)", width: 20 },
  { id: "staff_28", label: "Обмеження", width: 18 },
  { id: "staff_31", label: "Місце перебування", width: 24 },
  { id: "staff_32", label: "Примітки", width: 28 },
  { id: "staff_33", label: "Напрямок", width: 20 },
] as const;

const importantOverviewValue = (
  row: BackendPersonnelOverviewRow,
  columnId: (typeof IMPORTANT_OVERVIEW_COLUMNS)[number]["id"],
) =>
  columnId === "person"
    ? row.name
    : row.staffSheetColumns?.[columnId]?.trim() || "";

const importantHeaderColor = (index: number) => {
  if (index <= 3) return "#F4C20D";
  if (index <= 5) return "#70AD47";
  return "#8EA9DB";
};

export const buildImportantOverviewExportTitle = (
  filters: OverviewExportFilters,
) => {
  if (filters.activeFilters?.length) {
    return filters.activeFilters
      .map((filter) => `${filter.label}: ${filter.values.join(", ")}`)
      .join(" · ");
  }
  if (filters.unit != null || filters.status != null) {
    return `${formatOverviewExportFilterValue("unit", filters.unit ?? "ALL")} — ${formatOverviewExportFilterValue("status", filters.status ?? "ALL")}`;
  }
  return "Без додаткових фільтрів";
};

export const buildImportantOverviewExportSheetData = (
  rows: BackendPersonnelOverviewRow[],
  filters: OverviewExportFilters,
): SheetData => [
  [
    {
      value: buildImportantOverviewExportTitle(filters),
      fontWeight: "bold" as const,
      fontSize: 16,
      align: "center" as const,
      alignVertical: "center" as const,
      backgroundColor: "#EAF1EE",
      textColor: EXPORT_COLORS.title,
      height: 32,
      columnSpan: IMPORTANT_OVERVIEW_COLUMNS.length,
    },
    ...Array.from({ length: IMPORTANT_OVERVIEW_COLUMNS.length - 1 }, () => null),
  ],
  IMPORTANT_OVERVIEW_COLUMNS.map((column, index) => ({
    value: column.label,
    fontWeight: "bold" as const,
    fontSize: 10,
    align: "center" as const,
    alignVertical: "center" as const,
    wrap: true,
    height: 36,
    backgroundColor: importantHeaderColor(index),
    borderColor: "#53605A",
    borderStyle: "thin" as const,
  })),
  ...rows.map((row, rowIndex) =>
    IMPORTANT_OVERVIEW_COLUMNS.map((column) => ({
      value: importantOverviewValue(row, column.id),
      fontSize: 10,
      align: "center" as const,
      alignVertical: "center" as const,
      wrap: true,
      height: 32,
      backgroundColor: rowIndex % 2 ? "#F7F9F8" : "#FFFFFF",
      borderColor: "#7D8983",
      borderStyle: "thin" as const,
    })),
  ),
];

export const buildImportantOverviewExportSheetOptions = () => ({
  columns: IMPORTANT_OVERVIEW_COLUMNS.map((column) => ({
    width: column.width,
  })),
  stickyRowsCount: 2,
  stickyColumnsCount: 3,
  showGridLines: false,
  orientation: "landscape" as const,
});

export type ImportantOverviewSummary = {
  unitLabel: string;
  staff: number;
  listed: number;
  trainingTrip: number;
  detached: number;
  treatment: number;
  vacation: number;
  awol: number;
  missing: number;
  killed: number;
  wounded300: number;
  absent: number;
  management: number;
  support: number;
  platoon: number;
  attached: number;
  onExit: number;
  battleReady: number;
  available: number;
  inRanks: number;
};

export const buildImportantOverviewSummary = (
  rosterRows: EjournalPreviewRow[],
  filters: OverviewExportFilters,
  archivePeriods: PbArchivePeriod[] = [],
): ImportantOverviewSummary => {
  const unitValues =
    filters.activeFilters?.find((filter) => filter.id === "unit")?.values ?? [];
  const selectedUnits = new Set(
    unitValues.map(normalizeRosterMatchText).filter(Boolean),
  );
  const people = extractBchsAwayPeopleFromDbRows(rosterRows).filter(
    (person) =>
      !selectedUnits.size ||
      selectedUnits.has(normalizeRosterMatchText(person.rosterUnit)),
  );
  const listedPeople = people.filter(hasBchsFullName);
  const count = (matches: (person: (typeof listedPeople)[number]) => boolean) =>
    listedPeople.filter(matches).length;
  const placeContains = (
    person: (typeof listedPeople)[number],
    value: string,
  ) => normalizeBchsText(person.medicalPlace).includes(value);

  const trainingTrip = count(
    (person) =>
      isBchsTrainingStatus(person.status) ||
      isBchsBusinessTripStatus(person.status),
  );
  const detached = count((person) => isBchsDetachedStatus(person.status));
  const treatment = count((person) => isBchsTreatmentStatus(person.status));
  const vacation = count((person) => isBchsVacationStatus(person.status));
  const awol = count((person) => isBchsAwolStatus(person.status));
  const missing = count((person) => isBchsMissingStatus(person.status));
  const killed = count((person) => isBchsKilledStatus(person.status));
  const listedNames = new Set(
    listedPeople.map((person) => canonicalName(person.fullName)).filter(Boolean),
  );
  const archiveMatches = (matches: (text: string) => boolean) => {
    const matchedNames = new Set<string>();
    for (const period of archivePeriods) {
      if (dateMs(period.returnDate)) continue;
      const name = canonicalName(period.fullName);
      if (!name || !listedNames.has(name)) continue;
      const text = normalizeBchsText(
        [period.absenceType, period.place].filter(Boolean).join(" "),
      );
      if (matches(text)) matchedNames.add(name);
    }
    return matchedNames.size;
  };
  const wounded300 = archiveMatches(
    (text) =>
      /(?:^|\D)300(?:\D|$)/.test(text) ||
      text.includes("по поран") ||
      (text.includes("ліку") && text.includes("поран")),
  );
  const absent =
    trainingTrip + detached + treatment + vacation + awol + missing + killed;

  return {
    unitLabel: unitValues.join(", ") || "1ПБ",
    staff: people.length,
    listed: listedPeople.length,
    trainingTrip,
    detached,
    treatment,
    vacation,
    awol,
    missing,
    killed,
    wounded300,
    absent,
    management: count((person) =>
      normalizeBchsText(person.roleType).includes("упр"),
    ),
    support: count((person) =>
      normalizeBchsText(person.roleType).includes("забезпеч"),
    ),
    platoon: count((person) => placeContains(person, "взвод")),
    attached: count(
      (person) =>
        normalizeBchsText(person.status).includes("приком") ||
        placeContains(person, "приком"),
    ),
    onExit: count((person) => placeContains(person, "на виконанні")),
    battleReady: count(
      (person) =>
        normalizeBchsText(person.combatReadiness) === "бг" ||
        normalizeBchsText(person.medicalPlace) === "бг",
    ),
    available: count((person) =>
      normalizeBchsText(person.status).includes("в строю"),
    ),
    inRanks: Math.max(0, listedPeople.length - absent),
  };
};

const summaryCell = (
  value: string | number,
  backgroundColor = "#FFFFFF",
  columnSpan?: number,
) => ({
  value,
  fontWeight: "bold" as const,
  fontSize: 11,
  align: columnSpan ? ("center" as const) : ("left" as const),
  alignVertical: "center" as const,
  backgroundColor,
  borderColor: "#111111",
  borderStyle: "thin" as const,
  height: 24,
  ...(columnSpan ? { columnSpan } : {}),
});

export const buildImportantOverviewSummarySheetData = (
  summary: ImportantOverviewSummary,
): SheetData => {
  const row = (label: string, value: number, color = "#FFFFFF") => [
    summaryCell(label, color),
    { ...summaryCell(value, color), align: "center" as const },
  ];
  return [
    [summaryCell(summary.unitLabel, "#FFFFFF", 2), null],
    row("За штатом", summary.staff),
    row("Всього по списку", summary.listed),
    [summaryCell("З них відсутні:", "#FFFFFF", 2), null],
    row("Навчання/відрядження", summary.trainingTrip, "#B4C7E7"),
    row("Відкомандировані", summary.detached),
    row("Лікування", summary.treatment, "#F4B6F4"),
    row("300", summary.wounded300, "#F4B6F4"),
    row("Відпустка/Лікувальна відпустка", summary.vacation, "#F4B6F4"),
    row("СЗЧ", summary.awol, "#C05AE6"),
    row("Зниклі безвісти", summary.missing, "#F4A261"),
    row("Загиблі", summary.killed, "#D00000"),
    row("Всього відсутніх:", summary.absent),
    [summaryCell("З них в строю:", "#FFFFFF", 2), null],
    row("Управління", summary.management, "#FF1F1F"),
    row("Забезпечення", summary.support, "#11BFEF"),
    row("Взводні (Наявність)", summary.platoon),
    row("Прикомандировані", summary.attached),
    row("На виході", summary.onExit, "#FFC000"),
    row("БГ", summary.battleReady, "#00B050"),
    row("В наявності", summary.available),
    row("Всього в строю:", summary.inRanks),
  ];
};

const importantExportSheetName = (value: string, index: number) => {
  const cleaned = value.replace(/[\\/*?:[\]]/g, " ").replace(/\s+/g, " ").trim();
  return (cleaned || `Статус ${index + 1}`).slice(0, 31);
};

export const buildImportantOverviewExportSheets = (
  rows: BackendPersonnelOverviewRow[],
  filters: OverviewExportFilters,
  allRows: BackendPersonnelOverviewRow[] = rows,
  rosterRows: EjournalPreviewRow[] = [],
  archivePeriods: PbArchivePeriod[] = [],
) => {
  const statusFilter = filters.activeFilters?.find(
    (filter) => filter.id === "status",
  );
  const unitFilter = filters.activeFilters?.find(
    (filter) => filter.id === "unit",
  );
  const selectedStatuses = statusFilter?.values.filter(Boolean) ?? [];
  const selectedUnits = new Set(
    (unitFilter?.values ?? []).map(normalizeRosterMatchText).filter(Boolean),
  );
  const unitRows = selectedUnits.size
    ? allRows.filter((row) =>
        selectedUnits.has(normalizeRosterMatchText(row.unit)),
      )
    : allRows;
  const statuses = selectedStatuses.length > 1 ? selectedStatuses : [""];
  const options = buildImportantOverviewExportSheetOptions();

  const statusSheets = statuses.map((selectedStatus, index) => {
    const normalizedStatus = normalizeRosterMatchText(selectedStatus);
    const sheetRows = selectedStatus
      ? unitRows.filter(
          (row) =>
            normalizeRosterMatchText(overviewStatusFilterLabel(row)) ===
            normalizedStatus,
        )
      : rows;
    const activeFilters = filters.activeFilters?.map((filter) =>
      filter.id === "status" && selectedStatus
        ? { ...filter, values: [selectedStatus] }
        : filter,
    );

    return {
      ...options,
      sheet: importantExportSheetName(
        selectedStatus || statusFilter?.values[0] || "Важливі",
        index,
      ),
      data: buildImportantOverviewExportSheetData(sheetRows, {
        ...filters,
        activeFilters,
      }),
    };
  });

  const executionRows = unitRows.filter((row) =>
    normalizeRosterMatchText(
      row.staffSheetColumns?.staff_31,
    ).includes("на виконанні"),
  );
  const executionFilters = [
    ...(unitFilter ? [unitFilter] : []),
    {
      id: "staff_31",
      label: "Місце перебування",
      values: ["На виконанні"],
    },
  ];
  const executionSheet = {
    ...options,
    sheet: "На виконанні",
    data: buildImportantOverviewExportSheetData(executionRows, {
      activeFilters: executionFilters,
    }),
  };

  const summarySheet = rosterRows.length
    ? [
        {
          sheet: "Підрахунок",
          data: buildImportantOverviewSummarySheetData(
            buildImportantOverviewSummary(rosterRows, filters, archivePeriods),
          ),
          columns: [{ width: 42 }, { width: 18 }],
          stickyRowsCount: 1,
          showGridLines: false,
        },
      ]
    : [];

  return [...summarySheet, ...statusSheets, executionSheet];
};

const exportFileNamePart = (value: string) =>
  value
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

export const buildImportantOverviewExportFileName = (
  filters: OverviewExportFilters,
  exportedAt: Date = new Date(),
) => {
  const unitValues =
    filters.activeFilters?.find((filter) => filter.id === "unit")?.values ?? [];
  const statusValues =
    filters.activeFilters?.find((filter) => filter.id === "status")?.values ?? [];
  const unitPart = unitValues.length
    ? unitValues.map(exportFileNamePart).join(", ")
    : "Усі роти";
  const statusPart = statusValues.length
    ? statusValues.map(exportFileNamePart).join(", ")
    : "Усі статуси";
  const datePart = [
    String(exportedAt.getDate()).padStart(2, "0"),
    String(exportedAt.getMonth() + 1).padStart(2, "0"),
    exportedAt.getFullYear(),
  ].join(".");

  return `1ПБ-${unitPart}-Статуси-${statusPart}-На виконанні-${datePart}.xlsx`;
};

export const buildOverviewExportSheetOptions = (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
) => ({
  columns: context.columns.map((column) => ({
    width: overviewExportColumnWidth(column.id),
  })),
  stickyRowsCount: 2,
  stickyColumnsCount: context.columns.length ? 1 : 0,
  showGridLines: false,
});

export const formatOverviewExportFilterValue = (
  kind: "unit" | "status",
  value: string,
) => {
  if (kind === "unit" && value === "ALL") return "Усі підрозділи";
  if (kind === "status" && value === "ALL") return "Усі статуси";
  return value.trim() || "—";
};

export const buildOverviewExportFilterHeaderText = (
  filters: OverviewExportFilters,
) => {
  if (filters.activeFilters?.length) {
    return filters.activeFilters
      .map((filter) => `${filter.label}: ${filter.values.join(", ")}`)
      .join(" · ");
  }
  return `${OVERVIEW_STAFF_COLUMN_HEADERS.unit}: ${formatOverviewExportFilterValue("unit", filters.unit ?? "ALL")} · ${OVERVIEW_STAFF_COLUMN_HEADERS.status}: ${formatOverviewExportFilterValue("status", filters.status ?? "ALL")}`;
};

export const buildOverviewExportSheetData = (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  filters: OverviewExportFilters,
): SheetData => {
  const columnCount = Math.max(context.columns.length, 1);
  const filterRow = [
    {
      value: buildOverviewExportFilterHeaderText(filters),
      fontWeight: "bold" as const,
      fontSize: 12,
      textColor: "#FFFFFF",
      backgroundColor: EXPORT_COLORS.title,
      align: "left" as const,
      alignVertical: "center" as const,
      wrap: true,
      height: 28,
      columnSpan: columnCount,
    },
    ...Array.from({ length: columnCount - 1 }, () => null),
  ];

  return [
    filterRow,
    context.columns.map((column) => ({
      value: column.label,
      fontWeight: "bold" as const,
      fontSize: 10,
      textColor: "#FFFFFF",
      backgroundColor: EXPORT_COLORS.header,
      borderColor: EXPORT_COLORS.border,
      borderStyle: "thin" as const,
      align: "center" as const,
      alignVertical: "center" as const,
      wrap: true,
      height: 34,
    })),
    ...context.rows.map((row, rowIndex) =>
      context.columns.map((column) => ({
        value: column.value(row),
        fontSize: 10,
        textColor: EXPORT_COLORS.text,
        backgroundColor: rowIndex % 2 ? EXPORT_COLORS.stripe : "#FFFFFF",
        borderColor: EXPORT_COLORS.border,
        borderStyle: "thin" as const,
        align: CENTERED_COLUMN_IDS.has(column.id)
          ? ("center" as const)
          : ("left" as const),
        alignVertical: "top" as const,
        wrap: true,
        height: 30,
      })),
    ),
  ];
};
