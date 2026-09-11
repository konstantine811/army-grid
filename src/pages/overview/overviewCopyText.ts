import type { BackendPersonnelOverviewRow } from "../../api";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { isBchsTreatmentStatus } from "../bchs/bchsCalc";
import { resolveOverviewPpdLocationExportRows } from "./overviewPpdLocationExport";
import { overviewStatusFilterLabel } from "./overviewRosterMerge";
import { isMorningLeaveSectionPerson } from "./overviewRotaBchsMorningExport";
import { capitalizeReportPosition } from "../documents/reportPosition";
import {
  isPlatoonCommanderPosition,
  isSectionCommanderPosition,
} from "./overviewRotaGudzExport";

const MANAGEMENT_SLOT_PATTERNS = [
  /командир\s+рот/i,
  /заступник\s+командир/i,
  /головн.*сержант/i,
  /сержант.*мат|мат\.?\s*забезп/i,
  /діловод/i,
] as const;

const uniquePeople = (rows: BackendPersonnelOverviewRow[]) => {
  const seen = new Set<string>();
  return rows.filter((row) => {
    const key = row.externalId || row.id || normalizeRosterMatchText(row.name);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

/** Коротка посада з колонки «Посада» (staff_5), без «Повної посади». */
const overviewRowPosition = (row: BackendPersonnelOverviewRow) =>
  capitalizeReportPosition(row.staffSheetColumns?.staff_5?.trim() || "");

const overviewRowPositionMatchText = (row: BackendPersonnelOverviewRow) =>
  [
    row.staffSheetColumns?.staff_5,
    row.staffSheetColumns?.staff_7,
    row.positionTitle,
  ]
    .map((value) => value?.trim())
    .filter(Boolean)
    .join(" ");

const overviewRowPlace = (row: BackendPersonnelOverviewRow) =>
  row.staffSheetColumns?.staff_31?.trim() || "";

const rotaCopySortKey = (
  row: BackendPersonnelOverviewRow,
  sourceOrder: number,
) => {
  const position = overviewRowPositionMatchText(row);
  const normalized = normalizeRosterMatchText(position);

  if (isPlatoonCommanderPosition(position)) {
    return { tier: 0, subOrder: 0, sourceOrder };
  }
  if (isSectionCommanderPosition(position)) {
    return { tier: 0, subOrder: 1, sourceOrder };
  }
  const managementIndex = MANAGEMENT_SLOT_PATTERNS.findIndex((pattern) =>
    pattern.test(normalized),
  );
  if (managementIndex >= 0) {
    return { tier: 1, subOrder: managementIndex, sourceOrder };
  }
  return { tier: 2, subOrder: sourceOrder, sourceOrder };
};

const sortOverviewRotaCopyRows = (
  rows: BackendPersonnelOverviewRow[],
) =>
  rows
    .map((row, sourceOrder) => ({ row, sourceOrder }))
    .sort((left, right) => {
      const leftKey = rotaCopySortKey(left.row, left.sourceOrder);
      const rightKey = rotaCopySortKey(right.row, right.sourceOrder);
      if (leftKey.tier !== rightKey.tier) return leftKey.tier - rightKey.tier;
      if (leftKey.subOrder !== rightKey.subOrder) {
        return leftKey.subOrder - rightKey.subOrder;
      }
      return leftKey.sourceOrder - rightKey.sourceOrder;
    })
    .map((entry) => entry.row);

/** Посада + ПІБ + місце перебування для всіх рядків поточного фільтра. */
export const buildOverviewRotaCopyText = (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
) => {
  const sortedRows = sortOverviewRotaCopyRows(context.rows);
  let lineNumber = 0;
  const lines = sortedRows.flatMap((row) => {
    const name = row.name?.trim();
    if (!name) return [];
    lineNumber += 1;
    return [
      [
        lineNumber,
        overviewRowPosition(row),
        name,
        overviewRowPlace(row),
      ].join(" - "),
    ];
  });

  return lines.join("\n");
};

const EMPTY_LOCATION_LABEL = "Без місця";

const overviewRowStatus = (row: BackendPersonnelOverviewRow) =>
  row.staffSheetColumns?.staff_21?.trim() || row.statusLabel?.trim() || "";

const isOverviewInServiceRow = (row: BackendPersonnelOverviewRow) =>
  normalizeRosterMatchText(overviewRowStatus(row)).includes("в строю");

const namedUnitPeople = (rows: BackendPersonnelOverviewRow[]) =>
  uniquePeople(rows).filter((row) => Boolean(row.name?.trim()));

const isOverviewPolygonDRow = (row: BackendPersonnelOverviewRow) => {
  const place = normalizeRosterMatchText(overviewRowPlace(row));
  return place.includes("полігон д") || place.includes("полигон д");
};

const isOverviewLeaveRow = (row: BackendPersonnelOverviewRow) =>
  isMorningLeaveSectionPerson(overviewRowStatus(row), overviewRowPlace(row));

const isOverviewTreatmentRow = (row: BackendPersonnelOverviewRow) => {
  if (isOverviewLeaveRow(row)) return false;
  const status = overviewRowStatus(row);
  const normalized = normalizeRosterMatchText(status);
  return (
    isBchsTreatmentStatus(status) ||
    /лік\s*(пор|хвор)/.test(normalized) ||
    normalized.startsWith("лік")
  );
};

export const collectOverviewInServiceRows = (
  rows: BackendPersonnelOverviewRow[],
) =>
  uniquePeople(rows).filter(
    (row) => isOverviewInServiceRow(row) && Boolean(row.name?.trim()),
  );

export const buildOverviewLocationGroups = (
  rows: BackendPersonnelOverviewRow[],
) => {
  const groups = new Map<
    string,
    { label: string; people: BackendPersonnelOverviewRow[] }
  >();
  collectOverviewInServiceRows(rows).forEach((row) => {
    const label = overviewRowPlace(row) || EMPTY_LOCATION_LABEL;
    const key = normalizeRosterMatchText(label) || EMPTY_LOCATION_LABEL;
    const existing = groups.get(key);
    if (existing) existing.people.push(row);
    else groups.set(key, { label, people: [row] });
  });
  return [...groups.values()]
    .map((group) => ({
      ...group,
      people: [...group.people].sort((left, right) =>
        left.name.localeCompare(right.name, "uk", { sensitivity: "base" }),
      ),
    }))
    .sort((left, right) => {
      if (right.people.length !== left.people.length) {
        return right.people.length - left.people.length;
      }
      return left.label.localeCompare(right.label, "uk", { sensitivity: "base" });
    });
};

/** Кількість усіх, хто в строю, по місцях — сума збігається з «В строю». */
export const buildOverviewLocationCountCopyText = (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
) => {
  const rows = resolveOverviewPpdLocationExportRows(context);
  const people = namedUnitPeople(rows);
  const groups = buildOverviewLocationGroups(rows);
  const total = groups.reduce((sum, group) => sum + group.people.length, 0);
  return [
    `В строю: ${total}`,
    ...groups.map((group) => `${group.label}: ${group.people.length}`),
    "",
    `Полігон Д: ${people.filter(isOverviewPolygonDRow).length}`,
    `Лікування: ${people.filter(isOverviewTreatmentRow).length}`,
    `Відпустка: ${people.filter(isOverviewLeaveRow).length}`,
  ].join("\n");
};

const textBlock = (label: string, rows: BackendPersonnelOverviewRow[]) => {
  const people = uniquePeople(rows);
  if (!people.length) return "";
  return [
    `${label}:`,
    ...people.map((row, index) => `${index + 1}\t${row.name.trim()}`),
  ].join("\n");
};

export const buildOverviewWhatsAppCopyText = (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
) => {
  const selectedStatuses =
    context.filters?.find((filter) => filter.id === "status")?.values ?? [];
  const statusGroups = selectedStatuses.length
    ? selectedStatuses.map((status) => ({
        label: status,
        rows: context.rows.filter(
          (row) =>
            normalizeRosterMatchText(overviewStatusFilterLabel(row)) ===
            normalizeRosterMatchText(status),
        ),
      }))
    : Array.from(
        context.rows.reduce((groups, row) => {
          const label = overviewStatusFilterLabel(row) || "Без статусу";
          const key = normalizeRosterMatchText(label);
          const group = groups.get(key);
          if (group) group.rows.push(row);
          else groups.set(key, { label, rows: [row] });
          return groups;
        }, new Map<string, { label: string; rows: BackendPersonnelOverviewRow[] }>()),
      ).map(([, group]) => group);

  const selectedUnits = new Set(
    (
      context.filters?.find((filter) => filter.id === "unit")?.values ?? []
    )
      .map(normalizeRosterMatchText)
      .filter(Boolean),
  );
  const executionRows = (context.allRows ?? context.rows).filter((row) => {
    if (
      selectedUnits.size &&
      !selectedUnits.has(normalizeRosterMatchText(row.unit))
    ) {
      return false;
    }
    return normalizeRosterMatchText(
      row.staffSheetColumns?.staff_31,
    ).includes("на виконанні");
  });

  return [
    ...statusGroups.map((group) => textBlock(group.label, group.rows)),
    textBlock("На виконанні", executionRows),
  ]
    .filter(Boolean)
    .join("\n\n");
};
