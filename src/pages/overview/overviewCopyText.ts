import type { BackendPersonnelOverviewRow } from "../../api";
import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import { overviewStatusFilterLabel } from "./overviewRosterMerge";

const uniquePeople = (rows: BackendPersonnelOverviewRow[]) => {
  const seen = new Set<string>();
  return rows.filter((row) => {
    const key = row.externalId || row.id || normalizeRosterMatchText(row.name);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
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
