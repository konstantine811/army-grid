import type { SciDataTableExportContext } from "@/components/sci/SciDataTable";
import type { SheetData } from "write-excel-file/browser";
import writeXlsxFile from "write-excel-file/browser";
import type { BackendPersonnelOverviewRow } from "../../api";
import { normalizeRosterMatchText } from "../personnel/fighterStatusImport";
import {
  looksLikePersonnelName,
  looksLikePersonnelRankTitle,
} from "../personnel/personnelUtils";
import { resolveOverviewPpdLocationExportRows } from "./overviewPpdLocationExport";
import { overviewStatusFilterLabel } from "./overviewRosterMerge";
import {
  isPlatoonCommanderPosition,
  isSectionCommanderPosition,
  resolveRotaGudzCallsign,
  resolveRotaGudzPersonName,
  resolveSelectedOverviewUnit,
} from "./overviewRotaGudzExport";

const EXPORT_COLUMNS = [
  { id: "index", label: "№", width: 6 },
  { id: "role", label: "Роль", width: 24 },
  { id: "rank", label: "Звання", width: 16 },
  { id: "name", label: "ПІБ", width: 36 },
  { id: "callsign", label: "Позивний", width: 14 },
  { id: "position", label: "Посада", width: 28 },
  { id: "status", label: "Статус", width: 18 },
  { id: "location", label: "Місце перебування", width: 24 },
] as const;

type CommanderExportCell = (typeof EXPORT_COLUMNS)[number]["id"];

export type OverviewCommanderRole =
  | "Командир роти"
  | "Командир взводу"
  | "Командир відділення";

export type OverviewCommanderExportRow = {
  role: OverviewCommanderRole;
  unit: string;
  platoon: string;
  section: string;
  rank: string;
  name: string;
  callsign: string;
  position: string;
  status: string;
  location: string;
};

export type OverviewCommanderSectionGroup = {
  title: string;
  commanders: OverviewCommanderExportRow[];
};

export type OverviewCommanderPlatoonGroup = {
  title: string;
  commanders: OverviewCommanderExportRow[];
  sections: OverviewCommanderSectionGroup[];
};

export type OverviewCommanderUnitGroup = {
  unit: string;
  companyCommanders: OverviewCommanderExportRow[];
  platoons: OverviewCommanderPlatoonGroup[];
};

const ROLE_ORDER: Record<OverviewCommanderRole, number> = {
  "Командир роти": 0,
  "Командир взводу": 1,
  "Командир відділення": 2,
};

const staffValue = (row: BackendPersonnelOverviewRow, columnNumber: number) =>
  row.staffSheetColumns?.[`staff_${columnNumber}`]?.trim() ?? "";

const normalizePosition = (value: string) => normalizeRosterMatchText(value);

const commanderPositionMatchText = (row: BackendPersonnelOverviewRow) =>
  [staffValue(row, 5), staffValue(row, 7), row.positionTitle ?? ""]
    .map((value) => value.trim())
    .filter(Boolean)
    .join(" ");

export const isDeputyCommanderPosition = (position: string) =>
  /заступник/.test(normalizePosition(position));

export const isCompanyCommanderPosition = (position: string) => {
  if (isDeputyCommanderPosition(position)) return false;
  return /командир\s+(?:\d+\s*)?(?:піхотн\S*\s+)?рот(?:а|и|і)(?:\s|$)/.test(
    normalizePosition(position),
  );
};

const isCommanderPlatoonPosition = (position: string) => {
  if (isDeputyCommanderPosition(position)) return false;
  if (isPlatoonCommanderPosition(position)) return true;
  return /командир(?:\s+\S+)*\s+взвод/.test(normalizePosition(position));
};

const isCommanderSectionPosition = (position: string) => {
  if (isDeputyCommanderPosition(position)) return false;
  return (
    isSectionCommanderPosition(position) ||
    /командир\s+\d+\s*відділен/i.test(position)
  );
};

export const resolveOverviewCommanderRole = (
  row: BackendPersonnelOverviewRow,
): OverviewCommanderRole | null => {
  const position = commanderPositionMatchText(row);
  if (isCompanyCommanderPosition(position)) return "Командир роти";
  if (isCommanderPlatoonPosition(position)) return "Командир взводу";
  if (isCommanderSectionPosition(position)) return "Командир відділення";
  return null;
};

const sanitizeRank = (value: string) => {
  const text = value.trim();
  if (!text || /^\d+$/.test(text)) return "";
  if (looksLikePersonnelRankTitle(text)) return text;
  if (looksLikePersonnelName(text)) return "";
  return "";
};

const resolveCommanderRank = (row: BackendPersonnelOverviewRow) =>
  sanitizeRank(staffValue(row, 13)) || sanitizeRank(row.rank?.trim() || "");

const resolveCommanderName = (row: BackendPersonnelOverviewRow) =>
  resolveRotaGudzPersonName(row) || row.name?.trim() || "";

const resolveCommanderStatus = (row: BackendPersonnelOverviewRow) =>
  staffValue(row, 21) || overviewStatusFilterLabel(row) || "";

const isDeadCommanderStatus = (status: string) =>
  /загибл|загинув|^200$/.test(normalizeRosterMatchText(status));

const resolveSectionTitle = (sectionRaw: string, positionRaw: string) => {
  const fromPosition = positionRaw.match(/(\d+)\s*відділен/i)?.[1];
  const fromSection = sectionRaw.match(/(\d+)/)?.[1];
  const number = fromPosition || fromSection;
  if (number) return `${number} відділення`;
  return sectionRaw.trim() || "Без відділення";
};

const resolveDisplayPosition = (
  row: BackendPersonnelOverviewRow,
  role: OverviewCommanderRole,
) => {
  const position = staffValue(row, 5);
  if (role !== "Командир відділення") return position;
  if (/\d/.test(position) && /відділен/i.test(position)) return position;
  const sectionNumber = staffValue(row, 4).match(/(\d+)/)?.[1];
  return sectionNumber ? `Командир ${sectionNumber} відділення` : position;
};

const compareUk = (left: string, right: string) =>
  left.localeCompare(right, "uk", { numeric: true, sensitivity: "base" });

const extractLeadingNumber = (value: string) => {
  const match = value.match(/(\d+)/);
  return match ? Number(match[1]) : null;
};

const isManagementPlatoon = (platoon: string) => {
  const normalized = normalizeRosterMatchText(platoon);
  return normalized === "ж" || normalized.includes("управлін");
};

const comparePlatoonTitles = (left: string, right: string) => {
  const leftManagement = isManagementPlatoon(left);
  const rightManagement = isManagementPlatoon(right);
  if (leftManagement !== rightManagement) return leftManagement ? -1 : 1;
  const leftNumber = extractLeadingNumber(left);
  const rightNumber = extractLeadingNumber(right);
  if (leftNumber != null && rightNumber != null && leftNumber !== rightNumber) {
    return leftNumber - rightNumber;
  }
  return compareUk(left, right);
};

export const buildOverviewCommanderExportRows = (
  rows: BackendPersonnelOverviewRow[],
): OverviewCommanderExportRow[] => {
  const seen = new Set<string>();
  const commanders: OverviewCommanderExportRow[] = [];

  rows.forEach((row) => {
    const role = resolveOverviewCommanderRole(row);
    if (!role) return;
    const name = resolveCommanderName(row);
    if (!name || !looksLikePersonnelName(name)) return;
    const status = resolveCommanderStatus(row);
    if (isDeadCommanderStatus(status)) return;
    const key = row.externalId || row.id || normalizeRosterMatchText(name);
    if (seen.has(key)) return;
    seen.add(key);
    commanders.push({
      role,
      unit: row.unit?.trim() || "",
      platoon: staffValue(row, 3),
      section: resolveSectionTitle(staffValue(row, 4), staffValue(row, 5)),
      rank: resolveCommanderRank(row),
      name,
      callsign: resolveRotaGudzCallsign(row),
      position: resolveDisplayPosition(row, role),
      status,
      location: staffValue(row, 31),
    });
  });

  return commanders.sort((left, right) => {
    const unitDelta = compareUk(left.unit, right.unit);
    if (unitDelta !== 0) return unitDelta;
    const roleDelta = ROLE_ORDER[left.role] - ROLE_ORDER[right.role];
    if (roleDelta !== 0) return roleDelta;
    const platoonDelta = comparePlatoonTitles(left.platoon, right.platoon);
    if (platoonDelta !== 0) return platoonDelta;
    const sectionDelta = compareUk(left.section, right.section);
    if (sectionDelta !== 0) return sectionDelta;
    return compareUk(left.name, right.name);
  });
};

const platoonGroupTitle = (platoon: string) => {
  const title = platoon.trim();
  if (!title) return "Без взводу";
  if (isManagementPlatoon(title)) return "Управління";
  return title;
};

export const groupOverviewCommandersByUnit = (
  commanders: OverviewCommanderExportRow[],
): OverviewCommanderUnitGroup[] => {
  const byUnit = new Map<string, OverviewCommanderExportRow[]>();
  commanders.forEach((commander) => {
    const unit = commander.unit || "Без підрозділу";
    const list = byUnit.get(unit);
    if (list) list.push(commander);
    else byUnit.set(unit, [commander]);
  });

  return [...byUnit.entries()]
    .sort(([left], [right]) => compareUk(left, right))
    .map(([unit, people]) => {
      const companyCommanders = people.filter(
        (person) => person.role === "Командир роти",
      );
      const platoonPeople = people.filter(
        (person) => person.role !== "Командир роти",
      );
      const byPlatoon = new Map<string, OverviewCommanderExportRow[]>();
      platoonPeople.forEach((person) => {
        const title = platoonGroupTitle(person.platoon);
        const list = byPlatoon.get(title);
        if (list) list.push(person);
        else byPlatoon.set(title, [person]);
      });

      const platoons = [...byPlatoon.entries()]
        .sort(([left], [right]) => comparePlatoonTitles(left, right))
        .map(([title, members]) => {
          const commandersInPlatoon = members.filter(
            (person) => person.role === "Командир взводу",
          );
          const bySection = new Map<string, OverviewCommanderExportRow[]>();
          members
            .filter((person) => person.role === "Командир відділення")
            .forEach((person) => {
              const sectionTitle = person.section || "Без відділення";
              const list = bySection.get(sectionTitle);
              if (list) list.push(person);
              else bySection.set(sectionTitle, [person]);
            });
          return {
            title,
            commanders: commandersInPlatoon,
            sections: [...bySection.entries()]
              .sort(([left], [right]) => compareUk(left, right))
              .map(([sectionTitle, sectionCommanders]) => ({
                title: sectionTitle,
                commanders: sectionCommanders,
              })),
          };
        });

      return { unit, companyCommanders, platoons };
    });
};

type SheetCell = {
  value: string | number;
  fontWeight?: "bold";
  fontSize: number;
  align: "left" | "center";
  alignVertical: "center";
  wrap: true;
  height: number;
  backgroundColor: string;
  textColor?: string;
  borderColor: string;
  borderStyle: "thin";
};

const columnHeaderRow = (): SheetCell[] =>
  EXPORT_COLUMNS.map((column) => ({
    value: column.label,
    fontWeight: "bold",
    fontSize: 10,
    align: "center",
    alignVertical: "center",
    wrap: true,
    height: 30,
    backgroundColor: "#39735C",
    textColor: "#FFFFFF",
    borderColor: "#53605A",
    borderStyle: "thin",
  }));

const groupHeaderRow = (title: string, backgroundColor: string): SheetCell[] =>
  EXPORT_COLUMNS.map((column, index) => ({
    value: index === 0 ? title : "",
    fontWeight: "bold",
    fontSize: 11,
    align: "left",
    alignVertical: "center",
    wrap: true,
    height: 26,
    backgroundColor,
    textColor: "#FFFFFF",
    borderColor: "#53605A",
    borderStyle: "thin",
  }));

const personRow = (
  commander: OverviewCommanderExportRow,
  index: number,
): SheetCell[] => {
  const values: Record<CommanderExportCell, string | number> = {
    index,
    role: commander.role,
    rank: commander.rank,
    name: commander.name,
    callsign: commander.callsign,
    position: commander.position,
    status: commander.status,
    location: commander.location,
  };
  return EXPORT_COLUMNS.map((column) => ({
    value: values[column.id],
    fontSize: 10,
    align: column.id === "name" || column.id === "position" ? "left" : "center",
    alignVertical: "center",
    wrap: true,
    height: 24,
    backgroundColor: index % 2 ? "#F7F9F8" : "#FFFFFF",
    borderColor: "#7D8983",
    borderStyle: "thin",
  }));
};

const appendPeople = (
  rows: SheetData,
  people: OverviewCommanderExportRow[],
  startIndex: number,
) => {
  let index = startIndex;
  people.forEach((person) => {
    rows.push(personRow(person, index));
    index += 1;
  });
  return index;
};

const spacerRow = (): SheetCell[] =>
  EXPORT_COLUMNS.map(() => ({
    value: "",
    fontSize: 10,
    align: "left",
    alignVertical: "center",
    wrap: true,
    height: 12,
    backgroundColor: "#FFFFFF",
    borderColor: "#FFFFFF",
    borderStyle: "thin",
  }));

const appendUnitTable = (rows: SheetData, group: OverviewCommanderUnitGroup) => {
  rows.push(groupHeaderRow(group.unit, "#1F4A3D"));
  rows.push(columnHeaderRow());
  let index = 1;
  if (group.companyCommanders.length) {
    rows.push(groupHeaderRow("Командування роти", "#2F5D4E"));
    index = appendPeople(rows, group.companyCommanders, index);
  }
  group.platoons.forEach((platoon) => {
    rows.push(groupHeaderRow(platoon.title, "#3D7A68"));
    index = appendPeople(rows, platoon.commanders, index);
    platoon.sections.forEach((section) => {
      rows.push(groupHeaderRow(section.title, "#6A9A88"));
      index = appendPeople(rows, section.commanders, index);
    });
  });
};

export const buildOverviewCommandersSheetData = (
  groups: OverviewCommanderUnitGroup[],
): SheetData => {
  const rows: SheetData = [];
  groups.forEach((group, groupIndex) => {
    if (groupIndex > 0) rows.push(spacerRow());
    appendUnitTable(rows, group);
  });
  return rows;
};

export const buildOverviewCommandersExportSheets = (
  commanders: OverviewCommanderExportRow[],
) => [
  {
    sheet: "Командири",
    columns: EXPORT_COLUMNS.map((column) => ({ width: column.width })),
    stickyRowsCount: 0,
    showGridLines: true,
    orientation: "landscape" as const,
    data: buildOverviewCommandersSheetData(
      groupOverviewCommandersByUnit(commanders),
    ),
  },
];

export const buildOverviewCommandersExportFileName = (
  unitLabel?: string | null,
  reportDate = new Date(),
) => {
  const dateLabel = new Intl.DateTimeFormat("uk-UA", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(reportDate);
  const unit = unitLabel?.replace(/\s+/g, " ").trim();
  return unit
    ? `Командири ${unit} ${dateLabel}.xlsx`
    : `Командири ${dateLabel}.xlsx`;
};

export const exportOverviewCommandersReport = async (
  context: SciDataTableExportContext<BackendPersonnelOverviewRow>,
  preferredRows?: BackendPersonnelOverviewRow[],
) => {
  const sourceRows = resolveOverviewPpdLocationExportRows(context, preferredRows);
  const commanders = buildOverviewCommanderExportRows(sourceRows);
  if (!commanders.length) {
    throw new Error("Немає командирів рот, взводів або відділень для експорту.");
  }

  await writeXlsxFile(buildOverviewCommandersExportSheets(commanders), {
    fontFamily: "Arial",
    fontSize: 10,
  }).toFile(
    buildOverviewCommandersExportFileName(
      resolveSelectedOverviewUnit(context.filters),
    ),
  );

  return { commanders, count: commanders.length };
};
