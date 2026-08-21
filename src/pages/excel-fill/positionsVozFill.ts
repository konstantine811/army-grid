import type { CellValue, ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { valueToDisplay } from "../../excelRoundTrip";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  resolveMorningGeneralListColumnLabel,
  resolvePersonRankTitle,
} from "../personnel/personnelUtils";

export type PositionsVozSourceKind =
  | "rank"
  | "name"
  | "callsign"
  | "vos"
  | "position"
  | "fullPosition"
  | "positionGroup"
  | "roleType"
  | "customKeys";

export type PositionsVozFillRule = {
  id: string;
  enabled: boolean;
  /** 1-based Excel column in template. */
  targetColumn: number;
  label: string;
  source: PositionsVozSourceKind;
  /** For source = customKeys: try these roster keys in order. */
  customKeys?: string[];
  onlyIfEmpty?: boolean;
};

export type PositionsVozPerson = {
  name: string;
  rank: string;
  callsign: string;
  vos: string;
  position: string;
  fullPosition: string;
  roleType: string;
  positionGroup: string;
  raw: EjournalPreviewRow;
};

export type PositionsVozChange = {
  rowNumber: number;
  column: number;
  label: string;
  person: string;
  from: string;
  to: string;
};

export type PositionsVozFillResult = {
  changes: PositionsVozChange[];
  matchedRows: number;
  unmatchedTemplateRows: number;
  unmatchedDbPeople: number;
  dbPeople: number;
  templatePeople: number;
};

export const POSITIONS_VOZ_RULES_STORAGE_KEY = "army-grid.positions-voz.fill-rules.v4";

export const DEFAULT_POSITIONS_VOZ_RULES: PositionsVozFillRule[] = [
  {
    id: "rank",
    enabled: true,
    targetColumn: 3,
    label: "Звання",
    source: "rank",
    onlyIfEmpty: false,
  },
  {
    id: "name",
    enabled: true,
    targetColumn: 4,
    label: "ПІБ",
    source: "name",
    onlyIfEmpty: true,
  },
  {
    id: "callsign",
    enabled: true,
    targetColumn: 5,
    label: "Позивний",
    source: "callsign",
    onlyIfEmpty: false,
  },
  {
    id: "fullPosition",
    enabled: true,
    targetColumn: 8,
    label: "Повна посада",
    source: "fullPosition",
    onlyIfEmpty: true,
  },
  {
    id: "vos",
    enabled: true,
    targetColumn: 9,
    label: "ВОС",
    source: "vos",
    onlyIfEmpty: true,
  },
  {
    id: "positionGroup",
    enabled: false,
    targetColumn: 6,
    label: "Група посади",
    source: "positionGroup",
    onlyIfEmpty: true,
  },
];

const normalizeName = (value: string) =>
  valueToDisplay(value as CellValue)
    .replace(/[ʼ’']/g, "")
    .replace(/\([^)]*\)/g, " ")
    .replace(/[.,;:№#"/\\|()[\]{}]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLocaleLowerCase("uk-UA");

const cellText = (value: unknown) => valueToDisplay(value as CellValue).trim();

/** Prefer exact key match so "посада" does not steal "повна_посада". */
const pickRowValue = (row: EjournalPreviewRow, keys: string[], fuzzy = true) => {
  const entries = Object.entries(row).filter(([key]) => !key.startsWith("__"));
  const norms = keys.map((wanted) => normalizeName(wanted).replace(/\s+/g, "_"));

  for (const wantedNorm of norms) {
    for (const [key, value] of entries) {
      const keyNorm = normalizeName(key).replace(/\s+/g, "_");
      if (keyNorm === wantedNorm || keyNorm === `roster__${wantedNorm}`) {
        const text = cellText(value);
        if (text) return text;
      }
    }
  }

  if (!fuzzy) return "";

  for (const wantedNorm of norms) {
    if (wantedNorm.length < 4) continue;
    for (const [key, value] of entries) {
      const keyNorm = normalizeName(key).replace(/\s+/g, "_");
      // Do not let short "посада" match "повна_посада".
      if (wantedNorm === "посада" && keyNorm.includes("повна")) continue;
      if (keyNorm.includes(wantedNorm) || wantedNorm.includes(keyNorm)) {
        const text = cellText(value);
        if (text) return text;
      }
    }
  }
  return "";
};

const getSheetCellByExcelColumn = (
  sheet: { columnIndexes: number[] },
  values: CellValue[],
  excelColumn1Based: number,
) => {
  const originalIndex = excelColumn1Based - 1;
  const compactIndex = sheet.columnIndexes.indexOf(originalIndex);
  if (compactIndex < 0) return "";
  return cellText(values[compactIndex]);
};

const derivePositionGroup = (position: string, roleType: string, fullPosition: string) => {
  const text = normalizeName(`${position} ${roleType} ${fullPosition}`);
  if (/охорон/.test(text)) return "ОХОРОНА";
  if (/мех|вод|водій|екіпаж/.test(text)) return "МЕХ-ВОД";
  if (/зв.?яз|радіо|реб/.test(text)) return "ЗВ'ЯЗОК";
  if (/піхот|стріл|гранат|кулемет/.test(text)) return "ПІХОТА";
  if (/упр|штаб|команд/.test(text)) return "УПРАВЛІННЯ";
  if (/медик|санітар/.test(text)) return "МЕДИКИ";
  if (/забезпеч|кухар|майстер|вмтз/.test(text)) return "ЗАБЕЗПЕЧЕННЯ";
  if (/брез/.test(text)) return "БРЕЗ";
  return roleType || position || "";
};

/** Empty or date / arrival-note junk in position cells. */
const isUnusablePositionValue = (value: string) => {
  const text = value.replace(/\s+/g, " ").trim();
  if (!text) return true;
  if (/T\d{2}:\d{2}:\d{2}/.test(text)) return true;
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) return true;
  if (/\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}/.test(text)) return true;
  return false;
};

const isPlaceholderUnit = (value: string) => {
  const text = normalizeName(value);
  return !text || text === "ж" || text === "-" || text === "—";
};

export const POSITIONS_VOZ_ABSENT_LABEL = "відсутній у списку";

/**
 * ПОВНА ПОСАДА → посада → підрозділ (коли посади немає або там дата;
 * у підрозділі часто «БРЕЗ» / «ПРИКОМАНДИРОВАНІ»).
 * Якщо нічого немає взагалі → «відсутній у списку».
 */
const resolveFillPosition = (
  fullPosition: string,
  position: string,
  unit: string,
) => {
  if (!isUnusablePositionValue(fullPosition)) return fullPosition.trim();
  if (!isUnusablePositionValue(position)) return position.trim();
  if (!isPlaceholderUnit(unit)) return unit.trim();
  return POSITIONS_VOZ_ABSENT_LABEL;
};

export const rosterRowsToPositionsPeople = (
  rows: EjournalPreviewRow[],
): PositionsVozPerson[] => {
  const people: PositionsVozPerson[] = [];
  const seen = new Set<string>();

  for (const row of rows) {
    const name =
      pickRowValue(row, ["піб", "прізвище", "column_14"]) ||
      cellText(row["ПІБ"]) ||
      cellText(row["прізвище"]);
    if (!name || name.length < 3) continue;
    const key = normalizeName(name);
    if (!key || seen.has(key)) continue;
    seen.add(key);

    const position = pickRowValue(row, ["посада"]);
    const rawFullPosition = pickRowValue(
      row,
      ["повна_посада", "повна посада", "повнапосада"],
      false,
    ) || pickRowValue(row, ["column_7"], false);
    const unit = pickRowValue(row, ["підрозділ"]);
    const fullPosition = resolveFillPosition(rawFullPosition, position, unit);
    const roleType = pickRowValue(row, ["тип в", "тип_в", "тип_вс", "column_22"]);
    const vos = pickRowValue(row, ["вос", "воз", "column_6"]);
    const callsign = pickRowValue(row, ["позивний", "позив", "column_15"]);
    const rank =
      resolvePersonRankTitle(row) ||
      pickRowValue(row, ["звання", "шпк_факт", "column_13"]);

    people.push({
      name,
      rank,
      callsign,
      vos,
      position,
      fullPosition,
      roleType,
      positionGroup: derivePositionGroup(position, roleType, fullPosition),
      raw: row,
    });
  }

  return people;
};

export const parseRosterLatestToPeople = (latest: {
  sheet: { columns?: unknown } | null;
  rows: Array<{ id?: string; excelRowNumber?: number | null; values?: unknown }>;
}): PositionsVozPerson[] => {
  if (!latest.sheet) return [];
  const rows = latest.rows.map((row) => ({
    __dbRowId: row.id,
    __rowNumber: row.excelRowNumber ?? undefined,
    ...(row.values && typeof row.values === "object" && !Array.isArray(row.values)
      ? (row.values as Record<string, unknown>)
      : {}),
  })) as EjournalPreviewRow[];

  return rosterRowsToPositionsPeople(rows);
};

const sourceValue = (person: PositionsVozPerson, rule: PositionsVozFillRule) => {
  switch (rule.source) {
    case "rank":
      return person.rank;
    case "name":
      return person.name;
    case "callsign":
      return person.callsign;
    case "vos":
      return person.vos;
    case "position":
      return person.position;
    case "fullPosition":
      return person.fullPosition;
    case "positionGroup":
      return person.positionGroup;
    case "roleType":
      return person.roleType;
    case "customKeys":
      return pickRowValue(person.raw, rule.customKeys ?? []);
    default:
      return "";
  }
};

export const loadPositionsVozRules = (): PositionsVozFillRule[] => {
  try {
    const raw = localStorage.getItem(POSITIONS_VOZ_RULES_STORAGE_KEY);
    if (!raw) return structuredClone(DEFAULT_POSITIONS_VOZ_RULES);
    const parsed = JSON.parse(raw) as PositionsVozFillRule[];
    if (!Array.isArray(parsed) || !parsed.length) {
      return structuredClone(DEFAULT_POSITIONS_VOZ_RULES);
    }
    return parsed.map((rule) => ({
      ...rule,
      enabled: Boolean(rule.enabled),
      targetColumn: Number(rule.targetColumn) || 1,
      onlyIfEmpty: Boolean(rule.onlyIfEmpty),
      customKeys: Array.isArray(rule.customKeys) ? rule.customKeys : [],
    }));
  } catch {
    return structuredClone(DEFAULT_POSITIONS_VOZ_RULES);
  }
};

export const savePositionsVozRules = (rules: PositionsVozFillRule[]) => {
  localStorage.setItem(POSITIONS_VOZ_RULES_STORAGE_KEY, JSON.stringify(rules, null, 2));
};

export const analyzePositionsVozFill = (
  template: ExcelWorkbookSnapshot | null,
  people: PositionsVozPerson[],
  rules: PositionsVozFillRule[],
): PositionsVozFillResult | null => {
  if (!template) return null;
  const sheet = template.sheets[0];
  if (!sheet) return null;

  const byName = new Map<string, PositionsVozPerson>();
  for (const person of people) {
    const key = normalizeName(person.name);
    if (key && !byName.has(key)) byName.set(key, person);
  }

  const enabledRules = rules.filter((rule) => rule.enabled && rule.targetColumn > 0);
  const changes: PositionsVozChange[] = [];
  let matchedRows = 0;
  let unmatchedTemplateRows = 0;
  const matchedNames = new Set<string>();

  // Prefer rawRows: this template has a title row, not a real header row.
  // Generic Excel snapshot header-detection can skip / rematerialize columns.
  const scanRows =
    sheet.rawRows.length > 0
      ? sheet.rawRows.map((values, index) => ({
          excelRowNumber: index + 1,
          values,
        }))
      : sheet.rows;

  for (const excelRow of scanRows) {
    if (excelRow.excelRowNumber < 2) continue;
    const name = getSheetCellByExcelColumn(sheet, excelRow.values, 4);
    if (!name || name.length < 3) continue;
    if (/посади|воз|вос/i.test(name)) continue;

    const person = byName.get(normalizeName(name));
    if (!person) {
      unmatchedTemplateRows += 1;
      for (const rule of enabledRules) {
        if (rule.source !== "fullPosition") continue;
        const current = getSheetCellByExcelColumn(
          sheet,
          excelRow.values,
          rule.targetColumn,
        );
        if (rule.onlyIfEmpty && current) continue;
        if (normalizeName(current) === normalizeName(POSITIONS_VOZ_ABSENT_LABEL)) {
          continue;
        }
        changes.push({
          rowNumber: excelRow.excelRowNumber,
          column: rule.targetColumn,
          label: rule.label || "Повна посада",
          person: name,
          from: current,
          to: POSITIONS_VOZ_ABSENT_LABEL,
        });
      }
      continue;
    }

    matchedRows += 1;
    matchedNames.add(normalizeName(person.name));

    for (const rule of enabledRules) {
      const next = sourceValue(person, rule).trim();
      if (!next) continue;
      const current = getSheetCellByExcelColumn(
        sheet,
        excelRow.values,
        rule.targetColumn,
      );
      if (rule.onlyIfEmpty && current) continue;
      if (normalizeName(current) === normalizeName(next)) continue;
      changes.push({
        rowNumber: excelRow.excelRowNumber,
        column: rule.targetColumn,
        label:
          rule.label ||
          resolveMorningGeneralListColumnLabel(`column_${rule.targetColumn}`),
        person: person.name,
        from: current,
        to: next,
      });
    }
  }

  return {
    changes,
    matchedRows,
    unmatchedTemplateRows,
    unmatchedDbPeople: Math.max(0, people.length - matchedNames.size),
    dbPeople: people.length,
    templatePeople: matchedRows + unmatchedTemplateRows,
  };
};
