import { hasRowData, type ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { cellValueToJson } from "../../shared/format";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import {
  getRosterPersonBirthDate,
  getRosterPersonName,
  getRosterPersonRnokpp,
} from "./personnelRosterMerge";
import { normalizeRosterMatchText } from "./fighterStatusImport";
import { looksLikePersonnelName } from "./personnelUtils";
import {
  ROSTER_ARCHIVE_FLAG_KEY,
  ROSTER_ARCHIVE_SOURCE_KEY,
  ROSTER_ARCHIVE_SOURCE_VALUE,
} from "./staffSheetArchiveMarker";

export {
  isPersonnelFromArchive,
  ROSTER_ARCHIVE_FLAG_KEY,
  ROSTER_ARCHIVE_SOURCE_KEY,
  ROSTER_ARCHIVE_SOURCE_VALUE,
} from "./staffSheetArchiveMarker";

/** Offset so archive excelRowNumber never collides with «Загальний список». */
export const ARCHIVE_ROSTER_EXCEL_ROW_OFFSET = 900_000;

type WorkbookSheet = ExcelWorkbookSnapshot["sheets"][number];

/**
 * Archive sheet column (1-based) → «Загальний список» column.
 * Layout differs: ПІБ is M(13), birth N(14), callsign P(16).
 */
const ARCHIVE_TO_ROSTER_COLUMN: ReadonlyArray<
  [archiveCol: number, rosterCol: number]
> = [
  [2, 2], // Рота → Підрозділ
  [3, 3],
  [4, 4],
  [5, 5],
  [6, 6],
  [8, 8],
  [10, 10],
  [11, 13], // Звання
  [13, 14], // ПІБ
  [14, 16], // Дата народження
  [15, 18], // Повних років
  [16, 15], // Позивний (in archive often unlabeled)
  [18, 21], // СТАТУС
  [19, 22],
  [20, 23],
  [21, 24],
  [22, 26], // КУРС БЗВП
  [23, 27], // Відрядження (БРЕЗ)
  [24, 28], // Обмеження / додаткова ознака
  [25, 29], // Підпорядкованість
  [26, 31], // Місце перебування
  [27, 32], // Примітки
  [28, 33], // Напрямок
  [29, 34], // Примітка 3
];

export const findArchiveSheet = (
  sheets: WorkbookSheet[],
): WorkbookSheet | undefined =>
  sheets.find((sheet) => /^архів$/i.test(sheet.sheetName.trim()));

const cellText = (values: unknown[], index0: number) =>
  String(values[index0] ?? "").trim();

const isLikelyBirthValue = (value: unknown) => {
  if (value instanceof Date) return true;
  if (typeof value === "number") return value >= 15_000 && value <= 60_000;
  const text = String(value ?? "").trim();
  return (
    /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(text) ||
    /^\d{4}-\d{1,2}-\d{1,2}$/.test(text)
  );
};

const pickRank = (values: unknown[]) => {
  const rankPattern =
    /(солдат|сержант|старшин|прапорщик|лейтенант|капітан|майор|полковник|генерал)/i;
  return [12, 10, 6, 11, 7]
    .map((index0) => cellText(values, index0))
    .find((value) => rankPattern.test(value));
};

const setNamedRosterFields = (
  row: EjournalPreviewRow,
  values: unknown[],
  name: string,
  nameColumn: 13 | 14,
) => {
  row.ПІБ = name;
  row.піб = name;
  row.column_14 = name;

  const rank =
    nameColumn === 13 ? cellText(values, 10) : pickRank(values) ?? "";
  if (rank) {
    row.column_13 = rank;
    row.звання = rank;
    row.Звання = rank;
  }

  const callSign =
    nameColumn === 13
      ? cellText(values, 15)
      : [values[14], values[16]]
          .filter((value) => !isLikelyBirthValue(value))
          .map((value) => String(value ?? "").trim())
          .find(Boolean) ?? "";
  if (callSign) {
    row.column_15 = callSign;
    row.позивний = callSign;
    row.Позивний = callSign;
  }

  const birth =
    nameColumn === 13
      ? values[13]
      : [values[14], values[15]].find(isLikelyBirthValue);
  if (birth != null && String(birth).trim()) {
    row.column_16 = cellValueToJson(birth);
  }

  const status = cellText(values, nameColumn === 13 ? 17 : 20);
  if (status) {
    row.column_21 = status;
    row.статус = status;
    row.СТАТУС = status;
  }
};

const normalizeArchiveRowToRosterShape = (
  sheetRow: WorkbookSheet["rows"][number],
): EjournalPreviewRow | null => {
  const values = sheetRow.values;
  const nameInM = cellText(values, 12);
  const nameInN = cellText(values, 13);
  const nameColumn = looksLikePersonnelName(nameInM)
    ? 13
    : looksLikePersonnelName(nameInN)
      ? 14
      : 0;
  if (!nameColumn) return null;
  const name = nameColumn === 13 ? nameInM : nameInN;

  const row: EjournalPreviewRow = {
    __rowNumber: sheetRow.excelRowNumber,
    [ROSTER_ARCHIVE_FLAG_KEY]: true,
    [ROSTER_ARCHIVE_SOURCE_KEY]: ROSTER_ARCHIVE_SOURCE_VALUE,
  };

  if (nameColumn === 13) {
    for (const [archiveCol, rosterCol] of ARCHIVE_TO_ROSTER_COLUMN) {
      const raw = values[archiveCol - 1];
      if (raw == null || String(raw).trim() === "") continue;
      row[`column_${rosterCol}`] = cellValueToJson(raw);
    }
  } else {
    // Older archive blocks largely preserve their original roster columns.
    values.forEach((raw, index0) => {
      if (raw == null || String(raw).trim() === "") return;
      row[`column_${index0 + 1}`] = cellValueToJson(raw);
    });
  }

  setNamedRosterFields(row, values, name, nameColumn);
  return row;
};

export const parseArchiveSheetToRosterRows = (
  sheet: WorkbookSheet,
): EjournalPreviewRow[] =>
  sheet.rows
    .filter((row) => hasRowData(row.values))
    .map(normalizeArchiveRowToRosterShape)
    .filter((row): row is EjournalPreviewRow => Boolean(row))
    .filter((row) => getRosterPersonName(row).length >= 3);

const personMatchKeys = (row: EjournalPreviewRow) => {
  const nameKey = normalizeRosterMatchText(getRosterPersonName(row));
  const birth = getRosterPersonBirthDate(row);
  const rnokpp = getRosterPersonRnokpp(row);
  return {
    nameKey,
    birth,
    rnokpp,
    nameBirth: nameKey && birth ? `${nameKey}|${birth}` : "",
  };
};

/** Archive people not already present in «Загальний список». */
export const pickArchiveOnlyRosterRows = (
  mainRows: EjournalPreviewRow[],
  archiveRows: EjournalPreviewRow[],
): EjournalPreviewRow[] => {
  if (!archiveRows.length) return [];

  const usedNames = new Set<string>();
  const usedRnokpp = new Set<string>();
  const usedNameBirth = new Set<string>();

  for (const row of mainRows) {
    const keys = personMatchKeys(row);
    if (keys.nameKey) usedNames.add(keys.nameKey);
    if (keys.rnokpp) usedRnokpp.add(keys.rnokpp);
    if (keys.nameBirth) usedNameBirth.add(keys.nameBirth);
  }

  const seen = new Set<string>();
  const result: EjournalPreviewRow[] = [];

  for (const row of archiveRows) {
    const keys = personMatchKeys(row);
    if (!keys.nameKey) continue;
    if (keys.rnokpp && usedRnokpp.has(keys.rnokpp)) continue;
    if (keys.nameBirth && usedNameBirth.has(keys.nameBirth)) continue;
    // Same ПІБ already in «Загальний список» → already present (no twin card).
    if (usedNames.has(keys.nameKey)) continue;

    const dedupeKey = keys.rnokpp || keys.nameBirth || keys.nameKey;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    result.push({
      ...row,
      __rowNumber:
        ARCHIVE_ROSTER_EXCEL_ROW_OFFSET + (Number(row.__rowNumber) || 0),
      [ROSTER_ARCHIVE_FLAG_KEY]: true,
      [ROSTER_ARCHIVE_SOURCE_KEY]: ROSTER_ARCHIVE_SOURCE_VALUE,
    });
  }

  return result;
};

export const appendArchiveOnlyToRosterRows = (
  mainRows: EjournalPreviewRow[],
  archiveSheet: WorkbookSheet | undefined,
): EjournalPreviewRow[] => {
  if (!archiveSheet) return mainRows;
  const archiveOnly = pickArchiveOnlyRosterRows(
    mainRows,
    parseArchiveSheetToRosterRows(archiveSheet),
  );
  if (!archiveOnly.length) return mainRows;
  return [...mainRows, ...archiveOnly];
};

/** DB import rows: archive-only people shaped like Загальний список cells. */
export const archiveOnlyRowsForPersonnelImport = (
  mainPreviewRows: EjournalPreviewRow[],
  archiveSheet: WorkbookSheet | undefined,
) => {
  if (!archiveSheet) return [];
  const archiveOnly = pickArchiveOnlyRosterRows(
    mainPreviewRows,
    parseArchiveSheetToRosterRows(archiveSheet),
  );

  return archiveOnly.map((row) => {
    const values: Record<string, unknown> = {
      [ROSTER_ARCHIVE_FLAG_KEY]: true,
      [ROSTER_ARCHIVE_SOURCE_KEY]: ROSTER_ARCHIVE_SOURCE_VALUE,
    };
    for (const [key, value] of Object.entries(row)) {
      if (key.startsWith("__") && key !== ROSTER_ARCHIVE_FLAG_KEY) continue;
      if (value == null || String(value).trim() === "") continue;
      values[key] = value;
    }
    return {
      excelRowNumber:
        Number(row.__rowNumber) || ARCHIVE_ROSTER_EXCEL_ROW_OFFSET,
      values,
    };
  });
};
