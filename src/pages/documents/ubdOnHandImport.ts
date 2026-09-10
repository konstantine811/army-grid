import type { BackendPersonDocument } from "../../api";
import type { ExcelWorkbookSnapshot, ExcelSheetSnapshot } from "../../excelRoundTrip";
import { valueToDisplay } from "../../excelRoundTrip";
import {
  anketaNameKeyVariants,
  normalizeAnketaNameKey,
} from "../anketa-data/anketaPersonMatch";
import { overviewNameMatchesQuery } from "../overview/overviewNameSearch";
import { looksLikePersonnelName } from "../personnel/personnelUtils";
import { looksLikePersonName } from "../soc-passport/socPassportFields";
import { getDocumentPersonName } from "./documentPersonName";

export type UbdOnHandImportRow = {
  sheetName: string;
  rowNumber: number;
  lastName: string;
  firstName: string;
  patronymic: string;
  fullName: string;
  nameKey: string;
  note: string;
};

export type UbdOnHandDocumentMatch = {
  document: BackendPersonDocument;
  documentPersonName: string;
};

export type UbdOnHandMatchResult = {
  row: UbdOnHandImportRow;
  documents: UbdOnHandDocumentMatch[];
};

export type UbdOnHandMatchReport = {
  imported: UbdOnHandImportRow[];
  matched: UbdOnHandMatchResult[];
  unmatched: UbdOnHandImportRow[];
  matchedDocumentIds: Set<string>;
  matchedNameKeys: Set<string>;
};

const cellText = (value: unknown) =>
  valueToDisplay(value as Parameters<typeof valueToDisplay>[0])
    .replace(/\s+/g, " ")
    .trim();

const normalizeHeader = (value: unknown) =>
  cellText(value).toLocaleLowerCase("uk-UA").replace(/\s+/g, " ");

const headerIncludesLastName = (header: string) =>
  header.includes("прізвищ") || header.includes("призвищ");

const assembleFullName = (
  lastName: string,
  firstName: string,
  patronymic: string,
) =>
  [lastName, firstName, patronymic]
    .map((part) => part.trim())
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();

const toImportRow = (options: {
  sheetName: string;
  rowNumber: number;
  fullName: string;
  lastName?: string;
  firstName?: string;
  patronymic?: string;
  note?: string;
}): UbdOnHandImportRow | null => {
  const fullName = options.fullName.replace(/\s+/g, " ").trim();
  if (!looksLikePersonnelName(fullName) && !looksLikePersonName(fullName)) {
    return null;
  }
  const parts = fullName.split(/\s+/).filter(Boolean);
  const lastName = options.lastName?.trim() || parts[0] || "";
  const firstName = options.firstName?.trim() || parts[1] || "";
  const patronymic =
    options.patronymic?.trim() ||
    (parts.length > 2 ? parts.slice(2).join(" ") : "");
  return {
    sheetName: options.sheetName,
    rowNumber: options.rowNumber,
    lastName,
    firstName,
    patronymic,
    fullName,
    nameKey: normalizeAnketaNameKey(fullName),
    note: String(options.note ?? "").trim(),
  };
};

const findSplitNameHeaderRow = (rows: unknown[][]) => {
  for (let index = 0; index < Math.min(rows.length, 12); index += 1) {
    const row = rows[index] ?? [];
    const headers = row.map((cell) => normalizeHeader(cell));
    const hasLast = headers.some((header) => headerIncludesLastName(header));
    const hasFirst = headers.some(
      (header) =>
        header.includes("ім") &&
        !headerIncludesLastName(header) &&
        !header.includes("батьк"),
    );
    if (hasLast && hasFirst) return index;
  }
  return -1;
};

const resolveSplitNameColumns = (headerRow: unknown[]) => {
  const headers = headerRow.map((cell) => normalizeHeader(cell));
  const find = (parts: string[], exclude: string[] = []) =>
    headers.findIndex((header) => {
      if (!header) return false;
      if (exclude.some((part) => header.includes(part))) return false;
      return parts.every((part) => header.includes(part));
    });
  const last = headers.findIndex((header) => headerIncludesLastName(header));
  return {
    last,
    first: find(["ім"], ["прізвищ", "призвищ", "батьк"]),
    patronymic: find(["батьк"]),
    note: headers.findIndex((header) => header.length > 0) >= 0
      ? Math.max(
          find(["приміт"]),
          find(["підрозд"]),
          find(["рот"]),
          headers.length - 1,
        )
      : -1,
  };
};

const parseSplitNameSheet = (
  sheet: ExcelSheetSnapshot,
  rows: unknown[][],
): UbdOnHandImportRow[] => {
  const headerIndex = findSplitNameHeaderRow(rows);
  if (headerIndex < 0) return [];
  const columns = resolveSplitNameColumns(rows[headerIndex] ?? []);
  const result: UbdOnHandImportRow[] = [];
  const seen = new Set<string>();

  for (let rowIndex = headerIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const lastName = cellText(row[columns.last]);
    const firstName = cellText(row[columns.first]);
    const patronymic = cellText(row[columns.patronymic]);
    const fullName = assembleFullName(lastName, firstName, patronymic);
    const note =
      columns.note >= 0
        ? cellText(row[columns.note])
        : row
            .slice(Math.max(columns.patronymic, columns.first, columns.last) + 1)
            .map((cell) => cellText(cell))
            .filter(Boolean)
            .join(" · ");
    const parsed = toImportRow({
      sheetName: sheet.sheetName,
      rowNumber: rowIndex + 1,
      fullName,
      lastName,
      firstName,
      patronymic,
      note,
    });
    if (!parsed || seen.has(parsed.nameKey)) continue;
    seen.add(parsed.nameKey);
    result.push(parsed);
  }
  return result;
};

const parseFullNameSheet = (
  sheet: ExcelSheetSnapshot,
  rows: unknown[][],
): UbdOnHandImportRow[] => {
  const result: UbdOnHandImportRow[] = [];
  const seen = new Set<string>();

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    let fullName = "";
    for (const cell of row) {
      const text = cellText(cell);
      if (!text) continue;
      if (/^\d+$/.test(text)) continue;
      if (/убд|загиб|звільн|сзч/i.test(text) && text.length < 40) continue;
      if (looksLikePersonnelName(text) || looksLikePersonName(text)) {
        fullName = text;
        break;
      }
    }
    const parsed = toImportRow({
      sheetName: sheet.sheetName,
      rowNumber: rowIndex + 1,
      fullName,
    });
    if (!parsed || seen.has(parsed.nameKey)) continue;
    seen.add(parsed.nameKey);
    result.push(parsed);
  }
  return result;
};

const sheetRows = (sheet: ExcelSheetSnapshot) =>
  (sheet.rawRows?.length
    ? sheet.rawRows
    : [
        ...sheet.headerRows,
        ...sheet.rows.map((row) => row.values),
      ]) as unknown[][];

export const parseUbdOnHandWorkbook = (
  workbook: ExcelWorkbookSnapshot,
): UbdOnHandImportRow[] => {
  const merged = new Map<string, UbdOnHandImportRow>();

  for (const sheet of workbook.sheets) {
    const rows = sheetRows(sheet);
    const parsed = findSplitNameHeaderRow(rows) >= 0
      ? parseSplitNameSheet(sheet, rows)
      : parseFullNameSheet(sheet, rows);
    for (const row of parsed) {
      if (!merged.has(row.nameKey)) merged.set(row.nameKey, row);
    }
  }

  return [...merged.values()].sort((left, right) =>
    left.fullName.localeCompare(right.fullName, "uk", {
      sensitivity: "base",
    }),
  );
};

const namesEquivalent = (left: string, right: string) => {
  const keyLeft = normalizeAnketaNameKey(left);
  const keyRight = normalizeAnketaNameKey(right);
  if (!keyLeft || !keyRight) return false;
  if (keyLeft === keyRight) return true;
  for (const variant of anketaNameKeyVariants(left)) {
    if (anketaNameKeyVariants(right).has(variant)) return true;
  }
  return overviewNameMatchesQuery(left, right);
};

const buildDocumentNameIndex = (documents: BackendPersonDocument[]) => {
  const byKey = new Map<string, BackendPersonDocument[]>();
  const entries: Array<{
    document: BackendPersonDocument;
    personName: string;
  }> = [];

  for (const document of documents) {
    const personName = getDocumentPersonName(document);
    entries.push({ document, personName });
    for (const key of anketaNameKeyVariants(personName)) {
      const list = byKey.get(key) ?? [];
      list.push(document);
      byKey.set(key, list);
    }
  }

  return { byKey, entries };
};

export const matchUbdOnHandAgainstDocuments = (
  imported: UbdOnHandImportRow[],
  documents: BackendPersonDocument[],
  options?: { documentTypes?: string[] },
): UbdOnHandMatchReport => {
  const allowedTypes = options?.documentTypes?.length
    ? new Set(options.documentTypes)
    : null;
  const scopedDocuments = allowedTypes
    ? documents.filter((document) => allowedTypes.has(document.type))
    : documents;
  const { byKey, entries } = buildDocumentNameIndex(scopedDocuments);

  const matched: UbdOnHandMatchResult[] = [];
  const unmatched: UbdOnHandImportRow[] = [];
  const matchedDocumentIds = new Set<string>();
  const matchedNameKeys = new Set<string>();

  for (const row of imported) {
    const hits = new Map<string, UbdOnHandDocumentMatch>();
    for (const key of anketaNameKeyVariants(row.fullName)) {
      for (const document of byKey.get(key) ?? []) {
        if (hits.has(document.id)) continue;
        hits.set(document.id, {
          document,
          documentPersonName: getDocumentPersonName(document),
        });
      }
    }
    if (!hits.size) {
      for (const entry of entries) {
        if (!namesEquivalent(row.fullName, entry.personName)) continue;
        if (hits.has(entry.document.id)) continue;
        hits.set(entry.document.id, {
          document: entry.document,
          documentPersonName: entry.personName,
        });
      }
    }

    if (!hits.size) {
      unmatched.push(row);
      continue;
    }

    matchedNameKeys.add(row.nameKey);
    for (const hit of hits.values()) matchedDocumentIds.add(hit.document.id);
    matched.push({
      row,
      documents: [...hits.values()].sort((left, right) =>
        left.documentPersonName.localeCompare(right.documentPersonName, "uk"),
      ),
    });
  }

  return {
    imported,
    matched,
    unmatched,
    matchedDocumentIds,
    matchedNameKeys,
  };
};

export const formatUbdOnHandMatchSummary = (report: UbdOnHandMatchReport) => {
  const documentCount = report.matched.reduce(
    (sum, item) => sum + item.documents.length,
    0,
  );
  return [
    `імпорт: ${report.imported.length}`,
    `з документами: ${report.matched.length}`,
    `без документів: ${report.unmatched.length}`,
    `записів у журналі: ${documentCount}`,
  ].join(" · ");
};

export const formatUbdOnHandSheetBreakdown = (
  imported: UbdOnHandImportRow[],
) => {
  const counts = new Map<string, number>();
  for (const row of imported) {
    counts.set(row.sheetName, (counts.get(row.sheetName) ?? 0) + 1);
  }
  return [...counts.entries()]
    .sort(([left], [right]) => left.localeCompare(right, "uk"))
    .map(([sheetName, count]) => `${sheetName}: ${count}`)
    .join(" · ");
};
