import { api } from "../../api";
import { invalidatePersonnelCaches } from "../../data/idbDataCache";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import { applyAnketaEditsToRows, loadAnketaEdits, upsertAnketaCellEdit } from "../anketa-data/anketaEdits";
import {
  loadPersonnelIndexForAnketa,
  normalizeAnketaNameKey,
} from "../anketa-data/anketaPersonMatch";
import { loadAnketaSheetPreferCache } from "../anketa-data/anketaSheet";
import { isBlankPersonValue } from "./personEnrichment";
import {
  getPersonFieldValue,
  isPersistedEjournalRowId,
  looksLikePersonnelName,
  resolvePersonFieldKey,
} from "./personnelUtils";

export const VK_TPV_DOVIDKY_SHEET_NAMES = [
  "ВК",
  "запис на зброю",
  "ТПВ",
  "ДОВІДКИ",
  "ВІДСУТНІ",
] as const;

export type VkTpvDovidkyRecord = {
  fullName: string;
  nameKey: string;
  militaryId: string;
  rnokpp: string;
  sourceSheet: string;
};

export type VkTpvDovidkyMergeReport = {
  parsed: number;
  matchedPersonnel: number;
  personnelMilitaryUpdated: number;
  matchedAnketa: number;
  anketaRnokppUpdated: number;
  skippedNoPersonnel: number;
  skippedNoAnketa: number;
  skippedAmbiguousPersonnel: number;
  skippedPersonnelHasMilitary: number;
  skippedAnketaHasRnokpp: number;
  skippedNoData: number;
  errors: Array<{ name: string; message: string }>;
};

const cellText = (value: unknown): string => {
  if (value == null) return "";
  if (typeof value === "object" && value !== null && "_error" in value) {
    return "";
  }
  return String(value).replace(/\s+/g, " ").trim();
};

export const extractMilitaryIdFromText = (value: string): string => {
  const text = String(value ?? "").replace(/\s+/g, " ").trim();
  if (!text) return "";
  const match = text.match(/(АГ|УН|АВ)\s*(\d[\d\s]*)/iu);
  if (!match?.[2]) return "";
  const digits = match[2].replace(/\s+/g, "");
  if (!digits) return "";
  return `${match[1].toUpperCase()} ${digits}`;
};

export const extractRnokppFromText = (value: string): string => {
  const text = String(value ?? "").trim();
  if (!text || /без\s*іпн/i.test(text)) return "";
  for (const match of text.match(/\b(\d{10})\b/g) ?? []) {
    if (match.length === 10) return match;
  }
  return "";
};

const extractFromCells = (cells: unknown[]) => {
  let militaryId = "";
  let rnokpp = "";
  for (const cell of cells) {
    const text = cellText(cell);
    if (!text) continue;
    if (!militaryId) militaryId = extractMilitaryIdFromText(text);
    if (!rnokpp) rnokpp = extractRnokppFromText(text);
  }
  return { militaryId, rnokpp };
};

const normalizeHeader = (value: unknown) =>
  cellText(value).toLocaleLowerCase("uk-UA").replace(/\s+/g, " ");

const findHeaderRowIndex = (rows: unknown[][]) => {
  for (let index = 0; index < Math.min(rows.length, 25); index += 1) {
    const row = rows[index] ?? [];
    if (row.some((cell) => normalizeHeader(cell).includes("піб"))) {
      return index;
    }
  }
  return -1;
};

const findColumnIndex = (headerRow: unknown[], matchers: string[]) => {
  for (let index = 0; index < headerRow.length; index += 1) {
    const header = normalizeHeader(headerRow[index]);
    if (!header) continue;
    if (matchers.some((part) => header.includes(part))) return index;
  }
  return -1;
};

const findInnColumnIndex = (headerRow: unknown[]) => {
  for (let index = 0; index < headerRow.length; index += 1) {
    const header = normalizeHeader(headerRow[index]);
    if (!header || header.includes("вк")) continue;
    if (header === "інн" || header.startsWith("інн ")) return index;
  }
  return -1;
};

const findVkColumnIndex = (headerRow: unknown[]) => {
  for (let index = 0; index < headerRow.length; index += 1) {
    const header = normalizeHeader(headerRow[index]);
    if (!header) continue;
    if (header.includes("вк№") || header === "вк" || header.startsWith("вк ")) {
      return index;
    }
  }
  return -1;
};

const parseStandardSheet = (
  sheetName: string,
  rows: unknown[][],
): VkTpvDovidkyRecord[] => {
  const headerIndex = findHeaderRowIndex(rows);
  if (headerIndex < 0) return [];

  const headerRow = rows[headerIndex] ?? [];
  const nameCol = findColumnIndex(headerRow, ["піб"]);
  const cinkaCol = findColumnIndex(headerRow, ["№цинка", "цинка new"]);
  const vkCol = findVkColumnIndex(headerRow);
  const innCol = findInnColumnIndex(headerRow);
  if (nameCol < 0) return [];

  const records: VkTpvDovidkyRecord[] = [];
  for (let rowIndex = headerIndex + 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const fullName = cellText(row[nameCol]);
    if (!looksLikePersonnelName(fullName)) continue;

    const fromVkCol = vkCol >= 0 ? cellText(row[vkCol]) : "";
    const fromInnCol = innCol >= 0 ? cellText(row[innCol]) : "";
    const fromCinka = cinkaCol >= 0 ? cellText(row[cinkaCol]) : "";
    const fromRow = extractFromCells(row);
    const militaryId =
      extractMilitaryIdFromText(fromVkCol) ||
      extractMilitaryIdFromText(fromCinka) ||
      fromRow.militaryId;
    const rnokpp =
      extractRnokppFromText(fromInnCol) ||
      extractRnokppFromText(fromCinka) ||
      fromRow.rnokpp;
    if (!militaryId && !rnokpp) continue;

    records.push({
      fullName,
      nameKey: normalizeAnketaNameKey(fullName),
      militaryId,
      rnokpp,
      sourceSheet: sheetName,
    });
  }
  return records;
};

const parseVidutniSheet = (rows: unknown[][]): VkTpvDovidkyRecord[] => {
  const records: VkTpvDovidkyRecord[] = [];
  for (const row of rows) {
    const fullName = cellText(row[4]);
    if (!looksLikePersonnelName(fullName)) continue;
    const fromRow = extractFromCells(row);
    const rnokppCol = cellText(row[14]);
    const rnokpp = extractRnokppFromText(rnokppCol) || fromRow.rnokpp;
    const militaryId = fromRow.militaryId;
    if (!militaryId && !rnokpp) continue;
    records.push({
      fullName,
      nameKey: normalizeAnketaNameKey(fullName),
      militaryId,
      rnokpp,
      sourceSheet: "ВІДСУТНІ",
    });
  }
  return records;
};

export const parseVkTpvDovidkyWorkbook = (
  snapshot: ExcelWorkbookSnapshot,
): VkTpvDovidkyRecord[] => {
  const merged = new Map<string, VkTpvDovidkyRecord>();

  for (const sheet of snapshot.sheets) {
    const sheetName = sheet.sheetName.trim();
    if (
      !VK_TPV_DOVIDKY_SHEET_NAMES.some(
        (name) => name.toLocaleLowerCase("uk-UA") === sheetName.toLocaleLowerCase("uk-UA"),
      )
    ) {
      continue;
    }

    const rows = (sheet.rawRows?.length
      ? sheet.rawRows
      : [
          ...sheet.headerRows,
          ...sheet.rows.map((row) => row.values),
        ]) as unknown[][];
    const parsed =
      sheetName.toLocaleLowerCase("uk-UA") === "відсутні"
        ? parseVidutniSheet(rows)
        : parseStandardSheet(sheetName, rows);

    for (const record of parsed) {
      const existing = merged.get(record.nameKey);
      if (!existing) {
        merged.set(record.nameKey, record);
        continue;
      }
      merged.set(record.nameKey, {
        ...existing,
        militaryId: existing.militaryId || record.militaryId,
        rnokpp: existing.rnokpp || record.rnokpp,
        sourceSheet: existing.sourceSheet,
      });
    }
  }

  return [...merged.values()];
};

export const formatVkTpvDovidkyMergeReport = (report: VkTpvDovidkyMergeReport) => {
  const parts = [
    `записів у файлі: ${report.parsed}`,
    `ВК в ООС: ${report.personnelMilitaryUpdated}`,
    `ІПН в анкети: ${report.anketaRnokppUpdated}`,
    report.skippedNoPersonnel ? `без збігу в ООС: ${report.skippedNoPersonnel}` : "",
    report.skippedAmbiguousPersonnel
      ? `дублікати ПІБ: ${report.skippedAmbiguousPersonnel}`
      : "",
    report.skippedNoAnketa ? `без рядка анкети: ${report.skippedNoAnketa}` : "",
    report.errors.length ? `помилок: ${report.errors.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

export const mergeVkTpvDovidkyWorkbook = async (
  snapshot: ExcelWorkbookSnapshot,
  options?: {
    onProgress?: (done: number, total: number) => void;
  },
): Promise<VkTpvDovidkyMergeReport> => {
  const records = parseVkTpvDovidkyWorkbook(snapshot);
  const report: VkTpvDovidkyMergeReport = {
    parsed: records.length,
    matchedPersonnel: 0,
    personnelMilitaryUpdated: 0,
    matchedAnketa: 0,
    anketaRnokppUpdated: 0,
    skippedNoPersonnel: 0,
    skippedNoAnketa: 0,
    skippedAmbiguousPersonnel: 0,
    skippedPersonnelHasMilitary: 0,
    skippedAnketaHasRnokpp: 0,
    skippedNoData: 0,
    errors: [],
  };

  if (!records.length) return report;

  const [personnelIndex, anketaSnapshot, anketaEdits] = await Promise.all([
    loadPersonnelIndexForAnketa({ force: true }),
    loadAnketaSheetPreferCache(),
    loadAnketaEdits(),
  ]);

  const anketaRows = applyAnketaEditsToRows(
    anketaSnapshot?.rows ?? [],
    anketaEdits,
  );
  const anketaByName = new Map<string, (typeof anketaRows)[number]>();
  for (const row of anketaRows) {
    const key = normalizeAnketaNameKey(row.fullName);
    if (!key || anketaByName.has(key)) continue;
    anketaByName.set(key, row);
  }

  let personnelUpdated = false;

  for (let index = 0; index < records.length; index += 1) {
    const record = records[index]!;
    options?.onProgress?.(index + 1, records.length);

    if (!record.militaryId && !record.rnokpp) {
      report.skippedNoData += 1;
      continue;
    }

    const personnelMatches = personnelIndex.byName.get(record.nameKey);
    if (!personnelMatches?.length) {
      report.skippedNoPersonnel += 1;
    } else if (personnelMatches.length > 1) {
      report.skippedAmbiguousPersonnel += 1;
    } else {
      const match = personnelMatches[0]!;
      report.matchedPersonnel += 1;

      if (record.militaryId) {
        const current = getPersonFieldValue(match.row, ["військового", "квитка"]);
        if (!isBlankPersonValue(current)) {
          report.skippedPersonnelHasMilitary += 1;
        } else {
          const rowId = match.row.__dbRowId ? String(match.row.__dbRowId) : "";
          if (isPersistedEjournalRowId(rowId)) {
            try {
              const targetKey =
                resolvePersonFieldKey(match.row, ["військового", "квитка"]) ||
                "військового_квитка";
              await api.updateEjournalRowValues(rowId, {
                [targetKey]: record.militaryId,
              });
              report.personnelMilitaryUpdated += 1;
              personnelUpdated = true;
            } catch (error) {
              report.errors.push({
                name: record.fullName,
                message:
                  error instanceof Error
                    ? error.message
                    : "Не вдалося записати військовий квиток",
              });
            }
          }
        }
      }
    }

    const anketaRow = anketaByName.get(record.nameKey);
    if (!anketaRow) {
      report.skippedNoAnketa += 1;
    } else if (record.rnokpp) {
      report.matchedAnketa += 1;
      const currentRnokpp = String(anketaRow.rnokpp ?? "").trim();
      if (currentRnokpp) {
        report.skippedAnketaHasRnokpp += 1;
      } else {
        try {
          await upsertAnketaCellEdit({
            rowNumber: anketaRow.__rowNumber,
            columnId: "rnokpp",
            value: record.rnokpp,
            externalId: String(anketaRow.externalId ?? "").trim() || undefined,
            fullName: anketaRow.fullName,
          });
          report.anketaRnokppUpdated += 1;
        } catch (error) {
          report.errors.push({
            name: record.fullName,
            message:
              error instanceof Error
                ? error.message
                : "Не вдалося записати РНОКПП в анкету",
          });
        }
      }
    }
  }

  if (personnelUpdated) {
    await invalidatePersonnelCaches();
  }

  return report;
};
