import {
  pushStaffSheetEnrichmentToGoogle,
  STAFF_SHEET_ENRICHMENT_COLUMNS,
  STAFF_SHEET_ENRICHMENT_START_ROW,
  STAFF_SHEET_ENRICHMENT_VALUE_INDEX,
  staffSheetBirthDateColumn,
} from "../excel-fill/staffSheet";
import {
  getPersonExternalId,
  getPersonFieldValue,
  isLikelyCallSignToken,
  looksLikePersonBirthDate,
  resolvePersonBirthDate,
  extractBirthYear,
  computeFullYearsFromBirthDate,
} from "../personnel/personnelUtils";
import type { EjournalPreviewRow } from "../ejournal/ejournalTypes";
import { readRosterColumnValue } from "../excel-fill/rosterSourceSnapshot";
import { normalizeAnketaNameKey, loadPersonnelIndexForAnketa, resolvePersonnelRowForStaffRoster, type AnketaPersonnelIndex } from "./anketaPersonMatch";
import type { AnketaRow } from "./anketaSheet";
import {
  findMergedPersonnelRow,
  loadStaffSheetEnrichmentContext,
} from "./staffSheetEnrichmentContext";
import { rowHasListedQuestionnaire } from "../personnel/personAttachments";
import type { BackendPersonQuestionnaireMeta } from "../../api";
import type { VkTpvDovidkyNameEntry } from "../personnel/vkTpvDovidkyImport";
import {
  loadStaffSheetVkIndex,
  resolveStaffSheetMilitaryIdValue,
} from "./staffSheetMilitaryIdMerge";
import { runStaffSheetEnrichmentHeavy } from "./runStaffSheetHeavyJobs";
import { extractMilitaryIdFromText } from "../personnel/vkTpvDovidkyImport";

const cellText = (value: unknown) => String(value ?? "").trim();

const normalizeRnokpp = (value: string) => {
  const digits = value.replace(/\D/g, "");
  return digits.length === 10 ? digits : value.trim();
};

const resolveRnokpp = (
  personnelRow: EjournalPreviewRow,
  anketaRow: AnketaRow | null,
) => {
  const fromAnketa = normalizeRnokpp(String(anketaRow?.rnokpp ?? ""));
  if (fromAnketa.replace(/\D/g, "").length === 10) return fromAnketa;

  const fromOos = normalizeRnokpp(
    getPersonFieldValue(personnelRow, ["рнокпп_за_наявності"]),
  );
  if (fromOos.replace(/\D/g, "").length === 10) return fromOos;

  const rosterInn = normalizeRnokpp(readRosterColumnValue(personnelRow, 19));
  return rosterInn.replace(/\D/g, "").length === 10 ? rosterInn : "";
};

const sanitizeBirthDate = (value: string) => {
  const text = cellText(value);
  if (!text || isLikelyCallSignToken(text)) return "";
  return looksLikePersonBirthDate(text) ? text : "";
};

const resolveBirthDate = (
  personnelRow: EjournalPreviewRow,
  anketaRow: AnketaRow | null,
) => {
  const candidates = [
    cellText(anketaRow?.birthDate),
    resolvePersonBirthDate(personnelRow),
    readRosterColumnValue(personnelRow, 16),
    readRosterColumnValue(personnelRow, 17),
  ];
  for (const candidate of candidates) {
    const valid = sanitizeBirthDate(candidate);
    if (valid) return valid;
  }
  return "";
};

const hasPersonName = (row: EjournalPreviewRow) =>
  readRosterColumnValue(row, 14).trim().length >= 3;

export type StaffSheetEnrichmentReport = {
  rows: number;
  totalPositions: number;
  anketaYes: number;
  withMilitaryId: number;
  militaryIdFromVk: number;
  withBirthDate: number;
  withInn: number;
};

export type StaffSheetEnrichmentEntry = {
  excelRowNumber: number;
  values: string[];
};

export const buildStaffSheetEnrichmentEntries = (options: {
  rosterRows: EjournalPreviewRow[];
  mergedPersonnelRows?: EjournalPreviewRow[];
  personnelIndex?: AnketaPersonnelIndex;
  anketaRows: AnketaRow[];
  questionnaires: BackendPersonQuestionnaireMeta[];
  vkIndex?: Map<string, VkTpvDovidkyNameEntry>;
}): StaffSheetEnrichmentEntry[] => {
  const anketaIndex = new Map<string, AnketaRow>();
  const anketaByExternalId = new Map<string, AnketaRow>();
  for (const row of options.anketaRows) {
    const key = normalizeAnketaNameKey(row.fullName);
    if (key && !anketaIndex.has(key)) anketaIndex.set(key, row);
    const externalId = String(row.externalId ?? "").trim();
    if (externalId && !anketaByExternalId.has(externalId)) {
      anketaByExternalId.set(externalId, row);
    }
  }

  const entries: StaffSheetEnrichmentEntry[] = [];

  options.rosterRows.forEach((rosterRow, index) => {
    if (!hasPersonName(rosterRow)) return;

    const excelRowNumber =
      Number(rosterRow.__rowNumber) || STAFF_SHEET_ENRICHMENT_START_ROW + index;
    const personnelRow = options.personnelIndex
      ? resolvePersonnelRowForStaffRoster(rosterRow, options.personnelIndex)
      : findMergedPersonnelRow(
          rosterRow,
          options.mergedPersonnelRows ?? [],
        );
    const externalId = getPersonExternalId(personnelRow).trim();
    const rosterName = readRosterColumnValue(rosterRow, 14);
    const nameKey = normalizeAnketaNameKey(rosterName);
    const matchedAnketa =
      (externalId ? anketaByExternalId.get(externalId) : undefined) ??
      (nameKey ? anketaIndex.get(nameKey) : undefined) ??
      null;
    const hasQuestionnaire = rowHasListedQuestionnaire(
      personnelRow,
      options.questionnaires,
      {
        anketaExternalId: matchedAnketa?.externalId,
        anketaFullName: matchedAnketa?.fullName || rosterName,
        anketaBirthDate: matchedAnketa?.birthDate,
      },
    );
    const birthDate = resolveBirthDate(personnelRow, matchedAnketa);
    const inn = resolveRnokpp(personnelRow, matchedAnketa);
    const birthDateCol = staffSheetBirthDateColumn;
    // Рік і вік — звичайні числа (не формули): інакше Google з форматом «Дата»
    // показує 1980 як 21.06.1905 і фарбує червоним.
    const birthYear = birthDate ? extractBirthYear(birthDate) : null;
    const fullYears = birthDate ? computeFullYearsFromBirthDate(birthDate) : null;
    const militaryId = resolveStaffSheetMilitaryIdValue(
      readRosterColumnValue(rosterRow, 11),
      options.vkIndex?.get(normalizeAnketaNameKey(rosterName)),
    ).value;

    entries.push({
      excelRowNumber,
      values: STAFF_SHEET_ENRICHMENT_COLUMNS.map((columnNumber) => {
        if (columnNumber === 10) return hasQuestionnaire ? "так" : "";
        if (columnNumber === 11) return militaryId;
        if (columnNumber === birthDateCol) return birthDate;
        if (columnNumber === birthDateCol + 1) {
          return birthYear != null ? String(birthYear) : "";
        }
        if (columnNumber === birthDateCol + 2) {
          return fullYears != null ? String(fullYears) : "";
        }
        if (columnNumber === 19) return inn;
        return "";
      }),
    });
  });

  return entries;
};

const buildStaffSheetAnketaEntries = (
  entries: StaffSheetEnrichmentEntry[],
): StaffSheetEnrichmentEntry[] =>
  entries.map((entry) => ({
    excelRowNumber: entry.excelRowNumber,
    values: [entry.values[0] ?? ""],
  }));

export const buildStaffSheetEnrichmentReport = (
  entries: StaffSheetEnrichmentEntry[],
  rosterRows: EjournalPreviewRow[],
): StaffSheetEnrichmentReport => ({
  rows: entries.length,
  totalPositions: rosterRows.length,
  anketaYes: entries.filter((entry) => entry.values[0] === "так").length,
  withMilitaryId: entries.filter((entry) =>
    Boolean(
      String(
        entry.values[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.militaryId] ?? "",
      ).trim(),
    ),
  ).length,
  militaryIdFromVk: entries.filter((entry) => {
    const rosterRow = rosterRows.find(
      (row) => Number(row.__rowNumber) === entry.excelRowNumber,
    );
    if (!rosterRow) return false;
    const hadValue = Boolean(
      extractMilitaryIdFromText(readRosterColumnValue(rosterRow, 11)),
    );
    const mergedValue = String(
      entry.values[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.militaryId] ?? "",
    ).trim();
    return !hadValue && mergedValue.length > 0;
  }).length,
  withBirthDate: entries.filter((entry) =>
    looksLikePersonBirthDate(
      String(entry.values[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.birthDate] ?? ""),
    ),
  ).length,
  withInn: entries.filter(
    (entry) =>
      entry.values[STAFF_SHEET_ENRICHMENT_VALUE_INDEX.inn]?.replace(/\D/g, "")
        .length === 10,
  ).length,
});

export const pushStaffSheetAnketaToGoogle = async (options?: {
  onProgress?: (phase: string) => void;
}): Promise<Pick<StaffSheetEnrichmentReport, "rows" | "anketaYes">> => {
  options?.onProgress?.("Завантаження даних…");
  const context = await loadStaffSheetEnrichmentContext();

  options?.onProgress?.("Запис колонки «Анкета» у Google Sheet…");
  const entries = buildStaffSheetAnketaEntries(
    await runStaffSheetEnrichmentHeavy({
      rosterRows: context.rosterRows,
      personnelIndex: context.personnelIndex,
      anketaRows: context.anketaRows,
      questionnaires: context.questionnaires,
      vkIndex: new Map(),
    }),
  );

  if (!entries.length) {
    throw new Error("Немає рядків з ПІБ для оновлення Google Sheet «Штатка».");
  }

  await pushStaffSheetEnrichmentToGoogle(entries, { columns: [10] });

  return {
    rows: entries.length,
    anketaYes: entries.filter((entry) => entry.values[0] === "так").length,
  };
};

export const pushPersonnelEnrichmentToStaffSheet = async (options?: {
  onProgress?: (phase: string) => void;
}): Promise<StaffSheetEnrichmentReport> => {
  options?.onProgress?.("Завантаження даних…");
  const [context, personnelIndex] = await Promise.all([
    loadStaffSheetEnrichmentContext(),
    loadPersonnelIndexForAnketa(),
  ]);

  options?.onProgress?.("Запис у Google Sheet…");
  const entries = await runStaffSheetEnrichmentHeavy({
    rosterRows: context.rosterRows,
    personnelIndex,
    anketaRows: context.anketaRows,
    questionnaires: context.questionnaires,
    vkIndex: await loadStaffSheetVkIndex(),
  });

  if (!entries.length) {
    throw new Error("Немає рядків з ПІБ для оновлення Google Sheet «Штатка».");
  }

  await pushStaffSheetEnrichmentToGoogle(entries);

  return buildStaffSheetEnrichmentReport(entries, context.rosterRows);
};

export const formatStaffSheetEnrichmentReport = (
  report: StaffSheetEnrichmentReport,
) =>
  [
    `позицій: ${report.totalPositions}`,
    `з ПІБ: ${report.rows}`,
    `анкета так: ${report.anketaYes}`,
    report.withMilitaryId ? `ВК: ${report.withMilitaryId}` : "",
    report.militaryIdFromVk ? `ВК з файлу: ${report.militaryIdFromVk}` : "",
    `дата народження: ${report.withBirthDate}`,
    `ІПН: ${report.withInn}`,
  ]
    .filter(Boolean)
    .join(" · ");
