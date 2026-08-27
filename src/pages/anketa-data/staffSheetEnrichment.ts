import {
  pushStaffSheetEnrichmentToGoogle,
  STAFF_SHEET_ENRICHMENT_COLUMNS,
  STAFF_SHEET_ENRICHMENT_START_ROW,
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
import { normalizeAnketaNameKey } from "./anketaPersonMatch";
import type { AnketaRow } from "./anketaSheet";
import {
  findMergedPersonnelRow,
  loadStaffSheetEnrichmentContext,
} from "./staffSheetEnrichmentContext";
import { rowHasListedQuestionnaire } from "../personnel/personAttachments";
import type { BackendPersonQuestionnaireMeta } from "../../api";

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
  anketaYes: number;
  withBirthDate: number;
  withInn: number;
};

export type StaffSheetEnrichmentEntry = {
  excelRowNumber: number;
  values: string[];
};

export const buildStaffSheetEnrichmentEntries = (options: {
  rosterRows: EjournalPreviewRow[];
  mergedPersonnelRows: EjournalPreviewRow[];
  anketaRows: AnketaRow[];
  questionnaires: BackendPersonQuestionnaireMeta[];
}): StaffSheetEnrichmentEntry[] => {
  const anketaIndex = new Map<string, AnketaRow>();
  for (const row of options.anketaRows) {
    const key = normalizeAnketaNameKey(row.fullName);
    if (!key || anketaIndex.has(key)) continue;
    anketaIndex.set(key, row);
  }

  const entries: StaffSheetEnrichmentEntry[] = [];

  options.rosterRows.forEach((rosterRow, index) => {
    if (!hasPersonName(rosterRow)) return;

    const excelRowNumber =
      Number(rosterRow.__rowNumber) || STAFF_SHEET_ENRICHMENT_START_ROW + index;
    const personnelRow = findMergedPersonnelRow(
      rosterRow,
      options.mergedPersonnelRows,
    );
    const externalId = getPersonExternalId(personnelRow).trim();
    const matchedAnketa =
      [...options.anketaRows].find((row) => {
        const id = String(row.externalId ?? "").trim();
        return id && id === externalId;
      }) ??
      anketaIndex.get(
        normalizeAnketaNameKey(readRosterColumnValue(personnelRow, 14)),
      ) ??
      null;

    const rosterName = readRosterColumnValue(rosterRow, 14);
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

    entries.push({
      excelRowNumber,
      values: STAFF_SHEET_ENRICHMENT_COLUMNS.map((columnNumber) => {
        if (columnNumber === 10) return hasQuestionnaire ? "так" : "";
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

export const pushPersonnelEnrichmentToStaffSheet = async (options?: {
  onProgress?: (phase: string) => void;
}): Promise<StaffSheetEnrichmentReport> => {
  options?.onProgress?.("Завантаження даних…");
  const context = await loadStaffSheetEnrichmentContext();

  options?.onProgress?.("Запис у Google Sheet…");
  const entries = buildStaffSheetEnrichmentEntries({
    rosterRows: context.rosterRows,
    mergedPersonnelRows: context.mergedPersonnelRows,
    anketaRows: context.anketaRows,
    questionnaires: context.questionnaires,
  });

  if (!entries.length) {
    throw new Error("Немає рядків з ПІБ для оновлення Google Sheet «Штатка».");
  }

  await pushStaffSheetEnrichmentToGoogle(entries);

  return {
    rows: entries.length,
    anketaYes: entries.filter((entry) => entry.values[0] === "так").length,
    withBirthDate: entries.filter((entry) =>
      looksLikePersonBirthDate(String(entry.values[1] ?? "")),
    ).length,
    withInn: entries.filter(
      (entry) => entry.values[4]?.replace(/\D/g, "").length === 10,
    ).length,
  };
};

export const formatStaffSheetEnrichmentReport = (
  report: StaffSheetEnrichmentReport,
) =>
  [
    `рядків: ${report.rows}`,
    `анкета так: ${report.anketaYes}`,
    `дата народження: ${report.withBirthDate}`,
    `ІПН: ${report.withInn}`,
  ].join(" · ");
