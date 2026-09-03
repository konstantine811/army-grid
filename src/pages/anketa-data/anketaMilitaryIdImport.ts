import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";
import {
  buildVkTpvDovidkyNameIndex,
  extractMilitaryIdFromText,
  type VkTpvDovidkyNameEntry,
} from "../personnel/vkTpvDovidkyImport";
import {
  applyAnketaEditsToRows,
  bulkWriteAnketaCellEdits,
  loadAnketaEdits,
} from "./anketaEdits";
import { ANKETA_MISSING_VALUE_PRESETS } from "./anketaGaps";
import { normalizeAnketaNameKey } from "./anketaPersonMatch";
import {
  loadAnketaSheetPreferCache,
  type AnketaColumnKey,
  type AnketaRow,
} from "./anketaSheet";

/** Позначка «немає військового квитка» в колонці «Військовий квиток». */
export const ANKETA_MILITARY_ID_ABSENT_VALUE = "відсутній";

const ANKETA_RNOKPP_PLACEHOLDER_SET = new Set<string>([
  ...ANKETA_MISSING_VALUE_PRESETS,
  ANKETA_MILITARY_ID_ABSENT_VALUE,
  "квиток відсутній",
]);

export type AnketaMilitaryIdMergeAction =
  | "from_vk"
  | "kept_anketa"
  | "marked_absent"
  | "marked_not_found"
  | "unchanged";

export type AnketaRnokppMergeAction = "from_vk" | "kept_anketa" | "unchanged";

export type AnketaMilitaryIdMergeReport = {
  anketaRows: number;
  vkNamesInFile: number;
  updatedFromVk: number;
  keptAnketa: number;
  markedAbsent: number;
  markedNotFound: number;
  unchanged: number;
  rnokppUpdatedFromVk: number;
  rnokppKeptAnketa: number;
  rnokppUnchanged: number;
  skippedNoName: number;
  errors: Array<{ name: string; message: string }>;
};

export const normalizeAnketaRnokppDigits = (value: unknown) => {
  const digits = String(value ?? "").replace(/\D/g, "");
  return digits.length === 10 ? digits : "";
};

export const isAnketaRnokppNumber = (value: unknown) =>
  normalizeAnketaRnokppDigits(value).length === 10;

export const isAnketaRnokppReplaceable = (value: unknown) => {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return true;
  if (isAnketaRnokppNumber(trimmed)) return false;
  if (
    ANKETA_RNOKPP_PLACEHOLDER_SET.has(
      trimmed.toLocaleLowerCase("uk-UA"),
    )
  ) {
    return true;
  }
  return !isAnketaRnokppNumber(trimmed);
};

export const resolveAnketaRnokppValue = (
  currentValue: string,
  vkEntry: VkTpvDovidkyNameEntry | undefined,
): { value: string; action: AnketaRnokppMergeAction } => {
  const current = String(currentValue ?? "").trim();
  const vkRnokpp = normalizeAnketaRnokppDigits(vkEntry?.rnokpp);

  if (isAnketaRnokppNumber(current)) {
    const currentDigits = normalizeAnketaRnokppDigits(current);
    return { value: currentDigits, action: "kept_anketa" };
  }

  if (vkRnokpp) {
    if (current === vkRnokpp) {
      return { value: current, action: "unchanged" };
    }
    return { value: vkRnokpp, action: "from_vk" };
  }

  return { value: current, action: "unchanged" };
};

export const resolveAnketaMilitaryIdValue = (
  currentValue: string,
  vkEntry: VkTpvDovidkyNameEntry | undefined,
): { value: string; action: AnketaMilitaryIdMergeAction } => {
  const current = String(currentValue ?? "").trim();
  const currentId = extractMilitaryIdFromText(current);
  const vkId = extractMilitaryIdFromText(String(vkEntry?.militaryId ?? ""));

  if (vkId) {
    if (currentId === vkId) {
      return { value: vkId, action: "unchanged" };
    }
    return { value: vkId, action: "from_vk" };
  }

  if (currentId) {
    return { value: currentId, action: "kept_anketa" };
  }

  if (current) {
    return { value: current, action: "kept_anketa" };
  }

  if (vkEntry) {
    if (current === ANKETA_MILITARY_ID_ABSENT_VALUE) {
      return { value: current, action: "unchanged" };
    }
    return { value: ANKETA_MILITARY_ID_ABSENT_VALUE, action: "marked_absent" };
  }

  if (current === ANKETA_MILITARY_ID_ABSENT_VALUE) {
    return { value: current, action: "unchanged" };
  }
  return {
    value: ANKETA_MILITARY_ID_ABSENT_VALUE,
    action: "marked_not_found",
  };
};

export const planAnketaMilitaryIdMerges = (
  anketaRows: AnketaRow[],
  vkIndex: Map<string, VkTpvDovidkyNameEntry>,
) => {
  const edits: Array<{
    rowNumber: number;
    columnId: AnketaColumnKey;
    value: string;
    externalId?: string;
    fullName?: string;
  }> = [];
  const report: AnketaMilitaryIdMergeReport = {
    anketaRows: anketaRows.length,
    vkNamesInFile: vkIndex.size,
    updatedFromVk: 0,
    keptAnketa: 0,
    markedAbsent: 0,
    markedNotFound: 0,
    unchanged: 0,
    rnokppUpdatedFromVk: 0,
    rnokppKeptAnketa: 0,
    rnokppUnchanged: 0,
    skippedNoName: 0,
    errors: [],
  };

  for (const row of anketaRows) {
    const nameKey = normalizeAnketaNameKey(row.fullName);
    if (!nameKey) {
      report.skippedNoName += 1;
      continue;
    }

    const vkEntry = vkIndex.get(nameKey);
    const militaryResolved = resolveAnketaMilitaryIdValue(
      row.militaryId,
      vkEntry,
    );

    switch (militaryResolved.action) {
      case "from_vk":
        report.updatedFromVk += 1;
        break;
      case "kept_anketa":
        report.keptAnketa += 1;
        break;
      case "marked_absent":
        report.markedAbsent += 1;
        break;
      case "marked_not_found":
        report.markedNotFound += 1;
        break;
      case "unchanged":
        report.unchanged += 1;
        break;
    }

    if (
      militaryResolved.action !== "unchanged" &&
      militaryResolved.action !== "kept_anketa"
    ) {
      edits.push({
        rowNumber: row.__rowNumber,
        columnId: "militaryId",
        value: militaryResolved.value,
        externalId: String(row.externalId ?? "").trim() || undefined,
        fullName: row.fullName,
      });
    }

    const rnokppResolved = resolveAnketaRnokppValue(row.rnokpp, vkEntry);
    switch (rnokppResolved.action) {
      case "from_vk":
        report.rnokppUpdatedFromVk += 1;
        break;
      case "kept_anketa":
        report.rnokppKeptAnketa += 1;
        break;
      case "unchanged":
        report.rnokppUnchanged += 1;
        break;
    }

    if (rnokppResolved.action === "from_vk") {
      edits.push({
        rowNumber: row.__rowNumber,
        columnId: "rnokpp",
        value: rnokppResolved.value,
        externalId: String(row.externalId ?? "").trim() || undefined,
        fullName: row.fullName,
      });
    }
  }

  return { edits, report };
};

export const formatAnketaMilitaryIdMergeReport = (
  report: AnketaMilitaryIdMergeReport,
) => {
  const militaryWritten =
    report.updatedFromVk + report.markedAbsent + report.markedNotFound;
  const parts = [
    `рядків анкет: ${report.anketaRows}`,
    `ПІБ у файлі ВК: ${report.vkNamesInFile}`,
    militaryWritten ? `ВК записано: ${militaryWritten}` : "",
    report.updatedFromVk ? `ВК з файлу: ${report.updatedFromVk}` : "",
    report.markedAbsent ? `ВК відсутній (у файлі): ${report.markedAbsent}` : "",
    report.markedNotFound
      ? `ВК відсутній (немає у файлі): ${report.markedNotFound}`
      : "",
    report.keptAnketa ? `ВК залишено з анкети: ${report.keptAnketa}` : "",
    report.rnokppUpdatedFromVk
      ? `ІПН з файлу: ${report.rnokppUpdatedFromVk}`
      : "",
    report.rnokppKeptAnketa
      ? `ІПН залишено з анкети: ${report.rnokppKeptAnketa}`
      : "",
    report.skippedNoName ? `без ПІБ: ${report.skippedNoName}` : "",
    report.errors.length ? `помилок: ${report.errors.length}` : "",
  ].filter(Boolean);
  return parts.join(" · ");
};

export const mergeVkMilitaryIdToAnketa = async (
  snapshot: ExcelWorkbookSnapshot,
  options?: {
    onProgress?: (done: number, total: number) => void;
  },
): Promise<AnketaMilitaryIdMergeReport> => {
  const vkIndex = buildVkTpvDovidkyNameIndex(snapshot);
  const [anketaSnapshot, anketaEdits] = await Promise.all([
    loadAnketaSheetPreferCache(),
    loadAnketaEdits(),
  ]);

  const anketaRows = applyAnketaEditsToRows(
    anketaSnapshot?.rows ?? [],
    anketaEdits,
  );
  const { edits, report } = planAnketaMilitaryIdMerges(anketaRows, vkIndex);

  if (!edits.length) {
    options?.onProgress?.(anketaRows.length, anketaRows.length);
    return report;
  }

  const batchSize = 50;
  for (let index = 0; index < edits.length; index += batchSize) {
    const batch = edits.slice(index, index + batchSize);
    try {
      await bulkWriteAnketaCellEdits(
        batch.map(({ rowNumber, columnId, value, externalId, fullName }) => ({
          rowNumber,
          columnId,
          value,
          externalId,
          fullName,
        })),
      );
    } catch (error) {
      for (const item of batch) {
        report.errors.push({
          name: item.fullName ?? `#${item.rowNumber}`,
          message:
            error instanceof Error
              ? error.message
              : "Не вдалося записати дані з файлу ВК",
        });
      }
    }
    options?.onProgress?.(
      Math.min(index + batch.length, edits.length),
      edits.length,
    );
  }

  options?.onProgress?.(anketaRows.length, anketaRows.length);
  return report;
};
