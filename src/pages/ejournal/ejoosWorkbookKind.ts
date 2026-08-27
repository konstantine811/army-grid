import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";

export type WorkbookKind = "ejoos" | "pb_1pb" | "unknown";

const sheetNamesOf = (workbook: ExcelWorkbookSnapshot) =>
  workbook.sheets.map((sheet) => sheet.sheetName.trim().toLowerCase());

const hasAny = (names: string[], patterns: RegExp[]) =>
  patterns.some((pattern) => names.some((name) => pattern.test(name)));

/** Визначає тип Excel: шаблон/журнал ЕЖООС vs штатка 1ПБ (sh / Рух / archive). */
export const detectWorkbookKind = (
  workbook: ExcelWorkbookSnapshot,
): WorkbookKind => {
  const names = sheetNamesOf(workbook);
  const hasSh = hasAny(names, [/^sh$/i]);
  const looksLikeEjoos =
    hasAny(names, [/шпо|штатно.?посад/i]) &&
    hasAny(names, [/оос|облік\s*особов/i]) &&
    hasAny(names, [/табель/i]);
  // 1ПБ завжди має sh; ЕЖООС — ніколи. Рух/archive підтверджують, але не обовʼязкові.
  const looksLikePb = hasSh && !looksLikeEjoos;

  if (looksLikeEjoos && !looksLikePb) return "ejoos";
  if (looksLikePb && !looksLikeEjoos) return "pb_1pb";
  if (looksLikeEjoos && looksLikePb) {
    // Рідкісний гібрид — пріоритет ЕЖООС, якщо є ключові аркуші журналу.
    return "ejoos";
  }
  if (looksLikeEjoos) return "ejoos";
  if (looksLikePb) return "pb_1pb";
  return "unknown";
};

export const assertEjoosWorkbook = (workbook: ExcelWorkbookSnapshot): void => {
  const kind = detectWorkbookKind(workbook);
  if (kind === "ejoos") return;

  const sheets = workbook.sheets.map((s) => s.sheetName).join(", ");
  if (kind === "pb_1pb") {
    throw new Error(
      `Це файл 1ПБ (є аркуші sh / Рух / archive), а не ЕЖООС.\n` +
        `Завантажте шаблон ЕЖООС (ШПО, ООС, Виключені, Табель…), ` +
        `а 1ПБ обирайте окремо.\n` +
        `Аркуші файлу: ${sheets}`,
    );
  }
  throw new Error(
    `Файл не схожий на ЕЖООС (потрібні аркуші «1. ШПО», «2. ООС», «6. Табель»).\n` +
      `Аркуші: ${sheets}`,
  );
};

export const assertPbWorkbook = (workbook: ExcelWorkbookSnapshot): void => {
  const kind = detectWorkbookKind(workbook);
  if (kind === "pb_1pb") return;

  const sheets = workbook.sheets.map((s) => s.sheetName).join(", ");
  if (kind === "ejoos") {
    throw new Error(
      `Це файл ЕЖООС (ШПО / ООС / Табель), а не 1ПБ.\n` +
        `Для 1ПБ потрібні аркуші sh, Рух, archive.\n` +
        `Аркуші: ${sheets}`,
    );
  }
  throw new Error(
    `Файл не схожий на 1ПБ (потрібен аркуш sh; зазвичай ще Рух і archive).\nАркуші: ${sheets}`,
  );
};

export const ejoosDownloadFileName = (
  version: number,
  asOfDate?: string | null,
  fallback?: string | null,
) => {
  if (asOfDate) {
    const stamp = asOfDate.replaceAll(".", "-");
    return `ЄЖООС_v${version}_станом_на_${stamp}.xlsx`;
  }
  if (fallback && /ежоос|єжоос|ejoos/i.test(fallback)) return fallback;
  return `ЄЖООС_v${version}.xlsx`;
};
