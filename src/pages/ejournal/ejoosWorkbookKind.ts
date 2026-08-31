import JSZip from "jszip";
import type { ExcelWorkbookSnapshot } from "../../excelRoundTrip";

export type WorkbookKind = "ejoos" | "pb_1pb" | "unknown";

const sheetNamesOf = (workbook: ExcelWorkbookSnapshot) =>
  workbook.sheets.map((sheet) => sheet.sheetName.trim().toLowerCase());

const hasAny = (names: string[], patterns: RegExp[]) =>
  patterns.some((pattern) => names.some((name) => pattern.test(name)));

const kindFromNames = (names: string[]): WorkbookKind => {
  const lower = names.map((name) => name.trim().toLowerCase());
  const hasSh = hasAny(lower, [/^sh$/i]);
  const looksLikeEjoos =
    hasAny(lower, [/шпо|штатно.?посад/i]) &&
    hasAny(lower, [/оос|облік\s*особов/i]) &&
    hasAny(lower, [/табель/i]);
  const looksLikePb = hasSh && !looksLikeEjoos;

  if (looksLikeEjoos && !looksLikePb) return "ejoos";
  if (looksLikePb && !looksLikeEjoos) return "pb_1pb";
  if (looksLikeEjoos && looksLikePb) return "ejoos";
  if (looksLikeEjoos) return "ejoos";
  if (looksLikePb) return "pb_1pb";
  return "unknown";
};

/** Лише workbook.xml — без xlsx-populate, щоб експорт не валив вкладку. */
export async function readWorkbookSheetNames(
  file: Blob | File,
): Promise<string[]> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  if (!workbookXml) return [];
  return [...workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/gi)].map(
    (match) => match[1],
  );
}

export const detectWorkbookKindFromFile = async (file: Blob | File) =>
  kindFromNames(await readWorkbookSheetNames(file));

/** Визначає тип Excel: шаблон/журнал ЕЖООС vs штатка 1ПБ (sh / Рух / archive). */
export const detectWorkbookKind = (
  workbook: ExcelWorkbookSnapshot,
): WorkbookKind => kindFromNames(sheetNamesOf(workbook));

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
  _version: number,
  asOfDate?: string | null,
  _fallback?: string | null,
) => {
  if (asOfDate) {
    const stamp = asOfDate.replaceAll(".", "-");
    return `ЄЖООС_станом_на_${stamp}.xlsx`;
  }
  return "ЄЖООС.xlsx";
};
