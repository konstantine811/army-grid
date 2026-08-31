import JSZip from "jszip";

/**
 * Точковий патч колонки AE («Куди вибув») на аркуші «Виключені»:
 * lowercase + за можливості стиль як у «нормальних» клітинок AE.
 * Без xlsx-populate — лише правка XML у zip, щоб не ламати шаблон ЕЖООС.
 */

const EXCLUDED_SHEET_RE = /виключ/i;
const AE_CELL_RE = /<c\b([^>]*\br="AE(\d+)"[^>]*)(\/>|>[\s\S]*?<\/c>)/gi;

const escapeXml = (value: string) =>
  value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");

const decodeXml = (value: string) =>
  value
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#10;/g, "\n")
    .replace(/&#xA;/gi, "\n")
    .replace(/&amp;/g, "&");

const formatDestination = (value: string) =>
  value.replace(/\s+/g, " ").trim().toLocaleLowerCase("uk-UA");

const looksAllCaps = (value: string) => {
  const letters = value.replace(/[^\p{L}]/gu, "");
  if (!letters) return false;
  const upper = letters.toLocaleUpperCase("uk-UA");
  const lower = letters.toLocaleLowerCase("uk-UA");
  return letters === upper && letters !== lower;
};

const parseAttr = (attrs: string, name: string) =>
  attrs.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];

const parseSharedStrings = (xml: string): string[] => {
  const items: string[] = [];
  const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/gi;
  let match: RegExpExecArray | null;
  while ((match = siRe.exec(xml))) {
    const parts = [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)].map(
      (part) => decodeXml(part[1]),
    );
    items.push(parts.join(""));
  }
  return items;
};

const resolveCellText = (
  attrs: string,
  body: string,
  sharedStrings: string[],
): string => {
  const type = parseAttr(attrs, "t") || "";
  if (type === "inlineStr" || /<is\b/i.test(body)) {
    return [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
      .map((part) => decodeXml(part[1]))
      .join("");
  }
  if (type === "s") {
    const index = Number(
      body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1]?.trim() ?? "",
    );
    if (Number.isInteger(index) && sharedStrings[index] != null) {
      return sharedStrings[index];
    }
    return "";
  }
  const direct = body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/i)?.[1];
  return direct != null ? decodeXml(direct.trim()) : "";
};

const buildInlineCell = (row: number, styleId: string | undefined, text: string) => {
  const styleAttr = styleId != null && styleId !== "" ? ` s="${styleId}"` : "";
  const escaped = escapeXml(text);
  const space =
    text.startsWith(" ") || text.endsWith(" ") || text.includes("\n")
      ? ` xml:space="preserve"`
      : "";
  return `<c r="AE${row}"${styleAttr} t="inlineStr"><is><t${space}>${escaped}</t></is></c>`;
};

const resolveSheetPath = async (zip: JSZip, sheetNameRe: RegExp) => {
  const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
  const relsXml = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (!workbookXml || !relsXml) return null;

  const ridToTarget = new Map<string, string>();
  for (const rel of relsXml.matchAll(
    /\bId="([^"]+)"[^>]*\bTarget="([^"]+)"/g,
  )) {
    ridToTarget.set(rel[1], rel[2]);
  }
  for (const rel of relsXml.matchAll(
    /\bTarget="([^"]+)"[^>]*\bId="([^"]+)"/g,
  )) {
    if (!ridToTarget.has(rel[2])) ridToTarget.set(rel[2], rel[1]);
  }

  for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*)>/gi)) {
    const name = sheet[1].match(/\bname="([^"]+)"/)?.[1] ?? "";
    if (!sheetNameRe.test(name)) continue;
    const rid =
      sheet[1].match(/\br:id="([^"]+)"/)?.[1] ??
      sheet[1].match(/\bid="([^"]+)"/)?.[1];
    if (!rid) continue;
    const target = ridToTarget.get(rid);
    if (!target) continue;
    return target.startsWith("/")
      ? target.slice(1)
      : target.startsWith("xl/")
        ? target
        : `xl/${target.replace(/^\.\//, "")}`;
  }
  return null;
};

/**
 * Повертає новий blob: AE у «Виключені» → lowercase;
 * для ALL CAPS клітинок підставляє style з «нормальної» AE, якщо є.
 */
export async function patchExcludedDestinationInEjoosFile(
  file: Blob | File,
): Promise<Blob> {
  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const sheetPath = await resolveSheetPath(zip, EXCLUDED_SHEET_RE);
  if (!sheetPath) return file;

  const sheetXml = await zip.file(sheetPath)?.async("string");
  if (!sheetXml) return file;

  const sharedXml = await zip.file("xl/sharedStrings.xml")?.async("string");
  const sharedStrings = sharedXml ? parseSharedStrings(sharedXml) : [];

  type AeCell = {
    full: string;
    attrs: string;
    row: number;
    text: string;
    styleId?: string;
  };

  const cells: AeCell[] = [];
  AE_CELL_RE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = AE_CELL_RE.exec(sheetXml))) {
    const attrs = match[1];
    const row = Number(match[2]);
    if (!Number.isFinite(row) || row < 6) continue;
    const full = match[0];
    const body = full.endsWith("/>") ? "" : full.replace(/^<c\b[^>]*>/i, "").replace(/<\/c>$/i, "");
    const text = resolveCellText(attrs, body, sharedStrings);
    if (!text.trim()) continue;
    cells.push({
      full,
      attrs,
      row,
      text,
      styleId: parseAttr(attrs, "s"),
    });
  }

  if (!cells.length) return file;

  const styleRef =
    cells.find((cell) => !looksAllCaps(cell.text) && cell.styleId)?.styleId ??
    cells.find((cell) => cell.styleId)?.styleId;

  let nextXml = sheetXml;
  let changed = 0;
  for (const cell of cells) {
    const lowered = formatDestination(cell.text);
    const needsCase = lowered !== cell.text.trim() || looksAllCaps(cell.text);
    const styleId =
      looksAllCaps(cell.text) && styleRef ? styleRef : cell.styleId;
    if (!needsCase && styleId === cell.styleId) continue;
    const replacement = buildInlineCell(cell.row, styleId, lowered);
    if (replacement === cell.full) continue;
    nextXml = nextXml.replace(cell.full, replacement);
    changed += 1;
  }

  if (!changed) return file;
  zip.file(sheetPath, nextXml);
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  });
}
