import JSZip from "jszip";
import {
  repairBrokenCellOpenTags,
  stripStaleCalcChainFromZip,
} from "./ejoosWorkbookSanitize";

export type ZipCellWrite = {
  /** 1-based Excel row */
  row: number;
  /** 1-based Excel column */
  column: number;
  /** null clears the value while preserving the selected cell style. */
  value: string | number | null;
  /** Optional row whose cell style should be copied for consistent formatting. */
  styleSourceRow?: number;
  /**
   * Copy `s=` from this column (1-based) on the style-source row.
   * Needed when the target column is empty on neighbors (O/P last-rank).
   */
  styleSourceColumn?: number;
  /** Force wrapText alignment via styles.xml. */
  wrapText?: boolean;
  /**
   * Copy neighbor font/border `s`, then force wrap + center alignment.
   */
  copyNeighborStyle?: boolean;
  /**
   * Use the neighbor cell `s=` as-is (same font, align, borders).
   * Do not invent a new wrap/center xf — that is what made anketa fills bold.
   */
  keepNeighborStyle?: boolean;
  /** Absolute cellXfs index, if already resolved. */
  styleId?: string;
  /** Copy this row's `ht` instead of guessing height from text length. */
  heightSourceRow?: number;
  /** Index in xl/sharedStrings.xml when writing text as `t="s"`. */
  sharedStringIndex?: number;
  /** Keep the current value; only stamp neighbor `s` (font / align / wrap). */
  styleOnly?: boolean;
};

const escapeXml = (value: string) =>
  value
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\r\n/g, "&#10;")
    .replace(/\n/g, "&#10;")
    .replace(/\r/g, "&#10;");

const SST_NS =
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SST_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
const SST_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
const EMPTY_SST = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${SST_NS}" count="0" uniqueCount="0"></sst>`;

const appendSharedString = (
  sstXml: string,
  text: string,
  knownCount?: number,
) => {
  const xml = sstXml || EMPTY_SST;
  const space =
    text.startsWith(" ") || text.endsWith(" ") || /[\n\t]/.test(text)
      ? ` xml:space="preserve"`
      : "";
  const si = `<si><t${space}>${escapeXml(text).replace(/&#10;/g, "_x000A_")}</t></si>`;
  const uniqueCount =
    knownCount ?? (xml.match(/<si\b/gi) ?? []).length;
  const next = xml.includes("</sst>")
    ? xml.replace(/<\/sst>/i, `${si}</sst>`)
    : `${xml}${si}`;
  const withCounts = next.replace(/<sst\b([^>]*)>/i, (_match, attrs) => {
    const cleaned = String(attrs)
      .replace(/\s+count="[^"]*"/i, "")
      .replace(/\s+uniqueCount="[^"]*"/i, "");
    return `<sst${cleaned} count="${uniqueCount + 1}" uniqueCount="${uniqueCount + 1}">`;
  });
  return { xml: withCounts, index: uniqueCount };
};

const ensureSharedStringsPart = async (zip: JSZip, sstXml: string) => {
  zip.file("xl/sharedStrings.xml", sstXml);
  const contentTypes = await zip.file("[Content_Types].xml")?.async("string");
  if (contentTypes && !/sharedStrings\.xml/i.test(contentTypes)) {
    zip.file(
      "[Content_Types].xml",
      contentTypes.replace(
        /<\/Types>/i,
        `<Override PartName="/xl/sharedStrings.xml" ContentType="${SST_CONTENT_TYPE}"/></Types>`,
      ),
    );
  }
  const rels = await zip.file("xl/_rels/workbook.xml.rels")?.async("string");
  if (rels && !/sharedStrings\.xml/i.test(rels)) {
    const ids = [...rels.matchAll(/\bId="rId(\d+)"/g)].map((match) =>
      Number(match[1]),
    );
    const nextId = Math.max(0, ...ids) + 1;
    zip.file(
      "xl/_rels/workbook.xml.rels",
      rels.replace(
        /<\/Relationships>/i,
        `<Relationship Id="rId${nextId}" Type="${SST_REL_TYPE}" Target="sharedStrings.xml"/></Relationships>`,
      ),
    );
  }
};

const parseAttr = (attrs: string, name: string) =>
  attrs.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];

const decodeSharedText = (xml: string) =>
  [...xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/gi)]
    .map((part) =>
      part[1]
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#10;/g, "\n")
        .replace(/&#xA;/gi, "\n")
        .replace(/_x000A_/gi, "\n")
        .replace(/_x000D_/gi, "\n")
        .replace(/&amp;/g, "&"),
    )
    .join("");

const parseSharedStringList = (sstXml: string) =>
  [...sstXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/gi)].map((match) =>
    decodeSharedText(match[1]),
  );

const xfListFromStyles = (stylesXml: string) => {
  const body =
    stylesXml.match(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/i)?.[1] ?? "";
  return [...body.matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi)].map(
    (match) => match[0],
  );
};

const styleHasWrap = (stylesXml: string, styleId: string) => {
  const index = Number(styleId);
  if (!Number.isInteger(index) || index < 0) return false;
  const xf = xfListFromStyles(stylesXml)[index];
  return Boolean(xf && /wrapText="(?:1|true)"/i.test(xf));
};

const cellTextHasBreak = (
  cellXml: string,
  attrs: string,
  strings: string[],
) => {
  const type = parseAttr(attrs, "t") || "";
  if (type === "s") {
    const index = Number(cellXml.match(/<v\b[^>]*>(\d+)<\/v>/i)?.[1]);
    const text = Number.isInteger(index) ? strings[index] ?? "" : "";
    return text.includes("\n") || text.includes("\r");
  }
  return /(?:&#10;|&#xA;|_x000A_|\n|\r)/i.test(cellXml);
};

const mostFrequentId = (ids: string[]) => {
  const counts = new Map<string, number>();
  for (const id of ids) counts.set(id, (counts.get(id) ?? 0) + 1);
  let best = ids[0] ?? "";
  let bestCount = 0;
  for (const [id, count] of counts) {
    if (count > bestCount) {
      best = id;
      bestCount = count;
    }
  }
  return best || undefined;
};

const xfXmlForStyle = (stylesXml: string, styleId: string) => {
  const index = Number(styleId);
  if (!Number.isInteger(index) || index < 0) return "";
  return xfListFromStyles(stylesXml)[index] ?? "";
};

const fillsFromStyles = (stylesXml: string) => {
  const body = stylesXml.match(/<fills\b[^>]*>([\s\S]*?)<\/fills>/i)?.[1] ?? "";
  return [...body.matchAll(/<fill\b[^>]*(?:\/>|>[\s\S]*?<\/fill>)/gi)].map(
    (match) => match[0],
  );
};

/** Жовта / кольорова заливка шаблону — не «звичайний» стиль рядка даних. */
const styleHasHighlightFill = (stylesXml: string, styleId: string) => {
  const xf = xfXmlForStyle(stylesXml, styleId);
  if (!xf) return false;
  const fillId = Number(xf.match(/\bfillId="(\d+)"/i)?.[1] ?? 0);
  const fill = fillsFromStyles(stylesXml)[fillId] ?? "";
  if (!fill) return fillId > 0;
  if (/patternType="(?:none|gray125)"/i.test(fill)) return false;
  return /patternType="solid"|<fgColor\b|<bgColor\b/i.test(fill);
};

const styleIsRegularCentered = (stylesXml: string, styleId: string) => {
  const xf = xfXmlForStyle(stylesXml, styleId);
  if (!xf) return false;
  if (fontIdIsBold(stylesXml, xfFontId(xf))) return false;
  if (styleHasHighlightFill(stylesXml, styleId)) return false;
  return (
    /vertical="center"/i.test(xf) && /horizontal="center"/i.test(xf)
  );
};

const styleIsRegularFont = (stylesXml: string, styleId: string) => {
  const xf = xfXmlForStyle(stylesXml, styleId);
  if (!xf) return false;
  return !fontIdIsBold(stylesXml, xfFontId(xf));
};

/**
 * Стиль «як у всіх»: найчастіший звичайний (не bold) + центр у колонці.
 * Найближчий сусід часто вже зіпсований попереднім записом.
 */
const findNearestColumnStyleId = (
  sheetXml: string,
  stylesXml: string,
  columnLetter: string,
  preferredRow: number,
  skipRows: Iterable<number>,
  _needWrap: boolean,
) => {
  const skip = new Set(skipRows);
  const cellRe = new RegExp(
    `<c\\b([^<>]*\\br="${columnLetter}(\\d+)"(?![0-9A-Za-z])[^<>]*)(\\/\\s*>|>[\\s\\S]*?<\\/c>)`,
    "gi",
  );
  const centered: string[] = [];
  const regular: string[] = [];
  const rows: number[] = [];
  const ids = new Map<number, string>();
  for (const cell of sheetXml.matchAll(cellRe)) {
    const row = Number(cell[2]);
    if (!row || row < 6) continue;
    const styleId = parseAttr(cell[1], "s");
    if (!styleId) continue;
    if (!cellHasContent(cell[0])) continue;
    // Прочерки з правильним стилем теж шаблон; жирні вже зіпсовані — не вчимось.
    if (styleIsRegularCentered(stylesXml, styleId)) {
      centered.push(styleId);
      continue;
    }
    if (skip.has(row)) continue;
    if (styleHasHighlightFill(stylesXml, styleId)) continue;
    rows.push(row);
    ids.set(row, styleId);
    if (styleIsRegularFont(stylesXml, styleId)) regular.push(styleId);
  }
  const canonical = mostFrequentId(centered) || mostFrequentId(regular);
  if (canonical) return canonical;
  const sheetCentered: string[] = [];
  const anyCellRe =
    /<c\b([^<>]*\br="[A-Z]{1,3}(\d+)"(?![0-9A-Za-z])[^<>]*)(\/\s*>|>[\s\S]*?<\/c>)/gi;
  for (const cell of sheetXml.matchAll(anyCellRe)) {
    const row = Number(cell[2]);
    if (!row || row < 6) continue;
    const styleId = parseAttr(cell[1], "s");
    if (!styleId || !cellHasContent(cell[0])) continue;
    if (styleIsRegularCentered(stylesXml, styleId)) sheetCentered.push(styleId);
  }
  const fromSheet = mostFrequentId(sheetCentered);
  if (fromSheet) return fromSheet;
  if (!rows.length) return undefined;
  const target = preferredRow || 7;
  return ids.get(
    rows.reduce((best, row) =>
      Math.abs(row - target) < Math.abs(best - target) ? row : best,
    ),
  );
};

const applyStyleOnlyToCell = (
  sheetXml: string,
  write: ZipCellWrite,
  styleId: string | undefined,
) => {
  if (!styleId) return sheetXml;
  const ref = `${columnNumberToLetter(write.column)}${write.row}`;
  // Не захоплювати `/` самозакритої клітинки: інакше виходить `<c r="F53"/ s="1">`.
  const cellRe = new RegExp(
    `<c\\b([^/<>]*?\\br="${ref}"(?![0-9A-Za-z])[^/<>]*)(\\/\\s*>|>[\\s\\S]*?<\\/c>)`,
    "i",
  );
  return sheetXml.replace(cellRe, (_full, attrs, tail) => {
    const cleaned = String(attrs).replace(/\/\s*$/, "");
    const next = /\bs="/i.test(cleaned)
      ? cleaned.replace(/\bs="[^"]*"/i, `s="${styleId}"`)
      : `${cleaned} s="${styleId}"`;
    return `<c${next}${tail}`;
  });
};

/** Найближчий до цільового рядка ООС із `s` і wrap у колонці історії. */
const findWrapTemplateRow = (
  sheetXml: string,
  sstXml: string,
  columnLetter: string,
  skipRows: Iterable<number>,
  stylesXml = "",
) => {
  const skip = new Set(skipRows);
  const near = [...skip].filter((row) => row > 0);
  const target = near.length ? Math.min(...near) : 7;
  const strings = parseSharedStringList(sstXml);
  const cellRe = new RegExp(
    `<c\\b([^<>]*\\br="${columnLetter}(\\d+)"(?![0-9A-Za-z])[^<>]*)(?:\\/>|>[\\s\\S]*?<\\/c>)`,
    "gi",
  );
  const pick = (candidates: number[]) => {
    if (!candidates.length) return 0;
    return candidates.reduce((best, row) =>
      Math.abs(row - target) < Math.abs(best - target) ? row : best,
    );
  };
  const wrapWithBreak: number[] = [];
  const wrapOnly: number[] = [];
  const styled: number[] = [];
  for (const cell of sheetXml.matchAll(cellRe)) {
    const row = Number(cell[2]);
    if (!row || row < 7 || skip.has(row)) continue;
    const styleId = parseAttr(cell[1], "s");
    if (!styleId) continue;
    styled.push(row);
    const wrapped = styleHasWrap(stylesXml, styleId);
    if (wrapped) wrapOnly.push(row);
    const type = parseAttr(cell[1], "t") || "";
    let text = "";
    if (type === "s") {
      const index = Number(cell[0].match(/<v\b[^>]*>(\d+)<\/v>/i)?.[1]);
      text = Number.isInteger(index) ? strings[index] ?? "" : "";
    } else if (type === "inlineStr" || /<is\b/i.test(cell[0])) {
      text = decodeSharedText(cell[0]);
    }
    if ((text.includes("\n") || text.includes("\r")) && wrapped) {
      wrapWithBreak.push(row);
    }
  }
  return pick(wrapWithBreak) || pick(wrapOnly) || pick(styled);
};

export const columnNumberToLetter = (columnNumber: number) => {
  let label = "";
  let value = columnNumber;
  while (value > 0) {
    const rem = (value - 1) % 26;
    label = String.fromCharCode(65 + rem) + label;
    value = Math.floor((value - 1) / 26);
  }
  return label;
};

const columnLetterToNumber = (letter: string) => {
  let index = 0;
  for (const char of letter.toUpperCase()) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return index;
};

const buildCellXml = (
  ref: string,
  styleId: string | undefined,
  value: string | number | null,
  sharedStringIndex?: number,
) => {
  const styleAttr = styleId != null && styleId !== "" ? ` s="${styleId}"` : "";
  if (value === null) return `<c r="${ref}"${styleAttr}/>`;
  if (typeof value === "number" && Number.isFinite(value)) {
    return `<c r="${ref}"${styleAttr}><v>${value}</v></c>`;
  }
  if (sharedStringIndex != null) {
    return `<c r="${ref}"${styleAttr} t="s"><v>${sharedStringIndex}</v></c>`;
  }
  const text = String(value);
  const space =
    text.startsWith(" ") || text.endsWith(" ") || /[\n\t]/.test(text)
      ? ` xml:space="preserve"`
      : "";
  return `<c r="${ref}"${styleAttr} t="inlineStr"><is><t${space}>${escapeXml(text)}</t></is></c>`;
};

export const resolveSheetPath = async (
  zip: JSZip,
  sheetNameOrRe: string | RegExp,
) => {
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

  const matcher =
    typeof sheetNameOrRe === "string"
      ? (name: string) => name === sheetNameOrRe
      : (name: string) => sheetNameOrRe.test(name);

  for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*)>/gi)) {
    const name = sheet[1].match(/\bname="([^"]+)"/)?.[1] ?? "";
    if (!matcher(name)) continue;
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

/** Повний блок `<row r="N">…</row>` або self-closing. */
const findRowBlock = (sheetXml: string, rowNum: number) => {
  const startRe = new RegExp(
    `<row\\b[^>]*\\br="${rowNum}"(?![0-9])[^>]*>`,
    "i",
  );
  const startMatch = startRe.exec(sheetXml);
  if (!startMatch) return null;
  const start = startMatch.index;
  const openTag = startMatch[0];
  if (/\/\s*>$/.test(openTag)) {
    return { full: openTag, start, end: start + openTag.length };
  }
  const closeTag = "</row>";
  const closeAt = sheetXml.indexOf(closeTag, start + openTag.length);
  if (closeAt < 0) return null;
  const end = closeAt + closeTag.length;
  return { full: sheetXml.slice(start, end), start, end };
};

const replaceOrInsertCellInRow = (
  rowXml: string,
  column: number,
  cellXml: string,
) => {
  const openMatch = rowXml.match(/^<row\b[^>]*>/i);
  if (!openMatch) return rowXml;
  const openTag = openMatch[0];
  if (/\/\s*>$/.test(openTag)) {
    const openOnly = openTag.replace(/\s*\/\s*>$/, ">");
    return `${openOnly}${cellXml}</row>`;
  }

  const closeIdx = rowXml.lastIndexOf("</row>");
  if (closeIdx < 0) return rowXml;
  const inner = rowXml.slice(openTag.length, closeIdx);
  const cellRe = /<c\b([^<>]*?)(\/>|>[\s\S]*?<\/c>)/gi;
  const cells: Array<{ full: string; col: number; index: number }> = [];
  let match: RegExpExecArray | null;
  while ((match = cellRe.exec(inner))) {
    const attrs = match[1];
    const ref = parseAttr(attrs, "r") || "";
    const letters = ref.match(/^([A-Z]+)/i)?.[1] ?? "";
    const col = letters ? columnLetterToNumber(letters) : 0;
    cells.push({ full: match[0], col, index: match.index });
  }

  const existing = cells.find((cell) => cell.col === column);
  if (existing) {
    return (
      openTag +
      inner.slice(0, existing.index) +
      cellXml +
      inner.slice(existing.index + existing.full.length) +
      "</row>"
    );
  }

  const insertBefore = cells.find((cell) => cell.col > column);
  if (!insertBefore) {
    return `${openTag}${inner}${cellXml}</row>`;
  }
  return (
    openTag +
    inner.slice(0, insertBefore.index) +
    cellXml +
    inner.slice(insertBefore.index) +
    "</row>"
  );
};

const cellReferenceParts = (cellXml: string) => {
  const ref = cellXml.match(/<c\b[^<>]*\br="([A-Z]+)([0-9]+)"/i);
  if (!ref) return null;
  return {
    column: ref[1].toUpperCase(),
    row: Number(ref[2]),
    columnNumber: columnLetterToNumber(ref[1]),
  };
};

const cellHasContent = (cellXml: string) => /<(?:v|f|is)\b/i.test(cellXml);

const updateRowSpans = (openTag: string, cells: string[]) => {
  const columns = cells
    .map((cellXml) => cellReferenceParts(cellXml)?.columnNumber ?? 0)
    .filter(Boolean);
  if (!columns.length) return openTag.replace(/\s+spans="[^"]*"/i, "");
  const spans = `spans="${Math.min(...columns)}:${Math.max(...columns)}"`;
  if (/\bspans="[^"]*"/i.test(openTag)) {
    return openTag.replace(/\bspans="[^"]*"/i, spans);
  }
  return openTag.replace(/>$/, ` ${spans}>`);
};

const insertRowXml = (sheetXml: string, rowNumber: number, rowXml: string) => {
  for (const row of sheetXml.matchAll(
    /<row\b[^>]*\br="([0-9]+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi,
  )) {
    const existingRowNumber = Number(row[1]);
    if (existingRowNumber > rowNumber) {
      return `${sheetXml.slice(0, row.index)}${rowXml}${sheetXml.slice(row.index)}`;
    }
  }

  const sheetDataClose = sheetXml.search(/<\/sheetData>/i);
  if (sheetDataClose >= 0) {
    return `${sheetXml.slice(0, sheetDataClose)}${rowXml}${sheetXml.slice(sheetDataClose)}`;
  }
  return sheetXml;
};

const normalizeWorksheetRows = (sheetXml: string) => {
  const dataOpen = sheetXml.search(/<sheetData\b[^>]*>/i);
  const dataClose = sheetXml.search(/<\/sheetData>/i);
  const dataStart =
    dataOpen >= 0
      ? dataOpen + (sheetXml.match(/<sheetData\b[^>]*>/i)?.[0].length ?? 0)
      : 0;
  const dataEnd = dataClose >= 0 ? dataClose : sheetXml.length;
  const dataXml = sheetXml.slice(dataStart, dataEnd);
  const rowCells = new Map<number, Map<number, string>>();
  for (const cell of dataXml.matchAll(/<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi)) {
    const cellXml = cell[0];
    const ref = cellReferenceParts(cellXml);
    if (!ref) continue;
    if (ref.columnNumber > 256 && !cellHasContent(cellXml)) continue;
    const cells = rowCells.get(ref.row) ?? new Map<number, string>();
    cells.set(ref.columnNumber, cellXml);
    rowCells.set(ref.row, cells);
  }

  const existingRows = new Set<number>();
  let normalized = sheetXml.replace(
    /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi,
    (rowXml) => {
      const openMatch = rowXml.match(/^<row\b[^>]*>/i);
      if (!openMatch) return rowXml;
      const rowNumber = Number(parseAttr(openMatch[0], "r") || 0);
      if (!rowNumber) return rowXml;
      existingRows.add(rowNumber);

      const isSelfClosing = /\/\s*>$/.test(openMatch[0]);
      const inner = isSelfClosing
        ? ""
        : rowXml.slice(openMatch[0].length, rowXml.lastIndexOf("</row>"));
      const cellRe = /<c\b[^<>]*?(?:\/>|>[\s\S]*?<\/c>)/gi;
      const cells = [...(rowCells.get(rowNumber)?.entries() ?? [])]
        .sort(([a], [b]) => a - b)
        .map(([, cellXml]) => cellXml);
      const withoutCells = inner.replace(cellRe, "");
      return `${updateRowSpans(openMatch[0], cells)}${withoutCells}${cells.join("")}</row>`;
    },
  );

  [...rowCells.keys()]
    .filter((rowNumber) => !existingRows.has(rowNumber))
    .sort((a, b) => a - b)
    .forEach((rowNumber) => {
      const cells = [...(rowCells.get(rowNumber)?.entries() ?? [])]
        .sort(([a], [b]) => a - b)
        .map(([, cellXml]) => cellXml);
      normalized = insertRowXml(
        normalized,
        rowNumber,
        `<row r="${rowNumber}" ${updateRowSpans(">", cells).slice(1)}${cells.join("")}</row>`,
      );
    });

  return normalized;
};

const stripOrphanSharedFormulas = (sheetXml: string) => {
  const masterSharedFormulaIds = new Set<string>();
  for (const formula of sheetXml.matchAll(
    /<f\b((?:(?!\/>)[^<>])*)>([\s\S]*?)<\/f>/gi,
  )) {
    const attrs = formula[1];
    if (!/\bt="shared"/i.test(attrs)) continue;
    const si = parseAttr(attrs, "si");
    if (si && formula[2].trim()) masterSharedFormulaIds.add(si);
  }

  return sheetXml.replace(
    /<f\b([^<>]*?\bt="shared"[^<>]*?)(?:\/>|>[\s\S]*?<\/f>)/gi,
    (formulaXml, attrs) => {
      const si = parseAttr(String(attrs), "si");
      if (si && !masterSharedFormulaIds.has(si)) return "";
      return formulaXml;
    },
  );
};

type SheetDataSplit = {
  prefix: string;
  suffix: string;
  parts: Array<{ row?: number; xml: string }>;
  indexByRow: Map<number, number>;
};

const splitSheetDataRows = (sheetXml: string): SheetDataSplit | null => {
  const open = sheetXml.search(/<sheetData\b[^>]*>/i);
  const close = sheetXml.search(/<\/sheetData>/i);
  if (open < 0 || close < 0) return null;
  const openTag = sheetXml.match(/<sheetData\b[^>]*>/i)?.[0];
  if (!openTag) return null;
  const prefix = sheetXml.slice(0, open + openTag.length);
  const suffix = sheetXml.slice(close);
  const body = sheetXml.slice(open + openTag.length, close);
  const parts: Array<{ row?: number; xml: string }> = [];
  const indexByRow = new Map<number, number>();
  const rowRe = /<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi;
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = rowRe.exec(body))) {
    if (match.index > last) {
      parts.push({ xml: body.slice(last, match.index) });
    }
    const rowNum = Number(parseAttr(match[0], "r") || 0);
    if (rowNum) indexByRow.set(rowNum, parts.length);
    parts.push({ row: rowNum || undefined, xml: match[0] });
    last = match.index + match[0].length;
  }
  if (last < body.length) parts.push({ xml: body.slice(last) });
  return { prefix, suffix, parts, indexByRow };
};

const joinSheetData = (split: SheetDataSplit) =>
  `${split.prefix}${split.parts.map((part) => part.xml).join("")}${split.suffix}`;

const ownStyleFromRowXml = (rowXml: string, column: number, row: number) => {
  const ref = `${columnNumberToLetter(column)}${row}`;
  const cell = rowXml.match(
    new RegExp(`<c\\b([^<>]*?\\br="${ref}"(?![0-9A-Za-z])[^<>]*)`, "i"),
  );
  return cell ? parseAttr(cell[1], "s") : undefined;
};

const applyHeightToRowXml = (rowXml: string, ht: number) => {
  const height = String(Math.min(220, Math.max(15, ht)));
  return rowXml.replace(/^<row\b([^>]*)/i, (_open, attrs) => {
    const without = String(attrs)
      .replace(/\s+ht="[^"]*"/i, "")
      .replace(/\s+customHeight="[^"]*"/i, "");
    return `<row${without} ht="${height}" customHeight="1"`;
  });
};

const upsertCellInSheetXml = (
  sheetXml: string,
  write: ZipCellWrite,
): string => {
  const letter = columnNumberToLetter(write.column);
  const ref = `${letter}${write.row}`;
  const cellRe = new RegExp(
    `<c\\b([^<>]*?\\br="${ref}"(?![0-9A-Za-z])[^<>]*?)(\\/>|>[\\s\\S]*?<\\/c>)`,
    "i",
  );
  const existing = sheetXml.match(cellRe);
  const styleRef = write.styleSourceRow
    ? `${letter}${write.styleSourceRow}`
    : "";
  const styleCell = styleRef
    ? sheetXml.match(
        new RegExp(
          `<c\\b([^<>]*?\\br="${styleRef}"(?![0-9A-Za-z])[^<>]*?)(\\/>|>[\\s\\S]*?<\\/c>)`,
          "i",
        ),
      )
    : null;
  const styleId =
    write.styleId ||
    (styleCell
      ? parseAttr(styleCell[1], "s")
      : existing
        ? parseAttr(existing[1], "s")
        : undefined);
  const cellXml = buildCellXml(
    ref,
    styleId,
    write.value,
    write.sharedStringIndex,
  );

  if (existing) {
    return sheetXml.replace(existing[0], cellXml);
  }

  const rowBlock = findRowBlock(sheetXml, write.row);
  if (rowBlock) {
    const nextRow = replaceOrInsertCellInRow(
      rowBlock.full,
      write.column,
      cellXml,
    );
    return (
      sheetXml.slice(0, rowBlock.start) +
      nextRow +
      sheetXml.slice(rowBlock.end)
    );
  }

  // Рядки в sheetData мусять іти за зростанням `r`, інакше Excel відкидає
  // клітинки («Removed Records: Cell information»).
  return insertRowXml(
    sheetXml,
    write.row,
    `<row r="${write.row}">${cellXml}</row>`,
  );
};

/** Страховка: після всіх правок рядки sheetData впорядковані за `r`. */
const sortSheetDataRows = (sheetXml: string) => {
  const open = sheetXml.search(/<sheetData\b[^>]*>/i);
  const close = sheetXml.search(/<\/sheetData>/i);
  if (open < 0 || close < 0) return sheetXml;
  const openTag = sheetXml.match(/<sheetData\b[^>]*>/i)?.[0] ?? "<sheetData>";
  const body = sheetXml.slice(open + openTag.length, close);
  const rows = [...body.matchAll(/<row\b[^>]*(?:\/>|>[\s\S]*?<\/row>)/gi)].map(
    (match) => match[0],
  );
  if (rows.length < 2) return sheetXml;
  const numbers = rows.map((rowXml) => Number(parseAttr(rowXml, "r") || 0));
  const sorted = numbers.every(
    (value, index) => index === 0 || numbers[index - 1] <= value,
  );
  if (sorted) return sheetXml;
  const ordered = rows
    .map((rowXml, index) => ({ rowXml, order: numbers[index], index }))
    .sort((a, b) => a.order - b.order || a.index - b.index)
    .map((entry) => entry.rowXml)
    .join("");
  return (
    sheetXml.slice(0, open + openTag.length) + ordered + sheetXml.slice(close)
  );
};

const xfAppliesAlignment = (xfXml: string) =>
  /applyAlignment="(?:1|true)"/i.test(xfXml);

const isCenteredWrapXf = (xfXml: string) =>
  xfAppliesAlignment(xfXml) &&
  /wrapText="(?:1|true)"/i.test(xfXml) &&
  /vertical="center"/i.test(xfXml) &&
  /horizontal="center"/i.test(xfXml);

const fontsFromStyles = (stylesXml: string) => {
  const body = stylesXml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/i)?.[1] ?? "";
  return [...body.matchAll(/<font\b[^>]*(?:\/>|>[\s\S]*?<\/font>)/gi)].map(
    (match) => match[0],
  );
};

const fontIdIsBold = (stylesXml: string, fontId: string) => {
  const font = fontsFromStyles(stylesXml)[Number(fontId)];
  return Boolean(font && /<b\b/i.test(font));
};

const regularFontId = (stylesXml: string) => {
  const fonts = fontsFromStyles(stylesXml);
  const index = fonts.findIndex((font) => !/<b\b/i.test(font));
  return index >= 0 ? String(index) : "0";
};

const xfFontId = (xfXml: string) => xfXml.match(/\bfontId="(\d+)"/i)?.[1] ?? "0";

const withRegularFont = (xfXml: string, stylesXml: string) => {
  const fontId = xfFontId(xfXml);
  if (!fontIdIsBold(stylesXml, fontId)) return xfXml;
  const nextFont = regularFontId(stylesXml);
  let xf = xfXml.replace(/\bfontId="[^"]*"/i, `fontId="${nextFont}"`);
  if (/applyFont=/i.test(xf)) {
    xf = xf.replace(/applyFont="[^"]*"/i, `applyFont="1"`);
  } else {
    xf = xf.replace(/<xf\b/, `<xf applyFont="1"`);
  }
  return xf;
};

const withWrapAlignment = (xfXml: string) => {
  let xf = xfXml;
  if (/applyAlignment=/i.test(xf)) {
    xf = xf.replace(/applyAlignment="[^"]*"/i, `applyAlignment="1"`);
  } else {
    xf = xf.replace(/<xf\b/, `<xf applyAlignment="1"`);
  }
  if (/<alignment\b/i.test(xf)) {
    return xf.replace(
      /<alignment\b([^>]*)(?:\/>|><\/alignment>|>)/i,
      (_all, attrs: string) => {
        const cleaned = String(attrs)
          .replace(/\s+wrapText="[^"]*"/gi, "")
          .replace(/\s+vertical="[^"]*"/gi, "")
          .replace(/\s+horizontal="[^"]*"/gi, "")
          .replace(/\/\s*$/, "")
          .trimEnd();
        return `<alignment${cleaned} wrapText="1" vertical="center" horizontal="center"/>`;
      },
    );
  }
  const alignment = `<alignment wrapText="1" vertical="center" horizontal="center"/>`;
  if (/\/\s*>\s*$/.test(xf.trim())) {
    return xf.replace(/\/\s*>\s*$/, `>${alignment}</xf>`);
  }
  return xf.replace(/<xf\b([^>]*)>/, `<xf$1>${alignment}`);
};

const withCenterAlignment = (xfXml: string) => {
  let xf = xfXml;
  if (/applyAlignment=/i.test(xf)) {
    xf = xf.replace(/applyAlignment="[^"]*"/i, `applyAlignment="1"`);
  } else {
    xf = xf.replace(/<xf\b/, `<xf applyAlignment="1"`);
  }
  if (/<alignment\b/i.test(xf)) {
    return xf.replace(
      /<alignment\b([^>]*)(?:\/>|><\/alignment>|>)/i,
      (_all, attrs: string) => {
        const cleaned = String(attrs)
          .replace(/\s+vertical="[^"]*"/gi, "")
          .replace(/\s+horizontal="[^"]*"/gi, "")
          .replace(/\/\s*$/, "")
          .trimEnd();
        return `<alignment${cleaned} vertical="center" horizontal="center"/>`;
      },
    );
  }
  const alignment = `<alignment vertical="center" horizontal="center"/>`;
  if (/\/\s*>\s*$/.test(xf.trim())) {
    return xf.replace(/\/\s*>\s*$/, `>${alignment}</xf>`);
  }
  return xf.replace(/<xf\b([^>]*)>/, `<xf$1>${alignment}`);
};

const ensurePlainCenteredStyle = (
  stylesXml: string,
  sourceStyleId: string | undefined,
) => {
  const block = stylesXml.match(/<cellXfs\b([^>]*)>([\s\S]*?)<\/cellXfs>/i);
  if (!block) return { xml: stylesXml, styleId: sourceStyleId };
  const xfs = [
    ...block[2].matchAll(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/gi),
  ].map((match) => match[0]);
  const sourceIndex = Number(sourceStyleId);
  const sourceXf =
    Number.isInteger(sourceIndex) && sourceIndex >= 0 && sourceIndex < xfs.length
      ? xfs[sourceIndex]
      : undefined;
  const fallbackFont = regularFontId(stylesXml);
  const next = withRegularFont(
    withCenterAlignment(
      sourceXf ||
        `<xf numFmtId="0" fontId="${fallbackFont}" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"/>`,
    ),
    stylesXml,
  );
  if (sourceXf && next === sourceXf) {
    return { xml: stylesXml, styleId: String(sourceIndex) };
  }
  const existing = xfs.findIndex((xf) => xf === next);
  if (existing >= 0) return { xml: stylesXml, styleId: String(existing) };
  const xml = stylesXml.replace(
    /<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/i,
    `<cellXfs count="${xfs.length + 1}">${block[2]}${next}</cellXfs>`,
  );
  return { xml, styleId: String(xfs.length) };
};

const xfAttr = (xfXml: string | undefined, name: string, fallback: string) =>
  xfXml?.match(new RegExp(`\\b${name}="(\\d+)"`, "i"))?.[1] ?? fallback;

/** Новий xf: wrap + центр + звичайний шрифт. Не патчимо старий XML — Excel ігнорує alignment без applyAlignment="1". */
const buildWrapCenterRegularXf = (
  sourceXf: string | undefined,
  stylesXml: string,
) => {
  const numFmtId = xfAttr(sourceXf, "numFmtId", "0");
  const fillId = xfAttr(sourceXf, "fillId", "0");
  const borderId = xfAttr(sourceXf, "borderId", "0");
  const xfId = xfAttr(sourceXf, "xfId", "0");
  const fontId = regularFontId(stylesXml);
  const applyNumberFormat =
    numFmtId !== "0" || /applyNumberFormat="(?:1|true)"/i.test(sourceXf || "")
      ? ` applyNumberFormat="1"`
      : "";
  const applyFill = /applyFill="(?:1|true)"/i.test(sourceXf || "")
    ? ` applyFill="1"`
    : "";
  const applyBorder =
    borderId !== "0" || /applyBorder="(?:1|true)"/i.test(sourceXf || "")
      ? ` applyBorder="1"`
      : "";
  return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="${xfId}" applyFont="1"${applyNumberFormat}${applyFill}${applyBorder} applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>`;
};

const cellXfsBlock = (stylesXml: string) =>
  stylesXml.match(
    /<(?:[\w]+:)?cellXfs\b([^>]*)>([\s\S]*?)<\/(?:[\w]+:)?cellXfs>/i,
  );

const countXfChildren = (body: string) =>
  (body.match(/<(?:[\w]+:)?xf\b/gi) ?? []).length;

/** Копія стилю клітинки + wrap + центр. Не чіпаємо styles.xml, якщо розбір xf ненадійний. */
const ensureWrapOnStyle = (
  stylesXml: string,
  sourceStyleId: string | undefined,
) => {
  const block = cellXfsBlock(stylesXml);
  if (!block) {
    return { xml: stylesXml, styleId: sourceStyleId };
  }
  const xfs = [
    ...block[2].matchAll(
      /<(?:[\w]+:)?xf\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[\w]+:)?xf>)/gi,
    ),
  ].map((match) => match[0]);
  if (!xfs.length || xfs.length !== countXfChildren(block[2])) {
    return { xml: stylesXml, styleId: sourceStyleId };
  }
  const sourceIndex = Number(sourceStyleId);
  const sourceXf =
    Number.isInteger(sourceIndex) && sourceIndex >= 0 && sourceIndex < xfs.length
      ? xfs[sourceIndex]
      : undefined;
  if (
    sourceXf &&
    isCenteredWrapXf(sourceXf) &&
    !fontIdIsBold(stylesXml, xfFontId(sourceXf))
  ) {
    return { xml: stylesXml, styleId: String(sourceIndex) };
  }
  const wrapped = buildWrapCenterRegularXf(sourceXf, stylesXml);
  const existing = xfs.findIndex((xf) => xf === wrapped);
  if (existing >= 0) return { xml: stylesXml, styleId: String(existing) };
  const open = block[0].slice(0, block[0].indexOf(">") + 1);
  const close = block[0].match(/<\/(?:[\w]+:)?cellXfs>/i)?.[0];
  if (!close) return { xml: stylesXml, styleId: sourceStyleId };
  const nextOpen = /\bcount="/i.test(open)
    ? open.replace(/\bcount="[^"]*"/i, `count="${xfs.length + 1}"`)
    : open.replace(/>$/, ` count="${xfs.length + 1}">`);
  const xml = stylesXml.replace(
    block[0],
    `${nextOpen}${block[2]}${wrapped}${close}`,
  );
  return { xml, styleId: String(xfs.length) };
};

const setRowHeightValue = (
  sheetXml: string,
  rowNumber: number,
  ht: number,
) => {
  const height = String(Math.min(220, Math.max(15, ht)));
  const re = new RegExp(`<row\\b([^>]*\\br="${rowNumber}"(?![0-9])[^>]*)`, "i");
  return sheetXml.replace(re, (_open, attrs) => {
    const without = String(attrs)
      .replace(/\s+ht="[^"]*"/i, "")
      .replace(/\s+customHeight="[^"]*"/i, "");
    return `<row${without} ht="${height}" customHeight="1"`;
  });
};

const setRowHeightForLines = (
  sheetXml: string,
  rowNumber: number,
  lines: number,
) => setRowHeightValue(sheetXml, rowNumber, lines * 16);

const copyRowHeightFrom = (
  sheetXml: string,
  fromRow: number,
  toRow: number,
  minLines: number,
) => {
  const computed = Math.min(220, Math.max(18, minLines * 16));
  if (fromRow === toRow) {
    return setRowHeightValue(sheetXml, toRow, computed);
  }
  const source = sheetXml.match(
    new RegExp(`<row\\b([^>]*\\br="${fromRow}"(?![0-9])[^>]*)`, "i"),
  );
  const sourceHt = source ? Number(parseAttr(source[1], "ht")) : NaN;
  const ht =
    Number.isFinite(sourceHt) && sourceHt > 0
      ? minLines <= 3
        ? sourceHt
        : Math.max(sourceHt, computed)
      : computed;
  return setRowHeightValue(sheetXml, toRow, ht);
};

const yieldToUi = () =>
  new Promise<void>((resolve) => {
    window.setTimeout(resolve, 0);
  });

/**
 * Точковий запис текстових клітинок у аркуш .xlsx без xlsx-populate.
 * Оновлення йдуть по рядках у пам'яті, без копіювання всього XML на кожну клітинку.
 */
export async function applyInlineStringWritesToWorkbook(
  file: Blob | File,
  sheetNameOrRe: string | RegExp,
  writes: ZipCellWrite[],
  options?: {
    onProgress?: (done: number, total: number) => void;
    onStatus?: (message: string) => void;
  },
): Promise<Blob> {
  if (!writes.length) return file;

  const zip = await JSZip.loadAsync(await file.arrayBuffer());
  const originalStylesXml =
    (await zip.file("xl/styles.xml")?.async("string")) || "";
  let stylesXml = originalStylesXml;
  let sstXml = (await zip.file("xl/sharedStrings.xml")?.async("string")) || "";
  const sstIndexByText = new Map<string, number>();
  let sstUniqueCount = (sstXml.match(/<si\b/gi) ?? []).length;
  const withSharedString = (write: ZipCellWrite): ZipCellWrite => {
    if (write.value === null || typeof write.value === "number") return write;
    if (write.sharedStringIndex != null) return write;
    const text = String(write.value);
    const cached = sstIndexByText.get(text);
    if (cached != null) return { ...write, sharedStringIndex: cached };
    const added = appendSharedString(sstXml, text, sstUniqueCount);
    sstXml = added.xml;
    sstUniqueCount = added.index + 1;
    sstIndexByText.set(text, added.index);
    return { ...write, sharedStringIndex: added.index };
  };
  const sheetPath = await resolveSheetPath(zip, sheetNameOrRe);
  if (!sheetPath) {
    throw new Error("Не знайдено аркуш у файлі для запису клітинок.");
  }
  const loadedSheetXml = await zip.file(sheetPath)?.async("string");
  if (!loadedSheetXml) {
    throw new Error(`Не вдалося прочитати аркуш ${sheetPath}.`);
  }
  let sheetXml = loadedSheetXml;

  const skipStyleRows = new Set(writes.map((write) => write.row));
  const needsNeighborStyle = writes.some((write) => write.copyNeighborStyle);
  const xmlTemplateRow = needsNeighborStyle
    ? findWrapTemplateRow(sheetXml, sstXml, "D", skipStyleRows, stylesXml)
    : 0;
  const resolvedWrites = writes.map((write) => {
    if (!write.copyNeighborStyle) return write;
    const fallback =
      write.styleSourceRow && write.styleSourceRow >= 7
        ? write.styleSourceRow
        : 7;
    const template = xmlTemplateRow || fallback;
    return {
      ...write,
      styleSourceRow: write.styleSourceRow || template,
    };
  });
  const byRow = new Map<number, ZipCellWrite[]>();
  for (const write of resolvedWrites) {
    const list = byRow.get(write.row) ?? [];
    list.push(write);
    byRow.set(write.row, list);
  }
  const split = splitSheetDataRows(sheetXml);
  const existingRowNums = new Set(
    split ? [...split.indexByRow.keys()] : [],
  );
  const styleIdBySource = new Map<string, string | undefined>();
  const wrapStyleBySource = new Map<string, string | undefined>();
  const needsSstParse = writes.some((write) => write.styleOnly);
  const sstStrings = needsSstParse ? parseSharedStringList(sstXml) : [];
  const writeNeedsWrap = (write: ZipCellWrite, rowXml?: string) => {
    if (write.wrapText) return true;
    if (String(write.value ?? "").includes("\n")) return true;
    if (!write.styleOnly) return false;
    const ref = `${columnNumberToLetter(write.column)}${write.row}`;
    const haystack = rowXml || sheetXml;
    const existing = haystack.match(
      new RegExp(
        `<c\\b([^/<>]*?\\br="${ref}"(?![0-9A-Za-z])[^/<>]*)(\\/\\s*>|>[\\s\\S]*?<\\/c>)`,
        "i",
      ),
    );
    return Boolean(
      existing && cellTextHasBreak(existing[0], existing[1], sstStrings),
    );
  };
  const sourceStyleIdFor = (
    write: ZipCellWrite,
    rowXml?: string,
  ) => {
    if (write.styleId) return write.styleId;
    const needWrap = Boolean(
      write.copyNeighborStyle &&
        writeNeedsWrap(write, rowXml) &&
        !write.keepNeighborStyle,
    );
    const ownStyle = rowXml
      ? ownStyleFromRowXml(rowXml, write.column, write.row)
      : undefined;
    if (ownStyle && !write.copyNeighborStyle) return ownStyle;
    const styleColumn = write.styleSourceColumn || write.column;
    const cacheKey = write.copyNeighborStyle
      ? `canon:${styleColumn}`
      : `${write.styleSourceRow ?? ""}:${styleColumn}:${needWrap ? "w" : "n"}`;
    if (styleIdBySource.has(cacheKey)) {
      return styleIdBySource.get(cacheKey);
    }
    const letter = columnNumberToLetter(styleColumn);
    const styleId = write.copyNeighborStyle
      ? findNearestColumnStyleId(
          sheetXml,
          stylesXml,
          letter,
          write.styleSourceRow || 0,
          skipStyleRows,
          needWrap,
        )
      : (() => {
          const styleRef = write.styleSourceRow
            ? `${letter}${write.styleSourceRow}`
            : "";
          const styleCell = styleRef
            ? sheetXml.match(
                new RegExp(
                  `<c\\b([^<>]*?\\br="${styleRef}"(?![0-9A-Za-z])[^<>]*?)(\\/>|>[\\s\\S]*?<\\/c>)`,
                  "i",
                ),
              )
            : null;
          return styleCell ? parseAttr(styleCell[1], "s") : ownStyle;
        })();
    styleIdBySource.set(cacheKey, styleId);
    return styleId;
  };
  const styleIdFor = (write: ZipCellWrite, rowXml?: string) => {
    const sourceId = sourceStyleIdFor(write, rowXml);
    // Новий xf у styles.xml ламає Excel (зняття Font/Format/Style з усієї книги).
    if (write.keepNeighborStyle) return sourceId;
    const needsWrap = writeNeedsWrap(write, rowXml) || Boolean(write.wrapText);
    if (!needsWrap || !stylesXml) return sourceId;
    const cacheKey = sourceId ?? "_none";
    if (wrapStyleBySource.has(cacheKey)) {
      return wrapStyleBySource.get(cacheKey);
    }
    const ensured = ensureWrapOnStyle(stylesXml, sourceId);
    stylesXml = ensured.xml;
    wrapStyleBySource.set(cacheKey, ensured.styleId);
    return ensured.styleId;
  };

  const newRows: Array<{ row: number; xml: string }> = [];
  const rows = [...byRow.keys()].sort((a, b) => a - b);
  const report = (done: number) => options?.onProgress?.(done, rows.length);

  if (split) {
    for (let index = 0; index < rows.length; index += 1) {
      if (index > 0 && index % 80 === 0) {
        report(index);
        await yieldToUi();
      }
      const row = rows[index]!;
      const rowWrites = (byRow.get(row) ?? []).sort(
        (a, b) => a.column - b.column,
      );
      const partIndex = split.indexByRow.get(row);
      if (partIndex == null) {
        const cells = rowWrites
          .filter((write) => !write.styleOnly)
          .map((write) => {
            const next = withSharedString(write);
            return buildCellXml(
              `${columnNumberToLetter(next.column)}${next.row}`,
              styleIdFor(next),
              next.value,
              next.sharedStringIndex,
            );
          });
        newRows.push({
          row,
          xml: `<row r="${row}">${cells.join("")}</row>`,
        });
        continue;
      }
      let rowXml = split.parts[partIndex]!.xml;
      for (const write of rowWrites) {
        if (write.styleOnly) {
          rowXml = applyStyleOnlyToCell(rowXml, write, styleIdFor(write, rowXml));
          continue;
        }
        const next = withSharedString(write);
        const cellXml = buildCellXml(
          `${columnNumberToLetter(next.column)}${next.row}`,
          styleIdFor(next, rowXml),
          next.value,
          next.sharedStringIndex,
        );
        rowXml = replaceOrInsertCellInRow(rowXml, next.column, cellXml);
      }
      const lines = Math.max(
        1,
        ...rowWrites.map((write) => String(write.value ?? "").split("\n").length),
      );
      if (lines > 1) {
        rowXml = applyHeightToRowXml(rowXml, lines * 16);
      }
      split.parts[partIndex] = { row, xml: rowXml };
    }
    report(rows.length);
    if (newRows.length) {
      const maxExisting = existingRowNums.size
        ? Math.max(...existingRowNums)
        : 0;
      if (newRows[0]!.row > maxExisting) {
        split.parts.push({ xml: newRows.map((item) => item.xml).join("") });
      } else {
        for (const item of newRows) {
          const insertAt = split.parts.findIndex(
            (part) => part.row != null && part.row > item.row,
          );
          const entry = { row: item.row, xml: item.xml };
          if (insertAt < 0) split.parts.push(entry);
          else {
            split.parts.splice(insertAt, 0, entry);
            for (const [rowNum, partIndex] of [...split.indexByRow]) {
              if (partIndex >= insertAt) {
                split.indexByRow.set(rowNum, partIndex + 1);
              }
            }
          }
          split.indexByRow.set(item.row, split.parts.indexOf(entry));
        }
      }
    }
    sheetXml = joinSheetData(split);
    if (newRows.length) {
      sheetXml = sortSheetDataRows(sheetXml);
    }
  } else {
    for (const row of rows) {
      const rowWrites = (byRow.get(row) ?? []).sort(
        (a, b) => a.column - b.column,
      );
      if (!existingRowNums.has(row)) {
        const cells = rowWrites
          .filter((write) => !write.styleOnly)
          .map((write) => {
            const next = withSharedString(write);
            return buildCellXml(
              `${columnNumberToLetter(next.column)}${next.row}`,
              styleIdFor(next),
              next.value,
              next.sharedStringIndex,
            );
          });
        newRows.push({
          row,
          xml: `<row r="${row}">${cells.join("")}</row>`,
        });
        continue;
      }
      for (const write of rowWrites) {
        if (write.styleOnly) {
          sheetXml = applyStyleOnlyToCell(sheetXml, write, styleIdFor(write));
          continue;
        }
        const next = withSharedString(write);
        sheetXml = upsertCellInSheetXml(sheetXml, {
          ...next,
          styleId: styleIdFor(next),
        });
      }
    }
    if (newRows.length) {
      const block = newRows.map((item) => item.xml).join("");
      const maxExisting = existingRowNums.size
        ? Math.max(...existingRowNums)
        : 0;
      if (newRows[0]!.row > maxExisting) {
        const close = sheetXml.search(/<\/sheetData>/i);
        sheetXml =
          close >= 0
            ? `${sheetXml.slice(0, close)}${block}${sheetXml.slice(close)}`
            : `${sheetXml}${block}`;
      } else {
        for (const item of newRows) {
          sheetXml = insertRowXml(sheetXml, item.row, item.xml);
        }
      }
      sheetXml = normalizeWorksheetRows(sheetXml);
      sheetXml = sortSheetDataRows(sheetXml);
    }
  }
  sheetXml = stripOrphanSharedFormulas(sheetXml);
  sheetXml = repairBrokenCellOpenTags(sheetXml);
  if (stylesXml && stylesXml !== originalStylesXml) {
    zip.file("xl/styles.xml", stylesXml);
  }
  if (sstXml) await ensureSharedStringsPart(zip, sstXml);
  zip.file(sheetPath, sheetXml);
  await stripStaleCalcChainFromZip(zip);
  options?.onStatus?.("Стискаю файл…");
  await yieldToUi();
  return zip.generateAsync({
    type: "blob",
    mimeType:
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    compression: "DEFLATE",
  });
}

const collectSheetCells = (sheetXml: string) => {
  const cellRe =
    /<c\b([^<>]*\br="([A-Z]{1,3})(\d+)"(?![0-9A-Za-z])[^<>]*)(\/\s*>|>[\s\S]*?<\/c>)/gi;
  const cells: Array<{
    row: number;
    column: number;
    styleId: string | undefined;
    hasContent: boolean;
  }> = [];
  for (const cell of sheetXml.matchAll(cellRe)) {
    const row = Number(cell[3]);
    if (!row) continue;
    cells.push({
      row,
      column: columnLetterToNumber(cell[2]),
      styleId: parseAttr(cell[1], "s"),
      hasContent: cellHasContent(cell[0]),
    });
  }
  return cells;
};

/**
 * ООС: найчастіший звичайний стиль колонки.
 */
export async function applyCanonicalDataStylesToSheets(
  file: Blob | File,
  sheetNameOrRes: Array<string | RegExp>,
): Promise<{ blob: Blob; styledCells: number }> {
  let blob: Blob = file instanceof Blob ? file : new Blob([file]);
  let styledCells = 0;
  for (const sheetRef of sheetNameOrRes) {
    const zip = await JSZip.loadAsync(await blob.arrayBuffer());
    let stylesXml = (await zip.file("xl/styles.xml")?.async("string")) || "";
    const sheetPath = await resolveSheetPath(zip, sheetRef);
    if (!sheetPath) continue;
    const sheetXml = await zip.file(sheetPath)?.async("string");
    if (!sheetXml) continue;
    const cells = collectSheetCells(sheetXml).filter((cell) => cell.row >= 6);
    if (!cells.length) continue;
    const byColumn = new Map<number, typeof cells>();
    for (const cell of cells) {
      const list = byColumn.get(cell.column) ?? [];
      list.push(cell);
      byColumn.set(cell.column, list);
    }
    const writes: ZipCellWrite[] = [];
    for (const [column, columnCells] of byColumn) {
      const good = columnCells
        .filter(
          (cell) =>
            cell.hasContent &&
            cell.styleId &&
            styleIsRegularCentered(stylesXml, cell.styleId),
        )
        .map((cell) => cell.styleId!);
      const common = columnCells
        .filter((cell) => cell.hasContent && cell.styleId)
        .map((cell) => cell.styleId!);
      let styleId = mostFrequentId(good);
      if (!styleId) {
        const ensured = ensurePlainCenteredStyle(
          stylesXml,
          mostFrequentId(common),
        );
        stylesXml = ensured.xml;
        styleId = ensured.styleId;
      }
      if (!styleId) continue;
      for (const cell of columnCells) {
        if (cell.styleId === styleId) continue;
        if (!cell.hasContent && !cell.styleId) continue;
        writes.push({
          row: cell.row,
          column,
          value: null,
          styleOnly: true,
          styleId,
        });
      }
    }
    if (!writes.length) continue;
    zip.file("xl/styles.xml", stylesXml);
    const next = await zip.generateAsync({
      type: "blob",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      compression: "DEFLATE",
    });
    blob = await applyInlineStringWritesToWorkbook(next, sheetRef, writes);
    styledCells += writes.length;
  }
  return { blob, styledCells };
}
